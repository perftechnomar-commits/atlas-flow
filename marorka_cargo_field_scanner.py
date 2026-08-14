from __future__ import annotations

from datetime import date, datetime, timedelta
from io import BytesIO
import os
import re
from typing import Any
from urllib.parse import urlencode, urljoin
import xml.etree.ElementTree as ET

import pandas as pd
import requests
import streamlit as st
from requests.auth import HTTPBasicAuth, HTTPDigestAuth


st.set_page_config(page_title="Marorka Cargo Field Scanner", layout="wide")

SERVICE_ROOT = "https://online.marorka.com/Odata/v1/ODataService.svc"
ENDPOINTS = {
    "ReportData": f"{SERVICE_ROOT}/ReportData",
    "ReportPivots": f"{SERVICE_ROOT}/ReportPivots",
    "ShipPivots": f"{SERVICE_ROOT}/ShipPivots",
}

DATE_FIELD_CANDIDATES = {
    "ReportData": ["StartDateTimeGMT"],
    "ReportPivots": ["DateTime", "StartDateTimeGMT", "ReportDateTime", "Timestamp"],
    "ShipPivots": ["DateTime", "StartDateTimeGMT", "Timestamp"],
}

DEFAULT_KEYWORDS = [
    "cargo",
    "teu",
    "container",
    "deadweight",
    "dwt",
    "draft",
    "draught",
    "laden",
    "ballast",
    "load",
    "loaded",
    "loading",
    "discharge",
    "discharging",
    "weight",
    "tonne",
    "tonnes",
    "metric ton",
    "reefer",
]

REPORTDATA_SELECT = [
    "ShipName",
    "ReportType",
    "StartDateTimeGMT",
    "EndDateTimeGMT",
    "StateName",
    "ValueDescription",
    "ReportedValue",
]

REQUEST_TIMEOUT_SECONDS = 75
RETRYABLE_STATUSES = {429, 500, 502, 503, 504}


def read_secret(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name, os.getenv(name, default))
    except Exception:
        value = os.getenv(name, default)
    return str(value).strip() if value is not None else default


def request_auth(username: str, password: str, auth_method: str) -> Any:
    method = auth_method.lower()
    if method == "basic":
        return HTTPBasicAuth(username, password)
    if method == "digest":
        return HTTPDigestAuth(username, password)
    if method in {"none", "anonymous", "", "bearer"}:
        return None
    raise ValueError("MARORKA_AUTH_METHOD must be basic, digest, bearer, or none.")


def request_headers(token: str, auth_method: str) -> dict[str, str]:
    headers = {"Accept": "application/json"}
    if auth_method.lower() == "bearer":
        if not token:
            raise ValueError("MARORKA_TOKEN is required for bearer authentication.")
        headers["Authorization"] = f"Bearer {token}"
    return headers


def get_with_retry(session: requests.Session, url: str, auth: Any) -> requests.Response:
    last_error: Exception | None = None
    for attempt in range(1, 4):
        try:
            response = session.get(url, auth=auth, timeout=REQUEST_TIMEOUT_SECONDS)
            if response.status_code in RETRYABLE_STATUSES and attempt < 3:
                import time
                time.sleep(2 ** (attempt - 1))
                continue
            return response
        except (requests.ConnectionError, requests.Timeout, requests.exceptions.ChunkedEncodingError) as exc:
            last_error = exc
            if attempt >= 3:
                raise
            import time
            time.sleep(2 ** (attempt - 1))
    if last_error:
        raise last_error
    raise requests.RequestException("Request failed before receiving a response.")


def extract_odata_page(payload: Any) -> tuple[list[dict[str, Any]], str | None]:
    if isinstance(payload, list):
        return payload, None
    if not isinstance(payload, dict):
        raise ValueError("Unexpected OData payload type.")

    rows = payload.get("value")
    next_link = payload.get("@odata.nextLink") or payload.get("odata.nextLink")
    if rows is None and isinstance(payload.get("d"), dict):
        d = payload["d"]
        rows = d.get("results")
        next_link = next_link or d.get("__next")
    if rows is None:
        raise ValueError("Could not find OData rows in response.")
    return rows, next_link


def keyword_pattern(keywords: list[str]) -> re.Pattern[str]:
    escaped = [re.escape(k.strip()) for k in keywords if k.strip()]
    return re.compile("|".join(escaped), re.IGNORECASE) if escaped else re.compile(r"$^")


def matched_keywords(text: Any, keywords: list[str]) -> list[str]:
    value = str(text or "")
    low = value.casefold()
    return [kw for kw in keywords if kw.casefold() in low]


def build_filtered_url(source: str, start_date: date, date_field: str | None) -> str:
    endpoint = ENDPOINTS[source]
    params: dict[str, str] = {}
    if date_field:
        params["$filter"] = f"{date_field} gt DateTime'{start_date.isoformat()}'"
        params["$orderby"] = f"{date_field} desc"
    if source == "ReportData":
        params["$select"] = ",".join(REPORTDATA_SELECT)
    return f"{endpoint}?{urlencode(params)}" if params else endpoint


def fetch_source_sample(
    source: str,
    start_date: date,
    max_pages: int,
    username: str,
    password: str,
    token: str,
    auth_method: str,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    auth = request_auth(username, password, auth_method)
    headers = request_headers(token, auth_method)

    candidate_errors: list[str] = []
    selected_field: str | None = None
    first_url: str | None = None

    with requests.Session() as session:
        session.headers.update(headers)

        # Try known date fields so the scan stays recent and bounded.
        for candidate in DATE_FIELD_CANDIDATES[source]:
            trial_url = build_filtered_url(source, start_date, candidate)
            try:
                trial_response = get_with_retry(session, trial_url, auth)
                if trial_response.ok:
                    # Confirm it is parseable before accepting the candidate.
                    extract_odata_page(trial_response.json())
                    selected_field = candidate
                    first_url = trial_url
                    break
                candidate_errors.append(f"{candidate}: HTTP {trial_response.status_code}")
            except Exception as exc:
                candidate_errors.append(f"{candidate}: {exc}")

        # Last-resort unfiltered request. Keep page count low to avoid a huge scan.
        if first_url is None:
            first_url = build_filtered_url(source, start_date, None)

        rows: list[dict[str, Any]] = []
        pages = 0
        next_url = first_url
        seen: set[str] = set()
        total_bytes = 0

        while next_url and pages < max_pages:
            if next_url in seen:
                break
            seen.add(next_url)
            response = get_with_retry(session, next_url, auth)
            total_bytes += len(response.content)
            response.raise_for_status()
            page_rows, next_link = extract_odata_page(response.json())
            rows.extend(page_rows)
            pages += 1
            next_url = urljoin(next_url, next_link) if next_link else None

    df = pd.DataFrame(rows)
    if "__metadata" in df.columns:
        df = df.drop(columns=["__metadata"])

    diagnostics = {
        "source": source,
        "rows_scanned": len(df),
        "columns_seen": len(df.columns),
        "pages_scanned": pages,
        "downloaded_mb": round(total_bytes / 1024 / 1024, 2),
        "date_field_used": selected_field or "none (unfiltered fallback)",
        "first_url": first_url,
        "date_filter_attempt_errors": "; ".join(candidate_errors),
    }
    return df, diagnostics


def scan_dataframe(source: str, df: pd.DataFrame, keywords: list[str]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    field_matches: list[dict[str, Any]] = []
    value_matches: list[dict[str, Any]] = []
    reportdata_values: list[dict[str, Any]] = []

    # 1) Column / field names.
    for column in df.columns:
        hits = matched_keywords(column, keywords)
        if hits:
            field_matches.append(
                {
                    "Source": source,
                    "Match Type": "Column name",
                    "Field / Column": column,
                    "Matched Keywords": ", ".join(hits),
                    "Example Value": first_nonempty_value(df[column]),
                }
            )

    # 2) ReportData ValueDescription values are the most important dynamic-field scan.
    if source == "ReportData" and "ValueDescription" in df.columns:
        unique_values = sorted(df["ValueDescription"].dropna().astype(str).unique(), key=str.casefold)
        for value in unique_values:
            hits = matched_keywords(value, keywords)
            if hits:
                subset = df[df["ValueDescription"].astype(str) == value]
                reportdata_values.append(
                    {
                        "Source": source,
                        "ValueDescription": value,
                        "Matched Keywords": ", ".join(hits),
                        "Occurrences in Scan": int(len(subset)),
                        "Example Ship": first_nonempty_value(subset.get("ShipName", pd.Series(dtype=object))),
                        "Example Report Type": first_nonempty_value(subset.get("ReportType", pd.Series(dtype=object))),
                        "Example Reported Value": first_nonempty_value(subset.get("ReportedValue", pd.Series(dtype=object))),
                    }
                )

    # 3) String cell contents across the wide sources (and ReportData as backup).
    identity_candidates = [c for c in ["ShipName", "ReportType", "DateTime", "StartDateTimeGMT", "EndDateTimeGMT"] if c in df.columns]
    for column in df.columns:
        series = df[column]
        # Avoid converting clearly numeric columns to strings if not needed.
        if pd.api.types.is_numeric_dtype(series):
            continue
        text_series = series.astype("string")
        mask = pd.Series(False, index=df.index)
        for kw in keywords:
            mask = mask | text_series.str.contains(re.escape(kw), case=False, na=False)
        if not mask.any():
            continue

        for idx in df.index[mask][:50]:
            value = df.at[idx, column]
            hits = matched_keywords(value, keywords)
            record: dict[str, Any] = {
                "Source": source,
                "Match Type": "Cell value",
                "Field / Column": column,
                "Matched Keywords": ", ".join(hits),
                "Matched Value": str(value)[:500],
            }
            for ident in identity_candidates:
                record[ident] = df.at[idx, ident]
            value_matches.append(record)

    return (
        pd.DataFrame(field_matches),
        pd.DataFrame(reportdata_values),
        pd.DataFrame(value_matches),
    )


def first_nonempty_value(series: pd.Series) -> str:
    if series is None or len(series) == 0:
        return ""
    for value in series:
        if pd.notna(value) and str(value).strip():
            return str(value)[:300]
    return ""


def fetch_and_scan_metadata(
    keywords: list[str],
    username: str,
    password: str,
    token: str,
    auth_method: str,
) -> tuple[pd.DataFrame, int]:
    auth = request_auth(username, password, auth_method)
    headers = request_headers(token, auth_method)
    headers["Accept"] = "application/xml,text/xml,*/*"

    with requests.Session() as session:
        session.headers.update(headers)
        response = get_with_retry(session, f"{SERVICE_ROOT}/$metadata", auth)
        response.raise_for_status()
        content = response.content

    rows: list[dict[str, Any]] = []
    try:
        root = ET.fromstring(content)
        for elem in root.iter():
            local_tag = elem.tag.rsplit("}", 1)[-1]
            name = elem.attrib.get("Name", "")
            type_name = elem.attrib.get("Type", "")
            searchable = " ".join([local_tag, name, type_name])
            hits = matched_keywords(searchable, keywords)
            if hits:
                rows.append(
                    {
                        "Metadata Element": local_tag,
                        "Name": name,
                        "Type": type_name,
                        "Matched Keywords": ", ".join(hits),
                    }
                )
    except ET.ParseError:
        text = content.decode("utf-8", errors="replace")
        for line in text.splitlines():
            hits = matched_keywords(line, keywords)
            if hits:
                rows.append(
                    {
                        "Metadata Element": "raw line",
                        "Name": line.strip()[:500],
                        "Type": "",
                        "Matched Keywords": ", ".join(hits),
                    }
                )
    return pd.DataFrame(rows).drop_duplicates(), len(content)


def to_excel_bytes(
    metadata_matches: pd.DataFrame,
    field_matches: pd.DataFrame,
    value_description_matches: pd.DataFrame,
    cell_matches: pd.DataFrame,
    diagnostics: pd.DataFrame,
    all_columns: pd.DataFrame,
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheets = {
            "Metadata Matches": metadata_matches,
            "Field Matches": field_matches,
            "ValueDescriptions": value_description_matches,
            "Cell Matches": cell_matches,
            "Diagnostics": diagnostics,
            "All Columns Seen": all_columns,
        }
        for sheet_name, df in sheets.items():
            safe = df if not df.empty else pd.DataFrame({"Result": ["No matches"]})
            safe.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()


def main() -> None:
    st.title("Marorka Cargo / Capacity Field Scanner")
    st.caption(
        "Scans OData metadata plus live ReportData, ReportPivots and ShipPivots data for cargo-related field names and values. "
        "It does not modify any existing dashboard cache."
    )

    username = read_secret("MARORKA_USERNAME")
    password = read_secret("MARORKA_PASSWORD")
    token = read_secret("MARORKA_TOKEN")
    auth_method = read_secret("MARORKA_AUTH_METHOD", "basic")

    if auth_method.lower() in {"basic", "digest"} and (not username or not password):
        st.error("Add MARORKA_USERNAME and MARORKA_PASSWORD to Streamlit secrets first.")
        st.stop()

    default_start = date.today() - timedelta(days=30)
    start_date = st.date_input("Scan data from", value=default_start, max_value=date.today())
    max_pages = st.number_input(
        "Maximum pages per API source",
        min_value=1,
        max_value=100,
        value=15,
        step=1,
        help="Keep this modest for the first scan. Increase only if no useful matches appear.",
    )
    keyword_text = st.text_area(
        "Keywords (comma separated)",
        value=", ".join(DEFAULT_KEYWORDS),
        height=120,
    )
    keywords = [k.strip() for k in keyword_text.split(",") if k.strip()]

    if not st.button("Run cargo field scan", type="primary"):
        st.stop()

    metadata_matches = pd.DataFrame()
    field_frames: list[pd.DataFrame] = []
    value_desc_frames: list[pd.DataFrame] = []
    cell_frames: list[pd.DataFrame] = []
    diagnostics_rows: list[dict[str, Any]] = []
    all_columns_rows: list[dict[str, Any]] = []

    with st.status("Scanning Marorka APIs...", expanded=True) as status:
        try:
            st.write("Scanning OData service metadata...")
            metadata_matches, metadata_bytes = fetch_and_scan_metadata(
                keywords, username, password, token, auth_method
            )
            st.write(f"Metadata scan complete ({metadata_bytes / 1024:.1f} KB).")
        except Exception as exc:
            st.warning(f"Metadata scan failed: {exc}")

        for source in ENDPOINTS:
            st.write(f"Scanning {source}...")
            try:
                df, diagnostics = fetch_source_sample(
                    source,
                    start_date,
                    int(max_pages),
                    username,
                    password,
                    token,
                    auth_method,
                )
                diagnostics_rows.append(diagnostics)
                for column in df.columns:
                    all_columns_rows.append({"Source": source, "Column": str(column)})

                field_df, value_desc_df, cell_df = scan_dataframe(source, df, keywords)
                if not field_df.empty:
                    field_frames.append(field_df)
                if not value_desc_df.empty:
                    value_desc_frames.append(value_desc_df)
                if not cell_df.empty:
                    cell_frames.append(cell_df)

                st.write(
                    f"{source}: {len(df):,} rows, {len(df.columns):,} columns, "
                    f"{diagnostics['pages_scanned']} pages."
                )
            except Exception as exc:
                diagnostics_rows.append(
                    {
                        "source": source,
                        "error": str(exc),
                    }
                )
                st.warning(f"{source} scan failed: {exc}")

        status.update(label="Cargo field scan complete", state="complete")

    field_matches = pd.concat(field_frames, ignore_index=True) if field_frames else pd.DataFrame()
    value_description_matches = (
        pd.concat(value_desc_frames, ignore_index=True) if value_desc_frames else pd.DataFrame()
    )
    cell_matches = pd.concat(cell_frames, ignore_index=True) if cell_frames else pd.DataFrame()
    diagnostics_df = pd.DataFrame(diagnostics_rows)
    all_columns_df = pd.DataFrame(all_columns_rows).drop_duplicates()

    total_matches = (
        len(metadata_matches)
        + len(field_matches)
        + len(value_description_matches)
        + len(cell_matches)
    )

    if total_matches:
        st.success(f"Found {total_matches:,} cargo/capacity-related matches across metadata and live API samples.")
    else:
        st.warning(
            "No cargo/capacity-related matches were found in this scan window/sample. "
            "That does not prove the fields never occur; increase the date window/pages and scan again."
        )

    tabs = st.tabs([
        "Strong Matches",
        "ReportData ValueDescriptions",
        "Cell Matches",
        "All Columns",
        "Diagnostics",
    ])

    with tabs[0]:
        st.subheader("Metadata / schema matches")
        st.dataframe(metadata_matches, use_container_width=True, hide_index=True)
        st.subheader("Live API field-name matches")
        st.dataframe(field_matches, use_container_width=True, hide_index=True)

    with tabs[1]:
        st.dataframe(value_description_matches, use_container_width=True, hide_index=True)

    with tabs[2]:
        st.caption("String values containing one of the search terms. Capped at 50 examples per matching column.")
        st.dataframe(cell_matches, use_container_width=True, hide_index=True)

    with tabs[3]:
        st.dataframe(all_columns_df, use_container_width=True, hide_index=True)

    with tabs[4]:
        st.dataframe(diagnostics_df, use_container_width=True, hide_index=True)

    excel_bytes = to_excel_bytes(
        metadata_matches,
        field_matches,
        value_description_matches,
        cell_matches,
        diagnostics_df,
        all_columns_df,
    )
    st.download_button(
        "Download scan results as Excel",
        data=excel_bytes,
        file_name=f"marorka_cargo_field_scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
