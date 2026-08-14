from __future__ import annotations

from datetime import date, datetime, timedelta
from io import BytesIO
import os
import re
import time
from typing import Any
from urllib.parse import urlencode, urljoin

import pandas as pd
import requests
import streamlit as st
from requests.auth import HTTPBasicAuth, HTTPDigestAuth


st.set_page_config(page_title="Marorka Voyage / Cargo Structure Explorer", layout="wide")

SERVICE_ROOT = "https://online.marorka.com/Odata/v1/ODataService.svc"
ENDPOINTS = {
    "ReportData": f"{SERVICE_ROOT}/ReportData",
    "ReportPivots": f"{SERVICE_ROOT}/ReportPivots",
    "ShipPivots": f"{SERVICE_ROOT}/ShipPivots",
}

REQUEST_TIMEOUT_SECONDS = 75
RETRYABLE_STATUSES = {429, 500, 502, 503, 504}

DATE_FIELD_CANDIDATES = {
    "ReportData": ["StartDateTimeGMT"],
    "ReportPivots": ["DateTime", "StartDateTimeGMT", "ReportDateTime", "Timestamp"],
    "ShipPivots": ["DateTime", "StartDateTimeGMT", "Timestamp"],
}

REPORTDATA_SELECT = [
    "ReportId",
    "ShipName",
    "ReportType",
    "StartDateTimeGMT",
    "EndDateTimeGMT",
    "StateName",
    "ValueDescription",
    "ReportedValue",
]

STRUCTURE_KEYWORDS = [
    "voyage", "voy", "port", "berth", "terminal", "departure", "arrival", "origin", "destination",
    "cargo", "teu", "container", "deadweight", "dwt", "draft", "draught", "laden", "ballast",
    "load", "loaded", "loading", "discharge", "discharged", "discharging", "weight", "tonne",
    "reefer", "feu", "20ft", "40ft",
]

CARGO_ONLY_KEYWORDS = [
    "cargo", "teu", "container", "deadweight", "dwt", "draft", "draught", "laden", "ballast",
    "load", "loaded", "loading", "discharge", "discharged", "discharging", "weight", "tonne",
    "reefer", "feu", "20ft", "40ft",
]

PREFERRED_IDENTITY_FIELDS = [
    "ReportId", "ShipName", "ReportType", "DateTime", "StartDateTimeGMT", "EndDateTimeGMT",
    "ReportDateTime", "Timestamp", "State", "StateName", "Port", "PortName", "PortCode",
    "Voyage", "VoyageNo", "VoyageNumber", "VoyageId", "VoyageID",
]

PREFERRED_CARGO_FIELDS = [
    "Cargo", "CargoMT", "CargoWeight", "CargoTEU", "CargoType", "CargoAmount", "CargoUnit",
    "DraftFore", "DraftAft", "DraftMean", "Deadweight", "DWT", "BallastDraft", "DesignDraft",
]


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
                time.sleep(2 ** (attempt - 1))
                continue
            return response
        except (requests.ConnectionError, requests.Timeout, requests.exceptions.ChunkedEncodingError) as exc:
            last_error = exc
            if attempt >= 3:
                raise
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
        data = payload["d"]
        rows = data.get("results")
        next_link = next_link or data.get("__next")
    if rows is None:
        raise ValueError("Could not find OData rows in response.")
    return rows, next_link


def odata_quote(value: str) -> str:
    return str(value).replace("'", "''")


def build_source_url(
    source: str,
    vessel: str,
    start_date: date,
    end_date: date,
    date_field: str | None,
) -> str:
    endpoint = ENDPOINTS[source]
    filters: list[str] = []
    if vessel:
        filters.append(f"ShipName eq '{odata_quote(vessel)}'")
    if date_field:
        end_exclusive = end_date + timedelta(days=1)
        filters.append(f"{date_field} ge DateTime'{start_date.isoformat()}'")
        filters.append(f"{date_field} lt DateTime'{end_exclusive.isoformat()}'")

    params: dict[str, str] = {}
    if filters:
        params["$filter"] = " and ".join(filters)
    if date_field:
        params["$orderby"] = f"{date_field} asc"
    if source == "ReportData":
        params["$select"] = ",".join(REPORTDATA_SELECT)
    return f"{endpoint}?{urlencode(params)}" if params else endpoint


def fetch_source(
    source: str,
    vessel: str,
    start_date: date,
    end_date: date,
    max_pages: int,
    username: str,
    password: str,
    token: str,
    auth_method: str,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    auth = request_auth(username, password, auth_method)
    headers = request_headers(token, auth_method)

    selected_field: str | None = None
    first_url: str | None = None
    candidate_errors: list[str] = []

    with requests.Session() as session:
        session.headers.update(headers)

        for candidate in DATE_FIELD_CANDIDATES[source]:
            trial_url = build_source_url(source, vessel, start_date, end_date, candidate)
            try:
                trial = get_with_retry(session, trial_url, auth)
                if trial.ok:
                    extract_odata_page(trial.json())
                    selected_field = candidate
                    first_url = trial_url
                    break
                candidate_errors.append(f"{candidate}: HTTP {trial.status_code}")
            except Exception as exc:
                candidate_errors.append(f"{candidate}: {exc}")

        if first_url is None:
            first_url = build_source_url(source, vessel, start_date, end_date, None)

        rows: list[dict[str, Any]] = []
        pages = 0
        total_bytes = 0
        next_url = first_url
        seen: set[str] = set()
        started = time.perf_counter()

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

    meta = {
        "source": source,
        "rows": int(len(df)),
        "columns": int(len(df.columns)),
        "pages": int(pages),
        "date_field": selected_field or "unfiltered fallback",
        "downloaded_mb": round(total_bytes / 1024 / 1024, 2),
        "fetch_seconds": round(time.perf_counter() - started, 2),
        "hit_page_limit": bool(next_url),
        "candidate_errors": " | ".join(candidate_errors),
        "first_url": first_url,
    }
    return df, meta


def matching_columns(df: pd.DataFrame, keywords: list[str]) -> list[str]:
    keys = [k.casefold() for k in keywords]
    return [c for c in df.columns if any(k in str(c).casefold() for k in keys)]


def matching_text_rows(df: pd.DataFrame, columns: list[str], keywords: list[str]) -> pd.Series:
    if df.empty or not columns:
        return pd.Series(False, index=df.index)
    pattern = "|".join(re.escape(k) for k in keywords)
    mask = pd.Series(False, index=df.index)
    for column in columns:
        text = df[column].astype("string").fillna("")
        mask = mask | text.str.contains(pattern, case=False, regex=True, na=False)
    return mask


def discover_structure_columns(df: pd.DataFrame) -> list[str]:
    if df.empty:
        return []
    selected: list[str] = []
    for column in PREFERRED_IDENTITY_FIELDS + PREFERRED_CARGO_FIELDS + matching_columns(df, STRUCTURE_KEYWORDS):
        if column in df.columns and column not in selected:
            selected.append(column)
    return selected


def detect_datetime_column(df: pd.DataFrame) -> str | None:
    for candidate in ["DateTime", "StartDateTimeGMT", "ReportDateTime", "Timestamp", "EndDateTimeGMT"]:
        if candidate in df.columns:
            parsed = pd.to_datetime(df[candidate], errors="coerce", utc=True)
            if parsed.notna().any():
                return candidate
    for column in df.columns:
        if "date" in str(column).casefold() or "time" in str(column).casefold():
            parsed = pd.to_datetime(df[column], errors="coerce", utc=True)
            if parsed.notna().any():
                return column
    return None


def normalize_report_type(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").casefold()).strip()


def infer_voyage_segments(report_df: pd.DataFrame) -> pd.DataFrame:
    """Create exploratory voyage segments from report sequence.

    This is deliberately labelled inferred, not authoritative. A new voyage begins
    at a Departure report. Reports before the first Departure are grouped as
    Pre-departure / unknown. Arrival reports close the current voyage but subsequent
    port/cargo reports remain in the same group until the next Departure.
    """
    if report_df.empty:
        return report_df.copy()

    df = report_df.copy()
    dt_col = detect_datetime_column(df)
    if dt_col:
        df["_sort_dt"] = pd.to_datetime(df[dt_col], errors="coerce", utc=True)
        df = df.sort_values("_sort_dt", na_position="last").reset_index(drop=True)
    else:
        df = df.reset_index(drop=True)

    voyage_number = 0
    labels: list[str] = []
    phases: list[str] = []
    for _, row in df.iterrows():
        report_type = normalize_report_type(row.get("ReportType", ""))
        is_departure = "departure" in report_type
        is_arrival = "arrival" in report_type
        is_cargo = "cargo" in report_type
        is_port = any(token in report_type for token in ["port", "berth", "cargo", "arrival", "departure"])

        if is_departure:
            voyage_number += 1
        label = f"Inferred Voyage {voyage_number}" if voyage_number > 0 else "Pre-departure / unknown"
        labels.append(label)

        if is_departure:
            phases.append("Departure")
        elif is_arrival:
            phases.append("Arrival")
        elif is_cargo:
            phases.append("Cargo operation")
        elif is_port:
            phases.append("Port")
        else:
            phases.append("Voyage / sea")

    df.insert(0, "InferredVoyage", labels)
    df.insert(1, "InferredPhase", phases)
    return df.drop(columns=["_sort_dt"], errors="ignore")


def build_reportdata_cargo_events(reportdata: pd.DataFrame) -> pd.DataFrame:
    if reportdata.empty:
        return pd.DataFrame()
    df = reportdata.copy()
    cols = [c for c in ["ValueDescription", "ReportType", "StateName"] if c in df.columns]
    mask = matching_text_rows(df, cols, CARGO_ONLY_KEYWORDS)
    events = df.loc[mask].copy()
    dt_col = detect_datetime_column(events)
    if dt_col:
        events[dt_col] = pd.to_datetime(events[dt_col], errors="coerce", utc=True)
        events = events.sort_values(dt_col)
    preferred = [c for c in REPORTDATA_SELECT if c in events.columns]
    return events[preferred]


def build_reportdata_wide_cargo(reportdata_events: pd.DataFrame) -> pd.DataFrame:
    if reportdata_events.empty or "ValueDescription" not in reportdata_events.columns:
        return pd.DataFrame()
    df = reportdata_events.copy()
    identity = [c for c in ["ReportId", "ShipName", "ReportType", "StartDateTimeGMT", "EndDateTimeGMT", "StateName"] if c in df.columns]
    if not identity:
        return pd.DataFrame()
    df["ValueDescription"] = df["ValueDescription"].astype("string")
    df["_order"] = range(len(df))
    df = df.sort_values("_order").drop_duplicates([*identity, "ValueDescription"], keep="last")
    wide = df.pivot(index=identity, columns="ValueDescription", values="ReportedValue").reset_index()
    wide.columns.name = None
    return wide


def build_report_type_summary(report_df: pd.DataFrame) -> pd.DataFrame:
    if report_df.empty or "ReportType" not in report_df.columns:
        return pd.DataFrame()
    summary = (
        report_df.assign(ReportType=report_df["ReportType"].astype("string").fillna("(Blank)"))
        .groupby("ReportType", dropna=False)
        .size()
        .reset_index(name="Rows")
        .sort_values(["Rows", "ReportType"], ascending=[False, True])
    )
    return summary


def build_candidate_field_catalog(reportpivots: pd.DataFrame, shippivots: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for source_name, df in [("ReportPivots", reportpivots), ("ShipPivots", shippivots)]:
        for column in discover_structure_columns(df):
            nonempty = df[column].dropna() if column in df.columns else pd.Series(dtype=object)
            example = ""
            if not nonempty.empty:
                example = str(nonempty.iloc[0])
            rows.append({
                "Source": source_name,
                "Field": column,
                "Non-empty rows": int(nonempty.shape[0]),
                "Example": example[:160],
            })
    return pd.DataFrame(rows)


def safe_sheet_name(name: str) -> str:
    return re.sub(r"[\\/*?:\[\]]", "_", name)[:31]


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            if df is None:
                continue
            export_df = df.copy()
            for column in export_df.columns:
                if pd.api.types.is_datetime64_any_dtype(export_df[column]):
                    try:
                        export_df[column] = pd.to_datetime(export_df[column], errors="coerce", utc=True).dt.tz_localize(None)
                    except Exception:
                        pass
            export_df.to_excel(writer, index=False, sheet_name=safe_sheet_name(name))
            ws = writer.sheets[safe_sheet_name(name)]
            ws.freeze_panes = "A2"
            for cells in ws.columns:
                max_len = max((len(str(cell.value)) if cell.value is not None else 0) for cell in cells)
                ws.column_dimensions[cells[0].column_letter].width = min(max(max_len + 2, 12), 50)
    return output.getvalue()


def show_df(title: str, df: pd.DataFrame, row_limit: int = 500) -> None:
    st.subheader(title)
    if df is None or df.empty:
        st.info("No matching rows found in this scan.")
        return
    st.dataframe(df.head(row_limit), use_container_width=True, hide_index=True)
    if len(df) > row_limit:
        st.caption(f"Showing first {row_limit:,} of {len(df):,} rows. Excel export contains all rows.")


st.title("Marorka Voyage / Cargo Structure Explorer")
st.caption(
    "Explores how Marorka represents voyage boundaries, report sequence, ports and cargo fields for one vessel. "
    "It does not modify AtlasFlow or any dashboard cache. Inferred voyage groups are exploratory only."
)

username = read_secret("MARORKA_USERNAME")
password = read_secret("MARORKA_PASSWORD")
token = read_secret("MARORKA_TOKEN")
auth_method = read_secret("MARORKA_AUTH_METHOD", "basic")

if auth_method.lower() in {"basic", "digest"} and (not username or not password):
    st.error("MARORKA_USERNAME and MARORKA_PASSWORD are required in this app's Streamlit Secrets.")
    st.stop()

col1, col2, col3, col4 = st.columns([1.5, 1, 1, 1])
vessel = col1.text_input("Vessel", value="AGIOS DIMITRIOS")
default_end = date.today()
default_start = max(date(2026, 1, 1), default_end - timedelta(days=60))
start_date = col2.date_input("From", value=default_start)
end_date = col3.date_input("To", value=default_end)
max_pages = int(col4.number_input("Max pages / source", min_value=1, max_value=100, value=25, step=5))

include_shippivots = st.checkbox("Also scan ShipPivots for voyage/cargo context", value=True)

if start_date > end_date:
    st.error("From date must be before To date.")
    st.stop()

if st.button("Run voyage / cargo structure scan", type="primary"):
    if not vessel.strip():
        st.error("Enter a vessel name.")
        st.stop()

    diagnostics_rows: list[dict[str, Any]] = []
    results: dict[str, pd.DataFrame] = {}

    try:
        with st.spinner("Scanning ReportPivots..."):
            reportpivots, meta = fetch_source(
                "ReportPivots", vessel.strip(), start_date, end_date, max_pages,
                username, password, token, auth_method,
            )
            results["ReportPivots Raw"] = reportpivots
            diagnostics_rows.append(meta)

        with st.spinner("Scanning ReportData cargo / voyage records..."):
            reportdata, meta = fetch_source(
                "ReportData", vessel.strip(), start_date, end_date, max_pages,
                username, password, token, auth_method,
            )
            results["ReportData Raw Sample"] = reportdata
            diagnostics_rows.append(meta)

        shippivots = pd.DataFrame()
        if include_shippivots:
            with st.spinner("Scanning ShipPivots..."):
                shippivots, meta = fetch_source(
                    "ShipPivots", vessel.strip(), start_date, end_date, max_pages,
                    username, password, token, auth_method,
                )
                results["ShipPivots Raw"] = shippivots
                diagnostics_rows.append(meta)

        structure_cols = discover_structure_columns(reportpivots)
        report_timeline = reportpivots[structure_cols].copy() if structure_cols else reportpivots.copy()
        report_timeline = infer_voyage_segments(report_timeline)
        cargo_events = build_reportdata_cargo_events(reportdata)
        cargo_wide = build_reportdata_wide_cargo(cargo_events)
        field_catalog = build_candidate_field_catalog(reportpivots, shippivots)
        report_type_summary = build_report_type_summary(reportpivots)

        voyage_cols = [c for c in reportpivots.columns if any(k in str(c).casefold() for k in ["voyage", "port", "origin", "destination", "berth", "terminal"])]
        cargo_cols = [c for c in reportpivots.columns if any(k in str(c).casefold() for k in CARGO_ONLY_KEYWORDS)]

        st.success("Scan complete.")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("ReportPivots reports", f"{len(reportpivots):,}")
        m2.metric("ReportData rows scanned", f"{len(reportdata):,}")
        m3.metric("Cargo-related ReportData rows", f"{len(cargo_events):,}")
        m4.metric("Voyage/port candidate columns", f"{len(voyage_cols):,}")

        if voyage_cols:
            st.info("Explicit voyage/port-related ReportPivots columns found: " + ", ".join(voyage_cols))
        else:
            st.warning(
                "No explicit voyage/port-named ReportPivots columns were found in this sample. "
                "Use the inferred report sequence below to determine whether Departure/Arrival reports can define voyages."
            )

        if cargo_cols:
            st.info("Cargo/draft-related ReportPivots columns found: " + ", ".join(cargo_cols))

        show_df("1. Chronological report timeline", report_timeline)
        show_df("2. Report types found", report_type_summary)
        show_df("3. Candidate voyage / cargo fields", field_catalog)
        show_df("4. Cargo-related ReportData events", cargo_events)
        show_df("5. Cargo-related ReportData pivoted by report", cargo_wide)

        diagnostics = pd.DataFrame(diagnostics_rows)
        show_df("6. API diagnostics", diagnostics, row_limit=50)

        export_sheets = {
            "Report Timeline": report_timeline,
            "Report Types": report_type_summary,
            "Candidate Fields": field_catalog,
            "Cargo Events": cargo_events,
            "Cargo By Report": cargo_wide,
            "Diagnostics": diagnostics,
            "ReportPivots Raw": reportpivots,
            "ReportData Raw": reportdata,
        }
        if include_shippivots:
            export_sheets["ShipPivots Raw"] = shippivots

        excel_bytes = to_excel_bytes(export_sheets)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            "Download voyage / cargo structure Excel",
            data=excel_bytes,
            file_name=f"marorka_voyage_cargo_structure_{re.sub(r'[^A-Za-z0-9]+', '_', vessel.strip())}_{stamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.caption(
            "Interpretation note: 'Inferred Voyage' starts a new group at each report whose ReportType contains 'Departure'. "
            "This is deliberately exploratory; use the exported timeline to confirm the real Marorka voyage/report structure before building the production dashboard."
        )

    except requests.HTTPError as exc:
        status = exc.response.status_code if exc.response is not None else "unknown"
        st.error(f"Marorka API request failed with HTTP {status}.")
        if exc.response is not None and exc.response.request is not None:
            st.code(exc.response.request.url, language="text")
    except Exception as exc:
        st.exception(exc)
