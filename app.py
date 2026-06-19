from __future__ import annotations

import base64
from datetime import date, datetime, timedelta, timezone
from hashlib import sha256
from html import escape
from io import BytesIO
import hmac
import mimetypes
import os
import re
import time
from pathlib import Path
from typing import Any
from urllib.parse import urlencode, urljoin
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import streamlit as st
from requests.auth import HTTPBasicAuth, HTTPDigestAuth


# =============================================================================
# Configuration
# =============================================================================

APP_TITLE = "AtlasFlow"
APP_DIR = Path(__file__).resolve().parent
ODATA_ENDPOINT = "https://online.marorka.com/Odata/v1/ODataService.svc/ReportData"
MAX_ODATA_PAGES = 250
API_CACHE_TTL_SECONDS = 21600  # 6 hours
API_FULL_START_DATE = date(2026, 1, 1)
TABLE_PREVIEW_ROW_LIMIT = 1000
DISPLAY_DATETIME_FORMAT = "%d/%m/%Y %H:%M"

EXCLUDED_REPORT_TYPES = [
    "Intake Report",
    "Fuel Change Report",
]

SOURCE_COLUMNS = [
    "ReportId",
    "ShipName",
    "ReportType",
    "StartDateTimeGMT",
    "EndDateTimeGMT",
    "LapTime",
    "StateName",
    "ValueDescription",
    "ReportedValue",
]

PIVOT_IDENTITY_COLUMNS = [
    "ReportId",
    "ShipName",
    "ReportType",
    "StartDateTimeGMT",
    "EndDateTimeGMT",
    "LapTime",
    "StateName",
]

DEFAULT_DISPLAY_IDENTITY_COLUMNS = [
    "ShipName",
    "ReportType",
    "StartDateTimeGMT",
    "EndDateTimeGMT",
    "LapTime",
    "StateName",
]

VESSEL_GROUPS = {
    "Fleet 1": ["ATETI", "CMA CGM THALASSA", "CZECH", "DOLPHIN II", "GSL CHRISTEL ELISABETH", "GSL VINIA", "ORCA I", "MYNY", "SYDNEY EXPRESS"],
    "Fleet 2": ["AGIOS DIMITRIOS", "ELENI T", "MAIRA", "MELINA", "NEWYORKER", "NIKOLAS", "TORRANCE"],
    "Fleet 3": ["BREMERHAVEN EXPRESS", "CMA CGM ALCAZAR", "GSL ALICE", "GSL CHATEAU D'IF", "GSL ELEFTHERIA", "GSL MAREN", "GSL MELINA", "ISTANBUL EXPRESS"],
    "Fleet 4": ["ANTHEA Y", "COLOMBIA EXPRESS", "COSTA RICA EXPRESS", "JAMAICA EXPRESS", "MEXICO EXPRESS", "NICARAGUA EXPRESS", "PANAMA EXPRESS", "ZIM NORFOLK", "ZIM XIAMEN"],
    "Fleet 9": ["CMA CGM AMERICA", "CMA CGM SAMBHAR", "GSL ELENI", "GSL GRANIA", "GSL KALLIOPI", "GSL NINGBO", "MSC QINGDAO", "MSC TIANJIN"],
    "Fleet 10": ["CAPTAIN THANASIS I", "CMA CGM JAMAICA", "GSL CHRISTEN", "GSL NICOLETTA", "GSL VALERIE", "JULIE", "KUMASI", "MANET"],
    "Fleet 11": ["ATHENA", "EPAMINONDAS", "IAN H", "MARIANNA I", "MSC ROMA", "TINA I"],
    "Fleet 12": ["GSL DOROTHEA", "GSL KITHIRA", "GSL MARIA", "GSL MELITA", "GSL SYROS", "GSL TEGEA", "GSL TINOS", "GSL TRIPOLI"],
    "Fleet 14": ["GSL CHLOE", "GSL ELIZABETH", "GSL MAMITSA", "GSL MERCER", "GSL ROSSI", "GSL SUSAN", "TONSBERG"],
    "Fleet 15": ["GSL ALEXANDRA", "GSL ARCADIA", "GSL EFFIE", "GSL LYDIA", "GSL MYNY", "GSL SOFIA", "GSL VIOLETTA", "KOSTAS K", "MARIA Y"],
}

VESSEL_OPTIONS = sorted({v for vessels in VESSEL_GROUPS.values() for v in vessels})

st.set_page_config(page_title=APP_TITLE, layout="wide")


# =============================================================================
# Styling
# =============================================================================


def apply_custom_css() -> None:
    
    st.markdown(
        """
        <style>
        :root {
            --bg: #090A1A;
            --panel: #12152B;
            --panel-soft: #191D3A;
            --border: rgba(139, 92, 246, 0.34);
            --text-soft: #B8C0D9;
            --primary: #8B5CF6;
            --secondary: #22D3EE;
            --accent: #A78BFA;
        }

        .stApp {
            background:
                __BACKGROUND_IMAGE_LAYER__
                radial-gradient(circle at top left, rgba(139, 92, 246, 0.34), transparent 34rem),
                radial-gradient(circle at top right, rgba(255, 176, 0, 0.10), transparent 30rem),
                linear-gradient(180deg, rgba(255, 216, 74, 0.04), transparent 22rem),
                var(--bg);
            background-position: center center;
            background-size: cover;
            background-attachment: fixed;
        }

        header[data-testid="stHeader"], header[data-testid="stHeader"] > div,
        div[data-testid="stToolbar"], div[data-testid="stDecoration"] {
            background: transparent !important;
            border: 0 !important;
            box-shadow: none !important;
        }

        .block-container {
            padding-top: 3.2rem;
            padding-bottom: 3rem;
            max-width: 1380px;
        }

        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #11100A 0%, #050505 100%);
            border-right: 1px solid var(--border);
        }

        section[data-testid="stSidebar"] > div {
            padding-bottom: 8rem !important;
        }

        section[data-testid="stSidebar"] label {
            color: #F5EFD8 !important;
            font-weight: 700 !important;
        }

        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        textarea {
            background-color: rgba(13, 13, 9, 0.90) !important;
            border: 1px solid rgba(255, 216, 74, 0.18) !important;
            border-radius: 14px !important;
            box-shadow: none !important;
            outline: none !important;
        }

        div[data-baseweb="select"] > div:focus-within,
        div[data-baseweb="input"] > div:focus-within,
        textarea:focus {
            border-color: rgba(255, 216, 74, 0.88) !important;
            box-shadow: 0 0 0 1px rgba(255, 216, 74, 0.55) !important;
        }

        [data-baseweb="tag"] {
            background: linear-gradient(135deg, rgba(255, 216, 74, 0.22), rgba(255, 176, 0, 0.14)) !important;
            border: 1px solid rgba(255, 216, 74, 0.38) !important;
            color: #FFF7CC !important;
            border-radius: 999px !important;
        }
        [data-baseweb="tag"] span, [data-baseweb="tag"] svg { color: #FFF7CC !important; }

        .dashboard-hero {
            padding: 1.8rem 2rem;
            border: 1px solid var(--border);
            border-radius: 24px;
            background: rgba(5, 5, 5, 0.40);
            box-shadow: inset 0 1px 0 rgba(255,216,74,0.20);
            margin-bottom: 1.1rem;
        }

        .eyebrow {
            color: var(--cyan);
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-size: 0.78rem;
            font-weight: 800;
            margin-bottom: 0.35rem;
        }

        .dashboard-title {
            font-size: clamp(2.4rem, 4vw, 4.4rem);
            line-height: 1.02;
            font-weight: 950;
            color: #FFFBEA;
            margin: 0;
            text-shadow: 0 3px 16px rgba(0,0,0,0.88);
        }

        .dashboard-subtitle {
            color: var(--text-soft);
            font-size: 1rem;
            margin-top: 0.8rem;
            text-shadow: 0 2px 10px rgba(0,0,0,0.82);
        }

        .section-title {
            font-size: 1.35rem;
            font-weight: 850;
            color: #FFFBEA;
            margin: 1.5rem 0 0.75rem 0;
        }

        .api-load-caption, .atlas-pill {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            margin: -0.2rem 0 1.0rem 0;
            padding: 0.38rem 0.72rem;
            border: 1px solid rgba(255, 216, 74, 0.22);
            border-radius: 999px;
            background: rgba(13, 13, 9, 0.46);
            color: #B8B29F;
            font-size: 0.80rem;
            font-weight: 650;
            backdrop-filter: blur(6px);
        }

        .api-load-caption span, .atlas-pill span {
            color: #FFF7CC;
            font-weight: 800;
        }

        div[data-testid="stMetric"] {
            background: rgba(13, 13, 9, 0.42) !important;
            border: 1px solid rgba(255, 216, 74, 0.32) !important;
            border-radius: 18px !important;
            padding: 0.85rem 1rem !important;
            box-shadow: inset 0 1px 0 rgba(255,216,74,0.14) !important;
        }

        div[data-testid="stMetricLabel"] p {
            color: #F5EFD8 !important;
            font-weight: 800 !important;
        }

        div[data-testid="stMetricValue"] {
            color: #FFFBEA !important;
            font-weight: 950 !important;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid var(--border);
            border-radius: 18px;
            overflow: hidden;
            box-shadow: 0 14px 36px rgba(0,0,0,0.30);
        }

        button[data-baseweb="tab"] {
            color: #CFC6A5 !important;
            font-weight: 750 !important;
        }

        button[data-baseweb="tab"][aria-selected="true"] {
            color: #FFD84A !important;
        }

        div[data-baseweb="tab-highlight"] {
            background-color: #FFD84A !important;
        }

        .stDownloadButton button, .stButton button {
            border-radius: 14px !important;
            border: 1px solid rgba(255, 216, 74, 0.45) !important;
            background: linear-gradient(135deg, rgba(255, 216, 74, 0.98), rgba(255, 176, 0, 0.86)) !important;
            color: #121008 !important;
            font-weight: 850 !important;
        }
        </style>
        """.replace("__BACKGROUND_IMAGE_LAYER__", background_image_layer),
        unsafe_allow_html=True,
    )


def dashboard_background_image_layer(image_url: str) -> str:
    if not image_url:
        return ""
    safe_url = image_url.replace("\\", "\\\\").replace("'", "\\'")
    return (
        "linear-gradient(rgba(5, 5, 5, 0.78), rgba(5, 5, 5, 0.88)),\n"
        f"                url('{safe_url}'),\n"
    )


def dashboard_background_image_url() -> str:
    source = read_secret("DASHBOARD_BACKGROUND_IMAGE")
    if source and re.match(r"^(https?://|data:)", source, flags=re.IGNORECASE):
        return source

    image_path = Path(source).expanduser() if source else DEFAULT_BACKGROUND_IMAGE
    if not image_path.is_absolute():
        image_path = APP_DIR / image_path
    if source and not image_path.is_file():
        image_path = DEFAULT_BACKGROUND_IMAGE

    if not image_path.is_file():
        return ""

    mime_type = mimetypes.guess_type(image_path.name)[0] or "image/png"
    encoded_image = base64.b64encode(image_path.read_bytes()).decode("ascii")
    return f"data:{mime_type};base64,{encoded_image}"


def render_header(selected_group: str, selected_vessels: list[str], selected_variables: list[str]) -> None:
    vessel_text = "All selected vessels" if len(selected_vessels) != 1 else selected_vessels[0]
    variable_text = "No variables selected" if not selected_variables else f"{len(selected_variables):,} selected variables"
    st.markdown(
        f"""
        <div class="dashboard-hero">
            <div class="eyebrow">Marorka API Noon Reports</div>
            <h1 class="dashboard-title">AtlasFlow</h1>
            <div class="dashboard-subtitle">
                {escape(selected_group)} | {escape(vessel_text)} | {escape(variable_text)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_api_load_caption(metadata: dict[str, Any] | None) -> None:
    metadata = metadata or {}
    last_load = metadata.get("loaded_at_local") or metadata.get("loaded_at_utc") or "-"
    last_load_display = str(last_load).replace(" EEST", "").replace(" EET", "")
    st.markdown(
        f"""
        <div class="api-load-caption">
            Last API load: <span>{escape(last_load_display)} LT</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# Secrets/auth/API helpers
# =============================================================================


class MarorkaConfigError(RuntimeError):
    pass


def read_secret(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name, os.getenv(name, default))
    except Exception:
        value = os.getenv(name, default)
    return str(value).strip() if value is not None else default


def app_timezone() -> ZoneInfo:
    timezone_name = read_secret("APP_TIMEZONE", "Europe/Athens")
    try:
        return ZoneInfo(timezone_name)
    except Exception:
        return ZoneInfo("Europe/Athens")


def local_time_label(dt_utc: datetime | None = None) -> str:
    dt_utc = dt_utc or datetime.now(timezone.utc)
    if dt_utc.tzinfo is None:
        dt_utc = dt_utc.replace(tzinfo=timezone.utc)
    local_dt = dt_utc.astimezone(app_timezone())
    return local_dt.strftime("%d-%m-%Y %H:%M:%S %Z")


def get_query_param(name: str, default: str = "") -> str:
    try:
        value = st.query_params.get(name, default)
    except Exception:
        value = st.experimental_get_query_params().get(name, [default])
    if isinstance(value, list):
        return str(value[0]) if value else default
    return str(value) if value is not None else default


def is_warmup_request() -> bool:
    return get_query_param("warmup", "0") == "1"


def warmup_token_is_valid() -> bool:
    expected_token = read_secret("WARMUP_TOKEN")
    provided_token = get_query_param("token", "")
    return bool(expected_token) and hmac.compare_digest(provided_token, expected_token)


def require_dashboard_password() -> None:
    dashboard_password = read_secret("DASHBOARD_PASSWORD")
    if not dashboard_password:
        return

    if st.session_state.get("dashboard_authenticated"):
        return

    apply_custom_css()
    st.markdown(
        """
        <div class="dashboard-hero">
            <div class="eyebrow">Secure access</div>
            <h1 class="dashboard-title">AtlasFlow</h1>
            <div class="dashboard-subtitle">Enter your dashboard password to continue.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    entered_password = st.text_input("Password", type="password")

    if st.button("Sign in", type="primary"):
        if hmac.compare_digest(entered_password, dashboard_password):
            st.session_state["dashboard_authenticated"] = True
            st.rerun()
        st.error("Invalid password.")

    st.stop()


def request_auth(username: str, password: str, auth_method: str) -> Any:
    method = auth_method.lower()
    if method == "basic":
        return HTTPBasicAuth(username, password)
    if method == "digest":
        return HTTPDigestAuth(username, password)
    if method == "bearer":
        return None
    if method in {"none", "anonymous", ""}:
        return None
    raise MarorkaConfigError("Unsupported MARORKA_AUTH_METHOD. Use basic, digest, bearer, or none.")


def request_headers(token: str, auth_method: str) -> dict[str, str]:
    headers = {"Accept": "application/json"}
    if auth_method.lower() == "bearer":
        if not token:
            raise MarorkaConfigError("MARORKA_TOKEN is required for bearer auth.")
        headers["Authorization"] = f"Bearer {token}"
    return headers


def build_odata_url(start_date: date) -> str:
    start_text = start_date.strftime("%Y-%m-%d")
    params = {
        "$filter": f"StartDateTimeGMT gt DateTime'{start_text}'",
        "$select": ",".join(SOURCE_COLUMNS),
    }
    return f"{ODATA_ENDPOINT}?{urlencode(params)}"


def extract_odata_page(payload: Any) -> tuple[list[dict[str, Any]], str | None]:
    if isinstance(payload, list):
        return payload, None

    if not isinstance(payload, dict):
        raise ValueError("Could not parse OData response payload.")

    rows = payload.get("value")
    next_link = payload.get("@odata.nextLink") or payload.get("odata.nextLink")

    if rows is None and isinstance(payload.get("d"), dict):
        data = payload["d"]
        rows = data.get("results")
        next_link = next_link or data.get("__next")

    if rows is None:
        raise ValueError("Could not find OData rows in the API response.")

    return rows, next_link


def rows_to_dataframe(rows: list[dict[str, Any]]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    if "__metadata" in df.columns:
        df = df.drop(columns=["__metadata"])
    for column in SOURCE_COLUMNS:
        if column not in df.columns:
            df[column] = pd.NA
    return df[SOURCE_COLUMNS]


def compact_odata_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    compact_rows: list[dict[str, Any]] = []
    for row in rows:
        if row.get("ValueDescription") is None:
            continue
        if row.get("ReportType") in EXCLUDED_REPORT_TYPES:
            continue
        compact_rows.append({column: row.get(column) for column in SOURCE_COLUMNS})
    return compact_rows


def fetch_report_data(
    username: str,
    password: str,
    token: str,
    auth_method: str,
    start_date: date,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    started_at = time.perf_counter()
    next_url = build_odata_url(start_date)
    kept_rows: list[dict[str, Any]] = []
    seen_urls: set[str] = set()
    pages = 0
    total_bytes = 0
    scanned_rows = 0
    first_url = next_url
    auth = request_auth(username, password, auth_method)
    headers = request_headers(token, auth_method)

    with requests.Session() as session:
        session.headers.update(headers)
        for _ in range(MAX_ODATA_PAGES):
            if next_url in seen_urls:
                break
            seen_urls.add(next_url)

            response = session.get(next_url, auth=auth, timeout=90)
            total_bytes += len(response.content)
            response.raise_for_status()
            pages += 1

            page_rows, next_link = extract_odata_page(response.json())
            scanned_rows += len(page_rows)
            kept_rows.extend(compact_odata_rows(page_rows))

            if not next_link:
                break
            next_url = urljoin(next_url, next_link)

    loaded_at_utc = datetime.now(timezone.utc)
    metadata = {
        "loaded_at_utc": loaded_at_utc.strftime("%d-%m-%Y %H:%M:%S UTC"),
        "loaded_at_local": local_time_label(loaded_at_utc),
        "rows": len(kept_rows),
        "kept_rows": len(kept_rows),
        "scanned_rows": scanned_rows,
        "discarded_rows": max(scanned_rows - len(kept_rows), 0),
        "pages": pages,
        "downloaded_mb": round(total_bytes / 1024 / 1024, 2),
        "fetch_seconds": round(time.perf_counter() - started_at, 2),
        "first_url": first_url,
        "hit_page_limit": pages >= MAX_ODATA_PAGES,
    }
    return rows_to_dataframe(kept_rows), metadata


@st.cache_data(ttl=API_CACHE_TTL_SECONDS, show_spinner=False)
def cached_fetch_report_data(
    username: str,
    password: str,
    token: str,
    auth_method: str,
    start_date: date,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    return fetch_report_data(username, password, token, auth_method, start_date)


# =============================================================================
# Transform helpers
# =============================================================================


def normalize_text(value: Any) -> str:
    text = str(value).lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def parse_datetime_series(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce", utc=True)
    missing_mask = parsed.isna()

    if missing_mask.any():
        date_text = series.astype("string")
        dotnet_millis = date_text.str.extract(r"/Date\((-?\d+)").iloc[:, 0]
        dotnet_parsed = pd.to_datetime(
            pd.to_numeric(dotnet_millis, errors="coerce"),
            errors="coerce",
            unit="ms",
            utc=True,
        )
        parsed = parsed.mask(missing_mask, dotnet_parsed)

    return parsed


def parse_numeric_value(value: Any) -> Any:
    if pd.isna(value):
        return pd.NA
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return pd.NA

    duration_match = re.fullmatch(r"(-?\d+):([0-5]?\d)(?::([0-5]?\d))?", text)
    if duration_match:
        hours = int(duration_match.group(1))
        sign = -1 if hours < 0 else 1
        minutes = int(duration_match.group(2))
        seconds = int(duration_match.group(3) or 0)
        return sign * (abs(hours) + minutes / 60 + seconds / 3600)

    numeric_text = text.replace(" ", "")
    if re.fullmatch(r"-?\d+,\d+", numeric_text):
        numeric_text = numeric_text.replace(",", ".")
    else:
        numeric_text = numeric_text.replace(",", "")

    numeric_text = re.sub(r"[^0-9.\-]", "", numeric_text)
    if numeric_text in {"", "-", ".", "-."}:
        return pd.NA

    try:
        return float(numeric_text)
    except ValueError:
        return pd.NA


def parse_numeric_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.map(parse_numeric_value), errors="coerce")


def match_selected_vessels(raw_ship_names: pd.Series, selected_vessels: list[str]) -> pd.Series:
    selected_keys = {normalize_text(vessel) for vessel in selected_vessels}
    return raw_ship_names.map(normalize_text).isin(selected_keys)


@st.cache_data(ttl=API_CACHE_TTL_SECONDS, show_spinner=False)
def cached_prepare_long_data(raw_df: pd.DataFrame) -> pd.DataFrame:
    missing_columns = sorted(set(SOURCE_COLUMNS).difference(raw_df.columns))
    if missing_columns:
        raise ValueError(f"Missing expected API columns: {', '.join(missing_columns)}")

    df = raw_df.copy()
    df["StartDateTimeGMT"] = parse_datetime_series(df["StartDateTimeGMT"])
    df["EndDateTimeGMT"] = parse_datetime_series(df["EndDateTimeGMT"])
    df["LapTime"] = parse_numeric_series(df["LapTime"])
    df["ParsedValue"] = parse_numeric_series(df["ReportedValue"])
    df = df[df["ValueDescription"].notna() & ~df["ReportType"].isin(EXCLUDED_REPORT_TYPES)].copy()
    return df


def available_variables(df: pd.DataFrame) -> list[str]:
    if df.empty or "ValueDescription" not in df.columns:
        return []
    return sorted(df["ValueDescription"].dropna().astype(str).unique().tolist(), key=str.casefold)


def available_report_types(df: pd.DataFrame) -> list[str]:
    if df.empty or "ReportType" not in df.columns:
        return []
    return sorted(df["ReportType"].dropna().astype(str).unique().tolist(), key=str.casefold)


def dataframe_date_window(df: pd.DataFrame) -> tuple[date, date]:
    if df.empty or "StartDateTimeGMT" not in df.columns:
        today = date.today()
        return today, today
    dates = pd.to_datetime(df["StartDateTimeGMT"], errors="coerce", utc=True).dt.date.dropna()
    if dates.empty:
        today = date.today()
        return today, today
    return max(dates.min(), API_FULL_START_DATE), min(dates.max(), date.today())


def filter_long_data(
    df: pd.DataFrame,
    selected_vessels: list[str],
    selected_report_types: list[str],
    selected_start: date,
    selected_end: date,
) -> pd.DataFrame:
    if df.empty:
        return df

    start_timestamp = pd.Timestamp(selected_start, tz="UTC")
    end_timestamp = pd.Timestamp(selected_end + timedelta(days=1), tz="UTC")
    start_values = pd.to_datetime(df["StartDateTimeGMT"], errors="coerce", utc=True)

    filtered = df[
        match_selected_vessels(df["ShipName"], selected_vessels)
        & start_values.ge(start_timestamp)
        & start_values.lt(end_timestamp)
    ].copy()

    if selected_report_types:
        filtered = filtered[filtered["ReportType"].astype("string").isin(selected_report_types)].copy()

    return filtered


@st.cache_data(show_spinner=False)
def build_pivot_table(filtered_long_df: pd.DataFrame, selected_variables: tuple[str, ...]) -> pd.DataFrame:
    if filtered_long_df.empty:
        return pd.DataFrame(columns=PIVOT_IDENTITY_COLUMNS + list(selected_variables))

    if not selected_variables:
        return (
            filtered_long_df[PIVOT_IDENTITY_COLUMNS]
            .drop_duplicates()
            .sort_values(["ShipName", "EndDateTimeGMT"], ascending=[True, False])
            .reset_index(drop=True)
        )

    selected_long = filtered_long_df[
        filtered_long_df["ValueDescription"].astype("string").isin(list(selected_variables))
    ].copy()

    if selected_long.empty:
        base = filtered_long_df[PIVOT_IDENTITY_COLUMNS].drop_duplicates().copy()
        for variable in selected_variables:
            base[variable] = pd.NA
        return base.sort_values(["ShipName", "EndDateTimeGMT"], ascending=[True, False]).reset_index(drop=True)

    selected_long["ValueDescription"] = selected_long["ValueDescription"].astype(str)
    selected_long["_source_order"] = range(len(selected_long))
    selected_long = selected_long.sort_values("_source_order")
    selected_long = selected_long.drop_duplicates(
        [*PIVOT_IDENTITY_COLUMNS, "ValueDescription"],
        keep="last",
    )

    pivot_df = (
        selected_long
        .pivot(index=PIVOT_IDENTITY_COLUMNS, columns="ValueDescription", values="ParsedValue")
        .reset_index()
    )
    pivot_df.columns.name = None

    for variable in selected_variables:
        if variable not in pivot_df.columns:
            pivot_df[variable] = pd.NA

    return pivot_df.sort_values(["ShipName", "EndDateTimeGMT"], ascending=[True, False]).reset_index(drop=True)


# =============================================================================
# Display/export helpers
# =============================================================================


def format_display_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    display_df = df.copy()
    for column in ["StartDateTimeGMT", "EndDateTimeGMT"]:
        if column in display_df.columns:
            display_df[column] = pd.to_datetime(display_df[column], errors="coerce").dt.strftime(DISPLAY_DATETIME_FORMAT)
    for column in display_df.columns:
        if column in {"ReportId", "ShipName", "ReportType", "StartDateTimeGMT", "EndDateTimeGMT", "StateName"}:
            continue
        values = pd.to_numeric(display_df[column], errors="coerce")
        if values.notna().any():
            display_df[column] = values.map(lambda value: "-" if pd.isna(value) else f"{value:,.3f}")
    return display_df.fillna("-")


@st.cache_data(show_spinner=False)
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    safe_df = df.copy()
    for column in safe_df.columns:
        if pd.api.types.is_datetime64_any_dtype(safe_df[column]):
            safe_df[column] = pd.to_datetime(safe_df[column], errors="coerce").dt.tz_localize(None)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        safe_df.to_excel(writer, index=False, sheet_name="AtlasFlow")
        worksheet = writer.sheets["AtlasFlow"]
        for column_cells in worksheet.columns:
            max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 48)
    return output.getvalue()


def filter_digest(column: str) -> str:
    return sha256(column.encode("utf-8")).hexdigest()[:10]


def parse_optional_float(value: str) -> tuple[float | None, bool]:
    text = str(value or "").strip()
    if not text:
        return None, True
    normalized = text.replace(" ", "").replace(",", "")
    try:
        return float(normalized), True
    except ValueError:
        return None, False


def parse_optional_date(value: str) -> tuple[pd.Timestamp | None, bool]:
    text = str(value or "").strip()
    if not text:
        return None, True
    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce", utc=True)
    if pd.isna(parsed):
        return None, False
    return parsed, True


def unique_display_values(series: pd.Series, limit: int = 500) -> list[str]:
    values = series.astype("string").fillna("(Blank)").drop_duplicates().tolist()
    values = sorted(values, key=lambda value: str(value).casefold())
    return values[:limit]


def is_numeric_like(series: pd.Series) -> bool:
    values = pd.to_numeric(series, errors="coerce")
    return values.notna().any()


def render_column_filters(df: pd.DataFrame, filter_columns: list[str]) -> list[dict[str, Any]]:
    specs: list[dict[str, Any]] = []
    for column in filter_columns:
        if column not in df.columns:
            continue

        st.caption(f"Filter: {column}")
        series = df[column]
        digest = filter_digest(column)

        if pd.api.types.is_datetime64_any_dtype(series):
            left, right = st.columns(2)
            from_text = left.text_input("From", key=f"atlas_filter_{digest}_from", placeholder="dd/mm/yyyy")
            to_text = right.text_input("To", key=f"atlas_filter_{digest}_to", placeholder="dd/mm/yyyy")
            from_value, from_ok = parse_optional_date(from_text)
            to_value, to_ok = parse_optional_date(to_text)
            if not from_ok or not to_ok:
                st.warning(f"{column}: enter dates as dd/mm/yyyy or yyyy-mm-dd.")
            specs.append({"column": column, "kind": "datetime", "from": from_value, "to": to_value})
            continue

        if is_numeric_like(series):
            values = pd.to_numeric(series, errors="coerce").dropna()
            if not values.empty:
                st.caption(f"Loaded range: {values.min():,.3f} to {values.max():,.3f}")
            left, right = st.columns(2)
            min_text = left.text_input("Min", key=f"atlas_filter_{digest}_min", placeholder="no minimum")
            max_text = right.text_input("Max", key=f"atlas_filter_{digest}_max", placeholder="no maximum")
            minimum, min_ok = parse_optional_float(min_text)
            maximum, max_ok = parse_optional_float(max_text)
            if not min_ok or not max_ok:
                st.warning(f"{column}: enter numeric Min/Max values only.")
            if minimum is not None and maximum is not None and minimum > maximum:
                minimum, maximum = maximum, minimum
            specs.append({"column": column, "kind": "numeric", "min": minimum, "max": maximum})
            continue

        values_key = f"atlas_filter_{digest}_values"
        selected_values = st.multiselect(
            "Values",
            options=unique_display_values(series),
            key=values_key,
            help="Leave blank to include all values for this column.",
        )
        specs.append({"column": column, "kind": "categorical", "values": selected_values})

    return specs


def apply_column_filters(df: pd.DataFrame, specs: list[dict[str, Any]]) -> pd.DataFrame:
    filtered = df.copy()
    for spec in specs:
        column = spec.get("column")
        if column not in filtered.columns:
            continue

        kind = spec.get("kind")
        if kind == "numeric":
            values = pd.to_numeric(filtered[column], errors="coerce")
            minimum = spec.get("min")
            maximum = spec.get("max")
            if minimum is not None:
                filtered = filtered[values >= minimum]
                values = pd.to_numeric(filtered[column], errors="coerce")
            if maximum is not None:
                filtered = filtered[values <= maximum]

        elif kind == "datetime":
            values = pd.to_datetime(filtered[column], errors="coerce", utc=True)
            from_value = spec.get("from")
            to_value = spec.get("to")
            if from_value is not None:
                filtered = filtered[values >= from_value]
                values = pd.to_datetime(filtered[column], errors="coerce", utc=True)
            if to_value is not None:
                filtered = filtered[values < (to_value + pd.Timedelta(days=1))]

        elif kind == "categorical":
            selected_values = spec.get("values") or []
            if selected_values:
                values = filtered[column].astype("string").fillna("(Blank)")
                filtered = filtered[values.isin(selected_values)]

    return filtered


# =============================================================================
# Sidebar/session helpers
# =============================================================================


def request_signature(username: str, auth_method: str, start_date: date) -> dict[str, Any]:
    return {
        "endpoint": ODATA_ENDPOINT,
        "username_hash": sha256(username.encode("utf-8")).hexdigest()[:12],
        "auth_method": auth_method.lower(),
        "start_date": start_date.isoformat(),
    }


def raw_data_covers_request(
    loaded_signature: dict[str, Any] | None,
    metadata: dict[str, Any] | None,
    requested_signature: dict[str, Any],
    requested_start_date: date,
) -> bool:
    if not loaded_signature or not metadata:
        return False
    for key in ["endpoint", "username_hash", "auth_method"]:
        if loaded_signature.get(key) != requested_signature.get(key):
            return False
    loaded_start_text = metadata.get("loaded_start_date") or loaded_signature.get("start_date")
    try:
        loaded_start_date = date.fromisoformat(str(loaded_start_text))
    except ValueError:
        return False
    return loaded_start_date <= requested_start_date


def get_loaded_state() -> tuple[pd.DataFrame | None, pd.DataFrame | None, dict[str, Any] | None]:
    return (
        st.session_state.get("loaded_raw_df"),
        st.session_state.get("loaded_long_df"),
        st.session_state.get("loaded_metadata"),
    )


def set_loaded_raw_state(raw_df: pd.DataFrame, metadata: dict[str, Any], signature: dict[str, Any]) -> None:
    metadata = metadata.copy()
    metadata.setdefault("loaded_at_utc", datetime.now(timezone.utc).strftime("%d-%m-%Y %H:%M:%S UTC"))
    metadata.setdefault("loaded_at_local", local_time_label())
    metadata["loaded_start_date"] = signature["start_date"]
    st.session_state["loaded_raw_df"] = raw_df
    st.session_state["loaded_metadata"] = metadata
    st.session_state["loaded_request_signature"] = signature
    st.session_state.pop("loaded_long_df", None)
    st.session_state.pop("loaded_prepare_signature", None)


def set_loaded_long_state(long_df: pd.DataFrame, signature: dict[str, Any]) -> None:
    st.session_state["loaded_long_df"] = long_df
    st.session_state["loaded_prepare_signature"] = signature


def selected_vessel_controls() -> tuple[str, list[str]]:
    group_options = ["Single vessel", "All fleets"] + list(VESSEL_GROUPS.keys())
    selected_group = st.sidebar.selectbox("Fleet group", options=group_options)

    if selected_group == "Single vessel":
        vessel = st.sidebar.selectbox("Vessel to include", options=VESSEL_OPTIONS)
        return selected_group, [vessel]

    if selected_group == "All fleets":
        group_vessels = VESSEL_OPTIONS
    else:
        group_vessels = VESSEL_GROUPS[selected_group]

    vessels = st.sidebar.multiselect(
        "Vessels to include",
        options=group_vessels,
        default=group_vessels,
        help="This controls the displayed pivot table only. The API data is loaded broadly.",
    )

    if not vessels:
        st.sidebar.caption("No vessels selected manually, so all vessels in this fleet group are included.")
        vessels = group_vessels

    return selected_group, vessels


def sidebar_refresh_control() -> bool:
    refresh_requested = st.sidebar.button("Refresh API data", use_container_width=False)
    if refresh_requested:
        st.session_state["confirm_api_refresh"] = True

    refresh = False
    if st.session_state.get("confirm_api_refresh"):
        metadata = st.session_state.get("loaded_metadata") or {}
        last_load = metadata.get("loaded_at_local") or metadata.get("loaded_at_utc") or "-"
        last_load_display = str(last_load).replace(" EEST", "").replace(" EET", "")
        st.sidebar.warning(
            "Refresh will call the API and may take a while.\n\n"
            f"Last updated data was on: {last_load_display} LT"
        )
        col1, col2 = st.sidebar.columns(2)
        if col1.button("Confirm"):
            refresh = True
            st.session_state["confirm_api_refresh"] = False
        if col2.button("Cancel"):
            st.session_state["confirm_api_refresh"] = False
            st.rerun()
    return refresh


def render_date_slicer(df: pd.DataFrame) -> tuple[date, date]:
    min_date, max_date = dataframe_date_window(df)
    st.sidebar.markdown("### Period")
    if min_date >= max_date:
        st.sidebar.caption(f"Available data period: {min_date.strftime('%d/%m/%Y')}")
        return min_date, max_date
    selected_start, selected_end = st.sidebar.slider(
        "Report period",
        min_value=min_date,
        max_value=max_date,
        value=(min_date, max_date),
        format="DD/MM/YYYY",
        key="atlas_period_slicer",
    )
    return selected_start, selected_end


# =============================================================================
# Warmup
# =============================================================================


def run_warmup_if_requested() -> None:
    if not is_warmup_request():
        return

    apply_custom_css()
    if not warmup_token_is_valid():
        st.error("Invalid or missing warmup token.")
        st.stop()

    username = read_secret("MARORKA_USERNAME")
    password = read_secret("MARORKA_PASSWORD")
    token = read_secret("MARORKA_TOKEN")
    auth_method = read_secret("MARORKA_AUTH_METHOD", "basic")
    start_date = API_FULL_START_DATE

    if auth_method.lower() in {"basic", "digest"} and (not username or not password):
        st.error("Warmup failed: MARORKA_USERNAME and MARORKA_PASSWORD are required.")
        st.stop()

    if get_query_param("force", "0") == "1":
        cached_fetch_report_data.clear()
        cached_prepare_long_data.clear()
        build_pivot_table.clear()

    try:
        with st.spinner("Warming up API..."):
            raw_df, metadata = cached_fetch_report_data(username, password, token, auth_method, start_date)
            long_df = cached_prepare_long_data(raw_df)
    except requests.HTTPError as exc:
        status = exc.response.status_code if exc.response is not None else "unknown"
        st.error(f"Warmup failed: Marorka API request failed with status {status}.")
        st.stop()
    except (MarorkaConfigError, ValueError, requests.RequestException) as exc:
        st.error(f"Warmup failed: {exc}")
        st.stop()

    st.success("Warmup OK.")
    st.write(
        {
            "last_api_load_local": metadata.get("loaded_at_local"),
            "compact_api_rows": int(len(raw_df)),
            "long_rows": int(len(long_df)),
            "available_variables": int(long_df["ValueDescription"].nunique()) if "ValueDescription" in long_df.columns else 0,
            "force_refresh": get_query_param("force", "0") == "1",
        }
    )
    st.stop()


# =============================================================================
# Main app
# =============================================================================


def main() -> None:
    run_warmup_if_requested()
    require_dashboard_password()
    apply_custom_css()

    username = read_secret("MARORKA_USERNAME")
    password = read_secret("MARORKA_PASSWORD")
    token = read_secret("MARORKA_TOKEN")
    auth_method = read_secret("MARORKA_AUTH_METHOD", "basic")

    if auth_method.lower() in {"basic", "digest"} and (not username or not password):
        st.info("Add MARORKA_USERNAME and MARORKA_PASSWORD to .streamlit/secrets.toml or Streamlit Cloud Secrets.")
        st.stop()

    api_start_date = API_FULL_START_DATE
    api_end_date = date.today()

    refresh = sidebar_refresh_control()
    selected_group, selected_vessels = selected_vessel_controls()

    raw_signature = request_signature(username, auth_method, api_start_date)
    current_raw_signature = st.session_state.get("loaded_request_signature")
    raw_df, long_df, metadata = get_loaded_state()

    needs_raw_load = (
        refresh
        or raw_df is None
        or metadata is None
        or not raw_data_covers_request(current_raw_signature, metadata, raw_signature, api_start_date)
    )

    if needs_raw_load:
        if refresh:
            cached_fetch_report_data.clear()
            cached_prepare_long_data.clear()
            build_pivot_table.clear()
        try:
            with st.spinner("Refreshing API..." if refresh else "Loading API..."):
                raw_df, metadata = cached_fetch_report_data(
                    username=username,
                    password=password,
                    token=token,
                    auth_method=auth_method,
                    start_date=api_start_date,
                )
            set_loaded_raw_state(raw_df, metadata, raw_signature)
            long_df = None
        except requests.HTTPError as exc:
            status = exc.response.status_code if exc.response is not None else "unknown"
            st.error(f"Marorka API request failed with status {status}.")
            st.caption("If credentials are correct, try MARORKA_AUTH_METHOD = 'digest'.")
            if exc.response is not None and exc.response.request is not None:
                st.code(exc.response.request.url, language="text")
            st.stop()
        except (MarorkaConfigError, ValueError, requests.RequestException) as exc:
            st.error(str(exc))
            st.stop()

    prepare_signature = {**raw_signature, "prepare_version": "atlasflow_dynamic_pivot_v1"}
    current_prepare_signature = st.session_state.get("loaded_prepare_signature")
    raw_df = st.session_state.get("loaded_raw_df")
    long_df = st.session_state.get("loaded_long_df")
    metadata = st.session_state.get("loaded_metadata")

    if raw_df is None or metadata is None:
        st.info("Loading Marorka data automatically. Use Refresh API data to force a new API pull.")
        st.stop()

    if long_df is None or current_prepare_signature != prepare_signature:
        try:
            transform_started_at = time.perf_counter()
            long_df = cached_prepare_long_data(raw_df)
            set_loaded_long_state(long_df, prepare_signature)
            metadata = st.session_state.get("loaded_metadata")
            if isinstance(metadata, dict):
                metadata["prepare_seconds"] = round(time.perf_counter() - transform_started_at, 2)
                metadata["long_rows"] = int(len(long_df))
                metadata["available_variables"] = int(long_df["ValueDescription"].nunique()) if "ValueDescription" in long_df.columns else 0
                st.session_state["loaded_metadata"] = metadata
        except (ValueError, TypeError) as exc:
            st.error(str(exc))
            st.stop()

    if long_df.empty:
        render_header(selected_group, selected_vessels, [])
        render_api_load_caption(metadata)
        st.warning("No Marorka report values were returned for the loaded API window.")
        st.stop()

    selected_start, selected_end = render_date_slicer(long_df)

    report_type_options = available_report_types(long_df)
    selected_report_types = st.sidebar.multiselect(
        "Report types to include",
        options=report_type_options,
        default=report_type_options,
        help="Leave all selected to include the full report-type range.",
    )

    filtered_long_for_options = filter_long_data(
        long_df,
        selected_vessels=selected_vessels,
        selected_report_types=selected_report_types,
        selected_start=selected_start,
        selected_end=selected_end,
    )

    variable_options = available_variables(filtered_long_for_options)
    st.sidebar.markdown("### Pivot variables")
    selected_variables = st.sidebar.multiselect(
        "Variables to include and filter",
        options=variable_options,
        default=st.session_state.get("atlas_selected_variables", []),
        key="atlas_selected_variables",
        help=(
            "Every selected ValueDescription becomes a displayed table column. "
            "The same selected variables are also available below as filters."
        ),
    )

    identity_columns = st.sidebar.multiselect(
        "Report identity columns",
        options=PIVOT_IDENTITY_COLUMNS,
        default=DEFAULT_DISPLAY_IDENTITY_COLUMNS,
        help="These fixed report fields appear before the selected variable columns.",
    )
    if not identity_columns:
        identity_columns = DEFAULT_DISPLAY_IDENTITY_COLUMNS

    render_header(selected_group, selected_vessels, selected_variables)
    render_api_load_caption(metadata)

    if not selected_variables:
        st.info("Select one or more variables from the sidebar to build the AtlasFlow pivot table.")

    pivot_df = build_pivot_table(filtered_long_for_options, tuple(selected_variables))

    filter_column_options = [column for column in [*identity_columns, *selected_variables] if column in pivot_df.columns]
    with st.sidebar.expander("Filters for displayed columns", expanded=False):
        st.caption("Choose columns to filter. Selected variables are already part of the displayed table.")
        columns_to_filter = st.multiselect(
            "Columns to filter",
            options=filter_column_options,
            default=[],
            key="atlas_columns_to_filter",
        )
        filter_specs = render_column_filters(pivot_df, columns_to_filter)

    filtered_pivot_df = apply_column_filters(pivot_df, filter_specs)

    display_columns = []
    for column in [*identity_columns, *selected_variables]:
        if column in filtered_pivot_df.columns and column not in display_columns:
            display_columns.append(column)
    if not display_columns:
        display_columns = [column for column in DEFAULT_DISPLAY_IDENTITY_COLUMNS if column in filtered_pivot_df.columns]

    output_df = filtered_pivot_df[display_columns].copy()

    tab_table, tab_diagnostics, tab_raw = st.tabs(["Pivot Table", "API Diagnostics", "Long Data"])

    if metadata.get("hit_page_limit"):
        st.warning(
            "The API refresh reached the page safety limit. The loaded dataset may be incomplete. "
            "Check API Diagnostics before using the export."
        )

    with tab_table:
        st.markdown('<div class="section-title">AtlasFlow Pivot Table</div>', unsafe_allow_html=True)

        metric_cols = st.columns(4)
        metric_cols[0].metric("Displayed rows", f"{len(output_df):,}")
        metric_cols[1].metric("Selected variables", f"{len(selected_variables):,}")
        metric_cols[2].metric("Source long rows", f"{len(filtered_long_for_options):,}")
        metric_cols[3].metric("Available variables", f"{len(variable_options):,}")

        preview_df = output_df.head(TABLE_PREVIEW_ROW_LIMIT)
        st.dataframe(format_display_dataframe(preview_df), use_container_width=True, hide_index=True)
        if len(output_df) > TABLE_PREVIEW_ROW_LIMIT:
            st.caption(
                f"Showing first {TABLE_PREVIEW_ROW_LIMIT:,} of {len(output_df):,} rows. "
                "Use the Excel export for the full filtered pivot table."
            )

        export_ready = (
            st.session_state.get("atlas_export_signature") == sha256(
                pd.util.hash_pandas_object(output_df, index=True).values.tobytes()
            ).hexdigest()
            and "atlas_export_bytes" in st.session_state
        )

        if st.button("Prepare Excel download", type="primary"):
            with st.spinner("Preparing Excel file..."):
                signature = sha256(pd.util.hash_pandas_object(output_df, index=True).values.tobytes()).hexdigest()
                st.session_state["atlas_export_bytes"] = to_excel_bytes(output_df)
                st.session_state["atlas_export_signature"] = signature
            export_ready = True

        if export_ready:
            st.download_button(
                "Download AtlasFlow Excel",
                data=st.session_state["atlas_export_bytes"],
                file_name="atlasflow_pivot_table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.caption("Excel generation is prepared on demand so normal dashboard loads stay faster.")

    with tab_diagnostics:
        st.markdown('<div class="section-title">Diagnostics</div>', unsafe_allow_html=True)
        diagnostics = pd.DataFrame(
            {
                "Metric": [
                    "Selected vessels",
                    "API start date",
                    "API end date",
                    "Selected start",
                    "Selected end",
                    "API loaded at",
                    "API loaded local time",
                    "Kept raw rows",
                    "Original API rows scanned",
                    "Discarded rows",
                    "Long rows prepared",
                    "Filtered long rows",
                    "Pivot rows before column filters",
                    "Pivot rows after column filters",
                    "Available variables",
                    "Selected variables",
                    "API pages",
                    "Downloaded MB",
                    "API fetch seconds",
                    "Prepare seconds",
                    "Hit API page limit",
                ],
                "Value": [
                    ", ".join(selected_vessels),
                    api_start_date.isoformat(),
                    api_end_date.isoformat(),
                    selected_start.isoformat(),
                    selected_end.isoformat(),
                    metadata.get("loaded_at_utc", "-"),
                    metadata.get("loaded_at_local", "-"),
                    f"{metadata.get('kept_rows', metadata.get('rows', 0)):,}",
                    f"{metadata.get('scanned_rows', 0):,}",
                    f"{metadata.get('discarded_rows', 0):,}",
                    f"{len(long_df):,}",
                    f"{len(filtered_long_for_options):,}",
                    f"{len(pivot_df):,}",
                    f"{len(output_df):,}",
                    f"{len(variable_options):,}",
                    f"{len(selected_variables):,}",
                    f"{metadata.get('pages', 0):,}",
                    metadata.get("downloaded_mb", "-"),
                    metadata.get("fetch_seconds", "-"),
                    metadata.get("prepare_seconds", "-"),
                    str(metadata.get("hit_page_limit", "-")),
                ],
            }
        )
        st.dataframe(diagnostics, use_container_width=True, hide_index=True)

        with st.expander("First API URL", expanded=False):
            st.code(metadata.get("first_url", "-"), language="text")

        st.markdown('<div class="section-title">Available Variable Counts</div>', unsafe_allow_html=True)
        if st.button("Calculate variable counts"):
            value_counts = (
                filtered_long_for_options.get("ValueDescription", pd.Series(dtype="object"))
                .value_counts(dropna=False)
                .reset_index()
            )
            value_counts.columns = ["ValueDescription", "Rows"]
            st.dataframe(value_counts.head(500), use_container_width=True, hide_index=True)
        else:
            st.caption("Variable counts are calculated on demand so diagnostics do not slow normal loads.")

    with tab_raw:
        st.markdown('<div class="section-title">Filtered Long Data</div>', unsafe_allow_html=True)
        raw_preview_columns = [column for column in SOURCE_COLUMNS if column in filtered_long_for_options.columns]
        raw_preview = filtered_long_for_options[raw_preview_columns].head(TABLE_PREVIEW_ROW_LIMIT).copy()
        st.dataframe(format_display_dataframe(raw_preview), use_container_width=True, hide_index=True)
        if len(filtered_long_for_options) > TABLE_PREVIEW_ROW_LIMIT:
            st.caption(f"Showing first {TABLE_PREVIEW_ROW_LIMIT:,} of {len(filtered_long_for_options):,} long rows.")


if __name__ == "__main__":
    main()
