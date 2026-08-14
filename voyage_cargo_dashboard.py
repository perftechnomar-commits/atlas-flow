from __future__ import annotations

from contextlib import contextmanager
from datetime import date, datetime, timedelta, timezone
from io import BytesIO
import hmac
import json
import os
from pathlib import Path
import re
import time
from typing import Any
from urllib.parse import urlencode, urljoin
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import streamlit as st
from requests.auth import HTTPBasicAuth, HTTPDigestAuth

try:
    import fcntl
except ImportError:  # Streamlit Cloud is Linux; local Windows fallback is below.
    fcntl = None


# =============================================================================
# Configuration
# =============================================================================

APP_TITLE = "Cargo Voyage Dashboard"
APP_DIR = Path(__file__).resolve().parent
CACHE_DIR = APP_DIR / ".cargo_voyage_cache"
MANIFEST_FILE = CACHE_DIR / "manifest.json"
REFRESH_LOCK_FILE = CACHE_DIR / "refresh.lock"
REFRESH_STATUS_FILE = CACHE_DIR / "refresh_status.json"
SCHEMA_VERSION = "2026-08-14-cargo-voyage-monthly-v1"

SERVICE_ROOT = "https://online.marorka.com/Odata/v1/ODataService.svc"
ENDPOINTS = {
    "reportdata": f"{SERVICE_ROOT}/ReportData",
    "reportpivots": f"{SERVICE_ROOT}/ReportPivots",
    "shippivots": f"{SERVICE_ROOT}/ShipPivots",
}
SOURCE_LABELS = {
    "reportdata": "ReportData",
    "reportpivots": "ReportPivots",
    "shippivots": "ShipPivots",
}
DATE_FIELDS = {
    "reportdata": "StartDateTimeGMT",
    "reportpivots": "DateTime",
    "shippivots": "DateTime",
}
DEFAULT_CHUNK_DAYS = {
    "reportdata": 31,
    "reportpivots": 31,
    "shippivots": 7,
}
DEFAULT_OVERLAP_DAYS = 14
MAX_ODATA_PAGES_PER_WINDOW = 1000
REQUEST_TIMEOUT_SECONDS = 75
REQUEST_MAX_ATTEMPTS = 4
WARMUP_REPLAY_GUARD_SECONDS = 600

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

REPORTPIVOTS_SELECT = [
    "ShipName",
    "DateTime",
    "CompanyName",
    "IMONo",
    "DeparturePort",
    "ArrivalPort",
    "DraftFore",
    "DraftAft",
    "CargoWeight",
    "CargoTEU",
]

SHIPPIVOTS_SELECT = [
    "Id",
    "ShipName",
    "IMONo",
    "CompanyName",
    "DateTime",
    "State",
    "Latitude",
    "Longitude",
    "Cargo",
    "DraftAft",
    "DraftFore",
    "VoyageIdInternal",
    "VoyageId",
]

# Curated from the live Marorka scan performed on 14 Aug 2026.
CARGO_VALUE_DESCRIPTIONS = {
    "Port",
    "20ft DG Units",
    "20ft Empty Units",
    "20ft Empty Units Weight [tons]",
    "20ft Full Units",
    "20ft Full Units Weight [tons]",
    "20ft Reefer Units",
    "40ft DG Units",
    "40ft Empty Units",
    "40ft Empty Units Weight [tons]",
    "40ft Full Units",
    "40ft Full Units Weight [tons]",
    "40ft Reefer Units",
    "48/54ft Units Weight [tons]",
    "Air Draft [m]",
    "Ballast Amount [tons]",
    "Calculated Draft Aft [m]",
    "Calculated Draft Forward [m]",
    "Calculated Mean Draft [m]",
    "Cargo Checked: Bridges",
    "Cargo Checked: Lashings",
    "Cargo Operations Completed During Port Stay",
    "Cargo Weight [tons]",
    "Cargo Weight Added [MT]",
    "Cargo Weight Removed [MT]",
    "Commenced Cargo Operation Time [dd:mm:yyyy hh:mm]",
    "Completed Cargo Operation Time [dd:mm:yyyy hh:mm]",
    "DG Units Weight [tons]",
    "Dead Load [tons]",
    "Draft Aft [m] (m)",
    "Draft Forward [m] (m)",
    "FEU Discharged Units",
    "FEU Discharged Weight [tons]",
    "FEU Loaded Units",
    "FEU Loaded Weight [tons]",
    "Observed Draft Aft [m]",
    "Observed Draft Forward [m]",
    "Observed Mean Draft [m]",
    "Reefer Units Weight [tons]",
    "Reefers Discharged Units",
    "Reefers Discharged Weight [tons]",
    "Reefers Loaded Units",
    "Reefers Loaded Weight [tons]",
    "TEU Discharged Units",
    "TEU Discharged Weight [tons]",
    "TEU Loaded Units",
    "TEU Loaded Weight [tons]",
    "Total 20ft Units Weight (Full and Empty)",
    "Total 40 ft Units Weight (Full and Empty)",
    "Total Empty Units Weight (20 and 40ft) [tons]",
    "Total Full Units Weight (20 and 40ft) [tons]",
    "Total Number DG Units (20 and 40ft)",
    "Total Number Empty Units (20 and 40ft)",
    "Total Number Full Units (20 and 40ft)",
    "Total Number Reefer Units (20 and 40ft)",
    "Total Number of 20ft Units (Full and Empty)",
    "Total Number of 40ft Units (Full and Empty)",
    "Total Units Weight (All Categories)",
}
CARGO_VALUE_KEYS = {re.sub(r"[^a-z0-9]+", "", x.lower()) for x in CARGO_VALUE_DESCRIPTIONS}

OPERATION_FIELDS = [
    "Cargo Weight Added [MT]",
    "Cargo Weight Removed [MT]",
    "TEU Loaded Units",
    "TEU Discharged Units",
    "TEU Loaded Weight [tons]",
    "TEU Discharged Weight [tons]",
    "FEU Loaded Units",
    "FEU Discharged Units",
    "FEU Loaded Weight [tons]",
    "FEU Discharged Weight [tons]",
    "Reefers Loaded Units",
    "Reefers Discharged Units",
    "Reefers Loaded Weight [tons]",
    "Reefers Discharged Weight [tons]",
]
COMPOSITION_FIELDS = [
    "Cargo Weight [tons]",
    "20ft Full Units",
    "20ft Empty Units",
    "40ft Full Units",
    "40ft Empty Units",
    "20ft Reefer Units",
    "40ft Reefer Units",
    "Total Number Full Units (20 and 40ft)",
    "Total Number Empty Units (20 and 40ft)",
    "Total Number Reefer Units (20 and 40ft)",
    "Total Units Weight (All Categories)",
]


st.set_page_config(page_title=APP_TITLE, layout="wide")


# =============================================================================
# Styling
# =============================================================================

def apply_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            --ink:#0B1F33; --muted:#52657A; --teal:#006B68; --teal2:#0AAEA6;
            --line:#D9E6E5; --bg:#F7FBFA; --card:#FFFFFF;
        }
        html, body, .stApp { background:var(--bg)!important; color:var(--ink)!important; }
        .block-container { max-width:1440px; padding-top:2.7rem; padding-bottom:3rem; }
        section[data-testid="stSidebar"] { background:linear-gradient(180deg,#006A66,#003C39)!important; }
        section[data-testid="stSidebar"] * { color:white!important; }
        section[data-testid="stSidebar"] div[data-baseweb="select"] *,
        section[data-testid="stSidebar"] div[data-baseweb="input"] *,
        section[data-testid="stSidebar"] input { color:#0B1F33!important; -webkit-text-fill-color:#0B1F33!important; }
        .hero { background:white; border:1px solid var(--line); border-radius:15px; padding:1.45rem 1.7rem; box-shadow:0 4px 18px rgba(15,23,42,.07); margin-bottom:1rem; }
        .eyebrow { color:var(--teal); text-transform:uppercase; letter-spacing:.14em; font-size:.76rem; font-weight:800; }
        .hero h1 { margin:.25rem 0 .35rem; font-weight:450; font-size:clamp(2.4rem,4vw,3.7rem); color:var(--ink); }
        .hero p { margin:0; color:var(--muted); }
        .load-pill { display:inline-flex; gap:.45rem; align-items:center; border:1px solid var(--line); border-radius:999px; background:white; padding:.4rem .8rem; color:var(--muted); font-size:.82rem; margin:0 0 1rem; }
        .load-pill strong { color:var(--ink); }
        .section-title { font-size:1.45rem; font-weight:500; color:var(--ink); margin:1.35rem 0 .6rem; }
        div[data-testid="stMetric"] { background:white!important; border:1px solid var(--line)!important; border-radius:12px!important; box-shadow:0 2px 10px rgba(15,23,42,.05)!important; padding:.75rem 1rem!important; }
        div[data-testid="stMetricLabel"] * { color:var(--muted)!important; font-weight:700!important; }
        div[data-testid="stMetricValue"] * { color:var(--ink)!important; font-weight:500!important; }
        div[data-testid="stDataFrame"] { border:1px solid var(--line); border-radius:10px; overflow:hidden; background:white; }
        .stButton button, .stDownloadButton button { background:linear-gradient(135deg,var(--teal),var(--teal2))!important; color:white!important; border:0!important; border-radius:8px!important; }
        </style>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# Secrets / auth / utilities
# =============================================================================

class CargoConfigError(RuntimeError):
    pass


def read_secret(name: str, default: str = "") -> str:
    try:
        value = st.secrets.get(name, os.getenv(name, default))
    except Exception:
        value = os.getenv(name, default)
    return str(value).strip() if value is not None else default


def read_int_secret(name: str, default: int, minimum: int, maximum: int) -> int:
    try:
        value = int(read_secret(name, str(default)))
    except (TypeError, ValueError):
        value = default
    return min(max(value, minimum), maximum)


def full_start_date() -> date:
    raw = read_secret("CARGO_FULL_START_DATE", "2026-01-01")
    try:
        return date.fromisoformat(raw)
    except ValueError:
        return date(2026, 1, 1)


def app_timezone() -> ZoneInfo:
    try:
        return ZoneInfo(read_secret("APP_TIMEZONE", "Europe/Athens"))
    except Exception:
        return ZoneInfo("Europe/Athens")


def local_time_label(dt: datetime | None = None) -> str:
    dt = dt or datetime.now(timezone.utc)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(app_timezone()).strftime("%d-%m-%Y %H:%M:%S %Z")


def get_query_param(name: str, default: str = "") -> str:
    try:
        value = st.query_params.get(name, default)
    except Exception:
        value = st.experimental_get_query_params().get(name, [default])
    if isinstance(value, list):
        value = value[0] if value else default
    return str(value) if value is not None else default


def request_auth(username: str, password: str, auth_method: str) -> Any:
    method = auth_method.lower()
    if method == "basic":
        return HTTPBasicAuth(username, password)
    if method == "digest":
        return HTTPDigestAuth(username, password)
    if method in {"none", "anonymous", "", "bearer"}:
        return None
    raise CargoConfigError("MARORKA_AUTH_METHOD must be basic, digest, bearer, or none.")


def request_headers(token: str, auth_method: str) -> dict[str, str]:
    headers = {"Accept": "application/json"}
    if auth_method.lower() == "bearer":
        if not token:
            raise CargoConfigError("MARORKA_TOKEN is required for bearer authentication.")
        headers["Authorization"] = f"Bearer {token}"
    return headers


RETRYABLE_STATUSES = {429, 500, 502, 503, 504}
RETRYABLE_EXCEPTIONS = (
    requests.ConnectionError,
    requests.ReadTimeout,
    requests.Timeout,
    requests.exceptions.ChunkedEncodingError,
)


def request_with_retry(session: requests.Session, url: str, auth: Any) -> requests.Response:
    last_error: Exception | None = None
    for attempt in range(1, REQUEST_MAX_ATTEMPTS + 1):
        try:
            response = session.get(url, auth=auth, timeout=REQUEST_TIMEOUT_SECONDS)
            if response.status_code in RETRYABLE_STATUSES and attempt < REQUEST_MAX_ATTEMPTS:
                retry_after = response.headers.get("Retry-After")
                try:
                    sleep_seconds = float(retry_after) if retry_after else 2 ** (attempt - 1)
                except ValueError:
                    sleep_seconds = 2 ** (attempt - 1)
                time.sleep(min(sleep_seconds, 30.0))
                continue
            return response
        except RETRYABLE_EXCEPTIONS as exc:
            last_error = exc
            if attempt >= REQUEST_MAX_ATTEMPTS:
                raise
            time.sleep(min(2 ** (attempt - 1), 30.0))
    if last_error:
        raise last_error
    raise requests.RequestException("Marorka request failed before a response was received.")


def extract_odata_page(payload: Any) -> tuple[list[dict[str, Any]], str | None]:
    if isinstance(payload, list):
        return payload, None
    if not isinstance(payload, dict):
        raise ValueError("Unexpected OData response payload.")
    rows = payload.get("value")
    next_link = payload.get("@odata.nextLink") or payload.get("odata.nextLink")
    if rows is None and isinstance(payload.get("d"), dict):
        data = payload["d"]
        rows = data.get("results")
        next_link = next_link or data.get("__next")
    if rows is None:
        raise ValueError("Could not find OData rows in Marorka response.")
    return rows, next_link


def normalize_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def parse_dt(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce", utc=True)
    missing = parsed.isna()
    if missing.any():
        dotnet = series.astype("string").str.extract(r"/Date\((-?\d+)").iloc[:, 0]
        fallback = pd.to_datetime(pd.to_numeric(dotnet, errors="coerce"), unit="ms", errors="coerce", utc=True)
        parsed = parsed.mask(missing, fallback)
    return parsed


def parse_num(value: Any) -> Any:
    if value is None or pd.isna(value):
        return pd.NA
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(" ", "")
    if not text:
        return pd.NA
    if re.fullmatch(r"-?\d+,\d+", text):
        text = text.replace(",", ".")
    else:
        text = text.replace(",", "")
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in {"", "-", ".", "-."}:
        return pd.NA
    try:
        return float(text)
    except ValueError:
        return pd.NA


def fmt_number(value: Any, decimals: int = 1, suffix: str = "") -> str:
    if value is None or pd.isna(value):
        return "-"
    try:
        return f"{float(value):,.{decimals}f}{suffix}"
    except (TypeError, ValueError):
        return str(value)


def month_key(value: date | pd.Timestamp) -> str:
    return f"{value.year:04d}-{value.month:02d}"


def month_starts_between(start_date: date, end_date: date) -> list[str]:
    cursor = date(start_date.year, start_date.month, 1)
    end_month = date(end_date.year, end_date.month, 1)
    keys: list[str] = []
    while cursor <= end_month:
        keys.append(month_key(cursor))
        cursor = date(cursor.year + 1, 1, 1) if cursor.month == 12 else date(cursor.year, cursor.month + 1, 1)
    return keys


def partition_path(source_key: str, key: str) -> Path:
    return CACHE_DIR / source_key / f"{key}.parquet"


def read_manifest() -> dict[str, Any]:
    try:
        if not MANIFEST_FILE.is_file():
            return {"schema_version": SCHEMA_VERSION, "sources": {}}
        data = json.loads(MANIFEST_FILE.read_text(encoding="utf-8"))
        if data.get("schema_version") != SCHEMA_VERSION:
            return {"schema_version": SCHEMA_VERSION, "sources": {}}
        return data
    except Exception:
        return {"schema_version": SCHEMA_VERSION, "sources": {}}


def atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(f"{path.name}.{os.getpid()}.{int(time.time()*1000)}.tmp")
    tmp.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    os.replace(str(tmp), str(path))


def read_refresh_status() -> dict[str, Any]:
    try:
        return json.loads(REFRESH_STATUS_FILE.read_text(encoding="utf-8")) if REFRESH_STATUS_FILE.is_file() else {}
    except Exception:
        return {}


def update_refresh_status(**updates: Any) -> None:
    status = read_refresh_status()
    status.update(updates)
    status["updated_at_utc"] = datetime.now(timezone.utc).isoformat()
    atomic_write_json(REFRESH_STATUS_FILE, status)


def recent_successful_refresh() -> bool:
    status = read_refresh_status()
    if status.get("state") != "complete":
        return False
    raw = status.get("updated_at_utc")
    if not raw:
        return False
    try:
        dt = datetime.fromisoformat(str(raw).replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        age = (datetime.now(timezone.utc) - dt.astimezone(timezone.utc)).total_seconds()
        return 0 <= age <= WARMUP_REPLAY_GUARD_SECONDS
    except Exception:
        return False


@contextmanager
def refresh_lock() -> Any:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    if fcntl is not None:
        handle = REFRESH_LOCK_FILE.open("a+", encoding="utf-8")
        try:
            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError:
                yield False
                return
            yield True
        finally:
            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
            except Exception:
                pass
            handle.close()
        return

    fd: int | None = None
    try:
        try:
            fd = os.open(str(REFRESH_LOCK_FILE), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            yield False
            return
        yield True
    finally:
        if fd is not None:
            os.close(fd)
            try:
                REFRESH_LOCK_FILE.unlink()
            except FileNotFoundError:
                pass


# =============================================================================
# API refresh / monthly persistent snapshots
# =============================================================================

def build_window_url(source_key: str, start_date: date, end_exclusive: date) -> str:
    date_field = DATE_FIELDS[source_key]
    query_start = start_date - timedelta(days=1)
    params: dict[str, str] = {
        "$filter": (
            f"{date_field} gt DateTime'{query_start.isoformat()}' and "
            f"{date_field} lt DateTime'{end_exclusive.isoformat()}'"
        ),
        "$orderby": f"{date_field} asc",
    }
    if source_key == "reportdata":
        params["$select"] = ",".join(REPORTDATA_SELECT)
    elif source_key == "reportpivots":
        params["$select"] = ",".join(REPORTPIVOTS_SELECT)
    elif source_key == "shippivots":
        params["$select"] = ",".join(SHIPPIVOTS_SELECT)
    return f"{ENDPOINTS[source_key]}?{urlencode(params)}"


def compact_source_rows(source_key: str, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if source_key != "reportdata":
        return rows
    kept: list[dict[str, Any]] = []
    for row in rows:
        desc = row.get("ValueDescription")
        if desc is None:
            continue
        if normalize_key(desc) not in CARGO_VALUE_KEYS:
            continue
        kept.append({column: row.get(column) for column in REPORTDATA_SELECT})
    return kept


def normalize_source_frame(source_key: str, df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        if source_key == "reportdata":
            return pd.DataFrame(columns=REPORTDATA_SELECT)
        return df
    out = df.copy()
    if "__metadata" in out.columns:
        out = out.drop(columns=["__metadata"])
    date_field = DATE_FIELDS[source_key]
    if date_field in out.columns:
        out[date_field] = parse_dt(out[date_field])
    if source_key == "reportdata":
        if "EndDateTimeGMT" in out.columns:
            out["EndDateTimeGMT"] = parse_dt(out["EndDateTimeGMT"])
        if "ReportId" in out.columns:
            out["ReportId"] = pd.to_numeric(out["ReportId"], errors="coerce").astype("Int64")
    return out


def fetch_window(
    source_key: str,
    start_date: date,
    end_exclusive: date,
    username: str,
    password: str,
    token: str,
    auth_method: str,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    auth = request_auth(username, password, auth_method)
    headers = request_headers(token, auth_method)
    next_url = build_window_url(source_key, start_date, end_exclusive)
    first_url = next_url
    seen: set[str] = set()
    frames: list[pd.DataFrame] = []
    pages = scanned = kept = total_bytes = 0
    started = time.perf_counter()
    has_more = False

    with requests.Session() as session:
        session.headers.update(headers)
        for _ in range(MAX_ODATA_PAGES_PER_WINDOW):
            if next_url in seen:
                break
            seen.add(next_url)
            response = request_with_retry(session, next_url, auth)
            total_bytes += len(response.content)
            response.raise_for_status()
            page_rows, next_link = extract_odata_page(response.json())
            pages += 1
            scanned += len(page_rows)
            compact = compact_source_rows(source_key, page_rows)
            kept += len(compact)
            if compact:
                frames.append(pd.DataFrame(compact))
            if not next_link:
                has_more = False
                break
            has_more = True
            next_url = urljoin(next_url, next_link)

    if pages >= MAX_ODATA_PAGES_PER_WINDOW and has_more:
        raise RuntimeError(
            f"{SOURCE_LABELS[source_key]} reached the {MAX_ODATA_PAGES_PER_WINDOW:,}-page limit "
            f"for {start_date} to {end_exclusive}. Reduce the configured chunk size."
        )

    df = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
    df = normalize_source_frame(source_key, df)
    date_field = DATE_FIELDS[source_key]
    if date_field in df.columns:
        dates = pd.to_datetime(df[date_field], errors="coerce", utc=True)
        start_ts = pd.Timestamp(start_date, tz="UTC")
        end_ts = pd.Timestamp(end_exclusive, tz="UTC")
        df = df.loc[dates.ge(start_ts) & dates.lt(end_ts)].copy()

    return df, {
        "pages": pages,
        "scanned_rows": scanned,
        "kept_rows": int(len(df)),
        "downloaded_mb": round(total_bytes / 1024 / 1024, 2),
        "fetch_seconds": round(time.perf_counter() - started, 2),
        "first_url": first_url,
    }


def dedupe_source(source_key: str, df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    if source_key == "reportdata" and {"ReportId", "ValueDescription"}.issubset(df.columns):
        return df.drop_duplicates(["ReportId", "ValueDescription"], keep="last")
    if source_key == "shippivots" and "Id" in df.columns:
        return df.drop_duplicates(["Id"], keep="last")
    keys = [c for c in ["ShipName", DATE_FIELDS[source_key]] if c in df.columns]
    return df.drop_duplicates(keys, keep="last") if keys else df.drop_duplicates(keep="last")


def publish_month(source_key: str, key: str, fresh_df: pd.DataFrame, refresh_start: date) -> int:
    path = partition_path(source_key, key)
    path.parent.mkdir(parents=True, exist_ok=True)
    month_start = date(int(key[:4]), int(key[5:7]), 1)
    month_end = date(month_start.year + 1, 1, 1) if month_start.month == 12 else date(month_start.year, month_start.month + 1, 1)

    existing = pd.DataFrame()
    if path.is_file():
        try:
            existing = pd.read_parquet(path)
        except Exception:
            existing = pd.DataFrame()
    if not existing.empty:
        date_field = DATE_FIELDS[source_key]
        values = pd.to_datetime(existing[date_field], errors="coerce", utc=True)
        cutoff = pd.Timestamp(max(refresh_start, month_start), tz="UTC")
        existing = existing.loc[values.lt(cutoff)].copy()

    combined = pd.concat([existing, fresh_df], ignore_index=True, sort=False)
    combined = normalize_source_frame(source_key, combined)
    combined = dedupe_source(source_key, combined)
    date_field = DATE_FIELDS[source_key]
    if date_field in combined.columns:
        values = pd.to_datetime(combined[date_field], errors="coerce", utc=True)
        combined = combined.loc[
            values.ge(pd.Timestamp(month_start, tz="UTC")) & values.lt(pd.Timestamp(month_end, tz="UTC"))
        ].copy()
        combined = combined.sort_values([c for c in ["ShipName", date_field] if c in combined.columns])

    tmp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
    combined.to_parquet(tmp, index=False, compression="zstd")
    os.replace(str(tmp), str(path))
    return int(len(combined))


def refresh_source(
    source_key: str,
    username: str,
    password: str,
    token: str,
    auth_method: str,
    manifest: dict[str, Any],
    *,
    full_refresh: bool,
) -> dict[str, Any]:
    start_floor = full_start_date()
    source_meta = dict((manifest.get("sources") or {}).get(source_key) or {})
    refresh_start = start_floor
    if not full_refresh:
        latest = source_meta.get("latest_source_date")
        try:
            latest_date = date.fromisoformat(str(latest)) if latest else None
        except ValueError:
            latest_date = None
        if latest_date:
            overlap = read_int_secret("CARGO_INCREMENTAL_OVERLAP_DAYS", DEFAULT_OVERLAP_DAYS, 1, 60)
            refresh_start = max(start_floor, latest_date - timedelta(days=overlap))

    chunk_days = read_int_secret(
        f"CARGO_{source_key.upper()}_CHUNK_DAYS",
        DEFAULT_CHUNK_DAYS[source_key],
        1,
        62,
    )
    end_exclusive = date.today() + timedelta(days=1)
    windows: list[tuple[date, date]] = []
    cursor = refresh_start
    while cursor < end_exclusive:
        nxt = min(cursor + timedelta(days=chunk_days), end_exclusive)
        windows.append((cursor, nxt))
        cursor = nxt

    fresh_by_month: dict[str, list[pd.DataFrame]] = {}
    total_pages = total_scanned = total_kept = 0
    total_mb = total_seconds = 0.0
    latest_source_date: date | None = None

    for i, (window_start, window_end) in enumerate(windows, start=1):
        update_refresh_status(
            state="running",
            source=source_key,
            stage="fetching",
            window=f"{i}/{len(windows)}",
            window_start=window_start.isoformat(),
            window_end_exclusive=window_end.isoformat(),
        )
        frame, meta = fetch_window(source_key, window_start, window_end, username, password, token, auth_method)
        total_pages += int(meta["pages"])
        total_scanned += int(meta["scanned_rows"])
        total_kept += int(meta["kept_rows"])
        total_mb += float(meta["downloaded_mb"])
        total_seconds += float(meta["fetch_seconds"])
        if not frame.empty:
            date_field = DATE_FIELDS[source_key]
            dates = pd.to_datetime(frame[date_field], errors="coerce", utc=True)
            if dates.notna().any():
                dmax = dates.max().date()
                if latest_source_date is None or dmax > latest_source_date:
                    latest_source_date = dmax
            for key, group in frame.groupby(dates.dt.strftime("%Y-%m")):
                if isinstance(key, str) and key:
                    fresh_by_month.setdefault(key, []).append(group.copy())

    touched_months = month_starts_between(refresh_start, date.today())
    partition_rows: dict[str, int] = {}
    for key in touched_months:
        fresh = pd.concat(fresh_by_month.get(key, []), ignore_index=True, sort=False) if fresh_by_month.get(key) else pd.DataFrame()
        partition_rows[key] = publish_month(source_key, key, fresh, refresh_start)

    loaded_at = datetime.now(timezone.utc)
    return {
        "loaded_at_utc": loaded_at.isoformat(),
        "loaded_at_local": local_time_label(loaded_at),
        "refresh_start_date": refresh_start.isoformat(),
        "latest_source_date": (latest_source_date or date.fromisoformat(str(source_meta.get("latest_source_date", start_floor.isoformat())))).isoformat(),
        "pages": total_pages,
        "scanned_rows": total_scanned,
        "kept_rows_last_refresh": total_kept,
        "downloaded_mb": round(total_mb, 2),
        "fetch_seconds": round(total_seconds, 2),
        "chunk_days": chunk_days,
        "partitions": partition_rows,
    }


def refresh_all_sources(username: str, password: str, token: str, auth_method: str, *, full_refresh: bool) -> dict[str, Any]:
    manifest = read_manifest()
    manifest.setdefault("sources", {})
    for source_key in ["shippivots", "reportpivots", "reportdata"]:
        meta = refresh_source(source_key, username, password, token, auth_method, manifest, full_refresh=full_refresh)
        manifest["sources"][source_key] = meta
        manifest["schema_version"] = SCHEMA_VERSION
        manifest["updated_at_utc"] = datetime.now(timezone.utc).isoformat()
        atomic_write_json(MANIFEST_FILE, manifest)
    return manifest


# =============================================================================
# Snapshot reads / voyage model
# =============================================================================

@st.cache_data(show_spinner=False, ttl=900)
def available_vessels_from_snapshot(manifest_updated: str) -> list[str]:
    del manifest_updated
    manifest = read_manifest()
    months = sorted(((manifest.get("sources") or {}).get("shippivots") or {}).get("partitions", {}).keys(), reverse=True)
    vessels: set[str] = set()
    for key in months[:3]:
        path = partition_path("shippivots", key)
        if not path.is_file():
            continue
        try:
            series = pd.read_parquet(path, columns=["ShipName"])["ShipName"]
            vessels.update(series.dropna().astype(str).str.strip().tolist())
        except Exception:
            continue
    return sorted(v for v in vessels if v)


@st.cache_data(show_spinner=False, ttl=900)
def read_vessel_source(source_key: str, vessel: str, start_date: date, end_date: date, manifest_updated: str) -> pd.DataFrame:
    del manifest_updated
    frames: list[pd.DataFrame] = []
    for key in month_starts_between(start_date, end_date):
        path = partition_path(source_key, key)
        if not path.is_file():
            continue
        try:
            frame = pd.read_parquet(path, filters=[("ShipName", "==", vessel)])
        except Exception:
            try:
                frame = pd.read_parquet(path)
                frame = frame[frame.get("ShipName", pd.Series(dtype="string")).astype("string") == vessel]
            except Exception:
                continue
        if not frame.empty:
            frames.append(frame)
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True, sort=False)
    date_field = DATE_FIELDS[source_key]
    if date_field in df.columns:
        values = pd.to_datetime(df[date_field], errors="coerce", utc=True)
        df = df.loc[
            values.ge(pd.Timestamp(start_date, tz="UTC"))
            & values.lt(pd.Timestamp(end_date + timedelta(days=1), tz="UTC"))
        ].copy()
    return df


def voyage_table(ship_df: pd.DataFrame) -> pd.DataFrame:
    if ship_df.empty or "VoyageId" not in ship_df.columns:
        return pd.DataFrame()
    work = ship_df.copy()
    work["DateTime"] = pd.to_datetime(work["DateTime"], errors="coerce", utc=True)
    work["VoyageId"] = work["VoyageId"].astype("string").str.strip()
    work = work[work["VoyageId"].notna() & work["VoyageId"].ne("") & work["DateTime"].notna()].copy()
    if work.empty:
        return pd.DataFrame()
    grouped = (
        work.groupby(["VoyageId", "VoyageIdInternal"], dropna=False)
        .agg(
            VoyageStart=("DateTime", "min"),
            VoyageEnd=("DateTime", "max"),
            Samples=("DateTime", "count"),
            StartLatitude=("Latitude", "first"),
            StartLongitude=("Longitude", "first"),
            EndLatitude=("Latitude", "last"),
            EndLongitude=("Longitude", "last"),
        )
        .reset_index()
        .sort_values("VoyageStart", ascending=False)
    )
    grouped["DurationHours"] = (grouped["VoyageEnd"] - grouped["VoyageStart"]).dt.total_seconds() / 3600
    return grouped


def filter_to_voyage(df: pd.DataFrame, date_column: str, start: pd.Timestamp, end: pd.Timestamp, pad_hours: int = 2) -> pd.DataFrame:
    if df.empty or date_column not in df.columns:
        return df.iloc[0:0].copy()
    values = pd.to_datetime(df[date_column], errors="coerce", utc=True)
    pad = pd.Timedelta(hours=pad_hours)
    return df.loc[values.ge(start - pad) & values.le(end + pad)].copy()


def pivot_cargo_reports(reportdata_df: pd.DataFrame) -> pd.DataFrame:
    if reportdata_df.empty:
        return pd.DataFrame()
    work = reportdata_df.copy()
    work["ParsedValue"] = work["ReportedValue"].map(parse_num)
    work["_order"] = range(len(work))
    work = work.sort_values("_order").drop_duplicates(["ReportId", "ValueDescription"], keep="last")
    identity = ["ReportId", "ShipName", "ReportType", "StartDateTimeGMT", "EndDateTimeGMT", "StateName"]
    pivot = work.pivot(index=identity, columns="ValueDescription", values="ReportedValue").reset_index()
    pivot.columns.name = None
    for field in CARGO_VALUE_DESCRIPTIONS:
        if field in pivot.columns:
            pivot[field] = pivot[field].map(parse_num) if field not in {
                "Port",
                "Cargo Checked: Bridges",
                "Cargo Checked: Lashings",
                "Cargo Operations Completed During Port Stay",
                "Commenced Cargo Operation Time [dd:mm:yyyy hh:mm]",
                "Completed Cargo Operation Time [dd:mm:yyyy hh:mm]",
            } else pivot[field]
    return pivot.sort_values("StartDateTimeGMT")


def non_null_value(df: pd.DataFrame, column: str, *, last: bool = True) -> Any:
    if df.empty or column not in df.columns:
        return pd.NA
    values = df[column].dropna()
    if values.empty:
        return pd.NA
    return values.iloc[-1] if last else values.iloc[0]


def numeric_sum(df: pd.DataFrame, column: str) -> float | Any:
    if df.empty or column not in df.columns:
        return pd.NA
    values = pd.to_numeric(df[column], errors="coerce")
    return values.sum(min_count=1)


def build_report_timeline(reportpivots: pd.DataFrame, cargo_by_report: pd.DataFrame) -> pd.DataFrame:
    cargo_cols = [
        "ReportId", "ReportType", "StartDateTimeGMT", "EndDateTimeGMT", "StateName",
        "Port", "Cargo Weight [tons]", "TEU Loaded Units", "TEU Discharged Units",
        "FEU Loaded Units", "FEU Discharged Units", "Reefers Loaded Units", "Reefers Discharged Units",
        "Draft Forward [m] (m)", "Draft Aft [m] (m)",
    ]
    cargo_cols = [c for c in cargo_cols if c in cargo_by_report.columns]
    timeline = cargo_by_report[cargo_cols].copy() if cargo_cols else pd.DataFrame()
    if not timeline.empty:
        timeline = timeline.rename(columns={"StartDateTimeGMT": "DateTime", "StateName": "State"})
        timeline["Source"] = "ReportData"

    if not reportpivots.empty:
        rp_cols = [c for c in ["DateTime", "DeparturePort", "ArrivalPort", "CargoWeight", "CargoTEU", "DraftFore", "DraftAft"] if c in reportpivots.columns]
        rp = reportpivots[rp_cols].copy()
        rp["Source"] = "ReportPivots"
        timeline = pd.concat([timeline, rp], ignore_index=True, sort=False)
    if timeline.empty:
        return timeline
    timeline["DateTime"] = pd.to_datetime(timeline["DateTime"], errors="coerce", utc=True)
    return timeline.sort_values("DateTime")


def to_excel_bytes(
    overview: pd.DataFrame,
    timeline: pd.DataFrame,
    cargo_by_report: pd.DataFrame,
    reportpivots: pd.DataFrame,
    shippivots: pd.DataFrame,
) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        overview.to_excel(writer, index=False, sheet_name="Voyage Overview")
        timeline.to_excel(writer, index=False, sheet_name="Timeline")
        cargo_by_report.to_excel(writer, index=False, sheet_name="Cargo By Report")
        reportpivots.to_excel(writer, index=False, sheet_name="ReportPivots")
        shippivots.to_excel(writer, index=False, sheet_name="ShipPivots")
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = min(max(max(len(str(cell.value)) if cell.value is not None else 0 for cell in col) + 2, 11), 42)
                ws.column_dimensions[col[0].column_letter].width = width
    return out.getvalue()


# =============================================================================
# Warmup / access
# =============================================================================

def warmup_token_valid() -> bool:
    expected = read_secret("WARMUP_TOKEN")
    provided = get_query_param("token")
    return bool(expected) and hmac.compare_digest(expected, provided)


def run_warmup_if_requested() -> None:
    if get_query_param("warmup", "0") != "1":
        return
    apply_css()
    if not warmup_token_valid():
        st.error("Invalid or missing warmup token.")
        st.stop()

    username = read_secret("MARORKA_USERNAME")
    password = read_secret("MARORKA_PASSWORD")
    token = read_secret("MARORKA_TOKEN")
    auth_method = read_secret("MARORKA_AUTH_METHOD", "basic")
    if auth_method.lower() in {"basic", "digest"} and (not username or not password):
        st.error("MARORKA_USERNAME and MARORKA_PASSWORD are required.")
        st.stop()

    force = get_query_param("force", "0") == "1"
    full = get_query_param("full", "0") == "1"
    force_again = get_query_param("force_again", "0") == "1"
    started = time.perf_counter()

    manifest = read_manifest()
    has_snapshot = bool((manifest.get("sources") or {}).get("shippivots"))
    replay_guard = bool(force and not force_again and has_snapshot and recent_successful_refresh())

    try:
        if force and not replay_guard:
            with refresh_lock() as acquired:
                if not acquired:
                    st.info("Another Cargo Dashboard refresh is already running. Existing snapshots remain available.")
                    st.write(read_refresh_status())
                    st.stop()
                update_refresh_status(state="running", stage="starting", started_at_utc=datetime.now(timezone.utc).isoformat())
                with st.spinner("Refreshing ShipPivots, ReportPivots and cargo ReportData..."):
                    manifest = refresh_all_sources(username, password, token, auth_method, full_refresh=full)
                update_refresh_status(state="complete", stage="ready")
        elif not has_snapshot:
            with refresh_lock() as acquired:
                if not acquired:
                    st.info("Initial Cargo Dashboard snapshot is being created by another request.")
                    st.stop()
                update_refresh_status(state="running", stage="initial_bootstrap")
                with st.spinner("Creating the first Cargo Dashboard snapshot..."):
                    manifest = refresh_all_sources(username, password, token, auth_method, full_refresh=True)
                update_refresh_status(state="complete", stage="ready")
    except Exception as exc:
        update_refresh_status(state="failed", stage="failed", error=str(exc))
        st.error(f"Warmup failed: {exc}")
        st.stop()

    manifest = read_manifest()
    st.success(
        "Warmup OK. Cargo voyage snapshots are ready."
        if not replay_guard
        else "Warmup OK. Recent successful refresh reused; duplicate API refresh prevented."
    )
    st.write({
        "schema_version": manifest.get("schema_version"),
        "updated_at_utc": manifest.get("updated_at_utc"),
        "sources": {
            key: {
                "last_load": value.get("loaded_at_local"),
                "latest_source_date": value.get("latest_source_date"),
                "pages": value.get("pages"),
                "kept_rows_last_refresh": value.get("kept_rows_last_refresh"),
            }
            for key, value in (manifest.get("sources") or {}).items()
        },
        "force_requested": force,
        "full_refresh": full,
        "replay_guard_applied": replay_guard,
        "warmup_seconds": round(time.perf_counter() - started, 2),
    })
    st.stop()


def require_password() -> None:
    expected = read_secret("DASHBOARD_PASSWORD")
    if not expected or st.session_state.get("cargo_authenticated"):
        return
    apply_css()
    st.markdown('<div class="hero"><div class="eyebrow">Secure access</div><h1>Cargo Voyage Dashboard</h1><p>Enter the dashboard password to continue.</p></div>', unsafe_allow_html=True)
    entered = st.text_input("Password", type="password")
    if st.button("Sign in"):
        if hmac.compare_digest(entered, expected):
            st.session_state["cargo_authenticated"] = True
            st.rerun()
        st.error("Invalid password.")
    st.stop()


# =============================================================================
# Main UI
# =============================================================================

def main() -> None:
    run_warmup_if_requested()
    require_password()
    apply_css()

    manifest = read_manifest()
    sources = manifest.get("sources") or {}
    if not all(sources.get(key) for key in ["shippivots", "reportpivots", "reportdata"]):
        st.markdown('<div class="hero"><div class="eyebrow">Marorka cargo intelligence</div><h1>Cargo Voyage Dashboard</h1><p>No prepared snapshot is available yet.</p></div>', unsafe_allow_html=True)
        st.info("Run the warmup URL once with ?warmup=1&force=1&token=YOUR_WARMUP_TOKEN.")
        st.stop()

    manifest_updated = str(manifest.get("updated_at_utc", ""))
    vessels = available_vessels_from_snapshot(manifest_updated)
    if not vessels:
        st.error("No vessels were found in the prepared ShipPivots snapshot.")
        st.stop()

    selected_vessel = st.sidebar.selectbox("Vessel", vessels)
    available_start = full_start_date()
    latest_dates = []
    for meta in sources.values():
        try:
            latest_dates.append(date.fromisoformat(str(meta.get("latest_source_date"))))
        except Exception:
            pass
    available_end = min(max(latest_dates) if latest_dates else date.today(), date.today())

    period = st.sidebar.date_input(
        "Voyage search period",
        value=(max(available_start, available_end - timedelta(days=180)), available_end),
        min_value=available_start,
        max_value=available_end,
        format="DD/MM/YYYY",
    )
    if isinstance(period, (tuple, list)) and len(period) == 2:
        search_start, search_end = period
    else:
        search_start = search_end = period if isinstance(period, date) else available_end

    ship_df = read_vessel_source("shippivots", selected_vessel, search_start, search_end, manifest_updated)
    voyages = voyage_table(ship_df)
    if voyages.empty:
        st.warning("No VoyageId values were found for this vessel in the selected period.")
        st.stop()

    voyage_labels = {
        row.VoyageId: (
            f"{row.VoyageId} | {row.VoyageStart.strftime('%d/%m/%Y %H:%M')} → "
            f"{row.VoyageEnd.strftime('%d/%m/%Y %H:%M')}"
        )
        for row in voyages.itertuples()
    }
    selected_voyage_id = st.sidebar.selectbox(
        "Voyage",
        voyages["VoyageId"].astype(str).tolist(),
        format_func=lambda x: voyage_labels.get(x, x),
    )
    selected_voyage = voyages[voyages["VoyageId"].astype(str) == str(selected_voyage_id)].iloc[0]
    voyage_start = pd.Timestamp(selected_voyage["VoyageStart"])
    voyage_end = pd.Timestamp(selected_voyage["VoyageEnd"])

    rp_df = read_vessel_source(
        "reportpivots",
        selected_vessel,
        (voyage_start - pd.Timedelta(days=1)).date(),
        (voyage_end + pd.Timedelta(days=1)).date(),
        manifest_updated,
    )
    rd_df = read_vessel_source(
        "reportdata",
        selected_vessel,
        (voyage_start - pd.Timedelta(days=1)).date(),
        (voyage_end + pd.Timedelta(days=1)).date(),
        manifest_updated,
    )
    rp_voyage = filter_to_voyage(rp_df, "DateTime", voyage_start, voyage_end, 3)
    rd_voyage = filter_to_voyage(rd_df, "StartDateTimeGMT", voyage_start, voyage_end, 3)
    ship_voyage = ship_df[ship_df["VoyageId"].astype("string") == str(selected_voyage_id)].copy()
    cargo_by_report = pivot_cargo_reports(rd_voyage)

    # Overview values.
    cargo_weight = non_null_value(rp_voyage, "CargoWeight")
    if pd.isna(cargo_weight):
        cargo_weight = non_null_value(cargo_by_report, "Cargo Weight [tons]")
    cargo_teu = non_null_value(rp_voyage, "CargoTEU")
    departure_port = non_null_value(rp_voyage, "DeparturePort", last=False)
    arrival_port = non_null_value(rp_voyage, "ArrivalPort", last=True)
    draft_fore = non_null_value(rp_voyage, "DraftFore")
    draft_aft = non_null_value(rp_voyage, "DraftAft")
    if pd.isna(draft_fore):
        draft_fore = non_null_value(ship_voyage, "DraftFore")
    if pd.isna(draft_aft):
        draft_aft = non_null_value(ship_voyage, "DraftAft")
    mean_draft = (float(draft_fore) + float(draft_aft)) / 2 if pd.notna(draft_fore) and pd.notna(draft_aft) else pd.NA

    st.markdown(
        f'<div class="hero"><div class="eyebrow">Marorka cargo intelligence</div><h1>Cargo Voyage Dashboard</h1>'
        f'<p>{selected_vessel} | Voyage {selected_voyage_id} | '
        f'{voyage_start.strftime("%d/%m/%Y %H:%M")} → {voyage_end.strftime("%d/%m/%Y %H:%M")} GMT</p></div>',
        unsafe_allow_html=True,
    )
    newest_load = max((str(meta.get("loaded_at_local", "")) for meta in sources.values()), default="-")
    st.markdown(f'<div class="load-pill"><strong>Prepared data:</strong> {newest_load}</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Voyage Overview</div>', unsafe_allow_html=True)
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Cargo Weight", fmt_number(cargo_weight, 1, " MT"))
    c2.metric("Cargo TEU", fmt_number(cargo_teu, 0, " TEU"))
    c3.metric("Departure", "-" if pd.isna(departure_port) else str(departure_port))
    c4.metric("Arrival", "-" if pd.isna(arrival_port) else str(arrival_port))
    c5.metric("Duration", fmt_number(selected_voyage["DurationHours"] / 24, 1, " d"))
    c6.metric("Mean Draft", fmt_number(mean_draft, 2, " m"))

    st.markdown('<div class="section-title">Cargo Operations During Voyage</div>', unsafe_allow_html=True)
    op_values = {field: numeric_sum(cargo_by_report, field) for field in OPERATION_FIELDS}
    a1, a2, a3, a4 = st.columns(4)
    a1.metric("Cargo Loaded", fmt_number(op_values.get("Cargo Weight Added [MT]"), 1, " MT"))
    a2.metric("Cargo Discharged", fmt_number(op_values.get("Cargo Weight Removed [MT]"), 1, " MT"))
    a3.metric("TEU Loaded", fmt_number(op_values.get("TEU Loaded Units"), 0))
    a4.metric("TEU Discharged", fmt_number(op_values.get("TEU Discharged Units"), 0))
    b1, b2, b3, b4 = st.columns(4)
    b1.metric("FEU Loaded", fmt_number(op_values.get("FEU Loaded Units"), 0))
    b2.metric("FEU Discharged", fmt_number(op_values.get("FEU Discharged Units"), 0))
    b3.metric("Reefers Loaded", fmt_number(op_values.get("Reefers Loaded Units"), 0))
    b4.metric("Reefers Discharged", fmt_number(op_values.get("Reefers Discharged Units"), 0))

    st.markdown('<div class="section-title">Voyage / Report Timeline</div>', unsafe_allow_html=True)
    timeline = build_report_timeline(rp_voyage, cargo_by_report)
    timeline_display = timeline.copy()
    for col in ["DateTime", "EndDateTimeGMT"]:
        if col in timeline_display.columns:
            timeline_display[col] = pd.to_datetime(timeline_display[col], errors="coerce", utc=True).dt.strftime("%d/%m/%Y %H:%M")
    st.dataframe(timeline_display, use_container_width=True, hide_index=True, height=410)

    st.markdown('<div class="section-title">Cargo by Report</div>', unsafe_allow_html=True)
    if cargo_by_report.empty:
        st.info("No cargo ReportData rows were mapped to this voyage.")
    else:
        report_options = cargo_by_report["ReportId"].astype(str).tolist()
        report_labels = {}
        for row in cargo_by_report.itertuples(index=False):
            rid = str(getattr(row, "ReportId"))
            rt = str(getattr(row, "ReportType"))
            dt = pd.to_datetime(getattr(row, "StartDateTimeGMT"), errors="coerce", utc=True)
            report_labels[rid] = f"{dt.strftime('%d/%m/%Y %H:%M') if pd.notna(dt) else '-'} | {rt} | Report {rid}"
        selected_report_id = st.selectbox("Selected report", report_options, format_func=lambda x: report_labels.get(x, x))
        report_row = cargo_by_report[cargo_by_report["ReportId"].astype(str) == selected_report_id].iloc[0]

        tab1, tab2, tab3, tab4 = st.tabs(["Cargo Summary", "Containers", "Operations", "Draft & Ballast"])
        with tab1:
            summary_fields = ["Cargo Weight [tons]", "Total Units Weight (All Categories)", "Total Number Full Units (20 and 40ft)", "Total Number Empty Units (20 and 40ft)", "Total Number Reefer Units (20 and 40ft)"]
            summary = pd.DataFrame([{"Field": f, "Value": report_row.get(f, pd.NA)} for f in summary_fields if f in cargo_by_report.columns])
            st.dataframe(summary, use_container_width=True, hide_index=True)
        with tab2:
            fields = [f for f in COMPOSITION_FIELDS if f in cargo_by_report.columns and f != "Cargo Weight [tons]"]
            table = pd.DataFrame([{"Field": f, "Value": report_row.get(f, pd.NA)} for f in fields])
            st.dataframe(table, use_container_width=True, hide_index=True)
        with tab3:
            fields = [f for f in OPERATION_FIELDS if f in cargo_by_report.columns]
            fields += [f for f in ["Port", "Commenced Cargo Operation Time [dd:mm:yyyy hh:mm]", "Completed Cargo Operation Time [dd:mm:yyyy hh:mm]", "Cargo Operations Completed During Port Stay"] if f in cargo_by_report.columns]
            table = pd.DataFrame([{"Field": f, "Value": report_row.get(f, pd.NA)} for f in fields])
            st.dataframe(table, use_container_width=True, hide_index=True)
        with tab4:
            fields = [f for f in ["Draft Forward [m] (m)", "Draft Aft [m] (m)", "Observed Draft Forward [m]", "Observed Draft Aft [m]", "Observed Mean Draft [m]", "Calculated Draft Forward [m]", "Calculated Draft Aft [m]", "Calculated Mean Draft [m]", "Ballast Amount [tons]", "Dead Load [tons]", "Air Draft [m]"] if f in cargo_by_report.columns]
            table = pd.DataFrame([{"Field": f, "Value": report_row.get(f, pd.NA)} for f in fields])
            st.dataframe(table, use_container_width=True, hide_index=True)

    overview_df = pd.DataFrame([{
        "Vessel": selected_vessel,
        "VoyageId": selected_voyage_id,
        "VoyageIdInternal": selected_voyage.get("VoyageIdInternal"),
        "VoyageStart": voyage_start,
        "VoyageEnd": voyage_end,
        "DurationHours": selected_voyage["DurationHours"],
        "DeparturePort": departure_port,
        "ArrivalPort": arrival_port,
        "CargoWeightMT": cargo_weight,
        "CargoTEU": cargo_teu,
        "DraftFore": draft_fore,
        "DraftAft": draft_aft,
        "MeanDraft": mean_draft,
    }])
    export_bytes = to_excel_bytes(overview_df, timeline, cargo_by_report, rp_voyage, ship_voyage)
    st.download_button(
        "Download selected voyage as Excel",
        export_bytes,
        file_name=f"{selected_vessel.replace(' ', '_')}_{selected_voyage_id}_cargo_voyage.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("API / snapshot diagnostics"):
        st.json({
            "schema_version": manifest.get("schema_version"),
            "manifest_updated_at_utc": manifest.get("updated_at_utc"),
            "sources": sources,
            "selected_vessel_ship_rows": len(ship_df),
            "selected_voyage_ship_rows": len(ship_voyage),
            "selected_voyage_reportpivots_rows": len(rp_voyage),
            "selected_voyage_reportdata_rows": len(rd_voyage),
            "selected_voyage_cargo_reports": len(cargo_by_report),
        })


if __name__ == "__main__":
    main()
