from __future__ import annotations
from contextlib import contextmanager
from datetime import date, datetime, timedelta, timezone
from hashlib import sha256
from html import escape
from io import BytesIO
import gc
import hmac
import json
import os
import re
import time
from pathlib import Path
from typing import Any
from urllib.parse import urlencode, urljoin
from zoneinfo import ZoneInfo
import pandas as pd
import pyarrow as pa
import pyarrow.dataset as ds
import pyarrow.parquet as pq
import requests
import streamlit as st
from requests.auth import HTTPBasicAuth, HTTPDigestAuth
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

try:
    import fcntl
except ImportError:  # pragma: no cover - Streamlit Cloud runs on Linux.
    fcntl = None


# =============================================================================
# Configuration
# =============================================================================

APP_TITLE = "AtlasFlow"
APP_DIR = Path(__file__).resolve().parent
SNAPSHOT_DIR = APP_DIR / ".atlasflow_cache"
RAW_SNAPSHOT_FILE = SNAPSHOT_DIR / "reportdata_raw.parquet"
METADATA_SNAPSHOT_FILE = SNAPSHOT_DIR / "reportdata_metadata.json"
ODATA_ENDPOINT = "https://online.marorka.com/Odata/v1/ODataService.svc/ReportData"
REPORTPIVOTS_ENDPOINT = "https://online.marorka.com/Odata/v1/ODataService.svc/ReportPivots"
SHIPPIVOTS_ENDPOINT = "https://online.marorka.com/Odata/v1/ODataService.svc/ShipPivots"

SOURCE_CONFIGS = {
    "reportdata": {
        "label": "ReportData",
        "endpoint": ODATA_ENDPOINT,
        "snapshot_file": SNAPSHOT_DIR / "reportdata_raw.parquet",
        "metadata_file": SNAPSHOT_DIR / "reportdata_metadata.json",
        "datetime_candidates": ["StartDateTimeGMT", "DateTime", "dateTime", "Timestamp"],
    },
    "reportpivots": {
        "label": "ReportPivots",
        "endpoint": REPORTPIVOTS_ENDPOINT,
        "snapshot_file": SNAPSHOT_DIR / "reportpivots_raw.parquet",
        "metadata_file": SNAPSHOT_DIR / "reportpivots_metadata.json",
        "datetime_candidates": ["DateTime", "StartDateTimeGMT", "ReportDateTime", "Timestamp"],
    },
    "shippivots": {
        "label": "ShipPivots",
        "endpoint": SHIPPIVOTS_ENDPOINT,
        "snapshot_file": SNAPSHOT_DIR / "shippivots_raw.parquet",
        "metadata_file": SNAPSHOT_DIR / "shippivots_metadata.json",
        "datetime_candidates": ["DateTime", "StartDateTimeGMT", "Timestamp"],
    },
}
MAX_ODATA_PAGES = 1000
MAX_CONSECUTIVE_EMPTY_ODATA_PAGES = 2
API_CACHE_TTL_SECONDS = 21600  # 6 hours
API_FULL_START_DATE = date(2026, 1, 1)
TABLE_PREVIEW_ROW_LIMIT = 1000
DISPLAY_DATETIME_FORMAT = "%d/%m/%Y %H:%M"


# Persistent, user-ready multi-source snapshot settings.
ATLAS_SNAPSHOT_SCHEMA_VERSION = "2026-07-15-multisource-prepared-incremental-v1"
ATLAS_SNAPSHOT_GENERATIONS_TO_KEEP = 2
ATLAS_PREPARE_VERSION = "atlasflow_dynamic_pivot_v3_oil_stats_prepared_v1"
API_REQUEST_TIMEOUT_SECONDS = 60
API_REQUEST_MAX_ATTEMPTS = 3

DEFAULT_SOURCE_CHUNK_DAYS = {
    "reportdata": 31,
    "reportpivots": 31,
    "shippivots": 7,
}
DEFAULT_SOURCE_OVERLAP_DAYS = {
    "reportdata": 14,
    "reportpivots": 14,
    "shippivots": 10,
}
DEFAULT_SOURCE_FULL_REFRESH_MAX_MINUTES = {
    "reportdata": 240,
    "reportpivots": 240,
    "shippivots": 360,
}
DEFAULT_SOURCE_INCREMENTAL_REFRESH_MAX_MINUTES = {
    "reportdata": 45,
    "reportpivots": 60,
    "shippivots": 90,
}

EXCLUDED_REPORT_TYPES = [
    "Intake Report",
    "Fuel Change Report",
]

# ReportData is intentionally loaded in the same compact mode as the
# Performance KPIs app. Pulling every ValueDescription from ReportData is too
# broad for Streamlit Cloud and makes the source slow/fragile. These aliases
# are the KPI/calculation values that AtlasFlow needs now; we can expand this
# whitelist later in controlled groups.
PERFORMANCE_KPI_VALUE_ALIASES = {
    "Engine Distance [nm]": [
        "Engine Distance [nm]",
    ],
    "Distance Over Ground [nm]": [
        "Distance Over Ground [nm]",
    ],
    "Steaming Time Since Last Report [hh:mm]": [
        "Steaming Time Since Last Report [hh:mm]",
        "Steaming Time Since Last Report",
    ],
    "ME Load [%MCR]": [
        "ME Load [%MCR]",
        "ME Load [% MCR]",
    ],
    "Power from Torque Meter [kW]": [
        "Power from Torque Meter [kW]",
        "Total Shaft Power [kW] (kW)",
        "Total Shaft Power [kW]",
    ],
    "GPS Speed [kn]": [
        "GPS Speed [kn]",
        "GPS Speed",
        "Speed Over Ground [kn]",
        "Speed Over Ground",
    ],
    "Log Speed [kn]": [
        "Log Speed [kn]",
        "Log Speed",
        "Speed Through Water [kn]",
        "Speed Through Water",
    ],
    "Main Engine - HSHFO": ["Main Engine - HSHFO"],
    "Main Engine - HSLFO": ["Main Engine - HSLFO"],
    "Main Engine - MGO": ["Main Engine - MGO"],
    "Main Engine - ULSHFO": ["Main Engine - ULSHFO"],
    "Main Engine - ULSLFO": ["Main Engine - ULSLFO"],
    "Main Engine - VLSHFO": ["Main Engine - VLSHFO"],
    "Main Engine - VLSLFO": ["Main Engine - VLSLFO"],
    "Boiler - HSHFO": ["Boiler - HSHFO"],
    "Boiler - HSLFO": ["Boiler - HSLFO"],
    "Boiler - MGO": ["Boiler - MGO"],
    "Boiler - ULSHFO": ["Boiler - ULSHFO"],
    "Boiler - ULSLFO": ["Boiler - ULSLFO"],
    "Boiler - VLSHFO": ["Boiler - VLSHFO"],
    "Boiler - VLSLFO": ["Boiler - VLSLFO"],

    # Additional bunker/fuel consumption ValueDescriptions for ME/DG/Auxiliary
    # analysis. These are included in the ReportData API whitelist so bunker
    # consumption fields can be selected/exported and later used for derived
    # calculations without broad-loading all ReportData.
    "Main Engine Total Consumed": [
        "Main Engine Total Consumed",
        "ME Total Consumed",
        "Main Engine Consumption",
        "ME Consumption",
        "MEConsumed",
    ],
    "Diesel Generator Total Consumed": [
        "Diesel Generator Total Consumed",
        "DG Total Consumed",
        "DG Totals Consumed",
        "DGTotalsConsumed",
        "DGTotalConsumed",
        "Generator Total Consumed",
    ],
    "Auxiliary Engine Total Consumed": [
        "Auxiliary Engine Total Consumed",
        "Aux Engine Total Consumed",
        "Aux Total Consumed",
        "AuxConsumed",
    ],
    "Total Fuel Consumed": [
        "Total Fuel Consumed",
        "Total Consumed",
        "Total Consumption",
        "Bunker Consumption",
        "Fuel Consumption",
    ],
    "Diesel Generator - HSHFO": ["Diesel Generator - HSHFO", "DG - HSHFO", "Generator - HSHFO"],
    "Diesel Generator - HSLFO": ["Diesel Generator - HSLFO", "DG - HSLFO", "Generator - HSLFO"],
    "Diesel Generator - MGO": ["Diesel Generator - MGO", "DG - MGO", "Generator - MGO", "DG - MGO/MDO"],
    "Diesel Generator - ULSHFO": ["Diesel Generator - ULSHFO", "DG - ULSHFO", "Generator - ULSHFO"],
    "Diesel Generator - ULSLFO": ["Diesel Generator - ULSLFO", "DG - ULSLFO", "Generator - ULSLFO"],
    "Diesel Generator - VLSHFO": ["Diesel Generator - VLSHFO", "DG - VLSHFO", "Generator - VLSHFO"],
    "Diesel Generator - VLSLFO": ["Diesel Generator - VLSLFO", "DG - VLSLFO", "Generator - VLSLFO"],
    "Auxiliary Engine - HSHFO": ["Auxiliary Engine - HSHFO", "Aux Engine - HSHFO", "Aux - HSHFO"],
    "Auxiliary Engine - HSLFO": ["Auxiliary Engine - HSLFO", "Aux Engine - HSLFO", "Aux - HSLFO"],
    "Auxiliary Engine - MGO": ["Auxiliary Engine - MGO", "Aux Engine - MGO", "Aux - MGO", "Aux - MGO/MDO"],
    "Auxiliary Engine - ULSHFO": ["Auxiliary Engine - ULSHFO", "Aux Engine - ULSHFO", "Aux - ULSHFO"],
    "Auxiliary Engine - ULSLFO": ["Auxiliary Engine - ULSLFO", "Aux Engine - ULSLFO", "Aux - ULSLFO"],
    "Auxiliary Engine - VLSHFO": ["Auxiliary Engine - VLSHFO", "Aux Engine - VLSHFO", "Aux - VLSHFO"],
    "Auxiliary Engine - VLSLFO": ["Auxiliary Engine - VLSLFO", "Aux Engine - VLSLFO", "Aux - VLSLFO"],

    # Lub oil ROB/received and DG running hours for consumption aggregation.
    # These are pulled from ReportData so AtlasFlow can expose oil consumption
    # totals as selectable derived variables, not as engineered KPI metrics.
    "MELO ROB [ltr]": ["MELO ROB [ltr]"],
    "MELO Received [ltr]": ["MELO Received [ltr]"],
    "Cylinder Oil 1 ROB [ltr]": ["Cylinder Oil 1 ROB [ltr]"],
    "Cylinder Oil 1 Received [ltr]": ["Cylinder Oil 1 Received [ltr]"],
    "Cylinder Oil 2 ROB [ltr]": ["Cylinder Oil 2 ROB [ltr]"],
    "Cylinder Oil 2 Received [ltr]": ["Cylinder Oil 2 Received [ltr]"],
    "GELO ROB [ltr]": ["GELO ROB [ltr]", "GELO Grade ROB [ltr]"],
    "GELO Received [ltr]": ["GELO Received [ltr]"],
    "DG1 Running Hours [hh:mm]": ["DG1 Running Hours [hh:mm]"],
    "DG2 Running Hours [hh:mm]": ["DG2 Running Hours [hh:mm]"],
    "DG3 Running Hours [hh:mm]": ["DG3 Running Hours [hh:mm]"],
    "DG4 Running Hours [hh:mm]": ["DG4 Running Hours [hh:mm]"],
}

REPORTDATA_VALUE_WHITELIST = sorted(
    {alias for aliases in PERFORMANCE_KPI_VALUE_ALIASES.values() for alias in aliases},
    key=str.casefold,
)
REPORTDATA_VALUE_WHITELIST_KEYS = {
    re.sub(r"[^a-z0-9]+", "", value.lower()) for value in REPORTDATA_VALUE_WHITELIST
}

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

# Derived calculation setup. These columns are calculated from the same Marorka
# source values used in the Performance KPIs app, but are exposed here as normal
# AtlasFlow variables that can be selected, displayed, filtered, summarized, and exported.
DERIVED_VALUE_ALIASES = {
    "Engine Distance [nm]": [
        "Engine Distance [nm]",
    ],
    "Distance Over Ground [nm]": [
        "Distance Over Ground [nm]",
    ],
    "Steaming Time Since Last Report [hh:mm]": [
        "Steaming Time Since Last Report [hh:mm]",
        "Steaming Time Since Last Report",
    ],
    "ME Load [%MCR]": [
        "ME Load [%MCR]",
        "ME Load [% MCR]",
    ],
    "Power from Torque Meter [kW]": [
        "Power from Torque Meter [kW]",
        "Total Shaft Power [kW] (kW)",
        "Total Shaft Power [kW]",
    ],
    "Main Engine - HSHFO": ["Main Engine - HSHFO"],
    "Main Engine - HSLFO": ["Main Engine - HSLFO"],
    "Main Engine - MGO": ["Main Engine - MGO"],
    "Main Engine - ULSHFO": ["Main Engine - ULSHFO"],
    "Main Engine - ULSLFO": ["Main Engine - ULSLFO"],
    "Main Engine - VLSHFO": ["Main Engine - VLSHFO"],
    "Main Engine - VLSLFO": ["Main Engine - VLSLFO"],
    "Boiler - HSHFO": ["Boiler - HSHFO"],
    "Boiler - HSLFO": ["Boiler - HSLFO"],
    "Boiler - MGO": ["Boiler - MGO"],
    "Boiler - ULSHFO": ["Boiler - ULSHFO"],
    "Boiler - ULSLFO": ["Boiler - ULSLFO"],
    "Boiler - VLSHFO": ["Boiler - VLSHFO"],
    "Boiler - VLSLFO": ["Boiler - VLSLFO"],

    # Additional bunker/fuel consumption ValueDescriptions for ME/DG/Auxiliary
    # analysis. These are included in the ReportData API whitelist so bunker
    # consumption fields can be selected/exported and later used for derived
    # calculations without broad-loading all ReportData.
    "Main Engine Total Consumed": [
        "Main Engine Total Consumed",
        "ME Total Consumed",
        "Main Engine Consumption",
        "ME Consumption",
        "MEConsumed",
    ],
    "Diesel Generator Total Consumed": [
        "Diesel Generator Total Consumed",
        "DG Total Consumed",
        "DG Totals Consumed",
        "DGTotalsConsumed",
        "DGTotalConsumed",
        "Generator Total Consumed",
    ],
    "Auxiliary Engine Total Consumed": [
        "Auxiliary Engine Total Consumed",
        "Aux Engine Total Consumed",
        "Aux Total Consumed",
        "AuxConsumed",
    ],
    "Total Fuel Consumed": [
        "Total Fuel Consumed",
        "Total Consumed",
        "Total Consumption",
        "Bunker Consumption",
        "Fuel Consumption",
    ],
    "Diesel Generator - HSHFO": ["Diesel Generator - HSHFO", "DG - HSHFO", "Generator - HSHFO"],
    "Diesel Generator - HSLFO": ["Diesel Generator - HSLFO", "DG - HSLFO", "Generator - HSLFO"],
    "Diesel Generator - MGO": ["Diesel Generator - MGO", "DG - MGO", "Generator - MGO", "DG - MGO/MDO"],
    "Diesel Generator - ULSHFO": ["Diesel Generator - ULSHFO", "DG - ULSHFO", "Generator - ULSHFO"],
    "Diesel Generator - ULSLFO": ["Diesel Generator - ULSLFO", "DG - ULSLFO", "Generator - ULSLFO"],
    "Diesel Generator - VLSHFO": ["Diesel Generator - VLSHFO", "DG - VLSHFO", "Generator - VLSHFO"],
    "Diesel Generator - VLSLFO": ["Diesel Generator - VLSLFO", "DG - VLSLFO", "Generator - VLSLFO"],
    "Auxiliary Engine - HSHFO": ["Auxiliary Engine - HSHFO", "Aux Engine - HSHFO", "Aux - HSHFO"],
    "Auxiliary Engine - HSLFO": ["Auxiliary Engine - HSLFO", "Aux Engine - HSLFO", "Aux - HSLFO"],
    "Auxiliary Engine - MGO": ["Auxiliary Engine - MGO", "Aux Engine - MGO", "Aux - MGO", "Aux - MGO/MDO"],
    "Auxiliary Engine - ULSHFO": ["Auxiliary Engine - ULSHFO", "Aux Engine - ULSHFO", "Aux - ULSHFO"],
    "Auxiliary Engine - ULSLFO": ["Auxiliary Engine - ULSLFO", "Aux Engine - ULSLFO", "Aux - ULSLFO"],
    "Auxiliary Engine - VLSHFO": ["Auxiliary Engine - VLSHFO", "Aux Engine - VLSHFO", "Aux - VLSHFO"],
    "Auxiliary Engine - VLSLFO": ["Auxiliary Engine - VLSLFO", "Aux Engine - VLSLFO", "Aux - VLSLFO"],

    # Lub oil ROB/received and DG running hours for consumption aggregation.
    # These are pulled from ReportData so AtlasFlow can expose oil consumption
    # totals as selectable derived variables, not as engineered KPI metrics.
    "MELO ROB [ltr]": ["MELO ROB [ltr]"],
    "MELO Received [ltr]": ["MELO Received [ltr]"],
    "Cylinder Oil 1 ROB [ltr]": ["Cylinder Oil 1 ROB [ltr]"],
    "Cylinder Oil 1 Received [ltr]": ["Cylinder Oil 1 Received [ltr]"],
    "Cylinder Oil 2 ROB [ltr]": ["Cylinder Oil 2 ROB [ltr]"],
    "Cylinder Oil 2 Received [ltr]": ["Cylinder Oil 2 Received [ltr]"],
    "GELO ROB [ltr]": ["GELO ROB [ltr]", "GELO Grade ROB [ltr]"],
    "GELO Received [ltr]": ["GELO Received [ltr]"],
    "DG1 Running Hours [hh:mm]": ["DG1 Running Hours [hh:mm]"],
    "DG2 Running Hours [hh:mm]": ["DG2 Running Hours [hh:mm]"],
    "DG3 Running Hours [hh:mm]": ["DG3 Running Hours [hh:mm]"],
    "DG4 Running Hours [hh:mm]": ["DG4 Running Hours [hh:mm]"],
}

ME_FUEL_COLUMNS = [
    "Main Engine - HSHFO",
    "Main Engine - HSLFO",
    "Main Engine - MGO",
    "Main Engine - ULSHFO",
    "Main Engine - ULSLFO",
    "Main Engine - VLSHFO",
    "Main Engine - VLSLFO",
]

BOILER_FUEL_COLUMNS = [
    "Boiler - HSHFO",
    "Boiler - HSLFO",
    "Boiler - MGO",
    "Boiler - ULSHFO",
    "Boiler - ULSLFO",
    "Boiler - VLSHFO",
    "Boiler - VLSLFO",
]

DG_FUEL_COLUMNS = [
    "Diesel Generator - HSHFO",
    "Diesel Generator - HSLFO",
    "Diesel Generator - MGO",
    "Diesel Generator - ULSHFO",
    "Diesel Generator - ULSLFO",
    "Diesel Generator - VLSHFO",
    "Diesel Generator - VLSLFO",
]

AUXILIARY_FUEL_COLUMNS = [
    "Auxiliary Engine - HSHFO",
    "Auxiliary Engine - HSLFO",
    "Auxiliary Engine - MGO",
    "Auxiliary Engine - ULSHFO",
    "Auxiliary Engine - ULSLFO",
    "Auxiliary Engine - VLSHFO",
    "Auxiliary Engine - VLSLFO",
]

DERIVED_VARIABLES = [
    "Calculated Slip",
    "ME Consumption Total",
    "DG Consumption Total",
    "Auxiliary Engine Consumption Total",
    "Boiler Sum",
    "Total Fuel Consumption",
    "Consumption ME 24 Hours [MT]",
    "SFOC [gr/Kwh]",
    "MELO Consumption Total [ltr]",
    "CYLO Consumption Total [ltr]",
    "GELO Consumption Total [ltr]",
    "Total DG Running Hours [hh:mm]",
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
            --atlas-topbar-h: 67px;
            --atlas-sidebar-w: 324px;
            --atlas-ink: #0B1F33;
            --atlas-muted: #24364F;
            --atlas-soft: #64748B;
            --atlas-teal: #006B68;
            --atlas-teal-bright: #0AAEA6;
            --atlas-line: #D9E6E5;
            --atlas-bg: #FAFCFC;
            --atlas-chip: #DDF4F2;
        }

        html,
        body,
        .stApp {
            background: var(--atlas-bg) !important;
            color: var(--atlas-ink) !important;
            font-family: "Segoe UI", "Inter", "Aptos", Arial, sans-serif !important;
        }

        .stApp {
            background:
                linear-gradient(90deg, rgba(0, 107, 104, 0.035), transparent 24rem),
                linear-gradient(180deg, #FFFFFF 0%, #FBFDFD 46%, #F4FAF9 100%) !important;
        }

        header[data-testid="stHeader"] {
            left: 0 !important;
            right: 0 !important;
            width: 100vw !important;
            height: var(--atlas-topbar-h) !important;
            background: rgba(255, 255, 255, 0.98) !important;
            border-bottom: 1px solid rgba(15, 23, 42, 0.10) !important;
            box-shadow: 0 1px 12px rgba(15, 23, 42, 0.04) !important;
            z-index: 999990 !important;
        }

        header[data-testid="stHeader"] > div {
            height: var(--atlas-topbar-h) !important;
            background: transparent !important;
        }

        div[data-testid="stToolbar"] {
            top: 0.55rem !important;
            right: 1.55rem !important;
            z-index: 999995 !important;
        }

        div[data-testid="stDecoration"] {
            display: none !important;
        }

        .atlas-topbar-brand {
            position: fixed;
            top: 0;
            left: 0;
            height: var(--atlas-topbar-h);
            display: flex;
            align-items: center;
            gap: 0.95rem;
            padding-left: 1.35rem;
            z-index: 999996;
            pointer-events: none;
        }

        .atlas-menu-lines {
            width: 24px;
            height: 24px;
            position: relative;
            flex: 0 0 24px;
        }

        .atlas-menu-lines::before,
        .atlas-menu-lines::after,
        .atlas-menu-lines span {
            content: "";
            position: absolute;
            left: 2px;
            width: 18px;
            height: 2px;
            border-radius: 999px;
            background: var(--atlas-teal);
        }

        .atlas-menu-lines::before { top: 6px; }
        .atlas-menu-lines span { top: 11px; }
        .atlas-menu-lines::after { top: 16px; }

        .atlas-logo-mark {
            width: 34px;
            height: 34px;
            position: relative;
            flex: 0 0 34px;
        }

        .atlas-logo-mark::before,
        .atlas-logo-mark::after {
            content: "";
            position: absolute;
            border-radius: 18px 18px 18px 6px;
            transform: rotate(34deg);
            box-shadow: 0 4px 12px rgba(0, 107, 104, 0.18);
        }

        .atlas-logo-mark::before {
            width: 20px;
            height: 31px;
            left: 4px;
            top: 2px;
            background: linear-gradient(145deg, #013F43 0%, #008C86 62%, #19BFB5 100%);
        }

        .atlas-logo-mark::after {
            width: 18px;
            height: 24px;
            left: 15px;
            top: 10px;
            background: linear-gradient(145deg, #0FB5AD 0%, #006B68 100%);
            opacity: 0.94;
        }

        .atlas-brand-word {
            color: #07515A;
            font-size: 1.85rem;
            font-weight: 400;
            line-height: 1;
            letter-spacing: 0;
        }

        @media (min-width: 769px) {
            section[data-testid="stSidebar"] {
                width: var(--atlas-sidebar-w) !important;
                min-width: var(--atlas-sidebar-w) !important;
                top: var(--atlas-topbar-h) !important;
                height: calc(100vh - var(--atlas-topbar-h)) !important;
                background:
                    radial-gradient(circle at 18px 8px, rgba(35, 209, 199, 0.18), transparent 16rem),
                    linear-gradient(180deg, #006A66 0%, #004743 46%, #003C39 100%) !important;
                border-right: 1px solid rgba(255, 255, 255, 0.16) !important;
                box-shadow: none !important;
            }

            section[data-testid="stSidebar"] > div {
                height: calc(100vh - var(--atlas-topbar-h)) !important;
                padding: 1.25rem 1.25rem 6rem 1.25rem !important;
            }

            .block-container {
                max-width: none !important;
                padding: 3.55rem 1.9rem 3rem 2.35rem !important;
            }
        }

        @media (max-width: 768px) {
            .atlas-brand-word { font-size: 1.35rem; }
            .atlas-logo-mark { width: 28px; height: 28px; }
            .block-container { padding: 5.2rem 1rem 2rem 1rem !important; }
        }

        section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
            gap: 0.72rem !important;
        }

        section[data-testid="stSidebar"] h1,
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3,
        section[data-testid="stSidebar"] p,
        section[data-testid="stSidebar"] span,
        section[data-testid="stSidebar"] label,
        section[data-testid="stSidebar"] label *,
        section[data-testid="stSidebar"] [data-testid="stWidgetLabel"],
        section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] *,
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] *,
        section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] * {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            font-weight: 760 !important;
            letter-spacing: 0 !important;
        }

        section[data-testid="stSidebar"] h3 {
            font-size: 1.05rem !important;
            margin-top: 0.25rem !important;
            margin-bottom: 0.1rem !important;
        }

        section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] *,
        section[data-testid="stSidebar"] small {
            color: #D7FFFA !important;
            -webkit-text-fill-color: #D7FFFA !important;
        }

        section[data-testid="stSidebar"] div[data-baseweb="select"] > div,
        section[data-testid="stSidebar"] div[data-baseweb="input"],
        section[data-testid="stSidebar"] [data-testid="stMultiSelect"] div[data-baseweb="select"] > div,
        section[data-testid="stSidebar"] [data-testid="stSelectbox"] div[data-baseweb="select"] > div {
            min-height: 44px !important;
            border-radius: 9px !important;
            border: 1px solid rgba(8, 72, 70, 0.12) !important;
            background: #FFFFFF !important;
            box-shadow: none !important;
        }

        section[data-testid="stSidebar"] div[data-baseweb="select"] > div *,
        section[data-testid="stSidebar"] div[data-baseweb="input"] *,
        section[data-testid="stSidebar"] input {
            color: var(--atlas-ink) !important;
            -webkit-text-fill-color: var(--atlas-ink) !important;
        }

        section[data-testid="stSidebar"] div[data-testid="stButton"] button,
        section[data-testid="stSidebar"] div[data-testid="stButton"] button *,
        section[data-testid="stSidebar"] .stButton button,
        section[data-testid="stSidebar"] .stButton button * {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }

        section[data-testid="stSidebar"] div[data-baseweb="select"] svg,
        section[data-testid="stSidebar"] div[data-baseweb="input"] svg {
            color: #334155 !important;
            fill: #334155 !important;
        }

        section[data-testid="stSidebar"] [data-baseweb="tag"] {
            background: var(--atlas-chip) !important;
            border: 1px solid rgba(0, 107, 104, 0.22) !important;
            border-radius: 999px !important;
            box-shadow: none !important;
        }

        section[data-testid="stSidebar"] [data-baseweb="tag"] *,
        section[data-testid="stSidebar"] [data-baseweb="tag"] span,
        section[data-testid="stSidebar"] [data-baseweb="tag"] svg {
            color: #12313E !important;
            -webkit-text-fill-color: #12313E !important;
            fill: #12313E !important;
        }

        section[data-testid="stSidebar"] [data-testid="stExpander"] {
            border-radius: 8px !important;
            border: 1px solid rgba(255, 255, 255, 0.72) !important;
            background: rgba(255, 255, 255, 0.07) !important;
            box-shadow: none !important;
        }

        section[data-testid="stSidebar"] [data-testid="stExpander"] summary,
        section[data-testid="stSidebar"] [data-testid="stExpander"] summary * {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            font-weight: 760 !important;
        }



        /* Sidebar confirmation/warning cards need dark text on the light card.
           The general sidebar rule forces labels/text to white, so alerts must
           be overridden explicitly for readability. */
        section[data-testid="stSidebar"] div[data-testid="stAlert"],
        section[data-testid="stSidebar"] div[data-testid="stAlert"] > div,
        section[data-testid="stSidebar"] div[data-testid="stAlert"] [role="alert"] {
            background: rgba(255, 255, 255, 0.96) !important;
            border: 1px solid rgba(0, 107, 104, 0.28) !important;
            border-radius: 9px !important;
            box-shadow: 0 2px 10px rgba(15, 23, 42, 0.10) !important;
        }

        section[data-testid="stSidebar"] div[data-testid="stAlert"] *,
        section[data-testid="stSidebar"] div[data-testid="stAlert"] p,
        section[data-testid="stSidebar"] div[data-testid="stAlert"] span,
        section[data-testid="stSidebar"] div[data-testid="stAlert"] div {
            color: var(--atlas-ink) !important;
            -webkit-text-fill-color: var(--atlas-ink) !important;
            font-weight: 650 !important;
        }

        section[data-testid="stSidebar"] div[data-testid="stAlert"] svg {
            color: var(--atlas-teal) !important;
            fill: var(--atlas-teal) !important;
        }

        section[data-testid="stSidebar"] [data-testid="stTooltipHoverTarget"],
        section[data-testid="stSidebar"] [data-testid="stTooltipIcon"],
        section[data-testid="stSidebar"] button[aria-label="Help"],
        section[data-testid="stSidebar"] button[title="Help"] {
            position: relative !important;
            width: 18px !important;
            height: 18px !important;
            min-width: 18px !important;
            min-height: 18px !important;
            border-radius: 999px !important;
            background: #063F3C !important;
            border: 1.5px solid #BDF7EF !important;
            box-shadow: 0 0 0 2px rgba(255, 255, 255, 0.08) !important;
            opacity: 1 !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            overflow: hidden !important;
        }

        section[data-testid="stSidebar"] [data-testid="stTooltipHoverTarget"] svg,
        section[data-testid="stSidebar"] [data-testid="stTooltipIcon"] svg,
        section[data-testid="stSidebar"] button[aria-label="Help"] svg,
        section[data-testid="stSidebar"] button[title="Help"] svg {
            display: none !important;
        }

        section[data-testid="stSidebar"] [data-testid="stTooltipHoverTarget"]::after,
        section[data-testid="stSidebar"] [data-testid="stTooltipIcon"]::after,
        section[data-testid="stSidebar"] button[aria-label="Help"]::after,
        section[data-testid="stSidebar"] button[title="Help"]::after {
            content: "?" !important;
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            font: 800 12px/1 Arial, sans-serif !important;
        }

        section[data-testid="stSidebar"] [data-testid="stTooltipHoverTarget"]:hover,
        section[data-testid="stSidebar"] [data-testid="stTooltipIcon"]:hover,
        section[data-testid="stSidebar"] button[aria-label="Help"]:hover,
        section[data-testid="stSidebar"] button[title="Help"]:hover {
            background: var(--atlas-teal-bright) !important;
            border-color: #FFFFFF !important;
        }

        section[data-testid="stSidebar"] [data-testid="stTooltipIcon"] [data-testid="stTooltipHoverTarget"] {
            background: transparent !important;
            border: 0 !important;
            box-shadow: none !important;
            overflow: visible !important;
        }

        section[data-testid="stSidebar"] [data-testid="stTooltipIcon"] [data-testid="stTooltipHoverTarget"]::after,
        section[data-testid="stSidebar"] [data-testid="stTooltipIcon"] button[aria-label^="Help"]::after {
            content: none !important;
        }

        section[data-testid="stSidebar"] [data-testid="stTooltipIcon"] button[aria-label^="Help"] {
            background: transparent !important;
            border: 0 !important;
            box-shadow: none !important;
            padding: 0 !important;
            width: 18px !important;
            height: 18px !important;
        }

        section[data-testid="stSidebar"] [data-testid="stTooltipIcon"] button[aria-label^="Help"] svg {
            display: block !important;
            width: 13px !important;
            height: 13px !important;
            color: #FFFFFF !important;
            stroke: #FFFFFF !important;
            fill: none !important;
        }

        section[data-testid="stSidebar"] [data-testid="stTooltipIcon"]::after {
            content: none !important;
        }

        section[data-testid="stSidebar"] [data-testid="stTooltipIcon"] [data-testid="stTooltipHoverTarget"] {
            position: absolute !important;
            inset: 0 !important;
            width: 100% !important;
            height: 100% !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
        }

        section[data-testid="stSidebar"] [data-testid="stTooltipIcon"] button[aria-label^="Help"] {
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            width: 100% !important;
            height: 100% !important;
        }

        .dashboard-hero {
            box-sizing: border-box !important;
            min-height: 164px !important;
            margin: 0 0 0.82rem 0 !important;
            padding: 1.45rem 1.65rem 0.1rem 1.65rem !important;
            border-radius: 13px !important;
            border: 1px solid rgba(15, 23, 42, 0.11) !important;
            background: rgba(255, 255, 255, 0.93) !important;
            box-shadow: 0 2px 12px rgba(15, 23, 42, 0.12) !important;
            backdrop-filter: blur(4px) !important;
        }

        .eyebrow {
            color: var(--atlas-teal) !important;
            -webkit-text-fill-color: var(--atlas-teal) !important;
            font-size: 0.75rem !important;
            font-weight: 750 !important;
            letter-spacing: 0.14em !important;
            text-transform: uppercase !important;
            margin-bottom: 0.82rem !important;
        }

        .dashboard-title {
            color: var(--atlas-ink) !important;
            -webkit-text-fill-color: var(--atlas-ink) !important;
            font-family: "Segoe UI", "Inter", "Aptos Display", sans-serif !important;
            font-size: clamp(2.9rem, 3.6vw, 3.7rem) !important;
            font-weight: 400 !important;
            letter-spacing: 0 !important;
            line-height: 1.04 !important;
            margin: 0 !important;
            padding: 0 !important;
        }

        .dashboard-subtitle {
            color: var(--atlas-muted) !important;
            -webkit-text-fill-color: var(--atlas-muted) !important;
            font-size: 0.96rem !important;
            font-weight: 500 !important;
            margin-top: 0.75rem !important;
        }

        .api-load-caption,
        .atlas-pill {
            display: inline-flex !important;
            align-items: center !important;
            gap: 0.52rem !important;
            min-height: 37px !important;
            margin: 0 0 1.15rem 0 !important;
            padding: 0.4rem 0.9rem !important;
            border-radius: 999px !important;
            border: 1px solid rgba(15, 23, 42, 0.12) !important;
            background: rgba(255, 255, 255, 0.96) !important;
            box-shadow: 0 1px 8px rgba(15, 23, 42, 0.06) !important;
            color: var(--atlas-muted) !important;
            -webkit-text-fill-color: var(--atlas-muted) !important;
            font-size: 0.82rem !important;
            font-weight: 700 !important;
        }

        .api-load-caption strong,
        .api-load-caption span,
        .atlas-pill span {
            color: var(--atlas-ink) !important;
            -webkit-text-fill-color: var(--atlas-ink) !important;
            font-weight: 800 !important;
        }

        .api-load-clock {
            width: 17px;
            height: 17px;
            color: var(--atlas-teal);
            -webkit-text-fill-color: var(--atlas-teal);
            display: inline-flex;
            align-items: center;
            justify-content: center;
        }

        .api-load-clock svg {
            width: 17px;
            height: 17px;
            stroke: currentColor;
            stroke-width: 2.1;
            fill: none;
            stroke-linecap: round;
            stroke-linejoin: round;
        }

        div[data-testid="stTabs"] [data-baseweb="tab-list"] {
            gap: 1.35rem !important;
            border-bottom: 1px solid var(--atlas-line) !important;
        }

        button[data-baseweb="tab"] {
            height: 42px !important;
            padding: 0 0.45rem !important;
            color: #334155 !important;
            -webkit-text-fill-color: #334155 !important;
            font-size: 0.94rem !important;
            font-weight: 500 !important;
            letter-spacing: 0 !important;
        }

        button[data-baseweb="tab"][aria-selected="true"] {
            color: var(--atlas-teal) !important;
            -webkit-text-fill-color: var(--atlas-teal) !important;
            font-weight: 700 !important;
        }

        div[data-baseweb="tab-highlight"] {
            background-color: var(--atlas-teal) !important;
            height: 2px !important;
        }

        .section-title {
            color: var(--atlas-ink) !important;
            -webkit-text-fill-color: var(--atlas-ink) !important;
            font-size: 1.58rem !important;
            line-height: 1.22 !important;
            font-weight: 400 !important;
            letter-spacing: 0 !important;
            margin: 0.55rem 0 0.45rem 0 !important;
        }

        .stApp [data-testid="stMarkdownContainer"] p,
        .stApp [data-testid="stCaptionContainer"] p {
            color: var(--atlas-muted) !important;
            -webkit-text-fill-color: var(--atlas-muted) !important;
        }

        .stApp section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] *,
        .stApp section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        .stApp section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }

        .atlas-tabbar {
            display: flex;
            flex-wrap: wrap;
            align-items: flex-end;
            gap: 1.35rem;
            border-bottom: 1px solid var(--atlas-line);
            padding: 0 0 0.05rem 0;
            margin: 0.35rem 0 1rem 0;
        }

        .atlas-tabbar.compact {
            gap: 1.65rem;
            margin-top: 0.2rem;
            margin-bottom: 1.1rem;
        }

        .atlas-tablink {
            position: relative;
            display: inline-flex;
            align-items: center;
            min-height: 42px;
            padding: 0 0.45rem 0.58rem 0.45rem;
            color: #334155 !important;
            -webkit-text-fill-color: #334155 !important;
            font-size: 0.94rem !important;
            font-weight: 500 !important;
            line-height: 1.15 !important;
            text-decoration: none !important;
            border: 0 !important;
            background: transparent !important;
        }

        .atlas-tablink:hover {
            color: var(--atlas-teal) !important;
            -webkit-text-fill-color: var(--atlas-teal) !important;
            text-decoration: none !important;
        }

        .atlas-tablink.active {
            color: var(--atlas-teal) !important;
            -webkit-text-fill-color: var(--atlas-teal) !important;
            font-weight: 700 !important;
        }

        .atlas-tablink.active::after {
            content: "";
            position: absolute;
            left: 0.25rem;
            right: 0.25rem;
            bottom: -0.05rem;
            height: 2px;
            border-radius: 999px;
            background: var(--atlas-teal);
        }


        /* Native Streamlit radio widgets used as text-only tab bars.
           This avoids HTML links, so tab clicks never open a browser tab. */
        div[data-testid="stRadio"] > div[role="radiogroup"] {
            display: flex !important;
            flex-direction: row !important;
            flex-wrap: wrap !important;
            align-items: flex-end !important;
            gap: 1.35rem !important;
            border-bottom: 1px solid var(--atlas-line) !important;
            padding: 0 0 0.05rem 0 !important;
            margin: 0.35rem 0 1rem 0 !important;
        }

        div[data-testid="stRadio"] > div[role="radiogroup"] label {
            position: relative !important;
            min-height: 42px !important;
            padding: 0 0.45rem 0.58rem 0.45rem !important;
            margin: 0 !important;
            background: transparent !important;
            border: 0 !important;
            box-shadow: none !important;
            cursor: pointer !important;
            display: inline-flex !important;
            align-items: center !important;
        }

        div[data-testid="stRadio"] > div[role="radiogroup"] label div[data-baseweb="radio"],
        div[data-testid="stRadio"] > div[role="radiogroup"] label > div:has(input[type="radio"]),
        div[data-testid="stRadio"] > div[role="radiogroup"] label input[type="radio"] {
            display: none !important;
        }

        div[data-testid="stRadio"] > div[role="radiogroup"] label p,
        div[data-testid="stRadio"] > div[role="radiogroup"] label span,
        div[data-testid="stRadio"] > div[role="radiogroup"] label [data-testid="stMarkdownContainer"] * {
            color: #334155 !important;
            -webkit-text-fill-color: #334155 !important;
            font-size: 0.94rem !important;
            font-weight: 500 !important;
            line-height: 1.15 !important;
            text-decoration: none !important;
        }

        div[data-testid="stRadio"] > div[role="radiogroup"] label:hover p,
        div[data-testid="stRadio"] > div[role="radiogroup"] label:hover span,
        div[data-testid="stRadio"] > div[role="radiogroup"] label:hover [data-testid="stMarkdownContainer"] * {
            color: var(--atlas-teal) !important;
            -webkit-text-fill-color: var(--atlas-teal) !important;
        }

        div[data-testid="stRadio"] > div[role="radiogroup"] label:has(input[type="radio"]:checked) p,
        div[data-testid="stRadio"] > div[role="radiogroup"] label:has(input[type="radio"]:checked) span,
        div[data-testid="stRadio"] > div[role="radiogroup"] label:has(input[type="radio"]:checked) [data-testid="stMarkdownContainer"] * {
            color: var(--atlas-teal) !important;
            -webkit-text-fill-color: var(--atlas-teal) !important;
            font-weight: 700 !important;
        }

        div[data-testid="stRadio"] > div[role="radiogroup"] label:has(input[type="radio"]:checked)::after {
            content: "";
            position: absolute;
            left: 0.25rem;
            right: 0.25rem;
            bottom: -0.05rem;
            height: 2px;
            border-radius: 999px;
            background: var(--atlas-teal);
        }

        .atlas-metric-grid {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 1.35rem;
            margin: 0 0 0.4rem 0;
        }

        .atlas-metric-card {
            display: grid;
            grid-template-columns: 58px minmax(0, 1fr);
            align-items: center;
            gap: 1rem;
            min-height: 95px;
            padding: 1rem 1.05rem;
            background: #FFFFFF;
            border: 1px solid rgba(15, 23, 42, 0.11);
            border-radius: 9px;
            box-shadow: 0 3px 13px rgba(15, 23, 42, 0.08);
        }

        .atlas-metric-icon {
            width: 55px;
            height: 55px;
            border-radius: 8px;
            background: linear-gradient(135deg, #006F6A 0%, var(--atlas-teal-bright) 100%);
            color: #FFFFFF;
            display: inline-flex;
            align-items: center;
            justify-content: center;
        }

        .atlas-metric-icon svg {
            width: 32px;
            height: 32px;
            stroke: #FFFFFF;
            stroke-width: 2.2;
            fill: none;
            stroke-linecap: round;
            stroke-linejoin: round;
        }

        .atlas-metric-label {
            color: var(--atlas-muted);
            font-size: 0.88rem;
            line-height: 1.2;
            font-weight: 600;
            margin-bottom: 0.35rem;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .atlas-metric-value {
            color: var(--atlas-ink);
            font-size: 1.85rem;
            line-height: 1;
            font-weight: 400;
            letter-spacing: 0;
        }

        .atlas-table-frame,
        div[data-testid="stDataFrame"] {
            border: 1px solid rgba(15, 23, 42, 0.10) !important;
            border-radius: 7px !important;
            background: #FFFFFF !important;
            overflow: hidden;
            box-shadow: 0 2px 10px rgba(15, 23, 42, 0.08) !important;
        }

        .atlas-table-frame { margin-top: 0.4rem; }
        .atlas-table-scroll { width: 100%; overflow: auto; max-height: 560px; }

        .atlas-preview-table {
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            font-size: 0.88rem;
            color: var(--atlas-ink);
        }

        .atlas-preview-table thead th {
            position: sticky;
            top: 0;
            z-index: 1;
            text-align: left;
            padding: 0.62rem 1rem;
            color: #FFFFFF;
            background: linear-gradient(180deg, #006F6A 0%, #005D59 100%);
            border-right: 1px solid rgba(255, 255, 255, 0.20);
            font-weight: 750;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .atlas-preview-table tbody td {
            padding: 0.48rem 1rem;
            border-top: 1px solid #E7EDEF;
            border-right: 1px solid #E7EDEF;
            background: #FFFFFF;
            color: #12233D;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .atlas-preview-table tbody tr:nth-child(even) td { background: #FCFDFD; }
        .atlas-preview-table th:last-child,
        .atlas-preview-table td:last-child { border-right: 0; }

        .atlas-table-empty {
            margin-top: 1rem;
            padding: 1.1rem 1.2rem;
            border: 1px solid rgba(15, 23, 42, 0.10);
            border-radius: 9px;
            background: #FFFFFF;
            color: var(--atlas-muted);
            box-shadow: 0 2px 10px rgba(15, 23, 42, 0.06);
        }

        div[data-testid="stMetric"] {
            background: #FFFFFF !important;
            border: 1px solid rgba(15, 23, 42, 0.10) !important;
            border-radius: 9px !important;
            box-shadow: 0 2px 10px rgba(15, 23, 42, 0.08) !important;
        }

        [data-baseweb="popover"],
        [data-baseweb="menu"],
        [role="listbox"] {
            background: #FFFFFF !important;
            color: var(--atlas-ink) !important;
        }

        [role="option"],
        [role="option"] *,
        [data-baseweb="menu"] * {
            color: var(--atlas-ink) !important;
            -webkit-text-fill-color: var(--atlas-ink) !important;
        }

        div[data-testid="stAlert"],
        div[data-testid="stAlert"] > div,
        div[data-testid="stAlert"] [role="alert"] {
            background: rgba(255, 255, 255, 0.92) !important;
            border: 1px solid rgba(15, 118, 110, 0.18) !important;
            color: var(--atlas-ink) !important;
            border-radius: 9px !important;
            box-shadow: none !important;
        }

        /* Final high-specificity sidebar alert override: Streamlit's markdown
           container inside the alert was still inheriting the global sidebar
           white-text rule. Keep the alert card light and force readable text. */
        section[data-testid="stSidebar"] div[data-testid="stAlert"] [data-testid="stMarkdownContainer"],
        section[data-testid="stSidebar"] div[data-testid="stAlert"] [data-testid="stMarkdownContainer"] *,
        section[data-testid="stSidebar"] div[data-testid="stAlert"] [data-testid="stAlertContent"],
        section[data-testid="stSidebar"] div[data-testid="stAlert"] [data-testid="stAlertContent"] *,
        section[data-testid="stSidebar"] div[data-testid="stAlert"] [role="alert"],
        section[data-testid="stSidebar"] div[data-testid="stAlert"] [role="alert"] *,
        section[data-testid="stSidebar"] div[data-testid="stAlert"] p,
        section[data-testid="stSidebar"] div[data-testid="stAlert"] span {
            color: #0B1F33 !important;
            -webkit-text-fill-color: #0B1F33 !important;
            opacity: 1 !important;
            font-weight: 650 !important;
            text-shadow: none !important;
        }

        section[data-testid="stSidebar"] div[data-testid="stAlert"] {
            background: #FFFFFF !important;
        }

        .stButton button,
        .stDownloadButton button,
        div[data-testid="stButton"] button,
        div[data-testid="stDownloadButton"] button,
        button[kind="primary"],
        button[kind="secondary"] {
            border-radius: 8px !important;
            background: linear-gradient(135deg, #006F6A, var(--atlas-teal-bright)) !important;
            border: 1px solid rgba(0, 107, 104, 0.32) !important;
            box-shadow: none !important;
        }

        .stButton button,
        .stButton button *,
        .stDownloadButton button,
        .stDownloadButton button *,
        div[data-testid="stButton"] button,
        div[data-testid="stButton"] button *,
        div[data-testid="stDownloadButton"] button,
        div[data-testid="stDownloadButton"] button *,
        button[kind="primary"],
        button[kind="primary"] *,
        button[kind="secondary"],
        button[kind="secondary"] * {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            font-weight: 600 !important;
            opacity: 1 !important;
            text-shadow: none !important;
        }

        .stButton button:hover,
        .stDownloadButton button:hover,
        div[data-testid="stButton"] button:hover,
        div[data-testid="stDownloadButton"] button:hover,
        button[kind="primary"]:hover,
        button[kind="secondary"]:hover {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            filter: brightness(1.04);
        }

        @media (max-width: 1100px) {
            .atlas-metric-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        }

        @media (max-width: 640px) {
            .atlas-metric-grid { grid-template-columns: 1fr; }
            .dashboard-title { font-size: 2.55rem !important; }
            .dashboard-hero { min-height: 150px !important; }
        }
        </style>
        <div class="atlas-topbar-brand" aria-hidden="true">
            <div class="atlas-menu-lines"><span></span></div>
            <div class="atlas-logo-mark"></div>
            <div class="atlas-brand-word">Atlas Flow</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_header(selected_group: str, selected_vessels: list[str], selected_variables: list[str]) -> None:
    vessel_text = "All selected vessels" if len(selected_vessels) != 1 else selected_vessels[0]
    variable_text = "No variables selected" if not selected_variables else f"{len(selected_variables):,} selected variables"
    st.markdown(
        f"""
        <div class="dashboard-hero">
            <div class="eyebrow">Marorka API Explorer</div>
            <h1 class="dashboard-title">Atlas Flow</h1>
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
            <span class="api-load-clock" aria-hidden="true">
                <svg viewBox="0 0 24 24">
                    <circle cx="12" cy="12" r="9"></circle>
                    <path d="M12 7v5l3 2"></path>
                </svg>
            </span>
            <strong>Last API load:</strong> <span>{escape(last_load_display)} LT</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def slugify_tab_label(label: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", str(label).lower()).strip("-") or "tab"


def current_query_params_dict() -> dict[str, str]:
    try:
        items = st.query_params.to_dict()
    except Exception:
        try:
            raw_items = st.experimental_get_query_params()
            items = {key: str(value[0]) if isinstance(value, list) and value else str(value) for key, value in raw_items.items()}
        except Exception:
            items = {}
    return {str(key): str(value) for key, value in items.items()}


def get_tab_selection(param_name: str, options: list[str], default: str) -> str:
    slug_to_option = {slugify_tab_label(option): option for option in options}
    raw_value = get_query_param(param_name, slugify_tab_label(default)).strip().lower()
    if raw_value in slug_to_option:
        return slug_to_option[raw_value]
    if raw_value in options:
        return raw_value
    return default if default in options else options[0]


def render_text_tab_bar(
    options: list[str],
    selected: str,
    *,
    param_name: str,
    css_class: str = "",
    reset_params: list[str] | None = None,
) -> str:
    """Render a native Streamlit text-tab selector.

    Older AtlasFlow batches used HTML anchor links for the tab strip. Those
    looked correct, but browsers treated them as links and sometimes opened the
    app in a new tab. This version uses st.radio under the hood, styled as
    text-only tabs, so clicking a tab only updates Streamlit session state.
    """
    if not options:
        return selected

    key_suffix = slugify_tab_label(css_class) if css_class else "main"
    state_key = f"atlas_tab_{param_name}_{key_suffix}"
    if selected not in options:
        selected = options[0]
    if st.session_state.get(state_key) not in options:
        st.session_state[state_key] = selected

    previous_value = st.session_state.get(state_key, selected)
    choice = st.radio(
        " ",
        options=options,
        horizontal=True,
        label_visibility="collapsed",
        key=state_key,
    )

    if choice != previous_value:
        for reset_param in reset_params or []:
            reset_key = f"atlas_tab_{reset_param}_compact"
            st.session_state.pop(reset_key, None)
            if reset_param == "preview":
                st.session_state["atlas_reportdata_preview_mode"] = "Clean Dataset"
        st.session_state[state_key] = choice

    return choice


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
            <h1 class="dashboard-title">Atlas Flow</h1>
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




RETRYABLE_HTTP_STATUSES = {500, 502, 503, 504}
RETRYABLE_REQUEST_EXCEPTIONS = (
    requests.ConnectionError,
    requests.ReadTimeout,
    requests.Timeout,
    requests.exceptions.ChunkedEncodingError,
)


def request_with_retry(
    session: requests.Session,
    url: str,
    *,
    auth: Any,
    timeout: int = 90,
    max_attempts: int = 5,
    base_sleep_seconds: float = 2.0,
) -> requests.Response:
    """GET an OData page with retry/backoff for transient Marorka disconnects."""
    last_error: Exception | None = None

    for attempt in range(1, max_attempts + 1):
        try:
            response = session.get(url, auth=auth, timeout=timeout)
            if response.status_code in RETRYABLE_HTTP_STATUSES and attempt < max_attempts:
                time.sleep(base_sleep_seconds * (2 ** (attempt - 1)))
                continue
            return response
        except RETRYABLE_REQUEST_EXCEPTIONS as exc:
            last_error = exc
            if attempt >= max_attempts:
                raise
            time.sleep(base_sleep_seconds * (2 ** (attempt - 1)))

    if last_error is not None:
        raise last_error
    raise requests.RequestException("Marorka API request failed before a response was received.")


def odata_quote(value: str) -> str:
    return str(value).replace("'", "''")


def build_reportdata_value_filter() -> str:
    value_filters = [
        f"ValueDescription eq '{odata_quote(value)}'"
        for value in REPORTDATA_VALUE_WHITELIST
    ]
    return "(" + " or ".join(value_filters) + ")"


def build_odata_url(start_date: date) -> str:
    # Keep the OData request simple. The Marorka OData V1 ReportData endpoint
    # rejects long ValueDescription OR filters with 404. We therefore request
    # the date window only, then apply the KPI/consumption ValueDescription
    # whitelist locally inside compact_odata_rows() page-by-page before writing
    # to the Parquet snapshot. This preserves the same final dataset while
    # avoiding invalid/oversized OData URLs.
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


def should_continue_odata_paging(
    *,
    current_url: str,
    next_link: str | None,
    seen_urls: set[str],
    consecutive_empty_pages: int,
) -> tuple[bool, str | None, str | None]:
    """Return whether OData paging should continue plus next URL and stop reason.

    AtlasFlow follows the OData nextLink until the feed is exhausted, but still
    protects Streamlit Cloud from pagination loops or abnormal empty-page runs.
    """
    if not next_link:
        return False, None, "end_of_feed"

    if consecutive_empty_pages >= MAX_CONSECUTIVE_EMPTY_ODATA_PAGES:
        return False, None, "consecutive_empty_pages"

    resolved_next_url = urljoin(current_url, next_link)
    if resolved_next_url in seen_urls:
        return False, None, "repeated_next_link"

    return True, resolved_next_url, None


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
        value_description = row.get("ValueDescription")
        if value_description is None:
            continue
        if re.sub(r"[^a-z0-9]+", "", str(value_description).lower()) not in REPORTDATA_VALUE_WHITELIST_KEYS:
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
    consecutive_empty_pages = 0
    paging_stop_reason = "max_page_limit"
    first_url = next_url
    auth = request_auth(username, password, auth_method)
    headers = request_headers(token, auth_method)

    with requests.Session() as session:
        session.headers.update(headers)
        for _ in range(MAX_ODATA_PAGES):
            if next_url in seen_urls:
                paging_stop_reason = "repeated_current_url"
                break
            seen_urls.add(next_url)

            response = request_with_retry(session, next_url, auth=auth, timeout=90)
            total_bytes += len(response.content)
            response.raise_for_status()
            pages += 1

            page_rows, next_link = extract_odata_page(response.json())
            scanned_rows += len(page_rows)
            kept_rows.extend(compact_odata_rows(page_rows))
            consecutive_empty_pages = consecutive_empty_pages + 1 if len(page_rows) == 0 else 0

            should_continue, resolved_next_url, stop_reason = should_continue_odata_paging(
                current_url=next_url,
                next_link=next_link,
                seen_urls=seen_urls,
                consecutive_empty_pages=consecutive_empty_pages,
            )
            if not should_continue:
                paging_stop_reason = stop_reason
                break
            next_url = resolved_next_url or next_url

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
        "hit_page_limit": pages >= MAX_ODATA_PAGES and paging_stop_reason == "max_page_limit",
        "paging_stop_reason": paging_stop_reason,
        "max_pages": MAX_ODATA_PAGES,
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
# Multi-source wide OData helpers
# =============================================================================


def build_wide_odata_url(endpoint: str, start_date: date, datetime_column: str = "DateTime") -> str:
    start_text = start_date.strftime("%Y-%m-%d")
    params = {"$filter": f"{datetime_column} gt DateTime'{start_text}'"}
    return f"{endpoint}?{urlencode(params)}"


def fetch_wide_odata_source(
    source_key: str,
    username: str,
    password: str,
    token: str,
    auth_method: str,
    start_date: date,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    config = SOURCE_CONFIGS[source_key]
    endpoint = str(config["endpoint"])
    datetime_column = str(config.get("datetime_candidates", ["DateTime"])[0])
    next_url = build_wide_odata_url(endpoint, start_date, datetime_column)
    first_url = next_url
    seen_urls: set[str] = set()
    rows: list[dict[str, Any]] = []
    pages = 0
    total_bytes = 0
    consecutive_empty_pages = 0
    paging_stop_reason = "max_page_limit"
    started_at = time.perf_counter()
    auth = request_auth(username, password, auth_method)
    headers = request_headers(token, auth_method)

    with requests.Session() as session:
        session.headers.update(headers)
        for _ in range(MAX_ODATA_PAGES):
            if next_url in seen_urls:
                paging_stop_reason = "repeated_current_url"
                break
            seen_urls.add(next_url)
            response = request_with_retry(session, next_url, auth=auth, timeout=90)
            total_bytes += len(response.content)
            response.raise_for_status()
            pages += 1
            page_rows, next_link = extract_odata_page(response.json())
            rows.extend(page_rows)
            consecutive_empty_pages = consecutive_empty_pages + 1 if len(page_rows) == 0 else 0
            should_continue, resolved_next_url, stop_reason = should_continue_odata_paging(
                current_url=next_url,
                next_link=next_link,
                seen_urls=seen_urls,
                consecutive_empty_pages=consecutive_empty_pages,
            )
            if not should_continue:
                paging_stop_reason = stop_reason
                break
            next_url = resolved_next_url or next_url

    df = pd.DataFrame(rows)
    if "__metadata" in df.columns:
        df = df.drop(columns=["__metadata"])

    loaded_at_utc = datetime.now(timezone.utc)
    metadata = {
        "source": config["label"],
        "endpoint": endpoint,
        "loaded_at_utc": loaded_at_utc.strftime("%d-%m-%Y %H:%M:%S UTC"),
        "loaded_at_local": local_time_label(loaded_at_utc),
        "loaded_start_date": start_date.isoformat(),
        "rows": int(len(df)),
        "columns": int(len(df.columns)),
        "pages": pages,
        "downloaded_mb": round(total_bytes / 1024 / 1024, 2),
        "fetch_seconds": round(time.perf_counter() - started_at, 2),
        "first_url": first_url,
        "hit_page_limit": pages >= MAX_ODATA_PAGES and paging_stop_reason == "max_page_limit",
        "paging_stop_reason": paging_stop_reason,
        "max_pages": MAX_ODATA_PAGES,
    }
    return df, metadata


@st.cache_data(ttl=API_CACHE_TTL_SECONDS, show_spinner=False)
def cached_fetch_wide_odata_source(
    source_key: str,
    username: str,
    password: str,
    token: str,
    auth_method: str,
    start_date: date,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    return fetch_wide_odata_source(source_key, username, password, token, auth_method, start_date)


def source_signature(source_key: str, username: str, auth_method: str, start_date: date) -> dict[str, Any]:
    config = SOURCE_CONFIGS[source_key]
    return {
        "source": source_key,
        "endpoint": str(config["endpoint"]),
        "username_hash": sha256(username.encode("utf-8")).hexdigest()[:12],
        "auth_method": auth_method.lower(),
        "start_date": start_date.isoformat(),
    }


def save_source_snapshot(source_key: str, df: pd.DataFrame, metadata: dict[str, Any], signature: dict[str, Any]) -> None:
    try:
        config = SOURCE_CONFIGS[source_key]
        SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
        df.to_parquet(config["snapshot_file"], index=False)
        payload = {
            "metadata": metadata,
            "signature": signature,
            "saved_at_utc": datetime.now(timezone.utc).strftime("%d-%m-%Y %H:%M:%S UTC"),
        }
        Path(config["metadata_file"]).write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    except Exception:
        return


def load_source_snapshot(
    source_key: str,
    requested_signature: dict[str, Any],
    requested_start_date: date,
) -> tuple[pd.DataFrame, dict[str, Any], dict[str, Any]] | None:
    try:
        config = SOURCE_CONFIGS[source_key]
        snapshot_file = Path(config["snapshot_file"])
        metadata_file = Path(config["metadata_file"])
        if not snapshot_file.is_file() or not metadata_file.is_file():
            return None
        payload = json.loads(metadata_file.read_text(encoding="utf-8"))
        metadata = payload.get("metadata") or {}
        signature = payload.get("signature") or {}
        if not raw_data_covers_request(signature, metadata, requested_signature, requested_start_date):
            return None
        df = pd.read_parquet(snapshot_file)
        if not isinstance(df, pd.DataFrame):
            return None
        # Reject broken placeholder snapshots that can be created after an interrupted warmup.
        # If API metadata says rows exist, a one-column NoData parquet is not a usable source snapshot.
        if list(df.columns) == ["NoData"] and int(metadata.get("rows", 0) or 0) > 0:
            return None
        metadata = metadata.copy()
        metadata["loaded_from_snapshot"] = True
        metadata.setdefault("snapshot_saved_at_utc", payload.get("saved_at_utc", "-"))
        return df, metadata, signature
    except Exception:
        return None


def parse_wide_source_datetimes(df: pd.DataFrame) -> pd.DataFrame:
    parsed_df = df.copy()
    for column in parsed_df.columns:
        if "date" in str(column).lower() or "time" in str(column).lower():
            parsed = parse_datetime_series(parsed_df[column])
            if parsed.notna().any():
                parsed_df[column] = parsed
    return parsed_df


def detect_datetime_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for column in candidates:
        if column in df.columns:
            return column
    for column in df.columns:
        lower = str(column).lower()
        if "datetime" in lower or lower in {"date", "timestamp"}:
            return column
    return None


def filter_wide_source_data(
    df: pd.DataFrame,
    source_key: str,
    selected_vessels: list[str],
    selected_start: date,
    selected_end: date,
) -> pd.DataFrame:
    if df.empty:
        return df
    filtered = parse_wide_source_datetimes(df)
    if "ShipName" in filtered.columns and selected_vessels:
        filtered = filtered[match_selected_vessels(filtered["ShipName"], selected_vessels)].copy()
    datetime_column = detect_datetime_column(filtered, list(SOURCE_CONFIGS[source_key].get("datetime_candidates", [])))
    if datetime_column and datetime_column in filtered.columns:
        values = pd.to_datetime(filtered[datetime_column], errors="coerce", utc=True)
        start_timestamp = pd.Timestamp(selected_start, tz="UTC")
        end_timestamp = pd.Timestamp(selected_end + timedelta(days=1), tz="UTC")
        filtered = filtered[values.ge(start_timestamp) & values.lt(end_timestamp)].copy()
    return filtered


def load_or_fetch_source(
    source_key: str,
    username: str,
    password: str,
    token: str,
    auth_method: str,
    start_date: date,
    refresh: bool,
    auto_fetch: bool = False,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    sig = source_signature(source_key, username, auth_method, start_date)
    state_df_key = f"loaded_{source_key}_df"
    state_meta_key = f"loaded_{source_key}_metadata"
    state_sig_key = f"loaded_{source_key}_signature"
    df = st.session_state.get(state_df_key)
    metadata = st.session_state.get(state_meta_key)
    current_signature = st.session_state.get(state_sig_key)

    needs_load = (
        refresh
        or not isinstance(df, pd.DataFrame)
        or not isinstance(metadata, dict)
        or not raw_data_covers_request(current_signature, metadata, sig, start_date)
    )

    if needs_load and not refresh:
        snapshot = load_source_snapshot(source_key, sig, start_date)
        if snapshot is not None:
            df, metadata, snapshot_sig = snapshot
            st.session_state[state_df_key] = df
            st.session_state[state_meta_key] = metadata
            st.session_state[state_sig_key] = snapshot_sig
            needs_load = False

    if needs_load and not (refresh or auto_fetch):
        config = SOURCE_CONFIGS[source_key]
        empty_metadata = {
            "source": config["label"],
            "endpoint": str(config["endpoint"]),
            "loaded_at_utc": "-",
            "loaded_at_local": "No stored snapshot yet",
            "loaded_from_snapshot": False,
            "rows": 0,
            "columns": 0,
            "pages": 0,
            "first_url": "-",
            "needs_warmup": True,
        }
        return pd.DataFrame(), empty_metadata

    if needs_load:
        if refresh:
            cached_fetch_wide_odata_source.clear()
        df, metadata = fetch_wide_odata_source(source_key, username, password, token, auth_method, start_date)
        save_source_snapshot(source_key, df, metadata, sig)
        st.session_state[state_df_key] = df
        st.session_state[state_meta_key] = metadata
        st.session_state[state_sig_key] = sig

    return st.session_state[state_df_key], st.session_state[state_meta_key]


def render_wide_source_tab(source_label: str, df: pd.DataFrame, metadata: dict[str, Any], source_key: str, selected_vessels: list[str], selected_start: date, selected_end: date) -> pd.DataFrame:
    st.markdown(f'<div class="section-title">{escape(source_label)} Dataset</div>', unsafe_allow_html=True)
    render_api_load_caption(metadata)
    if list(df.columns) == ["NoData"] and int(metadata.get("rows", 0) or 0) > 0:
        st.error(
            f"{source_label} snapshot metadata shows {int(metadata.get('rows', 0)):,} API rows, "
            "but the stored parquet contains only a placeholder column. "
            f"Run the {source_key} warmup again with the latest app version."
        )
        return pd.DataFrame()

    filtered_df = filter_wide_source_data(df, source_key, selected_vessels, selected_start, selected_end)
    render_metric_cards(
        [
            ("Rows", f"{len(filtered_df):,}", "table_eye"),
            ("Columns", f"{len(filtered_df.columns):,}", "checked_columns"),
            ("API Rows", f"{metadata.get('rows', len(df)):,}", "database_rows"),
            ("API Pages", f"{metadata.get('pages', 0):,}", "numeric"),
        ]
    )

    default_columns = [c for c in ["ShipName", "DateTime", "State", "StateName", "GPSSpeed", "LogSpeed", "MEConsumed", "ShaftPower"] if c in filtered_df.columns]
    if not default_columns:
        default_columns = list(filtered_df.columns[: min(12, len(filtered_df.columns))])
    selected_columns = st.multiselect(
        f"{source_label} columns to preview/export",
        options=list(filtered_df.columns),
        default=default_columns,
        key=f"{source_key}_preview_columns",
    )
    if not selected_columns:
        selected_columns = default_columns
    output = filtered_df[selected_columns].copy() if selected_columns else filtered_df.copy()
    render_preview_table(output)
    if len(output) > TABLE_PREVIEW_ROW_LIMIT:
        st.caption(f"Showing first {TABLE_PREVIEW_ROW_LIMIT:,} of {len(output):,} rows. Export includes all filtered rows/columns selected above.")
    return output


def dataframe_memory_mb(df: Any) -> float:
    if not isinstance(df, pd.DataFrame):
        return 0.0
    try:
        return float(df.memory_usage(deep=True).sum()) / 1024 / 1024
    except Exception:
        return 0.0


def clear_wide_source_state(source_key: str) -> None:
    """Release wide-source DataFrames from this browser session.

    Snapshots remain on disk, so reopening a source reloads from Parquet rather
    than calling the API. This is the main Streamlit Cloud memory safeguard.
    """
    for suffix in ["df", "metadata", "signature"]:
        st.session_state.pop(f"loaded_{source_key}_{suffix}", None)


def clear_inactive_wide_sources(active_sources: set[str]) -> None:
    for source_key in ["reportpivots", "shippivots"]:
        if source_key not in active_sources:
            clear_wide_source_state(source_key)
    gc.collect()


def clear_stale_export_bytes(current_signature: str | None = None) -> None:
    """Remove large Excel byte buffers when they no longer match the current view."""
    if current_signature is None or st.session_state.get("atlas_export_signature") != current_signature:
        st.session_state.pop("atlas_export_bytes", None)
    if current_signature is None or st.session_state.get("atlas_multisource_export_signature") != current_signature:
        st.session_state.pop("atlas_multisource_export_bytes", None)
    gc.collect()


def current_memory_audit_rows(extra_frames: dict[str, pd.DataFrame] | None = None) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for key, value in st.session_state.items():
        if isinstance(value, pd.DataFrame):
            rows.append({"Object": f"session_state.{key}", "Rows": len(value), "Columns": len(value.columns), "Memory MB": round(dataframe_memory_mb(value), 2)})
        elif isinstance(value, (bytes, bytearray)):
            rows.append({"Object": f"session_state.{key}", "Rows": "-", "Columns": "-", "Memory MB": round(len(value) / 1024 / 1024, 2)})
    for name, frame in (extra_frames or {}).items():
        if isinstance(frame, pd.DataFrame):
            rows.append({"Object": name, "Rows": len(frame), "Columns": len(frame.columns), "Memory MB": round(dataframe_memory_mb(frame), 2)})
    if not rows:
        return pd.DataFrame(columns=["Object", "Rows", "Columns", "Memory MB"])
    return pd.DataFrame(rows).sort_values("Memory MB", ascending=False)


def wide_source_selected_columns(source_key: str, filtered_df: pd.DataFrame) -> list[str]:
    if filtered_df.empty:
        return []
    default_columns = [
        c for c in ["ShipName", "DateTime", "State", "StateName", "GPSSpeed", "LogSpeed", "MEConsumed", "ShaftPower"]
        if c in filtered_df.columns
    ]
    if not default_columns:
        default_columns = list(filtered_df.columns[: min(12, len(filtered_df.columns))])
    previous = st.session_state.get(f"{source_key}_preview_columns", default_columns)
    if not isinstance(previous, list):
        previous = default_columns
    selected_columns = [column for column in previous if column in filtered_df.columns]
    return selected_columns or default_columns


def load_wide_source_for_view(
    source_key: str,
    username: str,
    password: str,
    token: str,
    auth_method: str,
    start_date: date,
    refresh: bool,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Load one wide source only when the user opens it.

    This keeps ReportPivots and ShipPivots out of memory during normal Custom
    Analytics use. If no snapshot exists, the UI instructs the user to run the
    per-source warmup rather than live-loading a huge API in the app session.
    """
    return load_or_fetch_source(
        source_key,
        username,
        password,
        token,
        auth_method,
        start_date,
        refresh=refresh,
        auto_fetch=False,
    )


def build_wide_source_output_for_export(
    source_key: str,
    source_df: pd.DataFrame,
    selected_vessels: list[str],
    selected_start: date,
    selected_end: date,
) -> pd.DataFrame:
    filtered_df = filter_wide_source_data(source_df, source_key, selected_vessels, selected_start, selected_end)
    selected_columns = wide_source_selected_columns(source_key, filtered_df)
    if not selected_columns:
        return filtered_df.copy()
    return filtered_df[selected_columns].copy()


def to_multisource_excel_bytes(
    reportdata_df: pd.DataFrame,
    reportdata_summary_df: pd.DataFrame | None,
    reportpivots_df: pd.DataFrame,
    shippivots_df: pd.DataFrame,
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        write_table_sheet(writer, reportdata_df, "ReportData Clean", "ReportDataClean")
        if reportdata_summary_df is not None and not reportdata_summary_df.empty:
            write_table_sheet(writer, reportdata_summary_df, "ReportData Summary", "ReportDataSummary")
        if reportpivots_df is not None and not reportpivots_df.empty:
            write_table_sheet(writer, reportpivots_df, "ReportPivots", "ReportPivotsData")
        if shippivots_df is not None and not shippivots_df.empty:
            write_table_sheet(writer, shippivots_df, "ShipPivots", "ShipPivotsData")
    return output.getvalue()


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


def safe_divide(numerator: Any, denominator: Any) -> Any:
    numerator = pd.to_numeric(numerator, errors="coerce")
    denominator = pd.to_numeric(denominator, errors="coerce")
    denominator = denominator.mask(denominator == 0)
    return numerator / denominator


def sum_numeric_columns(df: pd.DataFrame, columns: list[str]) -> pd.Series:
    available_columns = [column for column in columns if column in df.columns]
    if not available_columns:
        return pd.Series(pd.NA, index=df.index, dtype="Float64")
    return df[available_columns].apply(pd.to_numeric, errors="coerce").sum(axis=1, min_count=1)


def calculation_alias_to_column() -> dict[str, str]:
    return {
        normalize_text(alias): column
        for column, aliases in DERIVED_VALUE_ALIASES.items()
        for alias in aliases
    }


def calculate_rob_consumption(
    df: pd.DataFrame,
    rob_column: str,
    received_column: str,
) -> pd.Series:
    """Calculate row-level consumption from ROB movement inside the current sample.

    The first report per vessel has no previous ROB inside the selected sample, so it
    remains blank. Negative values are treated as blank because they usually indicate
    correction/noise or a missing receipt entry rather than real consumption.
    """
    if rob_column not in df.columns:
        return pd.Series(pd.NA, index=df.index, dtype="Float64")

    work = df[["ShipName", "StartDateTimeGMT", "EndDateTimeGMT", rob_column]].copy()
    if received_column in df.columns:
        work[received_column] = pd.to_numeric(df[received_column], errors="coerce").fillna(0)
    else:
        work[received_column] = 0.0

    work[rob_column] = pd.to_numeric(work[rob_column], errors="coerce")
    work["_original_index"] = df.index
    work["_sort_date"] = pd.to_datetime(work["EndDateTimeGMT"], errors="coerce", utc=True)
    fallback_dates = pd.to_datetime(work["StartDateTimeGMT"], errors="coerce", utc=True)
    work["_sort_date"] = work["_sort_date"].fillna(fallback_dates)
    work = work.sort_values(["ShipName", "_sort_date", "_original_index"])
    previous_rob = work.groupby("ShipName", dropna=False)[rob_column].shift(1)
    consumption = previous_rob + work[received_column] - work[rob_column]
    consumption = consumption.where(consumption >= 0)
    consumption = consumption.where(previous_rob.notna() & work[rob_column].notna())
    result = pd.Series(pd.NA, index=df.index, dtype="Float64")
    result.loc[work["_original_index"]] = pd.to_numeric(consumption, errors="coerce").to_numpy()
    return result


def build_calculation_source_table(filtered_long_df: pd.DataFrame) -> pd.DataFrame:
    if filtered_long_df.empty:
        return pd.DataFrame(columns=PIVOT_IDENTITY_COLUMNS)

    alias_to_column = calculation_alias_to_column()
    source_long = filtered_long_df.copy()
    source_long["_value_key"] = source_long["ValueDescription"].map(normalize_text)
    source_long = source_long[
        source_long["_value_key"].isin(alias_to_column)
        & source_long["ParsedValue"].notna()
    ].copy()

    if source_long.empty:
        return filtered_long_df[PIVOT_IDENTITY_COLUMNS].drop_duplicates().copy()

    source_long["_canonical_column"] = source_long["_value_key"].map(alias_to_column)
    source_long["_source_order"] = range(len(source_long))
    source_long = source_long.sort_values("_source_order")
    source_long = source_long.drop_duplicates(
        [*PIVOT_IDENTITY_COLUMNS, "_canonical_column"],
        keep="last",
    )

    source_table = (
        source_long
        .pivot(index=PIVOT_IDENTITY_COLUMNS, columns="_canonical_column", values="ParsedValue")
        .reset_index()
    )
    source_table.columns.name = None
    return source_table


def add_performance_calculations(pivot_df: pd.DataFrame, source_table: pd.DataFrame) -> pd.DataFrame:
    if pivot_df.empty:
        for column in DERIVED_VARIABLES:
            if column not in pivot_df.columns:
                pivot_df[column] = pd.NA
        return pivot_df

    df = pivot_df.copy()
    if not source_table.empty:
        calculation_columns = [
            column for column in source_table.columns
            if column not in PIVOT_IDENTITY_COLUMNS and column not in df.columns
        ]
        if calculation_columns:
            df = df.merge(
                source_table[[*PIVOT_IDENTITY_COLUMNS, *calculation_columns]],
                on=PIVOT_IDENTITY_COLUMNS,
                how="left",
            )

    def numeric_column(column: str) -> pd.Series:
        if column not in df.columns:
            return pd.Series(pd.NA, index=df.index, dtype="Float64")
        return pd.to_numeric(df[column], errors="coerce")

    lap_time = numeric_column("LapTime")
    engine_distance = numeric_column("Engine Distance [nm]")
    distance_over_ground = numeric_column("Distance Over Ground [nm]")
    power = numeric_column("Power from Torque Meter [kW]")

    df["Calculated Slip"] = (1 - safe_divide(distance_over_ground, engine_distance)).round(3)

    me_sum = sum_numeric_columns(df, ME_FUEL_COLUMNS)
    official_me_total = numeric_column("Main Engine Total Consumed")
    df["ME Consumption Total"] = me_sum.fillna(official_me_total).round(3)

    dg_sum = sum_numeric_columns(df, DG_FUEL_COLUMNS)
    official_dg_total = numeric_column("Diesel Generator Total Consumed")
    df["DG Consumption Total"] = dg_sum.fillna(official_dg_total).round(3)

    aux_sum = sum_numeric_columns(df, AUXILIARY_FUEL_COLUMNS)
    official_aux_total = numeric_column("Auxiliary Engine Total Consumed")
    df["Auxiliary Engine Consumption Total"] = aux_sum.fillna(official_aux_total).round(3)

    df["Boiler Sum"] = sum_numeric_columns(df, BOILER_FUEL_COLUMNS).round(3)

    calculated_total = pd.concat(
        [
            pd.to_numeric(df["ME Consumption Total"], errors="coerce"),
            pd.to_numeric(df["DG Consumption Total"], errors="coerce"),
            pd.to_numeric(df["Auxiliary Engine Consumption Total"], errors="coerce"),
            pd.to_numeric(df["Boiler Sum"], errors="coerce"),
        ],
        axis=1,
    ).sum(axis=1, min_count=1)
    official_total = numeric_column("Total Fuel Consumed")
    df["Total Fuel Consumption"] = calculated_total.fillna(official_total).round(3)

    df["Consumption ME 24 Hours [MT]"] = safe_divide(df["ME Consumption Total"] * 24, lap_time).round(3)

    df["SFOC [gr/Kwh]"] = (
        safe_divide(df["Consumption ME 24 Hours [MT]"], power) / 0.000024
    ).round(3).fillna(0)

    # Oil consumption aggregation variables. These are row-level consumption
    # movements calculated inside the current selected sample, so Summary Analysis
    # can later Sum them by vessel/fleet/month/report type.
    df["MELO Consumption Total [ltr]"] = calculate_rob_consumption(
        df,
        "MELO ROB [ltr]",
        "MELO Received [ltr]",
    ).round(3)
    cylinder_oil_1_consumption = calculate_rob_consumption(
        df,
        "Cylinder Oil 1 ROB [ltr]",
        "Cylinder Oil 1 Received [ltr]",
    )
    cylinder_oil_2_consumption = calculate_rob_consumption(
        df,
        "Cylinder Oil 2 ROB [ltr]",
        "Cylinder Oil 2 Received [ltr]",
    )
    df["CYLO Consumption Total [ltr]"] = pd.concat(
        [cylinder_oil_1_consumption, cylinder_oil_2_consumption],
        axis=1,
    ).sum(axis=1, min_count=1).round(3)
    df["GELO Consumption Total [ltr]"] = calculate_rob_consumption(
        df,
        "GELO ROB [ltr]",
        "GELO Received [ltr]",
    ).round(3)
    df["Total DG Running Hours [hh:mm]"] = sum_numeric_columns(
        df,
        [
            "DG1 Running Hours [hh:mm]",
            "DG2 Running Hours [hh:mm]",
            "DG3 Running Hours [hh:mm]",
            "DG4 Running Hours [hh:mm]",
        ],
    ).round(3)

    return df


@st.cache_data(show_spinner=False)
def build_pivot_table(filtered_long_df: pd.DataFrame, selected_variables: tuple[str, ...]) -> pd.DataFrame:
    if filtered_long_df.empty:
        return pd.DataFrame(columns=PIVOT_IDENTITY_COLUMNS + list(selected_variables))

    calculation_source_table = build_calculation_source_table(filtered_long_df)

    api_selected_variables = [
        variable for variable in selected_variables
        if variable not in DERIVED_VARIABLES
    ]

    if not api_selected_variables:
        pivot_df = (
            filtered_long_df[PIVOT_IDENTITY_COLUMNS]
            .drop_duplicates()
            .sort_values(["ShipName", "EndDateTimeGMT"], ascending=[True, False])
            .reset_index(drop=True)
        )
    else:
        selected_long = filtered_long_df[
            filtered_long_df["ValueDescription"].astype("string").isin(api_selected_variables)
        ].copy()

        if selected_long.empty:
            pivot_df = filtered_long_df[PIVOT_IDENTITY_COLUMNS].drop_duplicates().copy()
            for variable in api_selected_variables:
                pivot_df[variable] = pd.NA
        else:
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

        for variable in api_selected_variables:
            if variable not in pivot_df.columns:
                pivot_df[variable] = pd.NA

    pivot_df = add_performance_calculations(pivot_df, calculation_source_table)

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


METRIC_ICON_SVGS = {
    "table_eye": '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="5" width="18" height="14" rx="1.8"></rect><path d="M3 10h18M8 5v14M16 5v14"></path><path d="M8.4 12.3c1.1-1.3 2.3-1.9 3.6-1.9s2.5.6 3.6 1.9c-1.1 1.3-2.3 1.9-3.6 1.9s-2.5-.6-3.6-1.9Z"></path><circle cx="12" cy="12.3" r="1.1"></circle></svg>',
    "checked_columns": '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="4" width="16" height="16" rx="2"></rect><path d="M9 4v16M15 4v16M7.2 12.2l1.9 1.9 3.8-4.4"></path></svg>',
    "database_rows": '<svg viewBox="0 0 24 24" aria-hidden="true"><ellipse cx="12" cy="5" rx="8" ry="3"></ellipse><path d="M4 5v6c0 1.7 3.6 3 8 3s8-1.3 8-3V5"></path><path d="M4 11v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6"></path><path d="M8 9h8M8 15h8"></path></svg>',
    "columns_plus": '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="13" height="16" rx="2"></rect><path d="M7.3 4v16M11.7 4v16M18 9v8M14 13h8"></path></svg>',
    "average": '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 19V5"></path><path d="M7 17V9M12 17V6M17 17v-5"></path><path d="M4 12h16"></path></svg>',
    "total": '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M17.5 5H7l6 7-6 7h10.5"></path></svg>',
    "numeric": '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 8h3v8M4 16h6M13 8h4l-4 8h4M20 8v8"></path></svg>',
    "missing": '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="4" width="16" height="16" rx="2" stroke-dasharray="3 2"></rect><path d="M10 9.5a2.2 2.2 0 1 1 3.3 1.9c-.8.5-1.3 1-1.3 2.1"></path><path d="M12 17h.01"></path></svg>',
    # Backwards-compatible names used by older cards.
    "table": '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="18" height="18" rx="1.8"></rect><path d="M3 9h18M3 15h18M9 3v18M15 3v18"></path></svg>',
    "list": '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 6h13M8 12h13M8 18h9"></path><path d="M3.5 6h.01M3.5 12h.01M3.5 18h.01"></path></svg>',
    "database": '<svg viewBox="0 0 24 24" aria-hidden="true"><ellipse cx="12" cy="5" rx="8" ry="3"></ellipse><path d="M4 5v6c0 1.7 3.6 3 8 3s8-1.3 8-3V5"></path><path d="M4 11v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6"></path></svg>',
    "nodes": '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="4" r="2.5"></circle><circle cx="5" cy="19" r="2.5"></circle><circle cx="19" cy="19" r="2.5"></circle><path d="M10.8 6.2 6.2 16.8M13.2 6.2l4.6 10.6M7.5 19h9"></path></svg>',
}

def render_metric_cards(cards: list[tuple[str, str, str]]) -> None:
    card_html = []
    for label, value, icon_name in cards:
        icon_svg = METRIC_ICON_SVGS.get(icon_name, METRIC_ICON_SVGS["table"])
        card_html.append(
            f'<div class="atlas-metric-card"><div class="atlas-metric-icon">{icon_svg}</div>'
            f'<div><div class="atlas-metric-label">{escape(label)}</div>'
            f'<div class="atlas-metric-value">{escape(value)}</div></div></div>'
        )
    st.markdown(f'<div class="atlas-metric-grid">{"".join(card_html)}</div>', unsafe_allow_html=True)


def render_preview_table(df: pd.DataFrame, row_limit: int = TABLE_PREVIEW_ROW_LIMIT) -> None:
    preview_df = format_display_dataframe(df.head(row_limit))
    if preview_df.empty:
        st.markdown('<div class="atlas-table-empty">No rows to display.</div>', unsafe_allow_html=True)
        return

    columns = [str(column) for column in preview_df.columns]
    header_html = "".join(f"<th>{escape(column)}</th>" for column in columns)
    rows_html: list[str] = []
    for row in preview_df.itertuples(index=False, name=None):
        cell_html = "".join(("<td>" + escape(str(value)) + "</td>") for value in row)
        rows_html.append(f"<tr>{cell_html}</tr>")

    table_html = (
        '<div class="atlas-table-frame"><div class="atlas-table-scroll">'
        f'<table class="atlas-preview-table"><thead><tr>{header_html}</tr></thead>'
        f'<tbody>{"".join(rows_html)}</tbody></table></div></div>'
    )
    st.markdown(table_html, unsafe_allow_html=True)


def make_unique_excel_columns(columns: list[Any]) -> list[str]:
    """Return Excel-table-safe, unique column names while preserving readable labels."""
    safe_columns: list[str] = []
    seen: dict[str, int] = {}

    for position, column in enumerate(columns, start=1):
        label = str(column).strip() if column is not None else ""
        label = re.sub(r"[\x00-\x1f]", "", label)
        if not label or label.lower() in {"nan", "nat", "none"}:
            label = f"Column {position}"
        label = label[:240]

        key = label.casefold()
        count = seen.get(key, 0)
        if count:
            suffix = f"_{count + 1}"
            label = f"{label[:240 - len(suffix)]}{suffix}"
            key = label.casefold()
        seen[key] = count + 1
        safe_columns.append(label)

    return safe_columns


def make_excel_safe_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    safe_df = df.copy()
    safe_df.columns = make_unique_excel_columns(list(safe_df.columns))
    safe_df = safe_df.replace([float("inf"), float("-inf")], pd.NA)
    for column in safe_df.columns:
        if pd.api.types.is_datetime64_any_dtype(safe_df[column]):
            safe_df[column] = pd.to_datetime(safe_df[column], errors="coerce").dt.tz_localize(None)
    return safe_df


def add_excel_table(worksheet: Any, table_name: str) -> None:
    if worksheet.max_row < 2 or worksheet.max_column < 1:
        return

    headers = [worksheet.cell(row=1, column=col).value for col in range(1, worksheet.max_column + 1)]
    if any(header is None or str(header).strip() == "" for header in headers):
        return
    if len({str(header).casefold() for header in headers}) != len(headers):
        return

    table_ref = f"A1:{get_column_letter(worksheet.max_column)}{worksheet.max_row}"
    table = Table(displayName=table_name, ref=table_ref)
    table.tableStyleInfo = TableStyleInfo(
        # Built-in table style kept neutral; explicit teal formatting below
        # gives the export a stable AtlasFlow look in Excel.
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    worksheet.add_table(table)


def apply_teal_excel_table_format(worksheet: Any) -> None:
    """Apply AtlasFlow teal styling to exported Excel tables."""
    if worksheet.max_row < 1 or worksheet.max_column < 1:
        return

    header_fill = PatternFill(fill_type="solid", fgColor="006B68")
    even_fill = PatternFill(fill_type="solid", fgColor="EAF7F5")
    odd_fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
    border_side = Side(style="thin", color="B7DCD8")
    cell_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = cell_border

    for row_number in range(2, worksheet.max_row + 1):
        fill = even_fill if row_number % 2 == 0 else odd_fill
        for cell in worksheet[row_number]:
            cell.fill = fill
            cell.border = cell_border
            cell.alignment = Alignment(vertical="center", wrap_text=False)


def autofit_excel_columns(worksheet: Any, max_width: int = 48) -> None:
    for column_cells in worksheet.columns:
        max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
        worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), max_width)


def write_table_sheet(writer: Any, df: pd.DataFrame, sheet_name: str, table_name: str) -> None:
    safe_df = make_excel_safe_dataframe(df)
    safe_df.to_excel(writer, index=False, sheet_name=sheet_name)
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes = "A2"
    autofit_excel_columns(worksheet)
    add_excel_table(worksheet, table_name)
    apply_teal_excel_table_format(worksheet)


def to_excel_bytes(clean_df: pd.DataFrame, pivot_analysis_df: pd.DataFrame | None = None) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        write_table_sheet(writer, clean_df, "Clean Dataset", "AtlasFlowCleanDataset")

        if pivot_analysis_df is not None and not pivot_analysis_df.empty:
            write_table_sheet(writer, pivot_analysis_df, "Summary Analysis", "AtlasFlowSummaryAnalysis")

    return output.getvalue()


def to_displayed_table_excel_bytes(display_df: pd.DataFrame, sheet_name: str = "Displayed Table") -> bytes:
    """Export only the table currently visible to the user.

    This avoids preparing multiple hidden sheets during a normal tab export and keeps
    memory usage lower on Streamlit Cloud.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        write_table_sheet(writer, display_df, sheet_name[:31], "AtlasFlowDisplayedTable")
    return output.getvalue()


def flatten_pivot_columns(df: pd.DataFrame) -> pd.DataFrame:
    flat_df = df.copy()
    if isinstance(flat_df.columns, pd.MultiIndex):
        flat_df.columns = [
            " | ".join(str(part) for part in column if str(part) not in {"", "nan", "NaT"})
            for column in flat_df.columns.to_flat_index()
        ]
    else:
        flat_df.columns = [str(column) for column in flat_df.columns]
    return flat_df


def build_summary_analysis(
    clean_df: pd.DataFrame,
    group_fields: list[str],
    value_fields: list[str],
    aggregation: str,
) -> pd.DataFrame:
    if clean_df.empty or not group_fields or not value_fields:
        return pd.DataFrame()

    valid_group_fields = [column for column in group_fields if column in clean_df.columns]
    valid_value_fields = [column for column in value_fields if column in clean_df.columns]
    if not valid_group_fields or not valid_value_fields:
        return pd.DataFrame()

    source_df = clean_df.copy()
    for column in valid_value_fields:
        source_df[column] = pd.to_numeric(source_df[column], errors="coerce")

    aggregation_map = {
        "Average": "mean",
        "Sum": "sum",
        "Count": "count",
        "Minimum": "min",
        "Maximum": "max",
        "Median": "median",
    }
    aggfunc = aggregation_map.get(aggregation, "mean")

    summary_df = (
        source_df
        .groupby(valid_group_fields, dropna=False, as_index=False)[valid_value_fields]
        .agg(aggfunc)
    )

    if aggregation != "Count":
        for column in valid_value_fields:
            summary_df[column] = pd.to_numeric(summary_df[column], errors="coerce").round(3)

    rename_map = {column: f"{aggregation} {column}" for column in valid_value_fields}
    return summary_df.rename(columns=rename_map)


def numeric_column_options(df: pd.DataFrame) -> list[str]:
    options: list[str] = []
    for column in df.columns:
        values = pd.to_numeric(df[column], errors="coerce")
        if values.notna().any():
            options.append(column)
    return options


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
# Descriptive statistics helpers
# =============================================================================


def dataframe_numeric_options(df: pd.DataFrame) -> list[str]:
    return [column for column in df.columns if pd.to_numeric(df[column], errors="coerce").notna().any()]


def dataframe_categorical_options(df: pd.DataFrame) -> list[str]:
    options: list[str] = []
    for column in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[column]):
            continue
        numeric_values = pd.to_numeric(df[column], errors="coerce")
        if numeric_values.notna().any():
            continue
        if df[column].astype("string").dropna().nunique() <= 250:
            options.append(column)
    return options


def detect_analysis_datetime_column(df: pd.DataFrame) -> str | None:
    preferred_columns = ["StartDateTimeGMT", "EndDateTimeGMT", "DateTime", "ReportDateTime", "Timestamp"]
    for column in preferred_columns:
        if column in df.columns and pd.to_datetime(df[column], errors="coerce", utc=True).notna().any():
            return column
    for column in df.columns:
        lower = str(column).lower()
        if ("date" in lower or "time" in lower) and pd.to_datetime(df[column], errors="coerce", utc=True).notna().any():
            return column
    return None


def build_descriptive_statistics(df: pd.DataFrame, metric_column: str) -> pd.DataFrame:
    values = pd.to_numeric(df[metric_column], errors="coerce")
    clean_values = values.dropna()
    if clean_values.empty:
        return pd.DataFrame()
    rows = [
        ("Rows", len(df)),
        ("Numeric values", int(clean_values.count())),
        ("Missing values", int(values.isna().sum())),
        ("Sum", clean_values.sum()),
        ("Mean", clean_values.mean()),
        ("Median", clean_values.median()),
        ("Std dev", clean_values.std()),
        ("Minimum", clean_values.min()),
        ("P10", clean_values.quantile(0.10)),
        ("P25", clean_values.quantile(0.25)),
        ("P75", clean_values.quantile(0.75)),
        ("P90", clean_values.quantile(0.90)),
        ("Maximum", clean_values.max()),
    ]
    return pd.DataFrame(rows, columns=["Statistic", "Value"])


def build_grouped_descriptive_statistics(df: pd.DataFrame, metric_column: str, group_column: str) -> pd.DataFrame:
    if group_column not in df.columns or metric_column not in df.columns:
        return pd.DataFrame()
    source = df[[group_column, metric_column]].copy()
    source[metric_column] = pd.to_numeric(source[metric_column], errors="coerce")
    source = source[source[metric_column].notna()].copy()
    if source.empty:
        return pd.DataFrame()
    grouped = (
        source
        .groupby(group_column, dropna=False)[metric_column]
        .agg(Count="count", Sum="sum", Mean="mean", Median="median", Minimum="min", Maximum="max")
        .reset_index()
        .sort_values("Sum", ascending=False)
    )
    for column in ["Sum", "Mean", "Median", "Minimum", "Maximum"]:
        grouped[column] = pd.to_numeric(grouped[column], errors="coerce").round(3)
    return grouped


def build_monthly_trend(df: pd.DataFrame, metric_column: str, datetime_column: str) -> pd.DataFrame:
    source = df[[datetime_column, metric_column]].copy()
    source[datetime_column] = pd.to_datetime(source[datetime_column], errors="coerce", utc=True)
    source[metric_column] = pd.to_numeric(source[metric_column], errors="coerce")
    source = source[source[datetime_column].notna() & source[metric_column].notna()].copy()
    if source.empty:
        return pd.DataFrame()
    source["Month"] = source[datetime_column].dt.to_period("M").astype(str)
    trend = (
        source
        .groupby("Month", as_index=False)[metric_column]
        .agg(Count="count", Sum="sum", Mean="mean", Median="median")
        .sort_values("Month")
    )
    for column in ["Sum", "Mean", "Median"]:
        trend[column] = pd.to_numeric(trend[column], errors="coerce").round(3)
    return trend


def render_descriptive_statistics_tab(
    custom_df: pd.DataFrame,
    reportpivots_df: pd.DataFrame,
    shippivots_df: pd.DataFrame,
) -> None:
    st.markdown('<div class="section-title">Descriptive Statistics</div>', unsafe_allow_html=True)
    st.caption("Analyze exactly the same filtered/export-ready tables from each source. No extra KPI logic is applied here.")

    sources = {
        "Custom Analytics": custom_df,
        "Noon & Manual Reports": reportpivots_df,
        "15-Minute Operations": shippivots_df,
    }
    available_sources = [label for label, df in sources.items() if isinstance(df, pd.DataFrame) and not df.empty]
    if not available_sources:
        st.info("No filtered source table is available for descriptive statistics yet.")
        return

    selected_source = st.selectbox("Source table", options=available_sources, key="atlas_descriptive_source")
    analysis_df = sources[selected_source].copy()
    numeric_options = dataframe_numeric_options(analysis_df)
    if not numeric_options:
        st.info("The selected source table has no numeric columns to analyze.")
        return

    metric_column = st.selectbox("Metric to analyze", options=numeric_options, key="atlas_descriptive_metric")
    group_options = ["None"] + dataframe_categorical_options(analysis_df)
    default_group_index = group_options.index("ShipName") if "ShipName" in group_options else 0
    group_column = st.selectbox("Optional group by", options=group_options, index=default_group_index, key="atlas_descriptive_group")

    stats_df = build_descriptive_statistics(analysis_df, metric_column)
    if stats_df.empty:
        st.info("No numeric values were found for the selected metric.")
        return

    values = pd.to_numeric(analysis_df[metric_column], errors="coerce")
    render_metric_cards(
        [
            ("Numeric Values", f"{values.notna().sum():,}", "numeric"),
            ("Total", f"{values.sum(skipna=True):,.3f}", "total"),
            ("Average", f"{values.mean(skipna=True):,.3f}", "average"),
            ("Missing", f"{values.isna().sum():,}", "missing"),
        ]
    )

    st.markdown('<div class="section-title">Overall statistics</div>', unsafe_allow_html=True)
    st.dataframe(format_display_dataframe(stats_df), use_container_width=True, hide_index=True)

    if group_column != "None":
        grouped_df = build_grouped_descriptive_statistics(analysis_df, metric_column, group_column)
        if not grouped_df.empty:
            st.markdown('<div class="section-title">Grouped statistics</div>', unsafe_allow_html=True)
            st.dataframe(format_display_dataframe(grouped_df.head(100)), use_container_width=True, hide_index=True)

    datetime_column = detect_analysis_datetime_column(analysis_df)
    if datetime_column:
        trend_df = build_monthly_trend(analysis_df, metric_column, datetime_column)
        if not trend_df.empty:
            st.markdown('<div class="section-title">Monthly trend</div>', unsafe_allow_html=True)
            st.dataframe(format_display_dataframe(trend_df), use_container_width=True, hide_index=True)
            chart_df = trend_df.set_index("Month")[["Sum", "Mean"]]
            st.line_chart(chart_df)

    outlier_values = values.dropna()
    if len(outlier_values) >= 4:
        q1 = outlier_values.quantile(0.25)
        q3 = outlier_values.quantile(0.75)
        iqr = q3 - q1
        lower = q1 - 1.5 * iqr
        upper = q3 + 1.5 * iqr
        outliers = analysis_df[(values < lower) | (values > upper)].copy()
        if not outliers.empty:
            st.markdown('<div class="section-title">Potential outlier rows</div>', unsafe_allow_html=True)
            display_cols = [column for column in ["ShipName", "ReportType", "StartDateTimeGMT", "DateTime", metric_column] if column in outliers.columns]
            if not display_cols:
                display_cols = list(outliers.columns[: min(8, len(outliers.columns))])
            st.dataframe(format_display_dataframe(outliers[display_cols].head(100)), use_container_width=True, hide_index=True)


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


def save_raw_snapshot(raw_df: pd.DataFrame, metadata: dict[str, Any], signature: dict[str, Any]) -> None:
    """Persist the latest successful raw API pull as a local fallback after app restarts."""
    try:
        SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
        raw_df.to_parquet(RAW_SNAPSHOT_FILE, index=False)
        snapshot_payload = {
            "metadata": metadata,
            "signature": signature,
            "saved_at_utc": datetime.now(timezone.utc).strftime("%d-%m-%Y %H:%M:%S UTC"),
        }
        METADATA_SNAPSHOT_FILE.write_text(json.dumps(snapshot_payload, indent=2, default=str), encoding="utf-8")
    except Exception:
        # Snapshot persistence is a speed fallback only; never break the app if it fails.
        return


def load_raw_snapshot(
    requested_signature: dict[str, Any],
    requested_start_date: date,
) -> tuple[pd.DataFrame, dict[str, Any], dict[str, Any]] | None:
    """Load the latest local raw-data snapshot if it covers the current API request."""
    try:
        if not RAW_SNAPSHOT_FILE.is_file() or not METADATA_SNAPSHOT_FILE.is_file():
            return None
        snapshot_payload = json.loads(METADATA_SNAPSHOT_FILE.read_text(encoding="utf-8"))
        metadata = snapshot_payload.get("metadata") or {}
        signature = snapshot_payload.get("signature") or {}
        if not raw_data_covers_request(signature, metadata, requested_signature, requested_start_date):
            return None
        raw_df = pd.read_parquet(RAW_SNAPSHOT_FILE)
        if not isinstance(raw_df, pd.DataFrame) or raw_df.empty:
            return None
        metadata = metadata.copy()
        metadata["loaded_from_snapshot"] = True
        metadata.setdefault("snapshot_saved_at_utc", snapshot_payload.get("saved_at_utc", "-"))
        return raw_df, metadata, signature
    except Exception:
        return None


def set_loaded_long_state(long_df: pd.DataFrame, signature: dict[str, Any]) -> None:
    st.session_state["loaded_long_df"] = long_df
    st.session_state["loaded_prepare_signature"] = signature


def activate_reportdata_snapshot(
    username: str,
    auth_method: str,
    start_date: date,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    """Load the latest ReportData snapshot and seed the shared transform cache once."""
    signature = request_signature(username, auth_method, start_date)
    snapshot = load_raw_snapshot(signature, start_date)
    if snapshot is None:
        raise FileNotFoundError("The refreshed ReportData snapshot could not be loaded.")

    raw_df, metadata, snapshot_signature = snapshot
    cached_prepare_long_data.clear()
    build_pivot_table.clear()
    long_df = cached_prepare_long_data(raw_df)

    set_loaded_raw_state(raw_df, metadata, snapshot_signature)
    prepare_signature = {
        **signature,
        "prepare_version": "atlasflow_dynamic_pivot_v3_oil_stats",
    }
    set_loaded_long_state(long_df, prepare_signature)

    active_metadata = st.session_state.get("loaded_metadata")
    if isinstance(active_metadata, dict):
        active_metadata["long_rows"] = int(len(long_df))
        active_metadata["available_variables"] = (
            int(long_df["ValueDescription"].nunique())
            if "ValueDescription" in long_df.columns
            else 0
        )
        st.session_state["loaded_metadata"] = active_metadata
        metadata = active_metadata

    return raw_df, long_df, metadata


def refresh_all_atlasflow_snapshots(
    username: str,
    password: str,
    token: str,
    auth_method: str,
    start_date: date,
) -> dict[str, dict[str, Any]]:
    """Refresh each API once into an atomic snapshot, then activate the new data."""
    results: dict[str, dict[str, Any]] = {}

    results["reportdata"] = fetch_report_data_to_snapshot(
        username, password, token, auth_method, start_date
    )
    results["reportpivots"] = fetch_wide_source_to_snapshot(
        "reportpivots", username, password, token, auth_method, start_date
    )
    results["shippivots"] = fetch_wide_source_to_snapshot(
        "shippivots", username, password, token, auth_method, start_date
    )

    cached_fetch_report_data.clear()
    cached_fetch_wide_odata_source.clear()
    clear_wide_source_state("reportpivots")
    clear_wide_source_state("shippivots")
    activate_reportdata_snapshot(username, auth_method, start_date)
    gc.collect()
    return results


def selected_vessel_controls() -> tuple[str, list[str]]:
    group_options = ["Single vessel", "All fleets"] + list(VESSEL_GROUPS.keys())
    selected_group = st.sidebar.selectbox("Fleet group", options=group_options, key="atlas_fleet_group")

    if selected_group == "Single vessel":
        vessel = st.sidebar.selectbox("Vessel to include", options=VESSEL_OPTIONS, key="atlas_single_vessel")
        st.session_state["atlas_last_fleet_group"] = selected_group
        return selected_group, [vessel]

    if selected_group == "All fleets":
        group_vessels = VESSEL_OPTIONS
    else:
        group_vessels = VESSEL_GROUPS[selected_group]

    vessel_key = "atlas_selected_vessels"
    last_group_key = "atlas_last_fleet_group"
    previous_group = st.session_state.get(last_group_key)

    # Streamlit keeps multiselect state after the first render, so changing from
    # Fleet 1 to All fleets could otherwise keep only the old Fleet 1 vessels.
    # Reset to the full new group whenever the fleet-group selector changes.
    if previous_group != selected_group:
        st.session_state[vessel_key] = list(group_vessels)
        st.session_state[last_group_key] = selected_group

    previous_vessels = st.session_state.get(vessel_key, group_vessels)
    if not isinstance(previous_vessels, list):
        previous_vessels = group_vessels
    valid_default_vessels = [vessel for vessel in previous_vessels if vessel in group_vessels]
    if not valid_default_vessels:
        valid_default_vessels = list(group_vessels)
    if valid_default_vessels != previous_vessels:
        st.session_state[vessel_key] = valid_default_vessels

    vessels = st.sidebar.multiselect(
        "Vessels to include",
        options=group_vessels,
        default=valid_default_vessels,
        key=vessel_key,
        help="This controls the displayed datasets only. The API data is loaded broadly.",
    )

    if not vessels:
        st.sidebar.caption("No vessels selected manually, so all vessels in this fleet group are included.")
        vessels = group_vessels

    return selected_group, list(vessels)

def sidebar_refresh_control() -> bool:
    refresh_requested = st.sidebar.button("Refresh all APIs", use_container_width=False)
    if refresh_requested:
        st.session_state["confirm_api_refresh"] = True

    refresh = False
    if st.session_state.get("confirm_api_refresh"):
        metadata = st.session_state.get("loaded_metadata") or {}
        last_load = metadata.get("loaded_at_local") or metadata.get("loaded_at_utc") or "-"
        last_load_display = str(last_load).replace(" EEST", "").replace(" EET", "")
        st.sidebar.warning(
            "Refresh will call all Atlas Flow APIs and may take a while.\n\n"
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


def add_months(base_date: date, months: int) -> date:
    month_index = base_date.month - 1 + months
    year = base_date.year + month_index // 12
    month = month_index % 12 + 1
    days_in_month = [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    day = min(base_date.day, days_in_month[month - 1])
    return date(year, month, day)


def first_day_of_month(value: date) -> date:
    return date(value.year, value.month, 1)


def last_day_previous_month(value: date) -> date:
    return first_day_of_month(value) - timedelta(days=1)


def clamp_period(start_value: date, end_value: date, min_date: date, max_date: date) -> tuple[date, date]:
    start_value = max(start_value, min_date)
    end_value = min(end_value, max_date)
    if start_value > end_value:
        return min_date, max_date
    return start_value, end_value


def dynamic_period_dates(period_label: str, min_date: date, max_date: date) -> tuple[date, date]:
    anchor_date = min(max_date, date.today())

    if period_label == "YTD":
        start_value = date(anchor_date.year, 1, 1)
        end_value = anchor_date
    elif period_label == "Previous month":
        end_value = last_day_previous_month(anchor_date)
        start_value = first_day_of_month(end_value)
    elif period_label == "Year to previous month":
        end_value = last_day_previous_month(anchor_date)
        start_value = date(end_value.year, 1, 1)
    elif period_label == "Previous 3 months":
        end_value = anchor_date
        start_value = add_months(anchor_date, -3) + timedelta(days=1)
    elif period_label == "Previous 6 months":
        end_value = anchor_date
        start_value = add_months(anchor_date, -6) + timedelta(days=1)
    elif period_label == "Previous 12 months":
        end_value = anchor_date
        start_value = add_months(anchor_date, -12) + timedelta(days=1)
    else:
        start_value = min_date
        end_value = max_date

    return clamp_period(start_value, end_value, min_date, max_date)


def render_date_slicer(df: pd.DataFrame) -> tuple[date, date]:
    min_date, max_date = dataframe_date_window(df)
    st.sidebar.markdown("### Period")
    if min_date >= max_date:
        st.sidebar.caption(f"Available data period: {min_date.strftime('%d/%m/%Y')}")
        return min_date, max_date

    period_mode = st.sidebar.selectbox(
        "Period preset",
        options=[
            "Custom range",
            "YTD",
            "Previous month",
            "Year to previous month",
            "Previous 3 months",
            "Previous 6 months",
            "Previous 12 months",
            "Full available period",
        ],
        index=0,
        key="atlas_period_preset",
        help="Use a dynamic preset or choose Custom range to control the period manually with the slider.",
    )

    if period_mode == "Custom range":
        selected_start, selected_end = st.sidebar.slider(
            "Report period",
            min_value=min_date,
            max_value=max_date,
            value=(min_date, max_date),
            format="DD/MM/YYYY",
            key="atlas_period_slicer",
        )
    else:
        selected_start, selected_end = dynamic_period_dates(period_mode, min_date, max_date)
        st.sidebar.caption(
            f"Selected period: {selected_start.strftime('%d/%m/%Y')} to {selected_end.strftime('%d/%m/%Y')}"
        )

    return selected_start, selected_end




# =============================================================================
# Streamlit Cloud-safe snapshot refresh helpers
# =============================================================================


def normalize_snapshot_values(df: pd.DataFrame) -> pd.DataFrame:
    """Store raw API values as nullable strings to keep Parquet schemas stable page by page."""
    if df.empty:
        return df.copy()
    safe_df = df.copy()
    for column in safe_df.columns:
        safe_df[column] = safe_df[column].astype("string")
    return safe_df


def write_parquet_pages(page_frames: list[pd.DataFrame], output_file: Path) -> int:
    """Write already-normalized page frames into one Parquet file with a stable schema."""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    writer = None
    row_count = 0
    try:
        for frame in page_frames:
            if frame.empty:
                continue
            normalized = normalize_snapshot_values(frame)
            table = pa.Table.from_pandas(normalized, preserve_index=False)
            if writer is None:
                writer = pq.ParquetWriter(output_file, table.schema, compression="zstd")
            writer.write_table(table)
            row_count += len(normalized)
    finally:
        if writer is not None:
            writer.close()
    return row_count


def fetch_report_data_to_snapshot(
    username: str,
    password: str,
    token: str,
    auth_method: str,
    start_date: date,
) -> dict[str, Any]:
    """Fetch ReportData page-by-page and write to Parquet without keeping the full dataset in memory."""
    started_at = time.perf_counter()
    next_url = build_odata_url(start_date)
    first_url = next_url
    seen_urls: set[str] = set()
    pages = 0
    total_bytes = 0
    scanned_rows = 0
    kept_rows_total = 0
    consecutive_empty_pages = 0
    paging_stop_reason = "max_page_limit"
    # Use a unique temporary file per warmup request. Streamlit Cloud can run
    # multiple warmup/browser sessions at the same time; a fixed .tmp.parquet
    # name can be deleted by another session before this request reaches replace().
    tmp_file = RAW_SNAPSHOT_FILE.with_name(
        f"{RAW_SNAPSHOT_FILE.stem}.{os.getpid()}.{int(time.time() * 1000)}.tmp.parquet"
    )
    tmp_file.parent.mkdir(parents=True, exist_ok=True)
    if tmp_file.exists():
        tmp_file.unlink()

    auth = request_auth(username, password, auth_method)
    headers = request_headers(token, auth_method)
    writer = None

    try:
        with requests.Session() as session:
            session.headers.update(headers)
            for _ in range(MAX_ODATA_PAGES):
                if next_url in seen_urls:
                    paging_stop_reason = "repeated_current_url"
                    break
                seen_urls.add(next_url)

                response = request_with_retry(session, next_url, auth=auth, timeout=90)
                total_bytes += len(response.content)
                response.raise_for_status()
                pages += 1

                page_rows, next_link = extract_odata_page(response.json())
                scanned_rows += len(page_rows)
                compact_rows = compact_odata_rows(page_rows)
                if compact_rows:
                    page_df = normalize_snapshot_values(rows_to_dataframe(compact_rows))
                    table = pa.Table.from_pandas(page_df, preserve_index=False)
                    if writer is None:
                        writer = pq.ParquetWriter(tmp_file, table.schema, compression="zstd")
                    writer.write_table(table)
                    kept_rows_total += len(page_df)
                    del page_df, table

                consecutive_empty_pages = consecutive_empty_pages + 1 if len(page_rows) == 0 else 0
                del page_rows, compact_rows
                gc.collect()

                should_continue, resolved_next_url, stop_reason = should_continue_odata_paging(
                    current_url=next_url,
                    next_link=next_link,
                    seen_urls=seen_urls,
                    consecutive_empty_pages=consecutive_empty_pages,
                )
                if not should_continue:
                    paging_stop_reason = stop_reason
                    break
                next_url = resolved_next_url or next_url
    finally:
        if writer is not None:
            writer.close()

    RAW_SNAPSHOT_FILE.parent.mkdir(parents=True, exist_ok=True)

    if kept_rows_total == 0:
        empty_df = normalize_snapshot_values(pd.DataFrame(columns=SOURCE_COLUMNS))
        empty_table = pa.Table.from_pandas(empty_df, preserve_index=False)
        pq.write_table(empty_table, tmp_file, compression="zstd")
        del empty_df, empty_table

    if not tmp_file.exists():
        raise FileNotFoundError(f"ReportData snapshot temporary file was not created: {tmp_file}")

    tmp_file.replace(RAW_SNAPSHOT_FILE)
    loaded_at_utc = datetime.now(timezone.utc)
    metadata = {
        "loaded_at_utc": loaded_at_utc.strftime("%d-%m-%Y %H:%M:%S UTC"),
        "loaded_at_local": local_time_label(loaded_at_utc),
        "rows": kept_rows_total,
        "kept_rows": kept_rows_total,
        "scanned_rows": scanned_rows,
        "discarded_rows": max(scanned_rows - kept_rows_total, 0),
        "pages": pages,
        "downloaded_mb": round(total_bytes / 1024 / 1024, 2),
        "fetch_seconds": round(time.perf_counter() - started_at, 2),
        "first_url": first_url,
        "hit_page_limit": pages >= MAX_ODATA_PAGES and paging_stop_reason == "max_page_limit",
        "paging_stop_reason": paging_stop_reason,
        "max_pages": MAX_ODATA_PAGES,
        "loaded_start_date": start_date.isoformat(),
        "snapshot_format": "parquet",
        "reportdata_mode": "atlasflow_consumption_oil_stats_whitelist",
        "value_description_whitelist_count": len(REPORTDATA_VALUE_WHITELIST),
    }
    signature = request_signature(username, auth_method, start_date)
    snapshot_payload = {
        "metadata": metadata,
        "signature": signature,
        "saved_at_utc": datetime.now(timezone.utc).strftime("%d-%m-%Y %H:%M:%S UTC"),
    }
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    METADATA_SNAPSHOT_FILE.write_text(json.dumps(snapshot_payload, indent=2, default=str), encoding="utf-8")
    return metadata


def fetch_wide_source_to_snapshot(
    source_key: str,
    username: str,
    password: str,
    token: str,
    auth_method: str,
    start_date: date,
) -> dict[str, Any]:
    """Fetch a wide OData source page-by-page and write to Parquet safely."""
    config = SOURCE_CONFIGS[source_key]
    endpoint = str(config["endpoint"])
    datetime_column = str(config.get("datetime_candidates", ["DateTime"])[0])
    next_url = build_wide_odata_url(endpoint, start_date, datetime_column)
    first_url = next_url
    seen_urls: set[str] = set()
    pages = 0
    total_bytes = 0
    row_count = 0
    consecutive_empty_pages = 0
    paging_stop_reason = "max_page_limit"
    all_columns: list[str] = []
    target_file = Path(config["snapshot_file"])
    # Use a unique temporary file per warmup request to avoid cross-session
    # collisions when several source warmups or browser tabs run concurrently.
    tmp_file = target_file.with_name(
        f"{target_file.stem}.{os.getpid()}.{int(time.time() * 1000)}.tmp.parquet"
    )
    tmp_file.parent.mkdir(parents=True, exist_ok=True)
    if tmp_file.exists():
        tmp_file.unlink()

    auth = request_auth(username, password, auth_method)
    headers = request_headers(token, auth_method)
    writer = None
    started_at = time.perf_counter()

    try:
        with requests.Session() as session:
            session.headers.update(headers)
            for _ in range(MAX_ODATA_PAGES):
                if next_url in seen_urls:
                    paging_stop_reason = "repeated_current_url"
                    break
                seen_urls.add(next_url)
                response = request_with_retry(session, next_url, auth=auth, timeout=90)
                total_bytes += len(response.content)
                response.raise_for_status()
                pages += 1

                page_rows, next_link = extract_odata_page(response.json())
                page_df = pd.DataFrame(page_rows)
                if "__metadata" in page_df.columns:
                    page_df = page_df.drop(columns=["__metadata"])

                if not page_df.empty:
                    if not all_columns:
                        all_columns = list(page_df.columns)
                    else:
                        for column in page_df.columns:
                            if column not in all_columns:
                                all_columns.append(column)
                        for column in all_columns:
                            if column not in page_df.columns:
                                page_df[column] = pd.NA
                        page_df = page_df[all_columns]

                    page_df = normalize_snapshot_values(page_df)
                    table = pa.Table.from_pandas(page_df, preserve_index=False)
                    if writer is None:
                        writer = pq.ParquetWriter(tmp_file, table.schema, compression="zstd")
                    else:
                        # New columns after the first page are rare; keep only first-page schema to avoid schema errors.
                        schema_names = writer.schema.names
                        page_df = page_df[[column for column in schema_names if column in page_df.columns]]
                        for column in schema_names:
                            if column not in page_df.columns:
                                page_df[column] = pd.NA
                        page_df = page_df[schema_names]
                        table = pa.Table.from_pandas(page_df, schema=writer.schema, preserve_index=False)
                    writer.write_table(table)
                    row_count += len(page_df)
                    del page_df, table

                consecutive_empty_pages = consecutive_empty_pages + 1 if len(page_rows) == 0 else 0
                del page_rows
                gc.collect()

                should_continue, resolved_next_url, stop_reason = should_continue_odata_paging(
                    current_url=next_url,
                    next_link=next_link,
                    seen_urls=seen_urls,
                    consecutive_empty_pages=consecutive_empty_pages,
                )
                if not should_continue:
                    paging_stop_reason = stop_reason
                    break
                next_url = resolved_next_url or next_url
    finally:
        if writer is not None:
            writer.close()

    target_file.parent.mkdir(parents=True, exist_ok=True)

    if row_count == 0:
        # Some wide endpoints can legally return no rows for the selected API window.
        # In that case, create an empty snapshot with discovered columns if possible.
        empty_columns = all_columns if all_columns else ["NoData"]
        empty_df = normalize_snapshot_values(pd.DataFrame(columns=empty_columns))
        empty_df.to_parquet(tmp_file, index=False, compression="zstd")
        del empty_df

    if not tmp_file.exists():
        raise FileNotFoundError(f"{config['label']} snapshot temporary file was not created: {tmp_file}")

    # Validate before replacing the previous good snapshot. Do not accept a placeholder
    # NoData file when API rows were written.
    try:
        parquet_columns = pq.ParquetFile(tmp_file).schema.names
    except Exception as exc:
        raise RuntimeError(f"{config['label']} snapshot validation failed before save: {exc}") from exc
    if row_count > 0 and parquet_columns == ["NoData"]:
        raise RuntimeError(f"{config['label']} snapshot validation failed: placeholder NoData file for {row_count:,} rows.")

    if target_file.exists():
        target_file.unlink()
    os.replace(str(tmp_file), str(target_file))
    loaded_at_utc = datetime.now(timezone.utc)
    metadata = {
        "source": config["label"],
        "endpoint": endpoint,
        "loaded_at_utc": loaded_at_utc.strftime("%d-%m-%Y %H:%M:%S UTC"),
        "loaded_at_local": local_time_label(loaded_at_utc),
        "loaded_start_date": start_date.isoformat(),
        "rows": int(row_count),
        "columns": int(len(all_columns)),
        "pages": pages,
        "downloaded_mb": round(total_bytes / 1024 / 1024, 2),
        "fetch_seconds": round(time.perf_counter() - started_at, 2),
        "first_url": first_url,
        "hit_page_limit": pages >= MAX_ODATA_PAGES and paging_stop_reason == "max_page_limit",
        "paging_stop_reason": paging_stop_reason,
        "max_pages": MAX_ODATA_PAGES,
        "snapshot_format": "parquet",
    }
    signature = source_signature(source_key, username, auth_method, start_date)
    payload = {
        "metadata": metadata,
        "signature": signature,
        "saved_at_utc": datetime.now(timezone.utc).strftime("%d-%m-%Y %H:%M:%S UTC"),
    }
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    Path(config["metadata_file"]).write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    return metadata




# =============================================================================
# Persistent prepared multi-source snapshots and incremental refresh
# =============================================================================


class AtlasRefreshAlreadyRunning(RuntimeError):
    pass


def read_int_secret(
    name: str,
    default: int,
    *,
    minimum: int = 0,
    maximum: int = 10000,
) -> int:
    try:
        value = int(read_secret(name, str(default)))
    except (TypeError, ValueError):
        value = default
    return min(max(value, minimum), maximum)


def source_secret_name(source_key: str, suffix: str) -> str:
    return f"ATLASFLOW_{source_key.upper()}_{suffix}"


def source_primary_datetime_column(source_key: str) -> str:
    if source_key == "reportdata":
        return "StartDateTimeGMT"
    candidates = list(SOURCE_CONFIGS[source_key].get("datetime_candidates", []))
    return str(candidates[0] if candidates else "DateTime")


def source_manifest_path(source_key: str) -> Path:
    return SNAPSHOT_DIR / f"{source_key}_prepared_manifest.json"


def source_lock_path(source_key: str) -> Path:
    return SNAPSHOT_DIR / f"{source_key}_refresh.lock"


def source_status_path(source_key: str) -> Path:
    return SNAPSHOT_DIR / f"{source_key}_refresh_status.json"


def source_snapshot_path(source_key: str, generation: str) -> Path:
    return SNAPSHOT_DIR / f"{source_key}_prepared_{generation}.parquet"


def source_data_signature(source_key: str) -> str:
    if source_key == "reportdata":
        signature_text = "|".join(
            [
                ATLAS_SNAPSHOT_SCHEMA_VERSION,
                ATLAS_PREPARE_VERSION,
                *REPORTDATA_VALUE_WHITELIST,
                *SOURCE_COLUMNS,
            ]
        )
    else:
        config = SOURCE_CONFIGS[source_key]
        signature_text = "|".join(
            [
                ATLAS_SNAPSHOT_SCHEMA_VERSION,
                source_key,
                str(config["endpoint"]),
                *map(str, config.get("datetime_candidates", [])),
            ]
        )
    return sha256(signature_text.encode("utf-8")).hexdigest()[:16]


def atlas_source_signature(
    source_key: str,
    username: str,
    auth_method: str,
    start_date: date,
) -> dict[str, Any]:
    config = SOURCE_CONFIGS[source_key]
    return {
        "source": source_key,
        "endpoint": str(config["endpoint"]),
        "username_hash": sha256(username.encode("utf-8")).hexdigest()[:12],
        "auth_method": auth_method.lower(),
        "start_date": start_date.isoformat(),
        "data_signature": source_data_signature(source_key),
    }


def source_signature_covers_request(
    stored_signature: dict[str, Any] | None,
    requested_signature: dict[str, Any] | None,
    metadata: dict[str, Any] | None,
    requested_start_date: date,
) -> bool:
    if not stored_signature or not requested_signature or not metadata:
        return False
    for key in ["endpoint", "username_hash", "auth_method"]:
        if stored_signature.get(key) != requested_signature.get(key):
            return False
    source_key = str(requested_signature.get("source") or stored_signature.get("source") or "reportdata")
    expected_data_signature = source_data_signature(source_key)
    if stored_signature.get("data_signature") != expected_data_signature:
        return False
    loaded_start_text = metadata.get("loaded_start_date") or stored_signature.get("start_date")
    try:
        loaded_start_date = date.fromisoformat(str(loaded_start_text))
    except ValueError:
        return False
    return loaded_start_date <= requested_start_date


def _atomic_write_text(path: Path, text_value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(
        f"{path.name}.{os.getpid()}.{int(time.time() * 1000)}.tmp"
    )
    try:
        temp_path.write_text(text_value, encoding="utf-8")
        os.replace(str(temp_path), str(path))
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass


def read_source_refresh_status(source_key: str) -> dict[str, Any] | None:
    try:
        path = source_status_path(source_key)
        if not path.is_file():
            return None
        payload = json.loads(path.read_text(encoding="utf-8"))
        return payload if isinstance(payload, dict) else None
    except Exception:
        return None


def update_source_refresh_status(source_key: str, **updates: Any) -> None:
    payload = read_source_refresh_status(source_key) or {}
    payload.update(updates)
    payload["source"] = source_key
    payload["updated_at_utc"] = datetime.now(timezone.utc).isoformat()
    payload.setdefault("pid", os.getpid())
    try:
        _atomic_write_text(
            source_status_path(source_key),
            json.dumps(payload, indent=2, default=str),
        )
    except Exception:
        return


def source_refresh_status_summary(source_key: str) -> str:
    status = read_source_refresh_status(source_key) or {}
    stage = str(status.get("stage", "refreshing"))
    refresh_mode = str(status.get("refresh_mode", "refresh"))
    chunk_index = int(status.get("chunk_index", 0) or 0)
    chunks_total = int(status.get("chunks_total", 0) or 0)
    chunk_start = status.get("chunk_start_date")
    chunk_end = status.get("chunk_end_date_exclusive")
    parts = [f"{source_key}: {refresh_mode} {stage}"]
    if chunk_index and chunks_total:
        parts.append(f"window {chunk_index} of {chunks_total}")
    if chunk_start and chunk_end:
        parts.append(f"{chunk_start} to {chunk_end}")
    return "; ".join(parts)


@contextmanager
def source_refresh_lock(source_key: str) -> Any:
    """Prevent duplicate refreshes of the same API source."""
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    lock_path = source_lock_path(source_key)

    if fcntl is not None:
        handle = lock_path.open("a+", encoding="utf-8")
        try:
            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError:
                yield False
                return
            handle.seek(0)
            handle.truncate()
            handle.write(
                json.dumps(
                    {
                        "source": source_key,
                        "pid": os.getpid(),
                        "started_at_utc": datetime.now(timezone.utc).isoformat(),
                    }
                )
            )
            handle.flush()
            yield True
        finally:
            try:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
            except Exception:
                pass
            handle.close()
        return

    lock_fd: int | None = None
    try:
        try:
            lock_fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            yield False
            return
        os.write(
            lock_fd,
            json.dumps(
                {
                    "source": source_key,
                    "pid": os.getpid(),
                    "started_at_utc": datetime.now(timezone.utc).isoformat(),
                }
            ).encode("utf-8"),
        )
        yield True
    finally:
        if lock_fd is not None:
            os.close(lock_fd)
            try:
                lock_path.unlink()
            except FileNotFoundError:
                pass


def read_source_manifest(source_key: str) -> dict[str, Any] | None:
    try:
        path = source_manifest_path(source_key)
        if not path.is_file():
            return None
        payload = json.loads(path.read_text(encoding="utf-8"))
        return payload if isinstance(payload, dict) else None
    except Exception:
        return None


def source_manifest_is_valid(
    source_key: str,
    manifest: dict[str, Any] | None,
    requested_signature: dict[str, Any],
    requested_start_date: date,
) -> bool:
    if not manifest:
        return False
    if manifest.get("snapshot_schema_version") != ATLAS_SNAPSHOT_SCHEMA_VERSION:
        return False
    if manifest.get("source") != source_key:
        return False
    metadata = manifest.get("metadata") or {}
    stored_signature = manifest.get("signature") or {}
    if not source_signature_covers_request(
        stored_signature,
        requested_signature,
        metadata,
        requested_start_date,
    ):
        return False
    snapshot_file = SNAPSHOT_DIR / str(manifest.get("prepared_file", ""))
    return bool(manifest.get("generation")) and snapshot_file.is_file()


@st.cache_data(show_spinner=False)
def cached_read_prepared_source_snapshot(
    source_key: str,
    generation: str,
    snapshot_file: str,
) -> pd.DataFrame:
    del source_key, generation  # Deliberate cache keys.
    return pd.read_parquet(snapshot_file)


def load_source_snapshot(
    source_key: str,
    requested_signature: dict[str, Any],
    requested_start_date: date,
) -> tuple[pd.DataFrame, dict[str, Any], dict[str, Any]] | None:
    # Older call sites do not include source/data_signature in the requested signature.
    requested_signature = dict(requested_signature or {})
    requested_signature.setdefault("source", source_key)
    requested_signature.setdefault("data_signature", source_data_signature(source_key))
    manifest = read_source_manifest(source_key)
    if not source_manifest_is_valid(
        source_key,
        manifest,
        requested_signature,
        requested_start_date,
    ):
        return None
    assert manifest is not None
    generation = str(manifest["generation"])
    snapshot_path = SNAPSHOT_DIR / str(manifest["prepared_file"])
    try:
        df = cached_read_prepared_source_snapshot(
            source_key,
            generation,
            str(snapshot_path),
        )
    except Exception:
        return None
    if not isinstance(df, pd.DataFrame):
        return None
    metadata = dict(manifest.get("metadata") or {})
    metadata["loaded_from_snapshot"] = True
    metadata["snapshot_generation"] = generation
    metadata.setdefault("snapshot_saved_at_utc", manifest.get("saved_at_utc", "-"))
    metadata.setdefault("snapshot_schema_version", ATLAS_SNAPSHOT_SCHEMA_VERSION)
    return df, metadata, dict(manifest.get("signature") or {})


def source_snapshot_info(
    source_key: str,
    username: str,
    auth_method: str,
    start_date: date,
) -> tuple[dict[str, Any], dict[str, Any]] | None:
    requested_signature = atlas_source_signature(
        source_key,
        username,
        auth_method,
        start_date,
    )
    manifest = read_source_manifest(source_key)
    if not source_manifest_is_valid(
        source_key,
        manifest,
        requested_signature,
        start_date,
    ):
        return None
    assert manifest is not None
    return dict(manifest.get("metadata") or {}), manifest


def build_refresh_windows(
    start_date: date,
    end_date_exclusive: date,
    chunk_days: int,
) -> list[tuple[date, date]]:
    if start_date >= end_date_exclusive:
        return []
    windows: list[tuple[date, date]] = []
    cursor = start_date
    while cursor < end_date_exclusive:
        next_cursor = min(cursor + timedelta(days=chunk_days), end_date_exclusive)
        windows.append((cursor, next_cursor))
        cursor = next_cursor
    return windows


def build_source_window_url(
    source_key: str,
    window_start: date,
    window_end_exclusive: date,
) -> str:
    config = SOURCE_CONFIGS[source_key]
    datetime_column = source_primary_datetime_column(source_key)
    # Query one day earlier because Marorka's OData V1 filter uses strict gt.
    query_start = window_start - timedelta(days=1)
    filter_text = (
        f"{datetime_column} gt DateTime'{query_start.isoformat()}' and "
        f"{datetime_column} lt DateTime'{window_end_exclusive.isoformat()}'"
    )
    params: dict[str, str] = {"$filter": filter_text}
    if source_key == "reportdata":
        params["$select"] = ",".join(SOURCE_COLUMNS)
    return f"{config['endpoint']}?{urlencode(params)}"


def trim_frame_to_window(
    df: pd.DataFrame,
    datetime_column: str,
    window_start: date,
    window_end_exclusive: date,
) -> pd.DataFrame:
    if df.empty or datetime_column not in df.columns:
        return df.iloc[0:0].copy()
    values = parse_datetime_series(df[datetime_column])
    start_ts = pd.Timestamp(window_start, tz="UTC")
    end_ts = pd.Timestamp(window_end_exclusive, tz="UTC")
    return df.loc[values.ge(start_ts) & values.lt(end_ts)].copy()


def prepare_reportdata_snapshot_frame(df: pd.DataFrame) -> pd.DataFrame:
    for column in SOURCE_COLUMNS:
        if column not in df.columns:
            df[column] = pd.NA
    prepared = df[SOURCE_COLUMNS].copy()
    prepared["ReportId"] = pd.to_numeric(prepared["ReportId"], errors="coerce").astype("Int64")
    prepared["StartDateTimeGMT"] = parse_datetime_series(prepared["StartDateTimeGMT"])
    prepared["EndDateTimeGMT"] = parse_datetime_series(prepared["EndDateTimeGMT"])
    prepared["LapTime"] = parse_numeric_series(prepared["LapTime"])
    prepared["ParsedValue"] = parse_numeric_series(prepared["ReportedValue"])
    prepared = prepared[
        prepared["ValueDescription"].notna()
        & ~prepared["ReportType"].isin(EXCLUDED_REPORT_TYPES)
    ].copy()
    for column in ["ShipName", "ReportType", "StateName", "ValueDescription", "ReportedValue"]:
        prepared[column] = prepared[column].astype("string")
    return prepared[[*SOURCE_COLUMNS, "ParsedValue"]]


def normalize_wide_snapshot_frame(source_key: str, df: pd.DataFrame) -> pd.DataFrame:
    prepared = df.copy()
    prepared.columns = [str(column) for column in prepared.columns]
    if "__metadata" in prepared.columns:
        prepared = prepared.drop(columns=["__metadata"])
    datetime_column = source_primary_datetime_column(source_key)
    for column in prepared.columns:
        if column == datetime_column:
            prepared[column] = parse_datetime_series(prepared[column])
        else:
            prepared[column] = prepared[column].astype("string")
    return prepared


def normalize_source_snapshot_frame(source_key: str, df: pd.DataFrame) -> pd.DataFrame:
    if source_key == "reportdata":
        return prepare_reportdata_snapshot_frame(df)
    return normalize_wide_snapshot_frame(source_key, df)


def deduplicate_source_window(source_key: str, df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    if source_key == "reportdata":
        report_ids = df["ReportId"].astype("string").fillna("")
        value_keys = df["ValueDescription"].map(normalize_text)
        with_id = df.loc[report_ids.str.len().gt(0)].copy()
        if not with_id.empty:
            with_id["_rid"] = report_ids.loc[with_id.index]
            with_id["_value_key"] = value_keys.loc[with_id.index]
            with_id = with_id.drop_duplicates(["_rid", "_value_key"], keep="last").drop(columns=["_rid", "_value_key"])
        without_id = df.loc[report_ids.str.len().eq(0)].drop_duplicates(keep="last")
        return pd.concat([with_id, without_id], ignore_index=True)
    if "ReportId" in df.columns and df["ReportId"].notna().any():
        return df.drop_duplicates(["ReportId"], keep="last").reset_index(drop=True)
    datetime_column = source_primary_datetime_column(source_key)
    keys = [column for column in ["ShipName", datetime_column] if column in df.columns]
    if len(keys) == 2:
        return df.drop_duplicates(keys, keep="last").reset_index(drop=True)
    return df.drop_duplicates(keep="last").reset_index(drop=True)


def fetch_source_window(
    source_key: str,
    username: str,
    password: str,
    token: str,
    auth_method: str,
    window_start: date,
    window_end_exclusive: date,
    *,
    deadline_monotonic: float,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    started_at = time.perf_counter()
    next_url = build_source_window_url(source_key, window_start, window_end_exclusive)
    first_url = next_url
    seen_urls: set[str] = set()
    frames: list[pd.DataFrame] = []
    pages = 0
    total_bytes = 0
    scanned_rows = 0
    consecutive_empty_pages = 0
    paging_stop_reason = "max_page_limit"
    auth = request_auth(username, password, auth_method)
    headers = request_headers(token, auth_method)
    datetime_column = source_primary_datetime_column(source_key)

    with requests.Session() as session:
        session.headers.update(headers)
        for _ in range(MAX_ODATA_PAGES):
            if time.perf_counter() >= deadline_monotonic:
                raise TimeoutError(f"{SOURCE_CONFIGS[source_key]['label']} refresh exceeded its safety time limit.")
            if next_url in seen_urls:
                paging_stop_reason = "repeated_current_url"
                break
            seen_urls.add(next_url)
            response = request_with_retry(
                session,
                next_url,
                auth=auth,
                timeout=API_REQUEST_TIMEOUT_SECONDS,
                max_attempts=API_REQUEST_MAX_ATTEMPTS,
            )
            total_bytes += len(response.content)
            response.raise_for_status()
            pages += 1
            page_rows, next_link = extract_odata_page(response.json())
            scanned_rows += len(page_rows)
            consecutive_empty_pages = consecutive_empty_pages + 1 if len(page_rows) == 0 else 0

            if source_key == "reportdata":
                page_rows = compact_odata_rows(page_rows)
            page_df = pd.DataFrame(page_rows)
            if not page_df.empty:
                if source_key == "reportdata":
                    for column in SOURCE_COLUMNS:
                        if column not in page_df.columns:
                            page_df[column] = pd.NA
                    page_df = page_df[SOURCE_COLUMNS]
                page_df = trim_frame_to_window(
                    page_df,
                    datetime_column,
                    window_start,
                    window_end_exclusive,
                )
                if not page_df.empty:
                    frames.append(page_df)

            should_continue, resolved_next_url, stop_reason = should_continue_odata_paging(
                current_url=next_url,
                next_link=next_link,
                seen_urls=seen_urls,
                consecutive_empty_pages=consecutive_empty_pages,
            )
            if not should_continue:
                paging_stop_reason = stop_reason or "end_of_feed"
                break
            next_url = resolved_next_url or next_url

    hit_page_limit = pages >= MAX_ODATA_PAGES and paging_stop_reason == "max_page_limit"
    if hit_page_limit:
        raise RuntimeError(
            f"{SOURCE_CONFIGS[source_key]['label']} reached {MAX_ODATA_PAGES:,} pages inside "
            f"the bounded window {window_start} to {window_end_exclusive}. Reduce the configured chunk size."
        )

    if frames:
        window_df = pd.concat(frames, ignore_index=True, sort=False)
    elif source_key == "reportdata":
        window_df = pd.DataFrame(columns=SOURCE_COLUMNS)
    else:
        window_df = pd.DataFrame()
    window_df = normalize_source_snapshot_frame(source_key, window_df)
    window_df = deduplicate_source_window(source_key, window_df)

    date_values = (
        pd.to_datetime(window_df.get(datetime_column), errors="coerce", utc=True)
        if datetime_column in window_df.columns
        else pd.Series(dtype="datetime64[ns, UTC]")
    )
    metadata = {
        "window_start_date": window_start.isoformat(),
        "window_end_date_exclusive": window_end_exclusive.isoformat(),
        "rows": int(len(window_df)),
        "scanned_rows": int(scanned_rows),
        "pages": int(pages),
        "downloaded_mb": round(total_bytes / 1024 / 1024, 2),
        "fetch_seconds": round(time.perf_counter() - started_at, 2),
        "first_url": first_url,
        "paging_stop_reason": paging_stop_reason,
        "hit_page_limit": False,
        "latest_source_date": date_values.max().date().isoformat() if not date_values.empty and date_values.notna().any() else None,
    }
    return window_df, metadata


def write_temp_chunk(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_parquet(path, index=False, compression="zstd")
    if not path.is_file() or path.stat().st_size <= 0:
        raise RuntimeError(f"Temporary snapshot chunk was not created: {path}")


def collect_source_chunks(
    source_key: str,
    username: str,
    password: str,
    token: str,
    auth_method: str,
    refresh_start_date: date,
    end_date_exclusive: date,
    chunk_days: int,
    max_duration_seconds: int,
    refresh_mode: str,
) -> tuple[Path, list[Path], list[str], dict[str, Any]]:
    work_dir = SNAPSHOT_DIR / f".refresh_{source_key}_{os.getpid()}_{int(time.time() * 1000)}"
    work_dir.mkdir(parents=True, exist_ok=False)
    windows = build_refresh_windows(refresh_start_date, end_date_exclusive, chunk_days)
    deadline = time.perf_counter() + max_duration_seconds
    chunk_files: list[Path] = []
    union_columns: list[str] = []
    total_rows = total_scanned = total_pages = 0
    total_downloaded_mb = total_fetch_seconds = 0.0
    first_url = "-"
    latest_source_date: str | None = None

    for index, (window_start, window_end) in enumerate(windows, start=1):
        update_source_refresh_status(
            source_key,
            state="running",
            stage="fetching",
            refresh_mode=refresh_mode,
            chunk_index=index,
            chunks_total=len(windows),
            chunk_start_date=window_start.isoformat(),
            chunk_end_date_exclusive=window_end.isoformat(),
            pages_completed=total_pages,
            rows_kept=total_rows,
        )
        frame, window_meta = fetch_source_window(
            source_key,
            username,
            password,
            token,
            auth_method,
            window_start,
            window_end,
            deadline_monotonic=deadline,
        )
        chunk_path = work_dir / f"chunk_{index:04d}.parquet"
        write_temp_chunk(frame, chunk_path)
        chunk_files.append(chunk_path)
        for column in frame.columns:
            if column not in union_columns:
                union_columns.append(str(column))
        total_rows += len(frame)
        total_scanned += int(window_meta["scanned_rows"])
        total_pages += int(window_meta["pages"])
        total_downloaded_mb += float(window_meta["downloaded_mb"])
        total_fetch_seconds += float(window_meta["fetch_seconds"])
        if first_url == "-":
            first_url = str(window_meta.get("first_url", "-"))
        latest_value = window_meta.get("latest_source_date")
        if latest_value and (latest_source_date is None or str(latest_value) > latest_source_date):
            latest_source_date = str(latest_value)
        del frame
        gc.collect()

    metadata = {
        "refresh_mode": refresh_mode,
        "refresh_api_start_date": refresh_start_date.isoformat(),
        "refresh_end_date_exclusive": end_date_exclusive.isoformat(),
        "chunk_days": int(chunk_days),
        "chunks_total": len(windows),
        "chunks_completed": len(chunk_files),
        "refresh_rows": int(total_rows),
        "scanned_rows": int(total_scanned),
        "discarded_rows": max(int(total_scanned) - int(total_rows), 0),
        "pages": int(total_pages),
        "downloaded_mb": round(total_downloaded_mb, 2),
        "fetch_seconds": round(total_fetch_seconds, 2),
        "prepare_seconds": 0,
        "first_url": first_url,
        "paging_stop_reason": "all_windows_completed",
        "hit_page_limit": False,
        "max_pages": MAX_ODATA_PAGES,
        "max_pages_per_window": MAX_ODATA_PAGES,
        "latest_source_date": latest_source_date,
    }
    return work_dir, chunk_files, union_columns, metadata


def source_snapshot_columns(source_key: str, existing_path: Path | None, fresh_columns: list[str]) -> list[str]:
    if source_key == "reportdata":
        return [*SOURCE_COLUMNS, "ParsedValue"]
    columns: list[str] = []
    if existing_path is not None and existing_path.is_file():
        try:
            for column in pq.ParquetFile(existing_path).schema.names:
                if column not in columns:
                    columns.append(column)
        except Exception:
            pass
    for column in fresh_columns:
        if column not in columns:
            columns.append(column)
    datetime_column = source_primary_datetime_column(source_key)
    if datetime_column not in columns:
        columns.insert(0, datetime_column)
    return columns or [datetime_column]


def source_arrow_schema(source_key: str, columns: list[str]) -> pa.Schema:
    if source_key == "reportdata":
        float_columns = {"LapTime", "ParsedValue"}
        datetime_columns = {"StartDateTimeGMT", "EndDateTimeGMT"}
        fields = []
        for column in columns:
            if column == "ReportId":
                field_type = pa.int64()
            elif column in float_columns:
                field_type = pa.float64()
            elif column in datetime_columns:
                field_type = pa.timestamp("ns", tz="UTC")
            else:
                field_type = pa.string()
            fields.append(pa.field(column, field_type, nullable=True))
        return pa.schema(fields)

    datetime_column = source_primary_datetime_column(source_key)
    return pa.schema(
        [
            pa.field(
                column,
                pa.timestamp("ns", tz="UTC") if column == datetime_column else pa.string(),
                nullable=True,
            )
            for column in columns
        ]
    )


def align_source_frame(
    source_key: str,
    frame: pd.DataFrame,
    columns: list[str],
) -> pd.DataFrame:
    aligned = frame.copy()
    for column in columns:
        if column not in aligned.columns:
            aligned[column] = pd.NA
    aligned = aligned[columns]
    if source_key == "reportdata":
        aligned["ReportId"] = pd.to_numeric(aligned["ReportId"], errors="coerce").astype("Int64")
        for column in ["StartDateTimeGMT", "EndDateTimeGMT"]:
            aligned[column] = pd.to_datetime(aligned[column], errors="coerce", utc=True)
        for column in ["LapTime", "ParsedValue"]:
            aligned[column] = pd.to_numeric(aligned[column], errors="coerce")
        for column in columns:
            if column not in {"ReportId", "StartDateTimeGMT", "EndDateTimeGMT", "LapTime", "ParsedValue"}:
                aligned[column] = aligned[column].astype("string")
        return aligned

    datetime_column = source_primary_datetime_column(source_key)
    for column in columns:
        if column == datetime_column:
            aligned[column] = pd.to_datetime(aligned[column], errors="coerce", utc=True)
        else:
            aligned[column] = aligned[column].astype("string")
    return aligned


def append_frame_to_parquet_writer(
    writer: pq.ParquetWriter,
    source_key: str,
    frame: pd.DataFrame,
    columns: list[str],
    schema: pa.Schema,
) -> int:
    if frame.empty:
        return 0
    aligned = align_source_frame(source_key, frame, columns)
    table = pa.Table.from_pandas(aligned, schema=schema, preserve_index=False)
    writer.write_table(table)
    return int(len(aligned))


def stream_existing_snapshot_before_cutoff(
    writer: pq.ParquetWriter,
    source_key: str,
    existing_path: Path,
    cutoff_date: date,
    columns: list[str],
    schema: pa.Schema,
) -> int:
    row_count = 0
    datetime_column = source_primary_datetime_column(source_key)
    cutoff_ts = pd.Timestamp(cutoff_date, tz="UTC")
    parquet_file = pq.ParquetFile(existing_path)
    for batch in parquet_file.iter_batches(batch_size=25000):
        frame = batch.to_pandas()
        if datetime_column in frame.columns:
            dates = pd.to_datetime(frame[datetime_column], errors="coerce", utc=True)
            frame = frame.loc[dates.isna() | dates.lt(cutoff_ts)].copy()
        row_count += append_frame_to_parquet_writer(
            writer,
            source_key,
            frame,
            columns,
            schema,
        )
        del frame
        gc.collect()
    return row_count


def stream_chunk_file(
    writer: pq.ParquetWriter,
    source_key: str,
    chunk_path: Path,
    columns: list[str],
    schema: pa.Schema,
) -> int:
    row_count = 0
    parquet_file = pq.ParquetFile(chunk_path)
    for batch in parquet_file.iter_batches(batch_size=25000):
        frame = batch.to_pandas()
        row_count += append_frame_to_parquet_writer(
            writer,
            source_key,
            frame,
            columns,
            schema,
        )
        del frame
        gc.collect()
    return row_count


def source_snapshot_latest_date(source_key: str, snapshot_path: Path) -> date | None:
    datetime_column = source_primary_datetime_column(source_key)
    try:
        parquet_file = pq.ParquetFile(snapshot_path)
        if datetime_column not in parquet_file.schema.names:
            return None
        latest: pd.Timestamp | None = None
        for batch in parquet_file.iter_batches(columns=[datetime_column], batch_size=100000):
            values = pd.to_datetime(batch.column(0).to_pandas(), errors="coerce", utc=True)
            if values.notna().any():
                batch_max = values.max()
                if latest is None or batch_max > latest:
                    latest = batch_max
        return latest.date() if latest is not None else None
    except Exception:
        return None


def snapshot_generation() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ") + f"-{os.getpid()}"


def cleanup_old_source_generations(source_key: str) -> None:
    try:
        files = sorted(
            SNAPSHOT_DIR.glob(f"{source_key}_prepared_*.parquet"),
            key=lambda path: path.stat().st_mtime,
            reverse=True,
        )
        for path in files[ATLAS_SNAPSHOT_GENERATIONS_TO_KEEP:]:
            try:
                path.unlink()
            except OSError:
                pass
        for temp_path in SNAPSHOT_DIR.glob(f".refresh_{source_key}_*"):
            try:
                if temp_path.is_dir() and time.time() - temp_path.stat().st_mtime > 3600:
                    for child in temp_path.iterdir():
                        child.unlink(missing_ok=True)
                    temp_path.rmdir()
            except OSError:
                pass
    except Exception:
        return


def cleanup_legacy_source_files(source_key: str) -> None:
    """Remove superseded fixed-name snapshots only after a new manifest is live."""
    try:
        config = SOURCE_CONFIGS[source_key]
        for key in ["snapshot_file", "metadata_file"]:
            legacy_path = Path(config[key])
            if legacy_path.is_file():
                legacy_path.unlink()
    except OSError:
        pass


def publish_source_snapshot(
    source_key: str,
    existing_manifest: dict[str, Any] | None,
    refresh_start_date: date,
    work_dir: Path,
    chunk_files: list[Path],
    fresh_columns: list[str],
    refresh_metadata: dict[str, Any],
    signature: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    existing_path: Path | None = None
    if existing_manifest is not None:
        candidate = SNAPSHOT_DIR / str(existing_manifest.get("prepared_file", ""))
        if candidate.is_file():
            existing_path = candidate

    columns = source_snapshot_columns(source_key, existing_path, fresh_columns)
    schema = source_arrow_schema(source_key, columns)
    generation = snapshot_generation()
    final_path = source_snapshot_path(source_key, generation)
    temp_path = final_path.with_name(f"{final_path.name}.tmp")
    temp_path.parent.mkdir(parents=True, exist_ok=True)
    writer = pq.ParquetWriter(temp_path, schema, compression="zstd")
    total_rows = 0
    try:
        if existing_path is not None:
            total_rows += stream_existing_snapshot_before_cutoff(
                writer,
                source_key,
                existing_path,
                refresh_start_date,
                columns,
                schema,
            )
        for chunk_path in chunk_files:
            total_rows += stream_chunk_file(
                writer,
                source_key,
                chunk_path,
                columns,
                schema,
            )
    finally:
        writer.close()

    try:
        if not temp_path.is_file() or temp_path.stat().st_size <= 0:
            raise RuntimeError(f"{SOURCE_CONFIGS[source_key]['label']} prepared snapshot was not created.")
        os.replace(str(temp_path), str(final_path))
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass

    latest_date = source_snapshot_latest_date(source_key, final_path)
    loaded_at_utc = datetime.now(timezone.utc)
    metadata = dict(refresh_metadata)
    metadata.update(
        {
            "source": SOURCE_CONFIGS[source_key]["label"],
            "endpoint": str(SOURCE_CONFIGS[source_key]["endpoint"]),
            "loaded_at_utc": loaded_at_utc.strftime("%d-%m-%Y %H:%M:%S UTC"),
            "loaded_at_local": local_time_label(loaded_at_utc),
            "loaded_start_date": API_FULL_START_DATE.isoformat(),
            "rows": int(total_rows),
            "kept_rows": int(total_rows),
            "columns": int(len(columns)),
            "snapshot_generation": generation,
            "snapshot_format": "prepared_parquet",
            "snapshot_schema_version": ATLAS_SNAPSHOT_SCHEMA_VERSION,
            "latest_source_date": latest_date.isoformat() if latest_date else refresh_metadata.get("latest_source_date"),
        }
    )
    saved_at_utc = loaded_at_utc.strftime("%d-%m-%Y %H:%M:%S UTC")
    manifest = {
        "snapshot_schema_version": ATLAS_SNAPSHOT_SCHEMA_VERSION,
        "source": source_key,
        "generation": generation,
        "prepared_file": final_path.name,
        "signature": signature,
        "metadata": metadata,
        "saved_at_utc": saved_at_utc,
    }
    _atomic_write_text(
        source_manifest_path(source_key),
        json.dumps(manifest, indent=2, default=str),
    )

    cached_read_prepared_source_snapshot.clear()
    if source_key == "reportdata":
        cached_prepare_long_data.clear()
        build_pivot_table.clear()
    else:
        clear_wide_source_state(source_key)
    cleanup_old_source_generations(source_key)
    cleanup_legacy_source_files(source_key)
    update_source_refresh_status(
        source_key,
        state="complete",
        stage="published",
        refresh_mode=metadata.get("refresh_mode"),
        rows_kept=total_rows,
        snapshot_generation=generation,
    )
    return metadata, manifest


def remove_refresh_work_dir(work_dir: Path | None) -> None:
    if work_dir is None:
        return
    try:
        if work_dir.is_dir():
            for child in work_dir.iterdir():
                try:
                    child.unlink()
                except OSError:
                    pass
            work_dir.rmdir()
    except OSError:
        pass


def refresh_source_snapshot(
    source_key: str,
    username: str,
    password: str,
    token: str,
    auth_method: str,
    *,
    full_refresh: bool,
    acquire_lock: bool = True,
) -> tuple[dict[str, Any], dict[str, Any]]:
    if source_key not in SOURCE_CONFIGS:
        raise ValueError(f"Unsupported AtlasFlow source: {source_key}")

    requested_signature = atlas_source_signature(
        source_key,
        username,
        auth_method,
        API_FULL_START_DATE,
    )

    def execute_refresh() -> tuple[dict[str, Any], dict[str, Any]]:
        existing_manifest = read_source_manifest(source_key)
        if not source_manifest_is_valid(
            source_key,
            existing_manifest,
            requested_signature,
            API_FULL_START_DATE,
        ):
            existing_manifest = None

        refresh_mode = "full"
        refresh_start_date = API_FULL_START_DATE
        if not full_refresh and existing_manifest is not None:
            existing_path = SNAPSHOT_DIR / str(existing_manifest["prepared_file"])
            latest_date_text = (existing_manifest.get("metadata") or {}).get("latest_source_date")
            try:
                latest_date = date.fromisoformat(str(latest_date_text)) if latest_date_text else None
            except ValueError:
                latest_date = None
            if latest_date is None:
                latest_date = source_snapshot_latest_date(source_key, existing_path)
            if latest_date is not None:
                overlap_days = read_int_secret(
                    source_secret_name(source_key, "INCREMENTAL_OVERLAP_DAYS"),
                    DEFAULT_SOURCE_OVERLAP_DAYS[source_key],
                    minimum=1,
                    maximum=90,
                )
                refresh_start_date = max(API_FULL_START_DATE, latest_date - timedelta(days=overlap_days))
                refresh_mode = "incremental"

        chunk_days = read_int_secret(
            source_secret_name(source_key, "REFRESH_CHUNK_DAYS"),
            DEFAULT_SOURCE_CHUNK_DAYS[source_key],
            minimum=1,
            maximum=62,
        )
        default_minutes = (
            DEFAULT_SOURCE_FULL_REFRESH_MAX_MINUTES[source_key]
            if refresh_mode == "full"
            else DEFAULT_SOURCE_INCREMENTAL_REFRESH_MAX_MINUTES[source_key]
        )
        max_minutes = read_int_secret(
            source_secret_name(
                source_key,
                "FULL_REFRESH_MAX_MINUTES" if refresh_mode == "full" else "INCREMENTAL_REFRESH_MAX_MINUTES",
            ),
            default_minutes,
            minimum=5,
            maximum=720,
        )
        end_date_exclusive = date.today() + timedelta(days=1)
        work_dir: Path | None = None
        try:
            update_source_refresh_status(
                source_key,
                state="running",
                stage="starting",
                refresh_mode=refresh_mode,
                refresh_start_date=refresh_start_date.isoformat(),
                end_date_exclusive=end_date_exclusive.isoformat(),
                chunk_days=chunk_days,
                max_minutes=max_minutes,
            )
            work_dir, chunk_files, fresh_columns, refresh_metadata = collect_source_chunks(
                source_key,
                username,
                password,
                token,
                auth_method,
                refresh_start_date,
                end_date_exclusive,
                chunk_days,
                max_minutes * 60,
                refresh_mode,
            )
            if (
                int(refresh_metadata.get("scanned_rows", 0) or 0) == 0
                or int(refresh_metadata.get("refresh_rows", 0) or 0) == 0
            ):
                if existing_manifest is not None:
                    metadata = dict(existing_manifest.get("metadata") or {})
                    metadata.update(
                        {
                            "refresh_mode": "no_changes",
                            "refresh_api_start_date": refresh_start_date.isoformat(),
                            "refresh_checked_at_local": local_time_label(),
                        }
                    )
                    update_source_refresh_status(source_key, state="complete", stage="no_changes", refresh_mode="no_changes")
                    return metadata, existing_manifest
                raise RuntimeError(f"{SOURCE_CONFIGS[source_key]['label']} returned zero usable rows during initial bootstrap.")

            update_source_refresh_status(
                source_key,
                state="running",
                stage="publishing",
                refresh_mode=refresh_mode,
                pages_completed=int(refresh_metadata.get("pages", 0) or 0),
                rows_kept=int(refresh_metadata.get("refresh_rows", 0) or 0),
            )
            metadata, manifest = publish_source_snapshot(
                source_key,
                existing_manifest if refresh_mode == "incremental" else None,
                refresh_start_date,
                work_dir,
                chunk_files,
                fresh_columns,
                refresh_metadata,
                requested_signature,
            )
            return metadata, manifest
        except Exception as exc:
            update_source_refresh_status(
                source_key,
                state="failed",
                stage="failed",
                refresh_mode=refresh_mode,
                error=str(exc),
            )
            raise
        finally:
            remove_refresh_work_dir(work_dir)

    if not acquire_lock:
        return execute_refresh()

    with source_refresh_lock(source_key) as lock_acquired:
        if not lock_acquired:
            existing = source_snapshot_info(source_key, username, auth_method, API_FULL_START_DATE)
            if existing is not None:
                metadata, manifest = existing
                metadata = dict(metadata)
                metadata["refresh_skipped_due_to_lock"] = True
                metadata["refresh_status"] = source_refresh_status_summary(source_key)
                return metadata, manifest
            raise AtlasRefreshAlreadyRunning(source_refresh_status_summary(source_key))
        return execute_refresh()


def migrate_legacy_source_snapshot(
    source_key: str,
    username: str,
    auth_method: str,
) -> tuple[dict[str, Any], dict[str, Any]] | None:
    config = SOURCE_CONFIGS[source_key]
    legacy_file = Path(config["snapshot_file"])
    legacy_metadata_file = Path(config["metadata_file"])
    if not legacy_file.is_file() or not legacy_metadata_file.is_file():
        return None
    try:
        legacy_df = pd.read_parquet(legacy_file)
        if source_key == "reportdata":
            prepared = prepare_reportdata_snapshot_frame(legacy_df)
        else:
            prepared = normalize_wide_snapshot_frame(source_key, legacy_df)
        if prepared.empty:
            return None
        work_dir = SNAPSHOT_DIR / f".refresh_{source_key}_migration_{os.getpid()}_{int(time.time() * 1000)}"
        work_dir.mkdir(parents=True, exist_ok=False)
        chunk_path = work_dir / "chunk_0001.parquet"
        write_temp_chunk(prepared, chunk_path)
        signature = atlas_source_signature(source_key, username, auth_method, API_FULL_START_DATE)
        metadata = {
            "refresh_mode": "legacy_migration",
            "refresh_api_start_date": API_FULL_START_DATE.isoformat(),
            "refresh_end_date_exclusive": (date.today() + timedelta(days=1)).isoformat(),
            "chunk_days": 0,
            "chunks_total": 1,
            "chunks_completed": 1,
            "refresh_rows": int(len(prepared)),
            "scanned_rows": int(len(prepared)),
            "pages": 0,
            "downloaded_mb": 0,
            "fetch_seconds": 0,
            "first_url": "legacy_snapshot",
            "paging_stop_reason": "legacy_migration",
            "hit_page_limit": False,
            "latest_source_date": None,
        }
        published = publish_source_snapshot(
            source_key,
            None,
            API_FULL_START_DATE,
            work_dir,
            [chunk_path],
            list(prepared.columns),
            metadata,
            signature,
        )
        remove_refresh_work_dir(work_dir)
        return published
    except Exception:
        return None


def ensure_source_snapshot(
    source_key: str,
    username: str,
    auth_method: str,
) -> tuple[dict[str, Any], dict[str, Any]] | None:
    existing = source_snapshot_info(source_key, username, auth_method, API_FULL_START_DATE)
    if existing is not None:
        return existing
    with source_refresh_lock(source_key) as lock_acquired:
        if not lock_acquired:
            return source_snapshot_info(source_key, username, auth_method, API_FULL_START_DATE)
        existing = source_snapshot_info(source_key, username, auth_method, API_FULL_START_DATE)
        if existing is not None:
            return existing
        return migrate_legacy_source_snapshot(source_key, username, auth_method)


def parse_wide_source_datetimes(df: pd.DataFrame) -> pd.DataFrame:
    """Prepared snapshots already type the primary filter datetime column."""
    if df.empty:
        return df
    parsed_df = df
    for column in parsed_df.columns:
        if pd.api.types.is_datetime64_any_dtype(parsed_df[column]):
            continue
        lower = str(column).lower()
        if "datetime" in lower or lower in {"date", "timestamp"}:
            parsed = parse_datetime_series(parsed_df[column])
            if parsed.notna().any():
                if parsed_df is df:
                    parsed_df = df.copy()
                parsed_df[column] = parsed
    return parsed_df


def get_loaded_state() -> tuple[pd.DataFrame | None, pd.DataFrame | None, dict[str, Any] | None]:
    return (
        None,
        st.session_state.get("loaded_long_df"),
        st.session_state.get("loaded_metadata"),
    )


def activate_reportdata_snapshot(
    username: str,
    auth_method: str,
    start_date: date,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    signature = atlas_source_signature("reportdata", username, auth_method, start_date)
    snapshot = load_source_snapshot("reportdata", signature, start_date)
    if snapshot is None:
        raise FileNotFoundError("The prepared ReportData snapshot could not be loaded.")
    long_df, metadata, snapshot_signature = snapshot
    st.session_state.pop("loaded_raw_df", None)
    st.session_state["loaded_long_df"] = long_df
    st.session_state["loaded_metadata"] = dict(metadata)
    st.session_state["loaded_request_signature"] = snapshot_signature
    st.session_state["loaded_prepare_signature"] = source_data_signature("reportdata")
    st.session_state["loaded_reportdata_generation"] = metadata.get("snapshot_generation")
    return pd.DataFrame(), long_df, metadata


def load_or_fetch_source(
    source_key: str,
    username: str,
    password: str,
    token: str,
    auth_method: str,
    start_date: date,
    refresh: bool,
    auto_fetch: bool = False,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    del auto_fetch
    requested_signature = atlas_source_signature(source_key, username, auth_method, start_date)
    state_df_key = f"loaded_{source_key}_df"
    state_meta_key = f"loaded_{source_key}_metadata"
    state_sig_key = f"loaded_{source_key}_signature"
    state_generation_key = f"loaded_{source_key}_generation"

    if refresh:
        refresh_source_snapshot(
            source_key,
            username,
            password,
            token,
            auth_method,
            full_refresh=False,
        )

    manifest = read_source_manifest(source_key)
    current_generation = manifest.get("generation") if isinstance(manifest, dict) else None
    df = st.session_state.get(state_df_key)
    metadata = st.session_state.get(state_meta_key)
    current_signature = st.session_state.get(state_sig_key)
    session_generation = st.session_state.get(state_generation_key)
    session_ready = (
        isinstance(df, pd.DataFrame)
        and isinstance(metadata, dict)
        and source_signature_covers_request(current_signature, requested_signature, metadata, start_date)
        and session_generation == current_generation
    )
    if session_ready:
        return df, metadata

    snapshot = load_source_snapshot(source_key, requested_signature, start_date)
    if snapshot is None:
        migrated = ensure_source_snapshot(source_key, username, auth_method)
        if migrated is not None:
            snapshot = load_source_snapshot(source_key, requested_signature, start_date)
    if snapshot is None:
        config = SOURCE_CONFIGS[source_key]
        return pd.DataFrame(), {
            "source": config["label"],
            "endpoint": str(config["endpoint"]),
            "loaded_at_utc": "-",
            "loaded_at_local": "No prepared snapshot yet",
            "loaded_from_snapshot": False,
            "rows": 0,
            "columns": 0,
            "pages": 0,
            "first_url": "-",
            "needs_warmup": True,
        }

    df, metadata, snapshot_signature = snapshot
    st.session_state[state_df_key] = df
    st.session_state[state_meta_key] = metadata
    st.session_state[state_sig_key] = snapshot_signature
    st.session_state[state_generation_key] = metadata.get("snapshot_generation")
    return df, metadata


def load_raw_snapshot(
    requested_signature: dict[str, Any],
    requested_start_date: date,
) -> tuple[pd.DataFrame, dict[str, Any], dict[str, Any]] | None:
    requested_signature = dict(requested_signature or {})
    requested_signature.setdefault("source", "reportdata")
    requested_signature.setdefault("data_signature", source_data_signature("reportdata"))
    return load_source_snapshot("reportdata", requested_signature, requested_start_date)


def refresh_all_atlasflow_snapshots(
    username: str,
    password: str,
    token: str,
    auth_method: str,
    start_date: date,
) -> dict[str, dict[str, Any]]:
    del start_date
    results: dict[str, dict[str, Any]] = {}
    errors: dict[str, str] = {}
    for source_key in ["reportdata", "reportpivots", "shippivots"]:
        try:
            metadata, _ = refresh_source_snapshot(
                source_key,
                username,
                password,
                token,
                auth_method,
                full_refresh=False,
            )
            results[source_key] = metadata
        except Exception as exc:
            errors[source_key] = str(exc)
        finally:
            gc.collect()
    if errors:
        raise RuntimeError(
            "AtlasFlow source refresh failures: "
            + "; ".join(f"{key}: {value}" for key, value in errors.items())
        )
    cached_fetch_report_data.clear()
    cached_fetch_wide_odata_source.clear()
    activate_reportdata_snapshot(username, auth_method, API_FULL_START_DATE)
    return results



# =============================================================================
# Monthly partitioned storage and cross-source monthly comparison (v2)
# =============================================================================

# The original prepared-snapshot layer is retained for backwards compatibility,
# but this version publishes immutable monthly partitions. Incremental warmups
# rewrite only the months touched by the overlap window. Monthly summaries are
# generated at publish time, so all vessels and full months can be compared
# without opening millions of 15-minute rows in a browser session.
ATLAS_SNAPSHOT_SCHEMA_VERSION = "2026-07-15-monthly-partitioned-comparison-v2"
ATLAS_MONTHLY_COMPARISON_VERSION = "2026-07-15-cross-source-monthly-v1"
ATLAS_RAW_INTERACTIVE_ROW_LIMIT = 50_000
ATLAS_REPORTPIVOTS_INTERACTIVE_ROW_LIMIT = 250_000

COMPARISON_SOURCE_LABELS = {
    "reportdata": "ReportData",
    "reportpivots": "ReportPivots",
    "shippivots": "ShipPivots",
}

MONTHLY_COMPARISON_METRICS: dict[str, dict[str, Any]] = {
    "Average Speed [kn]": {
        "aggregation": "mean",
        "candidates": [
            "GPSSpeed", "GPS Speed [kn]", "GPS Speed", "Speed Over Ground [kn]", "Speed Over Ground", "SOG",
            "LogSpeed", "Log Speed [kn]", "Log Speed", "Speed Through Water [kn]", "Speed Through Water", "STW",
            "Average Speed", "Vessel Speed",
        ],
    },
    "Average Shaft Power [kW]": {
        "aggregation": "mean",
        "candidates": [
            "ShaftPower", "Shaft Power", "Power from Torque Meter [kW]",
            "Total Shaft Power [kW]", "Total Shaft Power [kW] (kW)",
            "ME Power", "Main Engine Power",
        ],
    },
    "Average ME Load [%]": {
        "aggregation": "mean",
        "percentage": True,
        "candidates": [
            "ME Load [%MCR]", "ME Load [% MCR]", "MELoad",
            "ME Load", "Main Engine Load", "MainEngineLoad",
        ],
    },
    "Total ME Fuel [MT]": {
        "aggregation": "sum",
        "candidates": [
            "Main Engine Total Consumed", "ME Total Consumed",
            "MEConsumed", "Main Engine Consumption", "ME Consumption",
        ],
        "component_candidates": ME_FUEL_COLUMNS,
    },
    "Total DG Fuel [MT]": {
        "aggregation": "sum",
        "candidates": [
            "Diesel Generator Total Consumed", "DG Total Consumed",
            "DG Totals Consumed", "DGTotalsConsumed", "DGTotalConsumed",
            "DGConsumed", "Generator Total Consumed",
        ],
        "component_candidates": DG_FUEL_COLUMNS,
    },
    "Total Auxiliary Fuel [MT]": {
        "aggregation": "sum",
        "candidates": [
            "Auxiliary Engine Total Consumed", "Aux Engine Total Consumed",
            "Aux Total Consumed", "AuxConsumed",
        ],
        "component_candidates": AUXILIARY_FUEL_COLUMNS,
    },
    "Total Boiler Fuel [MT]": {
        "aggregation": "sum",
        "candidates": [
            "Boiler Total Consumed", "BoilerConsumed", "Boiler Consumption",
        ],
        "component_candidates": BOILER_FUEL_COLUMNS,
    },
    "Total Fuel [MT]": {
        "aggregation": "sum",
        "candidates": [
            "Total Fuel Consumed", "Total Consumed", "Total Consumption",
            "Bunker Consumption", "Fuel Consumption", "FuelConsumed",
        ],
    },
    "Total Distance [nm]": {
        "aggregation": "sum",
        "candidates": [
            "Distance Over Ground [nm]", "DistanceOverGround", "Distance Over Ground",
            "Engine Distance [nm]", "EngineDistance", "Sailed Distance", "Distance",
        ],
    },
    "Total Running Hours [h]": {
        "aggregation": "sum",
        "candidates": [
            "Steaming Time Since Last Report [hh:mm]",
            "Steaming Time Since Last Report", "RunningHours", "Running Hours",
            "LapTime", "Operating Hours",
        ],
    },
    "Average SFOC [g/kWh]": {
        "aggregation": "mean",
        "candidates": ["SFOC [g/kWh]", "SFOC [gr/Kwh]", "SFOC"],
    },
}


def source_partition_root(source_key: str) -> Path:
    return SNAPSHOT_DIR / "monthly" / source_key


def source_partition_file(source_key: str, month_key: str, generation: str) -> Path:
    return source_partition_root(source_key) / f"data_{month_key.replace('-', '')}_{generation}.parquet"


def source_summary_file(source_key: str, month_key: str, generation: str) -> Path:
    return source_partition_root(source_key) / f"summary_{month_key.replace('-', '')}_{generation}.parquet"


def month_start_from_key(month_key: str) -> date:
    year_text, month_text = month_key.split("-", 1)
    return date(int(year_text), int(month_text), 1)


def next_month_start(value: date) -> date:
    return date(value.year + 1, 1, 1) if value.month == 12 else date(value.year, value.month + 1, 1)


def month_keys_for_range(start_value: date, end_exclusive: date) -> list[str]:
    if start_value >= end_exclusive:
        return []
    cursor = date(start_value.year, start_value.month, 1)
    keys: list[str] = []
    while cursor < end_exclusive:
        keys.append(f"{cursor.year:04d}-{cursor.month:02d}")
        cursor = next_month_start(cursor)
    return keys


def manifest_partitions(manifest: dict[str, Any] | None) -> dict[str, dict[str, Any]]:
    if not isinstance(manifest, dict) or not isinstance(manifest.get("partitions"), dict):
        return {}
    return {
        str(month_key): dict(entry)
        for month_key, entry in manifest["partitions"].items()
        if isinstance(entry, dict)
    }


def partition_entry_path(entry: dict[str, Any], field: str = "file") -> Path:
    return SNAPSHOT_DIR / str(entry.get(field, ""))


def source_manifest_is_valid(
    source_key: str,
    manifest: dict[str, Any] | None,
    requested_signature: dict[str, Any],
    requested_start_date: date,
) -> bool:
    if not manifest:
        return False
    if manifest.get("snapshot_schema_version") != ATLAS_SNAPSHOT_SCHEMA_VERSION:
        return False
    if manifest.get("source") != source_key:
        return False
    metadata = manifest.get("metadata") or {}
    stored_signature = manifest.get("signature") or {}
    if not source_signature_covers_request(
        stored_signature,
        requested_signature,
        metadata,
        requested_start_date,
    ):
        return False
    partitions = manifest_partitions(manifest)
    if not partitions:
        return False
    for entry in partitions.values():
        if not partition_entry_path(entry, "file").is_file():
            return False
        if entry.get("summary_file") and not partition_entry_path(entry, "summary_file").is_file():
            return False
    return True


@st.cache_data(show_spinner=False)
def cached_read_partitioned_source_snapshot(
    source_key: str,
    generation: str,
    partition_files: tuple[str, ...],
) -> pd.DataFrame:
    del source_key, generation
    frames = [pd.read_parquet(file_name) for file_name in partition_files]
    frames = [frame for frame in frames if isinstance(frame, pd.DataFrame) and not frame.empty]
    return pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()


@st.cache_data(show_spinner=False)
def cached_read_monthly_summary_files(
    generation_signature: str,
    summary_files: tuple[str, ...],
) -> pd.DataFrame:
    del generation_signature
    frames = [pd.read_parquet(file_name) for file_name in summary_files]
    frames = [frame for frame in frames if isinstance(frame, pd.DataFrame) and not frame.empty]
    return pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()


def load_source_snapshot(
    source_key: str,
    requested_signature: dict[str, Any],
    requested_start_date: date,
) -> tuple[pd.DataFrame, dict[str, Any], dict[str, Any]] | None:
    requested_signature = dict(requested_signature or {})
    requested_signature.setdefault("source", source_key)
    requested_signature.setdefault("data_signature", source_data_signature(source_key))
    manifest = read_source_manifest(source_key)
    if not source_manifest_is_valid(source_key, manifest, requested_signature, requested_start_date):
        return None
    assert manifest is not None
    partitions = manifest_partitions(manifest)
    total_rows = int((manifest.get("metadata") or {}).get("rows", 0) or 0)
    # Wide sources are intentionally read through load_wide_source_for_view(),
    # which applies predicate pushdown. Refuse a full wide-source materialization.
    if source_key != "reportdata" and total_rows > ATLAS_REPORTPIVOTS_INTERACTIVE_ROW_LIMIT:
        return None
    files = tuple(
        str(partition_entry_path(partitions[month_key], "file"))
        for month_key in sorted(partitions)
    )
    try:
        frame = cached_read_partitioned_source_snapshot(
            source_key,
            str(manifest.get("generation", "")),
            files,
        )
    except Exception:
        return None
    metadata = dict(manifest.get("metadata") or {})
    metadata["loaded_from_snapshot"] = True
    metadata["snapshot_generation"] = manifest.get("generation")
    metadata.setdefault("snapshot_saved_at_utc", manifest.get("saved_at_utc", "-"))
    metadata["partition_count"] = len(partitions)
    return frame, metadata, dict(manifest.get("signature") or {})


def source_snapshot_info(
    source_key: str,
    username: str,
    auth_method: str,
    start_date: date,
) -> tuple[dict[str, Any], dict[str, Any]] | None:
    signature = atlas_source_signature(source_key, username, auth_method, start_date)
    manifest = read_source_manifest(source_key)
    if not source_manifest_is_valid(source_key, manifest, signature, start_date):
        return None
    assert manifest is not None
    return dict(manifest.get("metadata") or {}), manifest


def _candidate_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized_columns = {normalize_text(column): str(column) for column in df.columns}
    for candidate in candidates:
        match = normalized_columns.get(normalize_text(candidate))
        if match is not None:
            return match
    for candidate in candidates:
        candidate_key = normalize_text(candidate)
        if len(candidate_key) < 6:
            continue
        for normalized_column, original_column in normalized_columns.items():
            if candidate_key in normalized_column or normalized_column in candidate_key:
                return original_column
    return None


def _reportdata_metric_rows(
    frame: pd.DataFrame,
    spec: dict[str, Any],
) -> tuple[pd.DataFrame, str | None]:
    if frame.empty or "ValueDescription" not in frame.columns:
        return pd.DataFrame(), None
    description_keys = frame["ValueDescription"].map(normalize_text)
    for candidate in list(spec.get("candidates") or []):
        mask = description_keys.eq(normalize_text(candidate))
        if mask.any():
            selected = frame.loc[mask, ["ShipName", "ReportId", "ParsedValue"]].copy()
            selected["MetricValue"] = pd.to_numeric(selected["ParsedValue"], errors="coerce")
            return selected, candidate
    components = list(spec.get("component_candidates") or [])
    component_keys = {normalize_text(value) for value in components}
    if component_keys:
        mask = description_keys.isin(component_keys)
        if mask.any():
            selected = frame.loc[mask, ["ShipName", "ReportId", "ParsedValue"]].copy()
            selected["MetricValue"] = pd.to_numeric(selected["ParsedValue"], errors="coerce")
            selected = (
                selected.groupby(["ShipName", "ReportId"], dropna=False, as_index=False)["MetricValue"]
                .sum(min_count=1)
            )
            return selected, " + ".join(components)
    return pd.DataFrame(), None


def _period_day_count(month_key: str) -> int:
    month_start = month_start_from_key(month_key)
    end_exclusive = min(next_month_start(month_start), date.today() + timedelta(days=1))
    return max((end_exclusive - month_start).days, 1)


def build_monthly_source_summary(
    source_key: str,
    partition_df: pd.DataFrame,
    month_key: str,
) -> pd.DataFrame:
    if partition_df.empty:
        return pd.DataFrame()
    datetime_column = source_primary_datetime_column(source_key)
    if datetime_column not in partition_df.columns or "ShipName" not in partition_df.columns:
        return pd.DataFrame()
    frame = partition_df.copy()
    frame[datetime_column] = pd.to_datetime(frame[datetime_column], errors="coerce", utc=True)
    frame = frame[frame[datetime_column].notna() & frame["ShipName"].notna()].copy()
    if frame.empty:
        return pd.DataFrame()
    frame["ShipName"] = frame["ShipName"].astype("string")
    frame["_ObservedDate"] = frame[datetime_column].dt.date
    grouped = frame.groupby("ShipName", dropna=False)
    records = grouped["ReportId"].nunique(dropna=True) if source_key == "reportdata" and "ReportId" in frame.columns else grouped.size()
    first_timestamp = grouped[datetime_column].min().reindex(records.index)
    last_timestamp = grouped[datetime_column].max().reindex(records.index)
    observed_days = grouped["_ObservedDate"].nunique(dropna=True).reindex(records.index)
    summary = pd.DataFrame(
        {
            "ShipName": records.index.astype(str),
            "Records": records.to_numpy(),
            "First Timestamp": first_timestamp.to_numpy(),
            "Last Timestamp": last_timestamp.to_numpy(),
            "Observed Days": observed_days.to_numpy(),
        }
    )
    summary.insert(0, "Source", COMPARISON_SOURCE_LABELS[source_key])
    summary.insert(0, "Month", month_key)
    period_days = _period_day_count(month_key)
    summary["Period Days"] = period_days
    summary["Month Complete"] = next_month_start(month_start_from_key(month_key)) <= date.today()
    summary["Day Coverage [%]"] = (
        pd.to_numeric(summary["Observed Days"], errors="coerce") / period_days * 100
    ).clip(upper=100).round(2)
    if source_key == "shippivots":
        unique_timestamps = grouped[datetime_column].nunique(dropna=True).reindex(records.index)
        observation_coverage = pd.Series(
            unique_timestamps.to_numpy() / max(period_days * 24 * 4, 1) * 100,
            index=summary.index,
            dtype="float64",
        )
        summary["Observation Coverage [%]"] = observation_coverage.clip(upper=100).round(2)
    else:
        summary["Observation Coverage [%]"] = pd.NA

    for metric_name, spec in MONTHLY_COMPARISON_METRICS.items():
        mapping_column = f"Mapping: {metric_name}"
        aggregation = str(spec.get("aggregation", "mean"))
        if source_key == "reportdata":
            metric_rows, mapping = _reportdata_metric_rows(frame, spec)
            if metric_rows.empty:
                summary[metric_name] = pd.NA
                summary[mapping_column] = pd.NA
                continue
            metric_group = metric_rows.groupby("ShipName", dropna=False)["MetricValue"]
        else:
            source_column = _candidate_column(frame, list(spec.get("candidates") or []))
            if source_column is None:
                summary[metric_name] = pd.NA
                summary[mapping_column] = pd.NA
                continue
            metric_rows = pd.DataFrame(
                {
                    "ShipName": frame["ShipName"],
                    "MetricValue": pd.to_numeric(frame[source_column], errors="coerce"),
                }
            )
            metric_group = metric_rows.groupby("ShipName", dropna=False)["MetricValue"]
            mapping = source_column
        aggregated = metric_group.sum(min_count=1) if aggregation == "sum" else metric_group.mean()
        values = summary["ShipName"].map(aggregated)
        if spec.get("percentage"):
            numeric = pd.to_numeric(values, errors="coerce")
            non_null = numeric.dropna()
            if not non_null.empty and non_null.abs().median() <= 1.5:
                numeric = numeric * 100
            values = numeric
        summary[metric_name] = pd.to_numeric(values, errors="coerce").round(3)
        summary[mapping_column] = mapping

    component_total = pd.concat(
        [
            pd.to_numeric(summary.get("Total ME Fuel [MT]"), errors="coerce"),
            pd.to_numeric(summary.get("Total DG Fuel [MT]"), errors="coerce"),
            pd.to_numeric(summary.get("Total Auxiliary Fuel [MT]"), errors="coerce"),
            pd.to_numeric(summary.get("Total Boiler Fuel [MT]"), errors="coerce"),
        ],
        axis=1,
    ).sum(axis=1, min_count=1)
    direct_total = pd.to_numeric(summary.get("Total Fuel [MT]"), errors="coerce")
    summary["Total Fuel [MT]"] = direct_total.fillna(component_total).round(3)
    return summary.sort_values("ShipName").reset_index(drop=True)


def _source_partition_columns(
    source_key: str,
    existing_manifest: dict[str, Any] | None,
    fresh_columns: list[str],
) -> list[str]:
    columns: list[str] = []
    for entry in manifest_partitions(existing_manifest).values():
        path = partition_entry_path(entry, "file")
        if not path.is_file():
            continue
        try:
            for column in pq.ParquetFile(path).schema.names:
                if column not in columns:
                    columns.append(column)
        except Exception:
            continue
    for column in fresh_columns:
        if column not in columns:
            columns.append(column)
    if source_key == "reportdata":
        ordered = [*SOURCE_COLUMNS, "ParsedValue"]
        return ordered
    datetime_column = source_primary_datetime_column(source_key)
    if datetime_column not in columns:
        columns.insert(0, datetime_column)
    return columns or [datetime_column]


def _split_fresh_chunks_to_month_files(
    source_key: str,
    chunk_files: list[Path],
    work_dir: Path,
    columns: list[str],
) -> dict[str, Path]:
    schema = source_arrow_schema(source_key, columns)
    datetime_column = source_primary_datetime_column(source_key)
    writers: dict[str, pq.ParquetWriter] = {}
    month_paths: dict[str, Path] = {}
    try:
        for chunk_path in chunk_files:
            parquet_file = pq.ParquetFile(chunk_path)
            for batch in parquet_file.iter_batches(batch_size=25_000):
                frame = batch.to_pandas()
                if datetime_column not in frame.columns:
                    continue
                dates = pd.to_datetime(frame[datetime_column], errors="coerce", utc=True)
                valid = dates.notna()
                if not valid.any():
                    continue
                frame = frame.loc[valid].copy()
                month_values = dates.loc[valid].dt.strftime("%Y-%m")
                for month_key in month_values.dropna().unique().tolist():
                    subset = frame.loc[month_values.eq(month_key)].copy()
                    if subset.empty:
                        continue
                    aligned = align_source_frame(source_key, subset, columns)
                    table = pa.Table.from_pandas(aligned, schema=schema, preserve_index=False)
                    if month_key not in writers:
                        path = work_dir / f"fresh_{month_key.replace('-', '')}.parquet"
                        month_paths[month_key] = path
                        writers[month_key] = pq.ParquetWriter(path, schema, compression="zstd")
                    writers[month_key].write_table(table)
                    del subset, aligned, table
                del frame, dates, month_values
                gc.collect()
    finally:
        for writer in writers.values():
            writer.close()
    return month_paths


def _read_month_frame(
    source_key: str,
    path: Path | None,
    *,
    before_date: date | None = None,
) -> pd.DataFrame:
    if path is None or not path.is_file():
        return pd.DataFrame()
    frame = pd.read_parquet(path)
    if before_date is not None and not frame.empty:
        datetime_column = source_primary_datetime_column(source_key)
        if datetime_column in frame.columns:
            values = pd.to_datetime(frame[datetime_column], errors="coerce", utc=True)
            frame = frame.loc[
                values.isna() | values.lt(pd.Timestamp(before_date, tz="UTC"))
            ].copy()
    return frame


def _partition_file_metadata(
    source_key: str,
    month_key: str,
    data_path: Path,
    summary_path: Path,
    frame: pd.DataFrame,
    summary: pd.DataFrame,
) -> dict[str, Any]:
    datetime_column = source_primary_datetime_column(source_key)
    values = (
        pd.to_datetime(frame[datetime_column], errors="coerce", utc=True)
        if datetime_column in frame.columns
        else pd.Series(dtype="datetime64[ns, UTC]")
    )
    return {
        "month": month_key,
        "file": str(data_path.relative_to(SNAPSHOT_DIR)),
        "summary_file": str(summary_path.relative_to(SNAPSHOT_DIR)),
        "rows": int(len(frame)),
        "summary_rows": int(len(summary)),
        "columns": int(len(frame.columns)),
        "min_datetime": (
            values.min().isoformat()
            if not values.empty and values.notna().any()
            else None
        ),
        "max_datetime": (
            values.max().isoformat()
            if not values.empty and values.notna().any()
            else None
        ),
    }


def _cleanup_partition_files(source_key: str, current_manifest: dict[str, Any]) -> None:
    root = source_partition_root(source_key)
    if not root.is_dir():
        return
    referenced = {
        str(partition_entry_path(entry, field).resolve())
        for entry in manifest_partitions(current_manifest).values()
        for field in ["file", "summary_file"]
        if entry.get(field)
    }
    grouped: dict[str, list[Path]] = {}
    for path in root.glob("*.parquet"):
        parts = path.stem.split("_")
        group_key = "_".join(parts[:2]) if len(parts) >= 2 else path.stem
        grouped.setdefault(group_key, []).append(path)
    for paths in grouped.values():
        paths.sort(key=lambda item: item.stat().st_mtime, reverse=True)
        protected = {str(path.resolve()) for path in paths[:2]} | referenced
        for path in paths:
            if str(path.resolve()) not in protected:
                try:
                    path.unlink()
                except OSError:
                    pass


def publish_source_snapshot(
    source_key: str,
    existing_manifest: dict[str, Any] | None,
    refresh_start_date: date,
    work_dir: Path,
    chunk_files: list[Path],
    fresh_columns: list[str],
    refresh_metadata: dict[str, Any],
    signature: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Publish immutable monthly source partitions and monthly summaries."""
    existing_partitions = manifest_partitions(existing_manifest)
    columns = _source_partition_columns(source_key, existing_manifest, fresh_columns)
    generation = snapshot_generation()
    source_partition_root(source_key).mkdir(parents=True, exist_ok=True)
    fresh_month_files = _split_fresh_chunks_to_month_files(
        source_key,
        chunk_files,
        work_dir,
        columns,
    )
    try:
        end_exclusive = date.fromisoformat(
            str(refresh_metadata.get("refresh_end_date_exclusive"))
        )
    except ValueError:
        end_exclusive = date.today() + timedelta(days=1)
    affected_months = month_keys_for_range(refresh_start_date, end_exclusive)
    new_partitions = dict(existing_partitions)

    for month_key in affected_months:
        old_entry = existing_partitions.get(month_key)
        old_path = partition_entry_path(old_entry, "file") if old_entry else None
        month_start = month_start_from_key(month_key)
        before_date = (
            refresh_start_date
            if month_start <= refresh_start_date < next_month_start(month_start)
            else None
        )
        old_frame = _read_month_frame(source_key, old_path, before_date=before_date)
        fresh_frame = _read_month_frame(source_key, fresh_month_files.get(month_key))
        combined = pd.concat([old_frame, fresh_frame], ignore_index=True, sort=False)
        if combined.empty:
            new_partitions.pop(month_key, None)
            continue
        combined = normalize_source_snapshot_frame(source_key, combined)
        combined = deduplicate_source_window(source_key, combined)
        datetime_column = source_primary_datetime_column(source_key)
        sort_columns = [
            column for column in [datetime_column, "ShipName"]
            if column in combined.columns
        ]
        if sort_columns:
            combined = combined.sort_values(sort_columns)
        data_path = source_partition_file(source_key, month_key, generation)
        summary_path = source_summary_file(source_key, month_key, generation)
        combined.to_parquet(data_path, index=False, compression="zstd")
        summary = build_monthly_source_summary(source_key, combined, month_key)
        if summary.empty:
            summary = pd.DataFrame(
                columns=[
                    "Month", "Source", "ShipName", "Records",
                    "First Timestamp", "Last Timestamp", "Observed Days",
                    "Period Days", "Day Coverage [%]", "Observation Coverage [%]",
                ]
            )
        summary.to_parquet(summary_path, index=False, compression="zstd")
        new_partitions[month_key] = _partition_file_metadata(
            source_key,
            month_key,
            data_path,
            summary_path,
            combined,
            summary,
        )
        del old_frame, fresh_frame, combined, summary
        gc.collect()

    if not new_partitions:
        raise RuntimeError(
            f"{SOURCE_CONFIGS[source_key]['label']} produced no monthly partitions."
        )

    total_rows = sum(
        int(entry.get("rows", 0) or 0)
        for entry in new_partitions.values()
    )
    latest_values = [
        entry.get("max_datetime")
        for entry in new_partitions.values()
        if entry.get("max_datetime")
    ]
    latest_source_date = (
        max(pd.Timestamp(value) for value in latest_values).date().isoformat()
        if latest_values
        else None
    )
    loaded_at_utc = datetime.now(timezone.utc)
    metadata = dict(refresh_metadata)
    metadata.update(
        {
            "source": SOURCE_CONFIGS[source_key]["label"],
            "endpoint": str(SOURCE_CONFIGS[source_key]["endpoint"]),
            "loaded_at_utc": loaded_at_utc.strftime("%d-%m-%Y %H:%M:%S UTC"),
            "loaded_at_local": local_time_label(loaded_at_utc),
            "loaded_start_date": API_FULL_START_DATE.isoformat(),
            "rows": int(total_rows),
            "kept_rows": int(total_rows),
            "columns": int(len(columns)),
            "partition_count": int(len(new_partitions)),
            "snapshot_generation": generation,
            "snapshot_format": "monthly_partitioned_parquet",
            "snapshot_schema_version": ATLAS_SNAPSHOT_SCHEMA_VERSION,
            "comparison_schema_version": ATLAS_MONTHLY_COMPARISON_VERSION,
            "latest_source_date": latest_source_date,
        }
    )
    saved_at_utc = loaded_at_utc.strftime("%d-%m-%Y %H:%M:%S UTC")
    manifest = {
        "snapshot_schema_version": ATLAS_SNAPSHOT_SCHEMA_VERSION,
        "comparison_schema_version": ATLAS_MONTHLY_COMPARISON_VERSION,
        "source": source_key,
        "generation": generation,
        "signature": signature,
        "partitions": {
            key: new_partitions[key]
            for key in sorted(new_partitions)
        },
        "metadata": metadata,
        "saved_at_utc": saved_at_utc,
    }
    _atomic_write_text(
        source_manifest_path(source_key),
        json.dumps(manifest, indent=2, default=str),
    )

    cached_read_prepared_source_snapshot.clear()
    cached_read_partitioned_source_snapshot.clear()
    cached_read_monthly_summary_files.clear()
    if source_key == "reportdata":
        cached_prepare_long_data.clear()
        build_pivot_table.clear()
    clear_wide_source_state(source_key)
    _cleanup_partition_files(source_key, manifest)
    cleanup_legacy_source_files(source_key)
    update_source_refresh_status(
        source_key,
        state="complete",
        stage="published",
        refresh_mode=metadata.get("refresh_mode"),
        rows_kept=total_rows,
        partition_count=len(new_partitions),
        snapshot_generation=generation,
    )
    return metadata, manifest


def _legacy_source_file(source_key: str) -> Path | None:
    manifest = read_source_manifest(source_key)
    if isinstance(manifest, dict) and manifest.get("prepared_file"):
        candidate = SNAPSHOT_DIR / str(manifest["prepared_file"])
        if candidate.is_file():
            return candidate
    candidate = Path(SOURCE_CONFIGS[source_key]["snapshot_file"])
    return candidate if candidate.is_file() else None


def migrate_legacy_source_snapshot(
    source_key: str,
    username: str,
    auth_method: str,
) -> tuple[dict[str, Any], dict[str, Any]] | None:
    # ReportData is compacted through a ValueDescription whitelist. This release
    # adds comparison fields (including speed), so an old compact snapshot cannot
    # be considered historically complete. Force one full ReportData bootstrap;
    # wide ReportPivots/ShipPivots snapshots can be partition-migrated safely.
    if source_key == "reportdata":
        return None
    legacy_file = _legacy_source_file(source_key)
    if legacy_file is None:
        return None
    try:
        parquet_file = pq.ParquetFile(legacy_file)
        fresh_columns = list(parquet_file.schema.names)
        if not fresh_columns:
            return None
        work_dir = SNAPSHOT_DIR / (
            f".refresh_{source_key}_migration_{os.getpid()}_"
            f"{int(time.time() * 1000)}"
        )
        work_dir.mkdir(parents=True, exist_ok=False)
        signature = atlas_source_signature(
            source_key,
            username,
            auth_method,
            API_FULL_START_DATE,
        )
        metadata = {
            "refresh_mode": "legacy_partition_migration",
            "refresh_api_start_date": API_FULL_START_DATE.isoformat(),
            "refresh_end_date_exclusive": (
                date.today() + timedelta(days=1)
            ).isoformat(),
            "chunk_days": 0,
            "chunks_total": 1,
            "chunks_completed": 1,
            "refresh_rows": int(parquet_file.metadata.num_rows),
            "scanned_rows": int(parquet_file.metadata.num_rows),
            "pages": 0,
            "downloaded_mb": 0,
            "fetch_seconds": 0,
            "first_url": "legacy_snapshot",
            "paging_stop_reason": "legacy_partition_migration",
            "hit_page_limit": False,
            "latest_source_date": None,
        }
        published = publish_source_snapshot(
            source_key,
            None,
            API_FULL_START_DATE,
            work_dir,
            [legacy_file],
            fresh_columns,
            metadata,
            signature,
        )
        remove_refresh_work_dir(work_dir)
        return published
    except Exception:
        return None


def ensure_source_snapshot(
    source_key: str,
    username: str,
    auth_method: str,
) -> tuple[dict[str, Any], dict[str, Any]] | None:
    existing = source_snapshot_info(
        source_key,
        username,
        auth_method,
        API_FULL_START_DATE,
    )
    if existing is not None:
        return existing
    with source_refresh_lock(source_key) as lock_acquired:
        if not lock_acquired:
            return source_snapshot_info(
                source_key,
                username,
                auth_method,
                API_FULL_START_DATE,
            )
        existing = source_snapshot_info(
            source_key,
            username,
            auth_method,
            API_FULL_START_DATE,
        )
        if existing is not None:
            return existing
        return migrate_legacy_source_snapshot(
            source_key,
            username,
            auth_method,
        )


def refresh_source_snapshot(
    source_key: str,
    username: str,
    password: str,
    token: str,
    auth_method: str,
    *,
    full_refresh: bool,
    acquire_lock: bool = True,
) -> tuple[dict[str, Any], dict[str, Any]]:
    if source_key not in SOURCE_CONFIGS:
        raise ValueError(f"Unsupported AtlasFlow source: {source_key}")
    requested_signature = atlas_source_signature(
        source_key,
        username,
        auth_method,
        API_FULL_START_DATE,
    )

    def execute_refresh() -> tuple[dict[str, Any], dict[str, Any]]:
        existing_manifest = read_source_manifest(source_key)
        if not source_manifest_is_valid(
            source_key,
            existing_manifest,
            requested_signature,
            API_FULL_START_DATE,
        ):
            migrated = migrate_legacy_source_snapshot(
                source_key,
                username,
                auth_method,
            )
            existing_manifest = (
                read_source_manifest(source_key)
                if migrated is not None
                else None
            )
        if not source_manifest_is_valid(
            source_key,
            existing_manifest,
            requested_signature,
            API_FULL_START_DATE,
        ):
            existing_manifest = None

        refresh_mode = "full"
        refresh_start_date = API_FULL_START_DATE
        if not full_refresh and existing_manifest is not None:
            latest_text = (existing_manifest.get("metadata") or {}).get(
                "latest_source_date"
            )
            try:
                latest_date = (
                    date.fromisoformat(str(latest_text))
                    if latest_text
                    else None
                )
            except ValueError:
                latest_date = None
            if latest_date is not None:
                overlap_days = read_int_secret(
                    source_secret_name(
                        source_key,
                        "INCREMENTAL_OVERLAP_DAYS",
                    ),
                    DEFAULT_SOURCE_OVERLAP_DAYS[source_key],
                    minimum=1,
                    maximum=90,
                )
                refresh_start_date = max(
                    API_FULL_START_DATE,
                    latest_date - timedelta(days=overlap_days),
                )
                refresh_mode = "incremental"

        chunk_days = read_int_secret(
            source_secret_name(source_key, "REFRESH_CHUNK_DAYS"),
            DEFAULT_SOURCE_CHUNK_DAYS[source_key],
            minimum=1,
            maximum=62,
        )
        default_minutes = (
            DEFAULT_SOURCE_FULL_REFRESH_MAX_MINUTES[source_key]
            if refresh_mode == "full"
            else DEFAULT_SOURCE_INCREMENTAL_REFRESH_MAX_MINUTES[source_key]
        )
        max_minutes = read_int_secret(
            source_secret_name(
                source_key,
                (
                    "FULL_REFRESH_MAX_MINUTES"
                    if refresh_mode == "full"
                    else "INCREMENTAL_REFRESH_MAX_MINUTES"
                ),
            ),
            default_minutes,
            minimum=5,
            maximum=720,
        )
        end_exclusive = date.today() + timedelta(days=1)
        work_dir: Path | None = None
        try:
            update_source_refresh_status(
                source_key,
                state="running",
                stage="starting",
                refresh_mode=refresh_mode,
                refresh_start_date=refresh_start_date.isoformat(),
                end_date_exclusive=end_exclusive.isoformat(),
                chunk_days=chunk_days,
                max_minutes=max_minutes,
            )
            (
                work_dir,
                chunk_files,
                fresh_columns,
                refresh_metadata,
            ) = collect_source_chunks(
                source_key,
                username,
                password,
                token,
                auth_method,
                refresh_start_date,
                end_exclusive,
                chunk_days,
                max_minutes * 60,
                refresh_mode,
            )
            if int(refresh_metadata.get("scanned_rows", 0) or 0) == 0:
                if existing_manifest is not None:
                    metadata = dict(existing_manifest.get("metadata") or {})
                    metadata.update(
                        {
                            "refresh_mode": "no_changes",
                            "refresh_api_start_date": refresh_start_date.isoformat(),
                            "refresh_checked_at_local": local_time_label(),
                        }
                    )
                    update_source_refresh_status(
                        source_key,
                        state="complete",
                        stage="no_changes",
                        refresh_mode="no_changes",
                    )
                    return metadata, existing_manifest
                raise RuntimeError(
                    f"{SOURCE_CONFIGS[source_key]['label']} returned zero rows "
                    "during initial bootstrap."
                )
            update_source_refresh_status(
                source_key,
                state="running",
                stage="partitioning",
                refresh_mode=refresh_mode,
                pages_completed=int(
                    refresh_metadata.get("pages", 0) or 0
                ),
                rows_kept=int(
                    refresh_metadata.get("refresh_rows", 0) or 0
                ),
            )
            return publish_source_snapshot(
                source_key,
                (
                    existing_manifest
                    if refresh_mode == "incremental"
                    else None
                ),
                refresh_start_date,
                work_dir,
                chunk_files,
                fresh_columns,
                refresh_metadata,
                requested_signature,
            )
        except Exception as exc:
            update_source_refresh_status(
                source_key,
                state="failed",
                stage="failed",
                refresh_mode=refresh_mode,
                error=str(exc),
            )
            raise
        finally:
            remove_refresh_work_dir(work_dir)

    if not acquire_lock:
        return execute_refresh()
    with source_refresh_lock(source_key) as lock_acquired:
        if not lock_acquired:
            existing = source_snapshot_info(
                source_key,
                username,
                auth_method,
                API_FULL_START_DATE,
            )
            if existing is not None:
                metadata, manifest = existing
                metadata = dict(metadata)
                metadata["refresh_skipped_due_to_lock"] = True
                metadata["refresh_status"] = source_refresh_status_summary(
                    source_key
                )
                return metadata, manifest
            raise AtlasRefreshAlreadyRunning(
                source_refresh_status_summary(source_key)
            )
        return execute_refresh()


def _partition_entries_for_period(
    manifest: dict[str, Any],
    selected_start: date,
    selected_end: date,
) -> list[dict[str, Any]]:
    wanted_months = set(
        month_keys_for_range(
            selected_start,
            selected_end + timedelta(days=1),
        )
    )
    partitions = manifest_partitions(manifest)
    return [
        partitions[month_key]
        for month_key in sorted(partitions)
        if month_key in wanted_months
    ]


def _dataset_filter_expression(
    source_key: str,
    schema_names: list[str],
    selected_vessels: list[str] | None,
    selected_start: date | None,
    selected_end: date | None,
) -> Any:
    expression = None
    if selected_vessels and "ShipName" in schema_names:
        expression = ds.field("ShipName").isin(list(selected_vessels))
    datetime_column = source_primary_datetime_column(source_key)
    if datetime_column in schema_names:
        if selected_start is not None:
            start_value = datetime.combine(
                selected_start,
                datetime.min.time(),
                tzinfo=timezone.utc,
            )
            start_expression = ds.field(datetime_column) >= pa.scalar(
                start_value,
                type=pa.timestamp("ns", tz="UTC"),
            )
            expression = (
                start_expression
                if expression is None
                else expression & start_expression
            )
        if selected_end is not None:
            end_value = datetime.combine(
                selected_end + timedelta(days=1),
                datetime.min.time(),
                tzinfo=timezone.utc,
            )
            end_expression = ds.field(datetime_column) < pa.scalar(
                end_value,
                type=pa.timestamp("ns", tz="UTC"),
            )
            expression = (
                end_expression
                if expression is None
                else expression & end_expression
            )
    return expression


def read_partitioned_source_slice(
    source_key: str,
    manifest: dict[str, Any],
    selected_vessels: list[str] | None,
    selected_start: date,
    selected_end: date,
    *,
    row_limit: int | None,
) -> tuple[pd.DataFrame, int, bool]:
    entries = _partition_entries_for_period(
        manifest,
        selected_start,
        selected_end,
    )
    files = [
        str(partition_entry_path(entry, "file"))
        for entry in entries
        if partition_entry_path(entry, "file").is_file()
    ]
    if not files:
        return pd.DataFrame(), 0, False
    dataset = ds.dataset(files, format="parquet")
    schema_names = list(dataset.schema.names)
    expression = _dataset_filter_expression(
        source_key,
        schema_names,
        selected_vessels,
        selected_start,
        selected_end,
    )
    try:
        matching_rows = int(dataset.count_rows(filter=expression))
    except Exception:
        matching_rows = sum(
            int(entry.get("rows", 0) or 0)
            for entry in entries
        )
    truncated = row_limit is not None and matching_rows > row_limit
    try:
        table = (
            dataset.head(row_limit, filter=expression)
            if truncated and row_limit is not None
            else dataset.to_table(filter=expression)
        )
        frame = table.to_pandas()
        del table
    except Exception:
        frames: list[pd.DataFrame] = []
        remaining = row_limit
        for file_name in files:
            parquet_file = pq.ParquetFile(file_name)
            for batch in parquet_file.iter_batches(batch_size=50_000):
                batch_frame = batch.to_pandas()
                batch_frame = filter_wide_source_data(
                    batch_frame,
                    source_key,
                    selected_vessels or [],
                    selected_start,
                    selected_end,
                )
                if batch_frame.empty:
                    continue
                if remaining is not None:
                    batch_frame = batch_frame.head(remaining)
                    remaining -= len(batch_frame)
                frames.append(batch_frame)
                if remaining is not None and remaining <= 0:
                    break
            if remaining is not None and remaining <= 0:
                break
        frame = (
            pd.concat(frames, ignore_index=True, sort=False)
            if frames
            else pd.DataFrame(columns=schema_names)
        )
    return frame, matching_rows, truncated


def load_wide_source_for_view(
    source_key: str,
    username: str,
    password: str,
    token: str,
    auth_method: str,
    start_date: date,
    refresh: bool,
    selected_vessels: list[str] | None = None,
    selected_start: date | None = None,
    selected_end: date | None = None,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    if refresh:
        refresh_source_snapshot(
            source_key,
            username,
            password,
            token,
            auth_method,
            full_refresh=False,
        )
    requested_signature = atlas_source_signature(
        source_key,
        username,
        auth_method,
        start_date,
    )
    manifest = read_source_manifest(source_key)
    if not source_manifest_is_valid(
        source_key,
        manifest,
        requested_signature,
        start_date,
    ):
        migrated = ensure_source_snapshot(
            source_key,
            username,
            auth_method,
        )
        manifest = (
            read_source_manifest(source_key)
            if migrated is not None
            else None
        )
    if not source_manifest_is_valid(
        source_key,
        manifest,
        requested_signature,
        start_date,
    ):
        config = SOURCE_CONFIGS[source_key]
        return pd.DataFrame(), {
            "source": config["label"],
            "endpoint": str(config["endpoint"]),
            "loaded_at_local": "No prepared snapshot yet",
            "rows": 0,
            "needs_warmup": True,
        }
    assert manifest is not None
    selected_start = selected_start or start_date
    selected_end = selected_end or date.today()
    row_limit = (
        ATLAS_RAW_INTERACTIVE_ROW_LIMIT
        if source_key == "shippivots"
        else ATLAS_REPORTPIVOTS_INTERACTIVE_ROW_LIMIT
    )
    frame, matching_rows, truncated = read_partitioned_source_slice(
        source_key,
        manifest,
        selected_vessels,
        selected_start,
        selected_end,
        row_limit=row_limit,
    )
    metadata = dict(manifest.get("metadata") or {})
    metadata["loaded_from_snapshot"] = True
    metadata["snapshot_generation"] = manifest.get("generation")
    metadata["snapshot_saved_at_utc"] = manifest.get(
        "saved_at_utc",
        "-",
    )
    metadata["view_rows_matching"] = int(matching_rows)
    metadata["view_rows_loaded"] = int(len(frame))
    metadata["view_truncated"] = bool(truncated)
    metadata["interactive_row_limit"] = int(row_limit)
    return frame, metadata


def render_wide_source_tab(
    source_label: str,
    df: pd.DataFrame,
    metadata: dict[str, Any],
    source_key: str,
    selected_vessels: list[str],
    selected_start: date,
    selected_end: date,
) -> pd.DataFrame:
    st.markdown(
        f'<div class="section-title">{escape(source_label)} Dataset</div>',
        unsafe_allow_html=True,
    )
    render_api_load_caption(metadata)
    filtered_df = filter_wide_source_data(
        df,
        source_key,
        selected_vessels,
        selected_start,
        selected_end,
    )
    matching_rows = int(
        metadata.get("view_rows_matching", len(filtered_df)) or 0
    )
    render_metric_cards(
        [
            ("Rows in selection", f"{matching_rows:,}", "table_eye"),
            ("Rows loaded", f"{len(filtered_df):,}", "checked_columns"),
            (
                "Stored source rows",
                f"{int(metadata.get('rows', 0) or 0):,}",
                "database_rows",
            ),
            (
                "Monthly partitions",
                f"{int(metadata.get('partition_count', 0) or 0):,}",
                "numeric",
            ),
        ]
    )
    if metadata.get("view_truncated"):
        st.warning(
            f"The raw selection contains {matching_rows:,} rows. AtlasFlow "
            f"loaded the first {int(metadata.get('interactive_row_limit', 0) or 0):,} "
            "rows to protect Streamlit Cloud memory. Use Monthly Comparison for "
            "complete all-vessel monthly analysis, or narrow the raw detail period."
        )
    default_columns = [
        column
        for column in [
            "ShipName",
            "DateTime",
            "State",
            "StateName",
            "GPSSpeed",
            "LogSpeed",
            "MEConsumed",
            "ShaftPower",
        ]
        if column in filtered_df.columns
    ]
    if not default_columns:
        default_columns = list(
            filtered_df.columns[: min(12, len(filtered_df.columns))]
        )
    selected_columns = st.multiselect(
        f"{source_label} columns to preview/export",
        options=list(filtered_df.columns),
        default=default_columns,
        key=f"{source_key}_preview_columns",
    )
    output = (
        filtered_df[selected_columns].copy()
        if selected_columns
        else filtered_df.copy()
    )
    render_preview_table(output)
    if len(output) > TABLE_PREVIEW_ROW_LIMIT:
        st.caption(
            f"Showing first {TABLE_PREVIEW_ROW_LIMIT:,} of "
            f"{len(output):,} loaded rows."
        )
    return output


def load_monthly_comparison_data(
    username: str,
    auth_method: str,
    selected_vessels: list[str],
    selected_start: date,
    selected_end: date,
) -> tuple[
    pd.DataFrame,
    dict[str, dict[str, Any]],
    list[str],
]:
    summary_files: list[str] = []
    generation_parts: list[str] = []
    source_metadata: dict[str, dict[str, Any]] = {}
    missing_sources: list[str] = []
    wanted_months = set(
        month_keys_for_range(
            selected_start,
            selected_end + timedelta(days=1),
        )
    )
    for source_key in ["reportdata", "reportpivots", "shippivots"]:
        requested_signature = atlas_source_signature(
            source_key,
            username,
            auth_method,
            API_FULL_START_DATE,
        )
        manifest = read_source_manifest(source_key)
        if not source_manifest_is_valid(
            source_key,
            manifest,
            requested_signature,
            API_FULL_START_DATE,
        ):
            missing_sources.append(COMPARISON_SOURCE_LABELS[source_key])
            continue
        assert manifest is not None
        source_metadata[source_key] = dict(
            manifest.get("metadata") or {}
        )
        generation_parts.append(
            f"{source_key}:{manifest.get('generation')}"
        )
        for month_key, entry in manifest_partitions(manifest).items():
            if month_key not in wanted_months:
                continue
            summary_path = partition_entry_path(entry, "summary_file")
            if summary_path.is_file():
                summary_files.append(str(summary_path))
    comparison = cached_read_monthly_summary_files(
        "|".join(generation_parts),
        tuple(sorted(summary_files)),
    )
    if (
        not comparison.empty
        and selected_vessels
        and "ShipName" in comparison.columns
    ):
        comparison = comparison[
            match_selected_vessels(
                comparison["ShipName"],
                selected_vessels,
            )
        ].copy()
    return comparison, source_metadata, missing_sources


def render_monthly_comparison_workspace(
    username: str,
    auth_method: str,
    selected_vessels: list[str],
    selected_start: date,
    selected_end: date,
) -> None:
    st.markdown(
        '<div class="section-title">Monthly Cross-Source Comparison</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "This workspace compares all selected vessels from prepared monthly "
        "summaries. It never loads the full 15-minute history into the browser."
    )
    comparison_df, source_metadata, missing_sources = (
        load_monthly_comparison_data(
            username,
            auth_method,
            selected_vessels,
            selected_start,
            selected_end,
        )
    )
    if missing_sources:
        st.warning(
            "Prepared monthly summaries are not available yet for: "
            + ", ".join(missing_sources)
            + ". Run those source warmups."
        )
    if comparison_df.empty:
        st.info(
            "No prepared monthly comparison rows match the selected vessels "
            "and period."
        )
        return

    available_months = sorted(
        comparison_df["Month"].dropna().astype(str).unique().tolist(),
        reverse=True,
    )
    complete_months = sorted(
        comparison_df.loc[
            comparison_df.get("Month Complete", False).fillna(False).astype(bool),
            "Month",
        ].dropna().astype(str).unique().tolist(),
        reverse=True,
    ) if "Month Complete" in comparison_df.columns else available_months
    default_months = complete_months or available_months[:1]
    selected_months = st.multiselect(
        "Calendar months",
        options=available_months,
        default=default_months,
        key="atlas_monthly_comparison_months",
        help=(
            "Monthly comparison always uses complete calendar-month summaries. "
            "The current partial month can be selected explicitly when needed."
        ),
    )
    if selected_months:
        comparison_df = comparison_df[
            comparison_df["Month"].astype(str).isin(selected_months)
        ].copy()

    with st.expander("Source freshness and storage", expanded=False):
        freshness_rows = []
        for source_key, source_meta in source_metadata.items():
            freshness_rows.append(
                {
                    "Source": COMPARISON_SOURCE_LABELS[source_key],
                    "Last API load": source_meta.get("loaded_at_local", "-"),
                    "Latest source date": source_meta.get("latest_source_date", "-"),
                    "Stored rows": int(source_meta.get("rows", 0) or 0),
                    "Monthly partitions": int(source_meta.get("partition_count", 0) or 0),
                    "Refresh mode": source_meta.get("refresh_mode", "-"),
                }
            )
        if freshness_rows:
            st.dataframe(pd.DataFrame(freshness_rows), use_container_width=True, hide_index=True)

    available_sources = sorted(
        comparison_df["Source"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )
    selected_sources = st.multiselect(
        "Sources to compare",
        options=available_sources,
        default=available_sources,
        key="atlas_monthly_comparison_sources",
    )
    if selected_sources:
        comparison_df = comparison_df[
            comparison_df["Source"].isin(selected_sources)
        ].copy()
    metric_options = [
        metric
        for metric in MONTHLY_COMPARISON_METRICS
        if metric in comparison_df.columns
        and pd.to_numeric(
            comparison_df[metric],
            errors="coerce",
        ).notna().any()
    ]
    if not metric_options:
        st.info(
            "The prepared source summaries do not yet share a standardized "
            "numeric metric."
        )
        return
    metric = st.selectbox(
        "Standardized metric",
        options=metric_options,
        key="atlas_monthly_comparison_metric",
    )
    view_mode = st.radio(
        "Comparison view",
        options=[
            "Side-by-side",
            "Source rows",
            "Data quality",
            "Metric mapping",
        ],
        horizontal=True,
        key="atlas_monthly_comparison_view",
    )
    render_metric_cards(
        [
            (
                "Vessels",
                f"{comparison_df['ShipName'].nunique():,}",
                "table_eye",
            ),
            (
                "Months",
                f"{comparison_df['Month'].nunique():,}",
                "checked_columns",
            ),
            (
                "Sources",
                f"{comparison_df['Source'].nunique():,}",
                "database_rows",
            ),
            (
                "Monthly source rows",
                f"{len(comparison_df):,}",
                "numeric",
            ),
        ]
    )

    if view_mode == "Side-by-side":
        displayed = comparison_df.pivot_table(
            index=["Month", "ShipName"],
            columns="Source",
            values=metric,
            aggfunc="first",
        ).reset_index()
        source_columns = [
            source
            for source in available_sources
            if source in displayed.columns
        ]
        if source_columns:
            numeric_block = displayed[source_columns].apply(
                pd.to_numeric,
                errors="coerce",
            )
            displayed["Source Range"] = (
                numeric_block.max(axis=1)
                - numeric_block.min(axis=1)
            )
            displayed["Source Mean"] = numeric_block.mean(axis=1)
            displayed["Relative Range [%]"] = (
                displayed["Source Range"]
                / displayed["Source Mean"].replace(0, pd.NA)
                * 100
            ).round(2)
        displayed = displayed.sort_values(
            ["Month", "ShipName"],
            ascending=[False, True],
        )
    elif view_mode == "Data quality":
        quality_columns = [
            column
            for column in [
                "Month",
                "ShipName",
                "Source",
                "Records",
                "Observed Days",
                "Period Days",
                "Month Complete",
                "Day Coverage [%]",
                "Observation Coverage [%]",
                "First Timestamp",
                "Last Timestamp",
            ]
            if column in comparison_df.columns
        ]
        displayed = comparison_df[quality_columns].sort_values(
            ["Month", "ShipName", "Source"],
            ascending=[False, True, True],
        )
    elif view_mode == "Metric mapping":
        mapping_column = f"Mapping: {metric}"
        mapping_columns = [
            column
            for column in ["Source", mapping_column]
            if column in comparison_df.columns
        ]
        displayed = (
            comparison_df[mapping_columns]
            .drop_duplicates()
            .sort_values("Source")
        )
    else:
        source_columns = [
            column
            for column in [
                "Month",
                "ShipName",
                "Source",
                metric,
                "Records",
                "Day Coverage [%]",
                "Observation Coverage [%]",
                f"Mapping: {metric}",
            ]
            if column in comparison_df.columns
        ]
        displayed = comparison_df[source_columns].sort_values(
            ["Month", "ShipName", "Source"],
            ascending=[False, True, True],
        )

    st.dataframe(
        format_display_dataframe(displayed),
        use_container_width=True,
        hide_index=True,
    )
    export_signature = sha256(
        "|".join(
            [
                metric,
                view_mode,
                selected_start.isoformat(),
                selected_end.isoformat(),
                ",".join(selected_vessels),
                ",".join(selected_sources),
                str(len(displayed)),
            ]
        ).encode("utf-8")
    ).hexdigest()
    if (
        st.session_state.get("atlas_monthly_export_signature")
        != export_signature
    ):
        st.session_state.pop("atlas_monthly_export_bytes", None)
    if st.button(
        "Prepare monthly comparison Excel",
        type="primary",
        disabled=displayed.empty,
    ):
        with st.spinner("Preparing monthly comparison workbook..."):
            st.session_state["atlas_monthly_export_bytes"] = (
                to_displayed_table_excel_bytes(
                    displayed,
                    sheet_name="Monthly Comparison",
                )
            )
            st.session_state[
                "atlas_monthly_export_signature"
            ] = export_signature
    if (
        st.session_state.get("atlas_monthly_export_signature")
        == export_signature
        and "atlas_monthly_export_bytes" in st.session_state
    ):
        st.download_button(
            "Download monthly comparison Excel",
            data=st.session_state["atlas_monthly_export_bytes"],
            file_name="atlasflow_monthly_comparison.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )


# =============================================================================
# Warmup
# =============================================================================


def run_warmup_if_requested() -> None:
    """Build or incrementally refresh one or all prepared AtlasFlow snapshots."""
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
    requested_source = get_query_param("source", "reportdata").strip().lower()
    force_refresh = get_query_param("force", "0") == "1"
    full_refresh = get_query_param("full", "0") == "1"

    if auth_method.lower() in {"basic", "digest"} and (not username or not password):
        st.error("Warmup failed: MARORKA_USERNAME and MARORKA_PASSWORD are required.")
        st.stop()

    valid_sources = ["reportdata", "reportpivots", "shippivots"]
    if requested_source == "all":
        requested_sources = valid_sources
    elif requested_source in valid_sources:
        requested_sources = [requested_source]
    else:
        st.error("Invalid warmup source. Use reportdata, reportpivots, shippivots, or all.")
        st.stop()

    warmup_started_at = time.perf_counter()
    results: dict[str, dict[str, Any]] = {}
    failures: dict[str, str] = {}

    for source_key in requested_sources:
        try:
            if force_refresh:
                with st.spinner(f"Refreshing and preparing {SOURCE_CONFIGS[source_key]['label']}..."):
                    metadata, manifest = refresh_source_snapshot(
                        source_key,
                        username,
                        password,
                        token,
                        auth_method,
                        full_refresh=full_refresh,
                    )
            else:
                existing = ensure_source_snapshot(source_key, username, auth_method)
                if existing is None:
                    with st.spinner(f"Creating the first prepared {SOURCE_CONFIGS[source_key]['label']} snapshot..."):
                        metadata, manifest = refresh_source_snapshot(
                            source_key,
                            username,
                            password,
                            token,
                            auth_method,
                            full_refresh=True,
                        )
                else:
                    metadata, manifest = existing

            # ReportData is used by the default workspace, so seed its shared read cache.
            if source_key == "reportdata":
                signature = atlas_source_signature(source_key, username, auth_method, API_FULL_START_DATE)
                load_source_snapshot(source_key, signature, API_FULL_START_DATE)

            results[source_key] = {
                "refresh_mode": metadata.get("refresh_mode", "snapshot_only"),
                "snapshot_generation": manifest.get("generation"),
                "last_api_load_local": metadata.get("loaded_at_local", "-"),
                "rows": int(metadata.get("rows", 0) or 0),
                "columns": int(metadata.get("columns", 0) or 0),
                "monthly_partitions": int(metadata.get("partition_count", 0) or 0),
                "api_pages_last_refresh": int(metadata.get("pages", 0) or 0),
                "refresh_api_start_date": metadata.get("refresh_api_start_date", "-"),
                "refresh_skipped_due_to_lock": bool(metadata.get("refresh_skipped_due_to_lock", False)),
            }
        except AtlasRefreshAlreadyRunning as exc:
            failures[source_key] = f"already running: {exc}"
        except requests.HTTPError as exc:
            status = exc.response.status_code if exc.response is not None else "unknown"
            failures[source_key] = f"Marorka HTTP status {status}"
        except (
            MarorkaConfigError,
            ValueError,
            FileNotFoundError,
            RuntimeError,
            TimeoutError,
            OSError,
            requests.RequestException,
        ) as exc:
            failures[source_key] = str(exc)
        finally:
            gc.collect()

    if results:
        st.success("AtlasFlow prepared snapshot warmup completed.")
        st.write(
            {
                "sources": results,
                "force_refresh": force_refresh,
                "full_refresh": full_refresh,
                "warmup_seconds": round(time.perf_counter() - warmup_started_at, 2),
            }
        )
    if failures:
        st.error(
            "Some AtlasFlow sources were not refreshed: "
            + "; ".join(f"{source}: {error}" for source, error in failures.items())
        )
        for source_key in failures:
            st.caption(source_refresh_status_summary(source_key))
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

    if refresh:
        try:
            with st.spinner("Refreshing AtlasFlow APIs one source at a time..."):
                refresh_all_atlasflow_snapshots(
                    username,
                    password,
                    token,
                    auth_method,
                    api_start_date,
                )
        except requests.HTTPError as exc:
            status = exc.response.status_code if exc.response is not None else "unknown"
            st.error(
                f"AtlasFlow refresh failed with status {status}. "
                "Existing session data and the last valid snapshots remain available."
            )
            st.stop()
        except (
            MarorkaConfigError,
            ValueError,
            FileNotFoundError,
            RuntimeError,
            requests.RequestException,
        ) as exc:
            st.error(
                "AtlasFlow refresh failed. Existing session data and the last valid "
                f"snapshots remain available. Details: {exc}"
            )
            st.stop()
        st.success("All AtlasFlow API snapshots refreshed successfully.")
        st.rerun()

    reportdata_signature = atlas_source_signature(
        "reportdata",
        username,
        auth_method,
        api_start_date,
    )
    reportdata_manifest = read_source_manifest("reportdata")
    current_generation = (
        reportdata_manifest.get("generation")
        if isinstance(reportdata_manifest, dict)
        else None
    )
    long_df = st.session_state.get("loaded_long_df")
    metadata = st.session_state.get("loaded_metadata")
    session_signature = st.session_state.get("loaded_request_signature")
    session_generation = st.session_state.get("loaded_reportdata_generation")
    session_ready = (
        isinstance(long_df, pd.DataFrame)
        and isinstance(metadata, dict)
        and source_signature_covers_request(
            session_signature,
            reportdata_signature,
            metadata,
            api_start_date,
        )
        and session_generation == current_generation
    )

    if not session_ready:
        snapshot = load_source_snapshot(
            "reportdata",
            reportdata_signature,
            api_start_date,
        )
        if snapshot is None:
            migrated = ensure_source_snapshot("reportdata", username, auth_method)
            if migrated is not None:
                snapshot = load_source_snapshot(
                    "reportdata",
                    reportdata_signature,
                    api_start_date,
                )
        if snapshot is None:
            st.warning(
                "No prepared ReportData snapshot is available yet. Run the AtlasFlow warmup first; "
                "normal users will not be forced to wait for the large API pull."
            )
            st.code(
                "https://atlas-flow.streamlit.app/?warmup=1&force=1&source=all&token=warmup-atlas-flow",
                language="text",
            )
            st.stop()

        long_df, metadata, snapshot_signature = snapshot
        st.session_state.pop("loaded_raw_df", None)
        st.session_state["loaded_long_df"] = long_df
        st.session_state["loaded_metadata"] = metadata
        st.session_state["loaded_request_signature"] = snapshot_signature
        st.session_state["loaded_prepare_signature"] = source_data_signature("reportdata")
        st.session_state["loaded_reportdata_generation"] = metadata.get("snapshot_generation")

    long_df = st.session_state.get("loaded_long_df")
    metadata = st.session_state.get("loaded_metadata")
    if not isinstance(long_df, pd.DataFrame) or not isinstance(metadata, dict):
        st.error("The prepared ReportData snapshot could not be loaded.")
        st.stop()

    if long_df.empty:
        render_header(selected_group, selected_vessels, [])
        render_api_load_caption(metadata)
        st.warning("No Marorka report values were returned for the loaded API window.")
        st.stop()

    selected_start, selected_end = render_date_slicer(long_df)

    # ReportType is handled with the rest of the displayed-column filters.
    selected_report_types: list[str] = []
    filtered_long_for_options = filter_long_data(
        long_df,
        selected_vessels=selected_vessels,
        selected_report_types=selected_report_types,
        selected_start=selected_start,
        selected_end=selected_end,
    )

    variable_options = sorted(
        set(available_variables(filtered_long_for_options)).union(DERIVED_VARIABLES),
        key=str.casefold,
    )
    st.sidebar.markdown("### Pivot variables")
    previous_selected_variables = st.session_state.get("atlas_selected_variables", [])
    if not isinstance(previous_selected_variables, list):
        previous_selected_variables = []
    valid_default_variables = [variable for variable in previous_selected_variables if variable in variable_options]
    selected_variables = st.sidebar.multiselect(
        "Variables to include and filter",
        options=variable_options,
        default=valid_default_variables,
        key="atlas_selected_variables",
        help=(
            "Every selected ValueDescription becomes a displayed table column. "
            "The same selected variables are also available below as filters."
        ),
    )

    identity_columns = st.sidebar.multiselect(
        "Pivot rows",
        options=PIVOT_IDENTITY_COLUMNS,
        default=["ShipName"],
        help="Choose the row fields that appear before the selected variable columns.",
    )
    if not identity_columns:
        identity_columns = ["ShipName"]

    workspace_options = [
        "Monthly Comparison",
        "Custom Analytics",
        "Noon & Manual Reports",
        "15-Minute Operations",
        "Descriptive Statistics",
        "Export Center",
        "API Diagnostics",
    ]

    render_header(selected_group, selected_vessels, selected_variables)
    render_api_load_caption(metadata)

    # Top navigation keeps the previous tab-like layout while preserving the
    # memory optimization: only the selected workspace loads its heavy data.
    workspace = get_tab_selection(
        "workspace",
        workspace_options,
        st.session_state.get("atlas_workspace", "Monthly Comparison"),
    )
    workspace = render_text_tab_bar(
        workspace_options,
        workspace,
        param_name="workspace",
        reset_params=["preview"],
    )
    st.session_state["atlas_workspace"] = workspace

    active_wide_sources: set[str] = set()
    if workspace == "Noon & Manual Reports":
        active_wide_sources.add("reportpivots")
    elif workspace == "15-Minute Operations":
        active_wide_sources.add("shippivots")
    elif workspace == "Descriptive Statistics":
        descriptive_source_hint = st.session_state.get("atlas_descriptive_source_selector", "Custom Analytics")
        if descriptive_source_hint == "Noon & Manual Reports":
            active_wide_sources.add("reportpivots")
        elif descriptive_source_hint == "15-Minute Operations":
            active_wide_sources.add("shippivots")
    elif workspace == "Export Center":
        # Wide sources are loaded inside the export button only, not during normal render.
        active_wide_sources = set()
    clear_inactive_wide_sources(active_wide_sources)

    if metadata.get("hit_page_limit"):
        st.warning(
            "The API refresh reached the maximum page safety limit before the feed ended. "
            "The loaded dataset may be incomplete. Check API Diagnostics before using the export."
        )

    if workspace == "Monthly Comparison":
        # The comparison workspace reads pre-aggregated monthly summaries and
        # does not need to build the report-level pivot or its filter controls.
        pivot_df = pd.DataFrame()
        filtered_pivot_df = pd.DataFrame()
        filter_specs: list[dict[str, Any]] = []
        display_columns: list[str] = []
        output_df = pd.DataFrame()
    else:
        if not selected_variables:
            st.info("Select one or more variables from the sidebar to build the AtlasFlow pivot table.")

        pivot_df = build_pivot_table(filtered_long_for_options, tuple(selected_variables))

        filter_column_options = [column for column in [*identity_columns, *selected_variables] if column in pivot_df.columns]
        with st.sidebar.expander("Filters for displayed columns", expanded=False):
            st.caption("Choose columns to filter. Selected variables are already part of the displayed table.")
            previous_filter_columns = st.session_state.get("atlas_columns_to_filter", [])
            if not isinstance(previous_filter_columns, list):
                previous_filter_columns = []
            valid_filter_columns = [column for column in previous_filter_columns if column in filter_column_options]
            if valid_filter_columns != previous_filter_columns:
                st.session_state["atlas_columns_to_filter"] = valid_filter_columns

            columns_to_filter = st.multiselect(
                "Columns to filter",
                options=filter_column_options,
                default=valid_filter_columns,
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

    # Shared Custom Analytics preview/export configuration. It is only fully rendered in
    # Custom Analytics, but Export Center can reuse the saved choices.
    summary_group_fields: list[str] = []
    summary_value_fields: list[str] = []
    summary_aggregation = st.session_state.get("atlas_export_summary_aggregation", "Average")
    preview_mode = st.session_state.get("atlas_reportdata_preview_mode", "Clean Dataset")
    displayed_table_df = output_df.copy()
    export_sheet_name = "Clean Dataset"
    current_export_signature = sha256(
        "|".join([
            preview_mode,
            ",".join(selected_vessels),
            selected_start.isoformat(),
            selected_end.isoformat(),
            ",".join(display_columns),
            ",".join(selected_variables),
            str(len(output_df)),
        ]).encode("utf-8")
    ).hexdigest()

    if workspace == "Monthly Comparison":
        render_monthly_comparison_workspace(
            username,
            auth_method,
            selected_vessels,
            selected_start,
            selected_end,
        )

    elif workspace == "Custom Analytics":
        st.markdown('<div class="section-title">Custom Analytics Preview & Export</div>', unsafe_allow_html=True)

        summary_builder_columns = [column for column in output_df.columns]
        summary_value_options = numeric_column_options(output_df)

        st.caption(
            "Choose which table you want to preview and export. The visible table below is the same table prepared for Excel."
        )
        preview_options = ["Clean Dataset", "Summary Analysis", "Source Data"]
        preview_mode = get_tab_selection(
            "preview",
            preview_options,
            st.session_state.get("atlas_reportdata_preview_mode", "Clean Dataset"),
        )
        preview_mode = render_text_tab_bar(
            preview_options,
            preview_mode,
            param_name="preview",
            css_class="compact",
        )
        st.session_state["atlas_reportdata_preview_mode"] = preview_mode

        if preview_mode == "Summary Analysis":
            st.markdown('<div class="section-title">Summary Builder</div>', unsafe_allow_html=True)
            builder_cols = st.columns(3)
            with builder_cols[0]:
                previous_summary_groups = st.session_state.get("atlas_export_summary_groups", [])
                if not isinstance(previous_summary_groups, list):
                    previous_summary_groups = []
                default_summary_groups = [column for column in ["ShipName", "ReportType"] if column in summary_builder_columns]
                valid_summary_group_defaults = [column for column in previous_summary_groups if column in summary_builder_columns]
                if not valid_summary_group_defaults and "atlas_export_summary_groups" not in st.session_state:
                    valid_summary_group_defaults = default_summary_groups
                if valid_summary_group_defaults != previous_summary_groups:
                    st.session_state["atlas_export_summary_groups"] = valid_summary_group_defaults
                summary_group_fields = st.multiselect(
                    "Group by fields",
                    options=summary_builder_columns,
                    default=valid_summary_group_defaults,
                    key="atlas_export_summary_groups",
                    help="Choose the fields that define each summary row.",
                )
            with builder_cols[1]:
                previous_summary_values = st.session_state.get("atlas_export_summary_values", [])
                if not isinstance(previous_summary_values, list):
                    previous_summary_values = []
                valid_summary_value_defaults = [column for column in previous_summary_values if column in summary_value_options]
                if valid_summary_value_defaults != previous_summary_values:
                    st.session_state["atlas_export_summary_values"] = valid_summary_value_defaults
                summary_value_fields = st.multiselect(
                    "Value fields",
                    options=summary_value_options,
                    default=valid_summary_value_defaults,
                    key="atlas_export_summary_values",
                    help="Choose one or more numeric columns to aggregate.",
                )
            with builder_cols[2]:
                summary_aggregation = st.selectbox(
                    "Aggregation",
                    options=["Average", "Sum", "Count", "Minimum", "Maximum", "Median"],
                    index=["Average", "Sum", "Count", "Minimum", "Maximum", "Median"].index(summary_aggregation)
                    if summary_aggregation in ["Average", "Sum", "Count", "Minimum", "Maximum", "Median"] else 0,
                    key="atlas_export_summary_aggregation",
                )
        else:
            summary_group_fields = st.session_state.get("atlas_export_summary_groups", [])
            summary_value_fields = st.session_state.get("atlas_export_summary_values", [])
            if not isinstance(summary_group_fields, list):
                summary_group_fields = []
            if not isinstance(summary_value_fields, list):
                summary_value_fields = []

        summary_can_build = bool(summary_group_fields and summary_value_fields)
        if preview_mode == "Summary Analysis" and summary_can_build:
            displayed_table_df = build_summary_analysis(
                output_df,
                group_fields=summary_group_fields,
                value_fields=summary_value_fields,
                aggregation=summary_aggregation,
            )
            export_sheet_name = "Summary Analysis"
        elif preview_mode == "Summary Analysis":
            displayed_table_df = pd.DataFrame()
            export_sheet_name = "Summary Analysis"
            st.info("Select at least one Group by field and one Value field to preview Summary Analysis.")
        elif preview_mode == "Source Data":
            source_columns = [column for column in [*SOURCE_COLUMNS, "ParsedValue"] if column in filtered_long_for_options.columns]
            displayed_table_df = filtered_long_for_options[source_columns].copy()
            export_sheet_name = "Source Data"
        else:
            displayed_table_df = output_df.copy()
            export_sheet_name = "Clean Dataset"

        render_metric_cards(
            [
                ("Displayed Rows", f"{len(displayed_table_df):,}", "table_eye"),
                ("Selected Variables", f"{len(selected_variables):,}", "checked_columns"),
                ("Source Rows", f"{len(filtered_long_for_options):,}", "database_rows"),
                ("Available Variables", f"{len(variable_options):,}", "columns_plus"),
            ]
        )

        render_preview_table(displayed_table_df)
        if len(displayed_table_df) > TABLE_PREVIEW_ROW_LIMIT:
            st.caption(
                f"Showing first {TABLE_PREVIEW_ROW_LIMIT:,} of {len(displayed_table_df):,} rows. "
                "Excel export includes the full displayed table."
            )

        export_signature_payload = "|".join([
            preview_mode,
            ",".join(selected_vessels),
            selected_start.isoformat(),
            selected_end.isoformat(),
            ",".join(display_columns),
            ",".join(selected_variables),
            str(len(output_df)),
            str(len(displayed_table_df)),
            ",".join(summary_group_fields),
            ",".join(summary_value_fields),
            str(summary_aggregation),
            ",".join(displayed_table_df.columns.astype(str).tolist()) if not displayed_table_df.empty else "empty",
        ])
        current_export_signature = sha256(export_signature_payload.encode("utf-8")).hexdigest()
        clear_stale_export_bytes(current_export_signature)

        export_ready = (
            st.session_state.get("atlas_export_signature") == current_export_signature
            and "atlas_export_bytes" in st.session_state
        )

        if st.button("Prepare displayed table Excel", type="primary", disabled=displayed_table_df.empty):
            with st.spinner("Preparing Excel file..."):
                st.session_state["atlas_export_bytes"] = to_displayed_table_excel_bytes(
                    displayed_table_df,
                    sheet_name=export_sheet_name,
                )
                st.session_state["atlas_summary_analysis_df"] = displayed_table_df if preview_mode == "Summary Analysis" else pd.DataFrame()
                st.session_state["atlas_export_signature"] = current_export_signature
                gc.collect()
            export_ready = True

        if export_ready:
            st.download_button(
                "Download displayed table Excel",
                data=st.session_state["atlas_export_bytes"],
                file_name="atlasflow_displayed_table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.caption("Excel generation is prepared on demand. The download will contain only the visible table above.")

    elif workspace == "Noon & Manual Reports":
        reportpivots_df, reportpivots_metadata = load_wide_source_for_view(
            "reportpivots", username, password, token, auth_method, api_start_date, refresh,
            selected_vessels, selected_start, selected_end
        )
        if reportpivots_metadata.get("needs_warmup"):
            st.info("No ReportPivots snapshot is available yet. Run the ReportPivots warmup URL first.")
            st.code("https://atlas-flow.streamlit.app/?warmup=1&force=1&source=reportpivots&token=warmup-atlas-flow", language="text")
        else:
            render_wide_source_tab(
                "Noon & Manual Reports",
                reportpivots_df,
                reportpivots_metadata,
                "reportpivots",
                selected_vessels,
                selected_start,
                selected_end,
            )

    elif workspace == "15-Minute Operations":
        shippivots_df, shippivots_metadata = load_wide_source_for_view(
            "shippivots", username, password, token, auth_method, api_start_date, refresh,
            selected_vessels, selected_start, selected_end
        )
        if shippivots_metadata.get("needs_warmup"):
            st.info("No ShipPivots snapshot is available yet. Run the ShipPivots warmup URL first.")
            st.code("https://atlas-flow.streamlit.app/?warmup=1&force=1&source=shippivots&token=warmup-atlas-flow", language="text")
        else:
            render_wide_source_tab(
                "15-Minute Operations",
                shippivots_df,
                shippivots_metadata,
                "shippivots",
                selected_vessels,
                selected_start,
                selected_end,
            )

    elif workspace == "Descriptive Statistics":
        st.markdown('<div class="section-title">Descriptive Statistics</div>', unsafe_allow_html=True)
        st.caption("Analyze one source at a time. This avoids loading all large API tables into memory together.")
        selected_source = st.selectbox(
            "Source table",
            options=["Custom Analytics", "Noon & Manual Reports", "15-Minute Operations"],
            key="atlas_descriptive_source_selector",
        )
        if selected_source == "Custom Analytics":
            analysis_df = output_df.copy()
        elif selected_source == "Noon & Manual Reports":
            source_df, source_metadata = load_wide_source_for_view(
                "reportpivots", username, password, token, auth_method, api_start_date, refresh,
                selected_vessels, selected_start, selected_end
            )
            if source_metadata.get("needs_warmup"):
                st.info("No ReportPivots snapshot is available yet. Run the ReportPivots warmup URL first.")
                st.stop()
            analysis_df = build_wide_source_output_for_export("reportpivots", source_df, selected_vessels, selected_start, selected_end)
        else:
            source_df, source_metadata = load_wide_source_for_view(
                "shippivots", username, password, token, auth_method, api_start_date, refresh,
                selected_vessels, selected_start, selected_end
            )
            if source_metadata.get("needs_warmup"):
                st.info("No ShipPivots snapshot is available yet. Run the ShipPivots warmup URL first.")
                st.stop()
            analysis_df = build_wide_source_output_for_export("shippivots", source_df, selected_vessels, selected_start, selected_end)

        numeric_options = dataframe_numeric_options(analysis_df)
        if not numeric_options:
            st.info("The selected source table has no numeric columns to analyze.")
        else:
            metric_column = st.selectbox("Metric to analyze", options=numeric_options, key="atlas_descriptive_metric")
            group_options = ["None"] + dataframe_categorical_options(analysis_df)
            default_group_index = group_options.index("ShipName") if "ShipName" in group_options else 0
            group_column = st.selectbox("Optional group by", options=group_options, index=default_group_index, key="atlas_descriptive_group")
            stats_df = build_descriptive_statistics(analysis_df, metric_column)
            values = pd.to_numeric(analysis_df[metric_column], errors="coerce")
            render_metric_cards(
                [
                    ("Numeric Values", f"{values.notna().sum():,}", "numeric"),
                    ("Total", f"{values.sum(skipna=True):,.3f}", "total"),
                    ("Average", f"{values.mean(skipna=True):,.3f}", "average"),
                    ("Missing", f"{values.isna().sum():,}", "missing"),
                ]
            )
            st.markdown('<div class="section-title">Overall statistics</div>', unsafe_allow_html=True)
            st.dataframe(format_display_dataframe(stats_df), use_container_width=True, hide_index=True)
            if group_column != "None":
                grouped_df = build_grouped_descriptive_statistics(analysis_df, metric_column, group_column)
                if not grouped_df.empty:
                    st.markdown('<div class="section-title">Grouped statistics</div>', unsafe_allow_html=True)
                    st.dataframe(format_display_dataframe(grouped_df.head(100)), use_container_width=True, hide_index=True)
            datetime_column = detect_analysis_datetime_column(analysis_df)
            if datetime_column:
                trend_df = build_monthly_trend(analysis_df, metric_column, datetime_column)
                if not trend_df.empty:
                    st.markdown('<div class="section-title">Monthly trend</div>', unsafe_allow_html=True)
                    st.dataframe(format_display_dataframe(trend_df), use_container_width=True, hide_index=True)
                    st.line_chart(trend_df.set_index("Month")[["Sum", "Mean"]])
        del analysis_df
        gc.collect()

    elif workspace == "Export Center":
        st.markdown('<div class="section-title">AtlasFlow Export Center</div>', unsafe_allow_html=True)
        st.caption(
            "The full workbook is prepared on demand. ReportPivots and ShipPivots are loaded only while creating the workbook, then released."
        )
        render_metric_cards(
            [
                ("Custom Analytics Rows", f"{len(output_df):,}", "table_eye"),
                ("Noon & Manual Rows", "loaded on demand", "database_rows"),
                ("15-Minute Rows", "loaded on demand", "checked_columns"),
            ]
        )

        summary_group_fields = st.session_state.get("atlas_export_summary_groups", [])
        summary_value_fields = st.session_state.get("atlas_export_summary_values", [])
        if not isinstance(summary_group_fields, list):
            summary_group_fields = []
        if not isinstance(summary_value_fields, list):
            summary_value_fields = []
        summary_can_build = bool(summary_group_fields and summary_value_fields)

        multisource_signature_payload = "|".join([
            current_export_signature,
            ",".join(selected_vessels),
            selected_start.isoformat(),
            selected_end.isoformat(),
            ",".join(display_columns),
            ",".join(selected_variables),
            ",".join(summary_group_fields),
            ",".join(summary_value_fields),
            str(summary_aggregation),
        ])
        multisource_signature = sha256(multisource_signature_payload.encode("utf-8")).hexdigest()
        multisource_ready = (
            st.session_state.get("atlas_multisource_export_signature") == multisource_signature
            and "atlas_multisource_export_bytes" in st.session_state
        )

        if st.button("Prepare full AtlasFlow workbook", type="primary"):
            with st.spinner("Loading source snapshots and preparing workbook..."):
                reportpivots_df, reportpivots_metadata = load_wide_source_for_view(
                    "reportpivots", username, password, token, auth_method, api_start_date, refresh=False,
                    selected_vessels=selected_vessels, selected_start=selected_start, selected_end=selected_end
                )
                shippivots_df, shippivots_metadata = load_wide_source_for_view(
                    "shippivots", username, password, token, auth_method, api_start_date, refresh=False,
                    selected_vessels=selected_vessels, selected_start=selected_start, selected_end=selected_end
                )
                reportpivots_output_df = build_wide_source_output_for_export(
                    "reportpivots", reportpivots_df, selected_vessels, selected_start, selected_end
                ) if not reportpivots_metadata.get("needs_warmup") else pd.DataFrame()
                shippivots_output_df = build_wide_source_output_for_export(
                    "shippivots", shippivots_df, selected_vessels, selected_start, selected_end
                ) if not shippivots_metadata.get("needs_warmup") else pd.DataFrame()
                summary_analysis_df = pd.DataFrame()
                if summary_can_build:
                    summary_analysis_df = build_summary_analysis(
                        output_df,
                        group_fields=summary_group_fields,
                        value_fields=summary_value_fields,
                        aggregation=summary_aggregation,
                    )
                st.session_state["atlas_multisource_export_bytes"] = to_multisource_excel_bytes(
                    output_df,
                    summary_analysis_df if not summary_analysis_df.empty else None,
                    reportpivots_output_df,
                    shippivots_output_df,
                )
                st.session_state["atlas_multisource_export_signature"] = multisource_signature
                st.session_state["atlas_summary_analysis_df"] = summary_analysis_df
                st.session_state["atlas_export_signature"] = current_export_signature
                del reportpivots_df, shippivots_df, reportpivots_output_df, shippivots_output_df, summary_analysis_df
                clear_inactive_wide_sources(set())
                gc.collect()
            multisource_ready = True

        if multisource_ready:
            st.download_button(
                "Download full AtlasFlow workbook",
                data=st.session_state["atlas_multisource_export_bytes"],
                file_name="atlasflow_full_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.caption("Run the individual warmups first if the workbook is missing Noon & Manual or 15-Minute sheets.")

    elif workspace == "API Diagnostics":
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
                    "Loaded from snapshot",
                    "Snapshot saved at",
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
                    "Paging stop reason",
                    "Max page safety limit",
                ],
                "Value": [
                    ", ".join(selected_vessels),
                    api_start_date.isoformat(),
                    api_end_date.isoformat(),
                    selected_start.isoformat(),
                    selected_end.isoformat(),
                    metadata.get("loaded_at_utc", "-"),
                    metadata.get("loaded_at_local", "-"),
                    str(metadata.get("loaded_from_snapshot", False)),
                    metadata.get("snapshot_saved_at_utc", "-"),
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
                    metadata.get("paging_stop_reason", "-"),
                    f"{metadata.get('max_pages', MAX_ODATA_PAGES):,}",
                ],
            }
        )
        st.dataframe(diagnostics, use_container_width=True, hide_index=True)

        with st.expander("First API URL", expanded=False):
            st.code(metadata.get("first_url", "-"), language="text")

        st.markdown('<div class="section-title">Memory Audit</div>', unsafe_allow_html=True)
        audit_df = current_memory_audit_rows({
            "local.filtered_long_for_options": filtered_long_for_options,
            "local.pivot_df": pivot_df,
            "local.output_df": output_df,
        })
        st.dataframe(audit_df, use_container_width=True, hide_index=True)
        if st.button("Clear wide-source memory and Excel buffers"):
            clear_inactive_wide_sources(set())
            clear_stale_export_bytes(None)
            st.success("Released inactive wide-source DataFrames and export byte buffers from this session.")
            st.rerun()

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

    # Release the largest temporary views created during this run.
    del pivot_df, filtered_pivot_df
    gc.collect()


if __name__ == "__main__":
    main()
