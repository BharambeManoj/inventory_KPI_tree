import os
import re
import threading
import webbrowser
from datetime import datetime, date
from functools import wraps

import numpy as np
import openpyxl
import pandas as pd
from dotenv import load_dotenv
from flask import (
    Flask,
    Response,
    jsonify,
    redirect,
    request,
    send_from_directory,
    session,
    url_for,
)
from flask_cors import CORS
from waitress import serve

# ==============================================================================
# 1. ENVIRONMENT & SECURITY CONFIGURATION
# ==============================================================================
load_dotenv(override=True)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

HOST = os.getenv("HOST", "127.0.0.1").strip()
if HOST == "0.0.0.0" or not HOST:
    HOST = "127.0.0.1"  # Never bind 0.0.0.0 on Bosch corporate network

PORT = int(os.getenv("PORT", "5000").strip())
FLASK_SECRET_KEY = os.getenv("FLASK_SECRET_KEY", "bosch-chakan-vm-login-kpi-2026")
AUTH_BYPASS_DEV = os.getenv("AUTH_BYPASS_DEV", "true").strip().lower() in (
    "true",
    "1",
    "yes",
)

INVENTORY_ACTUAL_VS_FORECAST_PATH = os.getenv(
    "INVENTORY_ACTUAL_VS_FORECAST_PATH", "./Inventory Actual Vs Forecast.xlsx"
)
KPI_LOP3_PATH = os.getenv("KPI_LOP3_PATH", "./KPI LOP3.xlsx")

app = Flask(__name__, static_folder=None)
app.secret_key = FLASK_SECRET_KEY
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=False,
)

CORS(
    app,
    resources={
        r"/api/*": {
            "origins": [
                f"http://localhost:{PORT}",
                f"http://127.0.0.1:{PORT}",
                "http://localhost:5000",
                "http://127.0.0.1:5000",
            ]
        }
    },
    supports_credentials=True,
)

WHITELISTED_ASSETS = {
    "index.html",
    "bosch-logo.svg",
    "bosch-logo.png",
    "favicon.ico",
    "InventoryKPIDashboard.jsx",
}


@app.after_request
def add_no_cache_headers(response):
    """Prevents browser from caching old JSON or HTML responses."""
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


# ==============================================================================
# 2. AUTHENTICATION MIDDLEWARE (BYPASSED WHEN AUTH_BYPASS_DEV=true)
# ==============================================================================
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if AUTH_BYPASS_DEV:
            if "user" not in session:
                session["user"] = {
                    "name": "Manoj Bharambe",
                    "preferred_username": "bhm3chk@bosch.com",
                    "plant": "ChkP 1810",
                }
            else:
                session["user"]["plant"] = "ChkP 1810"
            return f(*args, **kwargs)
        if "user" not in session:
            return jsonify({"error": "Unauthorized", "login_url": "/login"}), 401
        return f(*args, **kwargs)

    return decorated_function


# ==============================================================================
# 3. STRICT 2-WORKBOOK EXCEL PARSER (ZERO DUMMY DATA / ZERO FALLBACK SEED)
#    Workbook 1: 'Main' & 'CF' in Inventory Actual Vs Forecast.xlsx
#    Workbook 2: 'Sheet1' in KPI LOP3.xlsx
# ==============================================================================
DATA_STORE = {
    "Main": pd.DataFrame(),
    "CF": pd.DataFrame(),
    "Target": pd.DataFrame(),
    "Target_Dict": {},
    "LOP3": pd.DataFrame(),
    "last_refresh": "",
    "inv_mtime": 0.0,
    "lop3_mtime": 0.0,
    "source_info": {},
}

MONTH_ABBR_MAP = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "sept": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}


def sanitize_str(val: str, max_len: int = 400) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r"[<>\"'`;]", "", s)
    return s[:max_len]


def format_cd05100_pct(delta_pct: float) -> str:
    """
    CD-05100 Relative Deviation Percentage Rule:
      - < 10%  -> 1 decimal place, no space before % (e.g., +9.4%, -3.2%, 0.0%)
      - >= 10% -> 0 decimal places, no space before % (e.g., +15%, -12%)
    """
    if delta_pct is None or np.isnan(delta_pct) or np.isinf(delta_pct):
        return "0.0%"
    val = float(delta_pct)
    abs_val = abs(val)
    sign = "+" if val > 0 else ("-" if val < 0 else "")
    if abs_val < 10.0:
        return f"{sign}{abs_val:.1f}%"
    return f"{sign}{round(abs_val):.0f}%"


def resolve_workbook_path(env_path: str, fallback_names: list) -> str:
    candidates = [os.path.abspath(os.path.join(BASE_DIR, env_path))]
    for name in fallback_names:
        candidates.append(os.path.abspath(os.path.join(BASE_DIR, name)))
        candidates.append(os.path.abspath(os.path.join(BASE_DIR, "data", name)))

    for c in candidates:
        if os.path.isfile(c):
            return c

    lower_targets = {os.path.basename(env_path).lower()} | {
        n.lower() for n in fallback_names
    }
    for folder in [BASE_DIR, os.path.join(BASE_DIR, "data")]:
        if os.path.isdir(folder):
            for entry in os.listdir(folder):
                if entry.lower() in lower_targets:
                    full = os.path.join(folder, entry)
                    if os.path.isfile(full):
                        return full

    return os.path.abspath(os.path.join(BASE_DIR, env_path))


def normalize_vs(val) -> str:
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip().upper()
    if s in ("ESP/ABS", "ESP-ABS", "ESP_ABS", "ESP ABS"):
        return "ESP-ABS"
    return s


def parse_excel_month_header(col_header):
    """
    Parses Excel column headers like datetime(2024, 1, 1) or 'Jan-24', 'Feb-25'
    into (month_int, year_int, month_year_str 'MM-YYYY', display_label 'Jan-24').
    """
    if isinstance(col_header, (datetime, date, pd.Timestamp)):
        m = int(col_header.month)
        y = int(col_header.year)
        return m, y, f"{m:02d}-{y}", col_header.strftime("%b-%y")

    if col_header is None or pd.isna(col_header):
        return None

    s = str(col_header).strip()
    if not s:
        return None

    m1 = re.match(r"^([A-Za-z]{3,4})[-_ /\.](\d{2,4})$", s)
    if m1:
        m_name = m1.group(1).lower()[:3]
        yr_raw = m1.group(2)
        if m_name in MONTH_ABBR_MAP:
            m_int = MONTH_ABBR_MAP[m_name]
            y_int = int(yr_raw) + 2000 if len(yr_raw) == 2 else int(yr_raw)
            return m_int, y_int, f"{m_int:02d}-{y_int}", f"{m_name.capitalize()}-{str(y_int)[-2:]}"

    m2 = re.match(r"^(\d{1,2})[-/](\d{4})$", s)
    if m2:
        m_int = int(m2.group(1))
        y_int = int(m2.group(2))
        if 1 <= m_int <= 12:
            dt = datetime(y_int, m_int, 1)
            return m_int, y_int, f"{m_int:02d}-{y_int}", dt.strftime("%b-%y")

    m3 = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if m3:
        y_int = int(m3.group(1))
        m_int = int(m3.group(2))
        if 1 <= m_int <= 12:
            dt = datetime(y_int, m_int, 1)
            return m_int, y_int, f"{m_int:02d}-{y_int}", dt.strftime("%b-%y")

    return None


def clean_numeric(val) -> float:
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float, np.number)):
        if np.isnan(val) or np.isinf(val):
            return 0.0
        return float(val)
    s = str(val).strip()
    if s in ("", "-", "--", "N/A", "nan", "None", "#REF!", "#DIV/0!", "#VALUE!", "#NAME?"):
        return 0.0
    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def load_forecast_workbook_strictly(inv_path: str):
    """
    Strictly reads:
      1. 'Main' sheet from 'Inventory Actual Vs Forecast.xlsx' for monthly actuals/forecast values.
      2. 'CF' sheet from 'Inventory Actual Vs Forecast.xlsx' for CF02, CF05, CF09 strictly without changing values.
    No dummy data is generated.
    """
    if not os.path.isfile(inv_path):
        raise FileNotFoundError(f"Forecast Excel file not found at: {inv_path}")

    wb = openpyxl.load_workbook(inv_path, data_only=True, read_only=True)
    sheet_map = {s.strip().lower(): s for s in wb.sheetnames}
    if "main" not in sheet_map:
        wb.close()
        raise ValueError(f"'Main' sheet not found in {inv_path}. Available sheets: {wb.sheetnames}")
    if "cf" not in sheet_map:
        wb.close()
        raise ValueError(f"'CF' sheet not found in {inv_path}. Available sheets: {wb.sheetnames}")

    main_sheet_name = sheet_map["main"]
    cf_sheet_name = sheet_map["cf"]

    ws_main = wb[main_sheet_name]
    rows_iter = list(ws_main.iter_rows(values_only=True))

    ws_cf = wb[cf_sheet_name]
    cf_rows_iter = list(ws_cf.iter_rows(values_only=True))

    target_sheet_dict = {}
    if "target" in sheet_map:
        ws_tgt = wb[sheet_map["target"]]
        for r in list(ws_tgt.iter_rows(values_only=True))[1:]:
            if r and len(r) >= 4 and r[0] is not None:
                try:
                    yr = int(r[0])
                    vs_clean = normalize_vs(r[1])
                    inv_type = str(r[2]).strip().lower()
                    val = float(r[3])
                    target_sheet_dict[(yr, vs_clean, inv_type)] = val
                except Exception:
                    pass

    wb.close()

    if not rows_iter:
        raise ValueError(f"'{main_sheet_name}' sheet in {inv_path} is empty.")
    if not cf_rows_iter:
        raise ValueError(f"'{cf_sheet_name}' sheet in {inv_path} is empty.")

    # 1. Parse 'Main' sheet (dynamically locate Category & Value Stream columns)
    header_idx = 0
    col_main_cat = 0
    col_main_vs = 1
    for idx, r in enumerate(rows_iter[:15]):
        row_strs = [str(c).strip().lower() if c is not None else "" for c in r[:15]]
        if any("categor" in s for s in row_strs) or any("value stre" in s for s in row_strs):
            header_idx = idx
            for c_i, s_val in enumerate(row_strs):
                if "categor" in s_val:
                    col_main_cat = c_i
                elif "value stre" in s_val or s_val == "vs":
                    col_main_vs = c_i
            break

    header_row = rows_iter[header_idx]
    month_cols = {}
    for c_idx, h_val in enumerate(header_row):
        if c_idx <= max(col_main_cat, col_main_vs):
            continue
        parsed_m = parse_excel_month_header(h_val)
        if parsed_m:
            month_cols[c_idx] = parsed_m

    if not month_cols:
        raise ValueError(f"No monthly columns (e.g. Jan-24, Jan-25) found in '{main_sheet_name}' header row.")

    main_records = []
    last_main_cat = None
    for r in rows_iter[header_idx + 1 :]:
        if not r or len(r) <= max(col_main_cat, col_main_vs):
            continue
        raw_cat = r[col_main_cat]
        raw_vs = r[col_main_vs]
        if raw_cat is not None and str(raw_cat).strip() != "":
            last_main_cat = str(raw_cat).strip()
        if last_main_cat is None or raw_vs is None:
            continue

        cat_str = last_main_cat
        vs_str = normalize_vs(raw_vs)
        if not cat_str or not vs_str or cat_str.lower().startswith("categor"):
            continue

        if cat_str.lower() in ("ofggit", "0fggit"):
            cat_str = "0FGGit"

        for c_idx, (m_int, y_int, my_str, disp_lbl) in month_cols.items():
            val = clean_numeric(r[c_idx] if c_idx < len(r) else 0.0)
            main_records.append(
                {
                    "Attribute": cat_str,
                    "CAT-YEAR": f"{cat_str}-{y_int}-{vs_str}",
                    "Category": cat_str,
                    "Month": m_int,
                    "Month year": my_str,
                    "Month Label": disp_lbl,
                    "Value": round(val, 2),
                    "Value Stream": vs_str,
                    "Year": y_int,
                }
            )

    df_main = pd.DataFrame(main_records)
    if df_main.empty:
        raise ValueError(f"No valid data rows found in '{main_sheet_name}'.")

    # 2. Strictly Parse 'CF' sheet (Year | Category | Value Stream | CF02 | CF05 | CF09)
    cf_header_idx = 0
    for idx, r in enumerate(cf_rows_iter[:15]):
        row_strs = [str(c).strip().upper() if c is not None else "" for c in r[:10]]
        if "CF02" in row_strs or "CF05" in row_strs or "CF09" in row_strs:
            cf_header_idx = idx
            break

    cf_header = cf_rows_iter[cf_header_idx]
    col_year, col_cat, col_vs, col_cf02, col_cf05, col_cf09 = None, None, None, None, None, None
    for c_idx, h_val in enumerate(cf_header):
        if h_val is None:
            continue
        h_str = str(h_val).strip().lower()
        if h_str == "year" or "year" in h_str:
            if col_year is None:
                col_year = c_idx
        elif "categor" in h_str:
            if col_cat is None:
                col_cat = c_idx
        elif "value stre" in h_str or h_str == "vs":
            if col_vs is None:
                col_vs = c_idx
        elif h_str == "cf02" or "cf02" in h_str:
            col_cf02 = c_idx
        elif h_str == "cf05" or "cf05" in h_str:
            col_cf05 = c_idx
        elif h_str == "cf09" or "cf09" in h_str:
            col_cf09 = c_idx

    if col_year is None:
        col_year = 0
    if col_cat is None:
        col_cat = 1
    if col_vs is None:
        col_vs = 2
    if col_cf02 is None:
        col_cf02 = 3
    if col_cf05 is None:
        col_cf05 = 4
    if col_cf09 is None:
        col_cf09 = 5

    cf_records = []
    last_seen_year = 2025
    for r in cf_rows_iter[cf_header_idx + 1 :]:
        if not r or len(r) <= max(col_cat, col_vs):
            continue
        raw_cat = r[col_cat] if col_cat < len(r) else None
        raw_vs = r[col_vs] if col_vs < len(r) else None
        if raw_cat is None or raw_vs is None:
            continue

        cat_str = str(raw_cat).strip()
        vs_str = normalize_vs(raw_vs)
        if not cat_str or not vs_str or cat_str.lower().startswith("categor"):
            continue

        if cat_str.lower() in ("ofggit", "0fggit"):
            cat_str = "0FGGit"

        raw_yr = r[col_year] if col_year is not None and col_year < len(r) else None
        yr_int = last_seen_year
        if raw_yr is not None and str(raw_yr).strip() != "":
            try:
                yr_int = int(float(str(raw_yr).strip()))
                last_seen_year = yr_int
            except ValueError:
                yr_int = last_seen_year

        cf02_val = clean_numeric(r[col_cf02] if col_cf02 < len(r) else 0.0)
        cf05_val = clean_numeric(r[col_cf05] if col_cf05 < len(r) else 0.0)
        cf09_val = clean_numeric(r[col_cf09] if col_cf09 < len(r) else 0.0)

        cf_records.append(
            {
                "Category": cat_str,
                "Value Stream": vs_str,
                "CF02": cf02_val,
                "CF05": cf05_val,
                "CF09": cf09_val,
                "CAT-YEAR": f"{cat_str}-{yr_int}-{vs_str}",
                "Year": yr_int,
            }
        )

    df_cf = pd.DataFrame(cf_records)
    if df_cf.empty:
        raise ValueError(f"No valid data rows found in '{cf_sheet_name}' sheet.")

    # Derive Target table from Category == 'Target' in Main sheet
    target_records = []
    df_target_rows = df_main[df_main["Category"].str.strip().str.lower() == "target"]
    for yr in sorted(df_main["Year"].unique().tolist()):
        for vs_item in df_main["Value Stream"].unique():
            sub_t = df_target_rows[
                (df_target_rows["Year"] == yr) & (df_target_rows["Value Stream"] == vs_item)
            ]
            sub_t_nz = sub_t[sub_t["Value"] != 0]
            if not sub_t_nz.empty:
                pa_val = round(float(sub_t_nz["Value"].iloc[0]), 2)
            elif not sub_t.empty:
                pa_val = round(float(sub_t["Value"].mean()), 2)
            else:
                pa_val = 0.0
            for cat_item in df_main["Category"].unique():
                target_records.append(
                    {
                        "CAT-YEAR": f"{cat_item}-{yr}-{vs_item}",
                        "Category": cat_item,
                        "Value Stream": vs_item,
                        "Year": int(yr),
                        "Inv Type": cat_item,
                        "PA": pa_val,
                    }
                )
    df_target = pd.DataFrame(target_records)

    return df_main, df_cf, df_target, target_sheet_dict, main_sheet_name, cf_sheet_name


def load_lop3_sheet_strictly(lop3_path: str):
    """
    Strictly reads ONLY 'Sheet1' from 'KPI LOP3.xlsx'.
    No other sheet is touched. Zero dummy planners or categories exist anywhere.
    """
    if not os.path.isfile(lop3_path):
        raise FileNotFoundError(f"KPI LOP3 Excel file not found at: {lop3_path}")

    wb = openpyxl.load_workbook(lop3_path, data_only=True, read_only=True)
    sheet_map = {s.strip().lower(): s for s in wb.sheetnames}
    lop3_sheet_name = sheet_map.get("sheet1", wb.sheetnames[0])
    ws = wb[lop3_sheet_name]

    rows_iter = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows_iter:
        raise ValueError(f"'{lop3_sheet_name}' in {lop3_path} is empty.")

    header_idx = 0
    for idx, r in enumerate(rows_iter[:10]):
        row_strs = [str(c).strip().lower() if c is not None else "" for c in r[:6]]
        if "planner" in row_strs or "value stream" in row_strs:
            header_idx = idx
            break

    header_row = rows_iter[header_idx]
    month_cols = {}
    for c_idx, h_val in enumerate(header_row):
        if c_idx < 5:
            continue
        parsed_m = parse_excel_month_header(h_val)
        if parsed_m:
            month_cols[c_idx] = parsed_m

    if not month_cols:
        raise ValueError(f"No monthly columns found in '{lop3_sheet_name}' of {lop3_path}.")

    lop3_records = []
    for r in rows_iter[header_idx + 1 :]:
        if not r or len(r) < 5:
            continue
        raw_vs = r[0]
        raw_prod = r[1]
        raw_planner = r[2]
        raw_cat = r[3]
        raw_target = r[4]

        if raw_vs is None or raw_cat is None:
            continue

        vs_str = normalize_vs(raw_vs)
        prod_str = str(raw_prod).strip() if raw_prod is not None else ""
        planner_str = str(raw_planner).strip() if raw_planner is not None else ""
        cat_str = str(raw_cat).strip()
        target_val = clean_numeric(raw_target)

        if not vs_str or vs_str.lower() == "value stream":
            continue

        for c_idx, (m_int, y_int, my_str, disp_lbl) in month_cols.items():
            v = clean_numeric(r[c_idx] if c_idx < len(r) else 0.0)
            lop3_records.append(
                {
                    "Value Stream": vs_str,
                    "Product": prod_str,
                    "Planner": planner_str,
                    "Category": cat_str,
                    "Target": round(target_val, 2),
                    "Month": m_int,
                    "Year": y_int,
                    "Month year": my_str,
                    "Month Label": disp_lbl,
                    "Value": round(v, 2),
                }
            )

    df_lop3 = pd.DataFrame(lop3_records)
    if df_lop3.empty:
        raise ValueError(f"No valid rows parsed from '{lop3_sheet_name}' in {lop3_path}.")

    return df_lop3, lop3_sheet_name


def ensure_fresh_sheets():
    """Reloads Main and Sheet1 automatically if either Excel file was modified on disk."""
    inv_resolved = resolve_workbook_path(
        INVENTORY_ACTUAL_VS_FORECAST_PATH,
        ["Inventory Actual Vs Forecast.xlsx", "inventory actual vs forecast.xlsx"],
    )
    lop3_resolved = resolve_workbook_path(
        KPI_LOP3_PATH,
        ["KPI LOP3.xlsx", "kpi lop3.xlsx"],
    )

    inv_mtime = os.path.getmtime(inv_resolved) if os.path.isfile(inv_resolved) else -1.0
    lop3_mtime = os.path.getmtime(lop3_resolved) if os.path.isfile(lop3_resolved) else -1.0

    if (
        DATA_STORE["Main"].empty
        or DATA_STORE["LOP3"].empty
        or inv_mtime != DATA_STORE["inv_mtime"]
        or lop3_mtime != DATA_STORE["lop3_mtime"]
    ):
        reload_all_sheets()


def reload_all_sheets():
    inv_resolved = resolve_workbook_path(
        INVENTORY_ACTUAL_VS_FORECAST_PATH,
        ["Inventory Actual Vs Forecast.xlsx", "inventory actual vs forecast.xlsx"],
    )
    lop3_resolved = resolve_workbook_path(
        KPI_LOP3_PATH,
        ["KPI LOP3.xlsx", "kpi lop3.xlsx"],
    )

    df_main, df_cf, df_target, target_sheet_dict, main_sname, cf_sname = load_forecast_workbook_strictly(inv_resolved)
    df_lop3, lop3_sname = load_lop3_sheet_strictly(lop3_resolved)

    DATA_STORE["Main"] = df_main
    DATA_STORE["CF"] = df_cf
    DATA_STORE["Target"] = df_target
    DATA_STORE["Target_Dict"] = target_sheet_dict
    DATA_STORE["LOP3"] = df_lop3
    DATA_STORE["inv_mtime"] = os.path.getmtime(inv_resolved)
    DATA_STORE["lop3_mtime"] = os.path.getmtime(lop3_resolved)
    DATA_STORE["last_refresh"] = datetime.now().strftime("%d %b %Y, %I:%M:%S %p")
    DATA_STORE["source_info"] = {
        "inventory_file": inv_resolved,
        "lop3_file": lop3_resolved,
        "sheets_used": {
            "inventory_sheet": main_sname,
            "cf_sheet": cf_sname,
            "lop3_sheet": lop3_sname,
        },
        "main_categories": sorted(df_main["Category"].unique().tolist()),
        "cf_categories": sorted(df_cf["Category"].unique().tolist()),
        "lop3_planners": sorted(df_lop3["Planner"].unique().tolist()),
        "lop3_products": sorted(df_lop3["Product"].unique().tolist()),
    }


reload_all_sheets()


# ==============================================================================
# 4. KPI TREE COMPUTATION ENGINE (STRICTLY MAIN + CF + SHEET1)
# ==============================================================================
def find_matching_category(available_cats: list, candidates: list) -> str:
    lower_map = {c.lower().strip(): c for c in available_cats}
    for cand in candidates:
        if cand.lower().strip() in lower_map:
            return lower_map[cand.lower().strip()]
    return candidates[0]


def compute_dashboard_payload(filters: dict) -> dict:
    ensure_fresh_sheets()

    df_main = DATA_STORE["Main"]
    df_cf = DATA_STORE["CF"]
    df_target = DATA_STORE["Target"]
    target_sheet_dict = DATA_STORE.get("Target_Dict", {})
    df_lop3 = DATA_STORE["LOP3"]

    available_vs = sorted(df_main["Value Stream"].dropna().unique().tolist())
    preferred_order = ["ESP-ABS", "VDS", "ATC", "NBS", "OSS", "VM-IN"]
    available_vs = [v for v in preferred_order if v in available_vs] + [
        v for v in available_vs if v not in preferred_order
    ]

    vs = normalize_vs(filters.get("value_stream", "ESP-ABS"))
    if vs not in available_vs and available_vs:
        vs = available_vs[0]

    df_main_vs = df_main[df_main["Value Stream"] == vs]

    available_years = sorted(df_main_vs["Year"].dropna().unique().astype(int).tolist())
    raw_year_req = sanitize_str(filters.get("year", ""))
    req_years = [
        int(y.strip())
        for y in raw_year_req.split(",")
        if y.strip().isdigit() and int(y.strip()) in available_years
    ]
    if req_years:
        selected_years = sorted(list(dict.fromkeys(req_years)))
    else:
        non_zero_yrs = sorted(
            df_main_vs[df_main_vs["Value"] != 0]["Year"].dropna().unique().astype(int).tolist()
        )
        default_yr = non_zero_yrs[-1] if non_zero_yrs else (available_years[-1] if available_years else 2025)
        selected_years = [default_yr]

    year_filter_str = ",".join(str(y) for y in selected_years)

    df_main_vs_yr = df_main_vs[df_main_vs["Year"].isin(selected_years)].sort_values(["Year", "Month"])
    available_month_years = df_main_vs_yr["Month year"].dropna().unique().tolist()

    raw_month_req = sanitize_str(filters.get("month_year", ""))
    raw_bench_req = sanitize_str(filters.get("benchmark", ""))

    # Primary Month year filter (scopes both KPI Tree and Data Tables)
    req_months = [
        m.strip()
        for m in raw_month_req.split(",")
        if m.strip() and m.strip() in available_month_years
    ]
    selected_months = list(dict.fromkeys(req_months))
    month_year = ",".join(selected_months)

    # Sidebar Benchmark Comparison filter (CF02, CF05, CF09, PA)
    req_benchmarks = [
        b.strip()
        for b in raw_bench_req.split(",")
        if b.strip() in ("CF02", "CF05", "CF09", "PA")
    ]
    selected_benchmarks = list(dict.fromkeys(req_benchmarks))
    benchmark = ",".join(selected_benchmarks)

    # Non-zero months in the selected year(s) for KPI Tree & Summary Cards
    df_main_non_target_yr = df_main_vs_yr[
        df_main_vs_yr["Category"].str.strip().str.lower() != "target"
    ]
    nz_months_yr = (
        df_main_non_target_yr[df_main_non_target_yr["Value"] != 0]["Month year"]
        .unique()
        .tolist()
    )
    if not nz_months_yr:
        nz_months_yr = df_main_vs_yr[df_main_vs_yr["Value"] != 0]["Month year"].unique().tolist()

    # KPI Tree & Dashboard strictly scope to selected month(s) when set, otherwise all non-zero months
    effective_tree_months = selected_months if selected_months else (nz_months_yr if nz_months_yr else available_month_years)

    # Filter LOP3 strictly from Sheet1 ONLY for the selected Value Stream and Year(s)
    # STRICT FILTERING: Never fall back to different years or months to prevent showing false data!
    df_lop3_vs = df_lop3[df_lop3["Value Stream"] == vs].copy()
    df_lop3_vs_yr = df_lop3_vs[df_lop3_vs["Year"].isin(selected_years)].copy()

    lop3_months_available = df_lop3_vs_yr["Month year"].unique().tolist()
    valid_lop3_tree_months = [m for m in effective_tree_months if m in lop3_months_available]
    has_lop3_data = bool(valid_lop3_tree_months and not df_lop3_vs_yr.empty)

    active_lop3_month = (
        valid_lop3_tree_months[0]
        if len(valid_lop3_tree_months) == 1
        else (
            f"{len(valid_lop3_tree_months)} Months ({year_filter_str})"
            if valid_lop3_tree_months
            else (selected_months[0] if selected_months else year_filter_str)
        )
    )

    df_lop3_base_meta = (
        df_lop3_vs[["Value Stream", "Product", "Planner", "Category"]]
        .drop_duplicates()
        .copy()
    )

    if has_lop3_data:
        df_lop3_tree_month = df_lop3_vs_yr[df_lop3_vs_yr["Month year"].isin(valid_lop3_tree_months)].copy()
        if len(valid_lop3_tree_months) > 1 and not df_lop3_tree_month.empty:
            df_lop3_tree_month = (
                df_lop3_tree_month.groupby(
                    ["Value Stream", "Product", "Planner", "Category"],
                    as_index=False,
                    sort=False,
                )
                .agg({"Target": "mean", "Value": "mean"})
            )
            df_lop3_tree_month["Month year"] = active_lop3_month
            df_lop3_tree_month["Target"] = df_lop3_tree_month["Target"].round(2)
            df_lop3_tree_month["Value"] = df_lop3_tree_month["Value"].round(2)
    else:
        df_lop3_tree_month = df_lop3_base_meta.copy()
        df_lop3_tree_month["Month year"] = active_lop3_month
        df_lop3_tree_month["Target"] = 0.0
        df_lop3_tree_month["Value"] = 0.0

    raw_prod_req = sanitize_str(filters.get("product", "ALL")) or "ALL"
    raw_plan_req = sanitize_str(filters.get("planner", "ALL")) or "ALL"
    raw_cat_req = sanitize_str(filters.get("category", "ALL")) or "ALL"

    available_products = df_lop3_base_meta["Product"].dropna().unique().tolist()
    sel_products = [
        p.strip()
        for p in raw_prod_req.split(",")
        if p.strip() and p.strip() != "ALL" and p.strip() in available_products
    ]
    product_filter = ",".join(sel_products) if sel_products else "ALL"

    df_lop3_prod_opts = (
        df_lop3_base_meta[df_lop3_base_meta["Product"].isin(sel_products)]
        if sel_products
        else df_lop3_base_meta
    )
    available_planners = df_lop3_prod_opts["Planner"].dropna().unique().tolist()
    sel_planners = [
        p.strip()
        for p in raw_plan_req.split(",")
        if p.strip() and p.strip() != "ALL" and p.strip() in available_planners
    ]
    planner_filter = ",".join(sel_planners) if sel_planners else "ALL"

    df_lop3_plan_opts = (
        df_lop3_prod_opts[df_lop3_prod_opts["Planner"].isin(sel_planners)]
        if sel_planners
        else df_lop3_prod_opts
    )
    available_categories = df_lop3_plan_opts["Category"].dropna().unique().tolist()
    sel_categories = [
        c.strip()
        for c in raw_cat_req.split(",")
        if c.strip() and c.strip() != "ALL" and c.strip() in available_categories
    ]
    category_filter = ",".join(sel_categories) if sel_categories else "ALL"

    df_lop3_filtered = df_lop3_tree_month.copy()
    if sel_products:
        df_lop3_filtered = df_lop3_filtered[df_lop3_filtered["Product"].isin(sel_products)]
    if sel_planners:
        df_lop3_filtered = df_lop3_filtered[df_lop3_filtered["Planner"].isin(sel_planners)]
    if sel_categories:
        df_lop3_filtered = df_lop3_filtered[df_lop3_filtered["Category"].isin(sel_categories)]

    df_main_slice = df_main_vs_yr[df_main_vs_yr["Month year"].isin(effective_tree_months)]
    df_cf_vs = df_cf[df_cf["Value Stream"] == vs]
    df_cf_slice = df_cf_vs[df_cf_vs["Year"].isin(selected_years)]
    main_cats_present = df_main_vs["Category"].unique().tolist()
    cf_cats_present = df_cf_slice["Category"].unique().tolist() if not df_cf_slice.empty else df_cf_vs["Category"].unique().tolist()
    all_cats_present = list(dict.fromkeys(main_cats_present + cf_cats_present))

    avg_tns_cat = find_matching_category(all_cats_present, ["Avg TNS", "TNS"])
    avg_tns_row = df_main_slice[df_main_slice["Category"] == avg_tns_cat]
    current_avg_tns = float(avg_tns_row["Value"].mean()) if not avg_tns_row.empty else 0.0

    # Determine Target from Main sheet 'Target' row for the selected Year(s)
    target_cat_main = find_matching_category(main_cats_present, ["Target"])
    ototal_cat_cf = find_matching_category(cf_cats_present, ["ototal", "oTotal", "dTotal", "aTotal", "Target"])

    df_month_target_slice = df_main_vs_yr[
        (df_main_vs_yr["Month year"].isin(effective_tree_months))
        & (df_main_vs_yr["Category"].str.strip().str.lower() == target_cat_main.strip().lower())
    ]
    nz_month_targets = (
        df_month_target_slice[df_month_target_slice["Value"] > 0]["Value"]
        if not df_month_target_slice.empty
        else pd.Series(dtype=float)
    )
    if not nz_month_targets.empty:
        pa_sheet_target_val = float(nz_month_targets.mean())
    else:
        all_yr_targets = df_main_vs_yr[
            df_main_vs_yr["Category"].str.strip().str.lower() == target_cat_main.strip().lower()
        ]["Value"]
        nz_all = all_yr_targets[all_yr_targets > 0]
        pa_sheet_target_val = float(nz_all.mean()) if not nz_all.empty else 0.0

    # User Requirement: Compulsory take targets from the Main sheet of the forecast workbook
    # Do NOT consider CF02, CF05, CF09 as the target
    selected_target_val = pa_sheet_target_val
    active_target_label = "Target"

    def get_metric_row(candidates: list, custom_main_slice=None) -> dict:
        active_slice = custom_main_slice if custom_main_slice is not None else df_main_slice
        cat_code = find_matching_category(all_cats_present, candidates)
        cf_cat_code = find_matching_category(cf_cats_present, [cat_code] + candidates) if cf_cats_present else cat_code
        cf_r = df_cf_slice[df_cf_slice["Category"] == cf_cat_code]
        cf02 = float(cf_r["CF02"].mean()) if not cf_r.empty else 0.0
        cf05 = float(cf_r["CF05"].mean()) if not cf_r.empty else 0.0
        cf09 = float(cf_r["CF09"].mean()) if not cf_r.empty else 0.0

        main_cat_code = find_matching_category(main_cats_present, [cat_code] + candidates) if main_cats_present else cat_code
        m_r = active_slice[active_slice["Category"] == main_cat_code]
        if not m_r.empty:
            nz_m_r = m_r[m_r["Value"] != 0]["Value"]
            val = float(nz_m_r.mean()) if not nz_m_r.empty else float(m_r["Value"].mean())
        else:
            val = 0.0

        cat_lower = cat_code.strip().lower()
        is_ototal = cat_lower in ("ototal", "dtotal", "atotal")
        is_stotal = cat_lower in ("stotal", "total stock", "sstock")
        is_tns = cat_lower in ("avg tns", "tns")

        # Determine target from Main file:
        inv_type = None
        if is_ototal or is_stotal:
            inv_type = "overall"
        elif "rm" in cat_lower:
            inv_type = "rm"
        elif "wip" in cat_lower:
            inv_type = "wip"
        elif "fg" in cat_lower or "0fg" in cat_lower:
            inv_type = "fg"
        elif "indirect" in cat_lower or "ind" in cat_lower:
            inv_type = "ind. matl"

        tgt_yr = selected_years[0] if selected_years else 2026
        tgt_days = 0.0
        if inv_type and inv_type != "overall":
            tgt_days = target_sheet_dict.get((tgt_yr, vs, inv_type), 0.0)

        # Plant Root (ototal / Stotal) strictly takes target from the Main sheet Target row
        if is_ototal or is_stotal or tgt_days == 0.0:
            tgt_days = pa_sheet_target_val

        is_dio_metric = (
            is_ototal
            or cat_lower.startswith("o")
            or cat_lower.startswith("d")
            or cat_lower.startswith("a")
            or "dio" in cat_lower
        )
        is_stock_metric = (
            is_stotal
            or cat_lower.startswith("s")
            or "stock" in cat_lower
        )

        if is_tns:
            target_val = current_avg_tns
        elif is_dio_metric:
            target_val = tgt_days if tgt_days > 0 else pa_sheet_target_val
        elif is_stock_metric:
            active_dio_tgt = tgt_days if tgt_days > 0 else pa_sheet_target_val
            target_val = (active_dio_tgt / 30.0) * current_avg_tns if active_dio_tgt > 0 and current_avg_tns > 0 else 0.0
        else:
            target_val = pa_sheet_target_val

        # Target from the Main file is the benchmark for performance comparison
        bench_val = target_val

        if bench_val > 0:
            if is_tns:
                status = "positive" if round(val, 2) >= round(bench_val, 2) else "negative"
            else:
                status = "negative" if round(val, 2) > round(bench_val, 2) else "positive"
        else:
            status = "neutral"

        raw_delta_pct = (
            ((val - bench_val) / bench_val) * 100.0
            if bench_val > 0
            else 0.0
        )
        delta_pct = round(raw_delta_pct, 1)
        delta_pct_formatted = format_cd05100_pct(raw_delta_pct)

        return {
            "category": cat_code,
            "cf02": round(cf02, 2),
            "cf05": round(cf05, 2),
            "cf09": round(cf09, 2),
            "pa": round(target_val, 2),
            "target": round(target_val, 2),
            "target_val": round(target_val, 2),
            "values": round(val, 2),
            "benchmark_val": round(bench_val, 2),
            "benchmark_label": "Target",
            "delta_pct": delta_pct,
            "delta_pct_formatted": delta_pct_formatted,
            "status": status,
        }

    def build_tier12_card(card_id: str, title: str, tier: str, candidate_Rows: list) -> dict:
        rows = [get_metric_row(c_list) for c_list in candidate_Rows]
        card_status = "neutral"
        ototal_statuses = [
            r["status"]
            for r in rows
            if r["category"].strip().lower() in ("ototal", "dtotal", "atotal")
        ]
        if ototal_statuses:
            card_status = ototal_statuses[0]
        else:
            row_statuses = [r["status"] for r in rows if r["status"] in ("negative", "positive")]
            if "negative" in row_statuses:
                card_status = "negative"
            elif "positive" in row_statuses:
                card_status = "positive"

        return {
            "id": card_id,
            "title": title,
            "tier": tier,
            "status": card_status,
            "rows": rows,
            "planners": [],
        }

    vs_display = "ESP/ABS" if vs == "ESP-ABS" else vs
    tier1_root = build_tier12_card(
        "NODE_PLANT_ROOT",
        f"{vs_display}_DIO",
        "PLANT KPI",
        [
            ["Stotal", "sTotal", "Total Stock"],
            ["Avg TNS", "TNS"],
            ["ototal", "oTotal", "dTotal", "aTotal"],
        ],
    )

    tier2a_cards = [
        build_tier12_card(
            "NODE_FG_INV",
            "FG INVENTORY",
            "MONITORING KPI",
            [["sFG"], ["oFG", "dFG", "aFG"]],
        ),
        build_tier12_card(
            "NODE_WIP",
            "WIP",
            "MONITORING KPI",
            [["sWIP"], ["oWIP", "dWIP", "aWIP"]],
        ),
        build_tier12_card(
            "NODE_RM_INV",
            "RM INVENTORY",
            "MONITORING KPI",
            [["oRM", "dRM", "aRM"], ["sRM"]],
        ),
    ]

    tier2b_cards = [
        build_tier12_card(
            "NODE_FG_STORE",
            "FG STORE",
            "MONITORING KPI",
            [["oFGstore", "dFGstore", "aFGstore"], ["sFGstore"]],
        ),
        build_tier12_card(
            "NODE_FG_GIT",
            "FG GIT",
            "MONITORING KPI",
            [["0FGGit", "oFGGit", "dFGGit", "aFGGit"], ["sFGGit"]],
        ),
        build_tier12_card(
            "NODE_RM_STORE",
            "RM STORE",
            "MONITORING KPI",
            [["oRMstore", "dRMstore", "aRMstore"], ["sRMstore"]],
        ),
        build_tier12_card(
            "NODE_RM_GIT",
            "RM GIT",
            "MONITORING KPI",
            [["oRMGit", "dRMGit", "aRMGit"], ["sRMGit"]],
        ),
    ]

    if "4W" in available_products and "2W" in available_products:
        branch_products = ["4W", "2W"]
    elif len(available_products) >= 2:
        branch_products = available_products[:2]
    elif len(available_products) == 1:
        branch_products = [available_products[0], ""]
    else:
        branch_products = ["4W", "2W"]

    def build_lop3_branch(slot_key: str, actual_prod: str) -> dict:
        if not actual_prod:
            return {
                "id": f"NODE_LOP3_{slot_key}",
                "slot": slot_key,
                "product": "-",
                "title": "No Product",
                "val_lop3": 0.0,
                "target_val": 0.0,
                "lop3_dio": 0.0,
                "status": "neutral",
                "children": [],
            }

        df_p = df_lop3_filtered[df_lop3_filtered["Product"] == actual_prod]
        total_val = round(float(df_p["Value"].sum()), 2) if (has_lop3_data and not df_p.empty) else 0.0
        total_target = round(float(df_p["Target"].sum()), 2) if (has_lop3_data and not df_p.empty) else 0.0
        parent_dio = (
            round((total_val / current_avg_tns) * 30.0, 2)
            if (has_lop3_data and current_avg_tns > 0)
            else 0.0
        )

        if has_lop3_data and total_target > 0:
            parent_status = "negative" if total_val > total_target else "positive"
        else:
            parent_status = "neutral"

        children = []
        if not df_p.empty:
            for idx, (cat_name, grp) in enumerate(df_p.groupby("Category", sort=False)):
                c_val = round(float(grp["Value"].sum()), 2) if has_lop3_data else 0.0
                c_tgt = round(float(grp["Target"].sum()), 2) if has_lop3_data else 0.0
                c_dio = (
                    round((c_val / current_avg_tns) * 30.0, 2)
                    if (has_lop3_data and current_avg_tns > 0)
                    else 0.0
                )

                if has_lop3_data and c_tgt > 0:
                    c_status = "negative" if c_val > c_tgt else "positive"
                else:
                    c_status = "neutral"

                planners_list = []
                if has_lop3_data:
                    for _, r in grp.iterrows():
                        p_name = str(r["Planner"]).strip()
                        if not p_name:
                            continue
                        r_val = round(float(r["Value"]), 2)
                        r_tgt = round(float(r["Target"]), 2)
                        r_dio = (
                            round((r_val / current_avg_tns) * 30.0, 2)
                            if current_avg_tns > 0
                            else 0.0
                        )
                        planners_list.append(
                            {
                                "planner": p_name,
                                "product": str(r["Product"]),
                                "category": str(r["Category"]),
                                "month_year": str(r["Month year"]),
                                "val_lop3": r_val,
                                "target_val": r_tgt,
                                "lop3_dio": r_dio,
                                "status": "negative" if (r_tgt > 0 and r_val > r_tgt) else ("positive" if r_tgt > 0 else "neutral"),
                            }
                        )

                children.append(
                    {
                        "id": f"NODE_{slot_key}_{idx}",
                        "product": actual_prod,
                        "category": str(cat_name),
                        "val_lop3": c_val,
                        "target_val": c_tgt,
                        "lop3_dio": c_dio,
                        "status": c_status,
                        "has_data": has_lop3_data,
                        "planners": planners_list,
                    }
                )

        return {
            "id": f"NODE_LOP3_{slot_key}",
            "slot": slot_key,
            "product": actual_prod,
            "title": f"Product : {actual_prod}",
            "val_lop3": total_val,
            "target_val": total_target,
            "lop3_dio": parent_dio,
            "status": parent_status,
            "has_data": has_lop3_data,
            "no_data_message": (
                f"No data available for {active_lop3_month}. Please update KPI LOP3.xlsx and retry."
                if not has_lop3_data
                else ""
            ),
            "children": children,
        }

    branch_a = build_lop3_branch("4W", branch_products[0])
    branch_b = build_lop3_branch("2W", branch_products[1])

    stotal_row = get_metric_row(["Stotal", "sTotal", "sRM"])
    ototal_row = get_metric_row(["ototal", "oTotal", "dTotal", "aTotal"])
    srmstore_row = get_metric_row(["sRMstore"])

    if has_lop3_data:
        lop3_total_val = round(branch_a["val_lop3"] + branch_b["val_lop3"], 2)
        lop3_total_tgt = round(branch_a["target_val"] + branch_b["target_val"], 2)
        lop3_raw_delta = (
            ((lop3_total_val - lop3_total_tgt) / lop3_total_tgt) * 100.0
            if lop3_total_tgt > 0
            else 0.0
        )
        lop3_delta_pct = round(lop3_raw_delta, 1)
        lop3_delta_pct_formatted = format_cd05100_pct(lop3_raw_delta)
        lop3_status = (
            "negative"
            if (lop3_total_tgt > 0 and lop3_total_val > lop3_total_tgt)
            else ("positive" if lop3_total_tgt > 0 else "neutral")
        )
    else:
        lop3_total_val = 0.0
        lop3_total_tgt = 0.0
        lop3_delta_pct = 0.0
        lop3_delta_pct_formatted = "No data"
        lop3_status = "neutral"

    if ototal_row["status"] == "negative":
        ototal_color = "#e20015"  # Bosch Red 50
    elif ototal_row["status"] == "positive":
        ototal_color = "#00884a"  # Bosch Green 60
    else:
        ototal_color = "#bfc0c2"  # Bosch Gray 30

    active_bench_disp = "Target"
    summary_cards = [
        {
            "id": "CARD_STOTAL",
            "label": f"Total Plant Stock ({stotal_row['category']})",
            "code": stotal_row["category"],
            "scenario": "ACT",
            "value": stotal_row["values"],
            "unit": "MINR",
            "benchmark_label": "Target",
            "benchmark_value": stotal_row["benchmark_val"],
            "delta_pct": stotal_row["delta_pct"],
            "delta_pct_formatted": stotal_row["delta_pct_formatted"],
            "color": "#005691",
            "status": stotal_row["status"],
        },
        {
            "id": "CARD_OTOTAL",
            "label": f"Plant DIO ({ototal_row['category']})",
            "code": ototal_row["category"],
            "scenario": "TGT",
            "value": ototal_row["values"],
            "unit": "Days",
            "benchmark_label": "Target",
            "benchmark_value": ototal_row["benchmark_val"],
            "delta_pct": ototal_row["delta_pct"],
            "delta_pct_formatted": ototal_row["delta_pct_formatted"],
            "color": ototal_color,
            "status": ototal_row["status"],
            "target_specified": bool(ototal_row["benchmark_val"] > 0),
            "target_message": "" if ototal_row["benchmark_val"] > 0 else "No target specified",
        },
        {
            "id": "CARD_SRMSTORE",
            "label": "RM Store Stock (sRMstore)",
            "code": "sRMstore",
            "scenario": "CF",
            "value": srmstore_row["values"],
            "unit": "MINR",
            "benchmark_label": "Target",
            "benchmark_value": srmstore_row["benchmark_val"],
            "delta_pct": srmstore_row["delta_pct"],
            "delta_pct_formatted": srmstore_row["delta_pct_formatted"],
            "color": "#008ecf",
            "status": srmstore_row["status"],
        },
        {
            "id": "CARD_LOP3_TOTAL",
            "label": f"LOP3 Total ({branch_a['product']} + {branch_b['product']})",
            "code": "LOP3",
            "scenario": "FCM",
            "value": lop3_total_val,
            "unit": "MINR",
            "benchmark_label": "Sheet1 Target" if has_lop3_data else "No Target",
            "benchmark_value": lop3_total_tgt,
            "delta_pct": lop3_delta_pct,
            "delta_pct_formatted": lop3_delta_pct_formatted,
            "color": "#006eab" if has_lop3_data else "#64748b",
            "status": lop3_status,
            "has_data": has_lop3_data,
            "no_data_message": (
                f"No data available for {active_lop3_month}. Please update KPI LOP3.xlsx and retry."
                if not has_lop3_data
                else ""
            ),
        },
    ]

    # Build Trend Series strictly for the selected Year(s) ONLY, synchronized with Card Value & Benchmark
    trend_series = {}
    for cat_code, grp in df_main_vs_yr.sort_values(["Year", "Month"]).groupby("Category"):
        m_now = get_metric_row([cat_code])
        trend_series[cat_code] = {
            "labels": grp["Month year"].tolist(),
            "actuals": [round(float(x), 2) for x in grp["Value"].tolist()],
            "cf09": m_now["benchmark_val"],
            "pa": m_now["pa"],
            "benchmark_val": m_now["benchmark_val"],
            "card_value": m_now["values"],
            "status": m_now["status"],
            "active_months": effective_tree_months,
        }

    if has_lop3_data and not df_lop3_vs_yr.empty:
        for (prod_c, cat_c), grp in df_lop3_vs_yr.sort_values(["Year", "Month"]).groupby(["Product", "Category"]):
            key = f"LOP3:{prod_c}:{cat_c}"
            m_grp = (
                grp.groupby(["Year", "Month", "Month year"], as_index=False)
                .agg({"Value": "sum", "Target": "sum"})
                .sort_values(["Year", "Month"])
            )
            t_val = round(float(m_grp["Target"].mean()), 2) if not m_grp.empty else 0.0
            act_slice = m_grp[m_grp["Month year"].isin(valid_lop3_tree_months)]["Value"] if valid_lop3_tree_months else m_grp["Value"]
            c_val = round(float(act_slice.mean()), 2) if not act_slice.empty else 0.0
            c_stat = "negative" if (t_val > 0 and c_val > t_val) else ("positive" if t_val > 0 else "neutral")
            trend_series[key] = {
                "labels": m_grp["Month year"].tolist(),
                "actuals": [round(float(x), 2) for x in m_grp["Value"].tolist()],
                "cf09": t_val,
                "pa": t_val,
                "benchmark_val": t_val,
                "card_value": c_val,
                "status": c_stat,
                "has_data": True,
                "active_months": valid_lop3_tree_months,
            }
    else:
        for (prod_c, cat_c), grp in df_lop3_base_meta.groupby(["Product", "Category"], sort=False):
            key = f"LOP3:{prod_c}:{cat_c}"
            trend_series[key] = {
                "labels": [],
                "actuals": [],
                "cf09": 0.0,
                "pa": 0.0,
                "benchmark_val": 0.0,
                "card_value": 0.0,
                "status": "neutral",
                "has_data": False,
                "active_months": [],
            }

    # ==========================================================================
    # MAIN & LOP3 TABLES VIEW (Scoped by Year & Month Filters)
    # ==========================================================================
    effective_table_months = effective_tree_months
    df_main_table_slice = df_main_vs_yr[df_main_vs_yr["Month year"].isin(effective_table_months)]
    table_month_label = (
        ",".join(effective_table_months)
        if len(effective_table_months) <= 3
        else f"{len(effective_table_months)} Months ({year_filter_str})"
    )

    main_table_records = []
    df_cf_table_source = df_cf_slice if not df_cf_slice.empty else df_cf_vs
    for _, r in df_cf_table_source.iterrows():
        cat_c = str(r["Category"])
        m_info = get_metric_row([cat_c], custom_main_slice=df_main_table_slice)
        main_table_records.append(
            {
                "Category": cat_c,
                "Value Stream": str(r["Value Stream"]),
                "Month year": table_month_label,
                "CAT-YEAR": str(r["CAT-YEAR"]),
                "CF02": m_info["cf02"],
                "CF05": m_info["cf05"],
                "CF09": m_info["cf09"],
                "Target": m_info["benchmark_val"],
                "PA_Target": m_info["benchmark_val"],
                "Actual_Value": m_info["values"],
                "Status": m_info["status"],
                "Delta_Pct_Formatted": m_info["delta_pct_formatted"],
            }
        )

    # LOP3 Table rows filtered strictly by selected_table_months (NO HISTORICAL FALLBACK)
    valid_lop3_table_months = [
        m for m in effective_table_months if m in lop3_months_available
    ]
    lop3_table_records = []
    if valid_lop3_table_months and not df_lop3_vs_yr.empty:
        df_lop3_table_df = df_lop3_vs_yr[
            df_lop3_vs_yr["Month year"].isin(valid_lop3_table_months)
        ].copy()
        if sel_products:
            df_lop3_table_df = df_lop3_table_df[df_lop3_table_df["Product"].isin(sel_products)]
        if sel_planners:
            df_lop3_table_df = df_lop3_table_df[df_lop3_table_df["Planner"].isin(sel_planners)]
        if sel_categories:
            df_lop3_table_df = df_lop3_table_df[df_lop3_table_df["Category"].isin(sel_categories)]

        df_tns_table_row = df_main_table_slice[df_main_table_slice["Category"] == avg_tns_cat]
        nz_tns_table = df_tns_table_row[df_tns_table_row["Value"] != 0]["Value"] if not df_tns_table_row.empty else pd.Series(dtype=float)
        table_avg_tns = float(nz_tns_table.mean()) if not nz_tns_table.empty else current_avg_tns

        for _, r in df_lop3_table_df.iterrows():
            v_val = round(float(r["Value"]), 2)
            t_val = round(float(r["Target"]), 2)
            d_val = round((v_val / table_avg_tns) * 30.0, 2) if table_avg_tns > 0 else 0.0
            lop3_table_records.append(
                {
                    "Value Stream": str(r["Value Stream"]),
                    "Product": str(r["Product"]),
                    "Planner": str(r["Planner"]),
                    "Category": str(r["Category"]),
                    "Month year": str(r["Month year"]),
                    "Target": t_val,
                    "ValLOP3": v_val,
                    "LOP3 DIO": d_val,
                    "Status": "negative" if (t_val > 0 and v_val > t_val) else ("positive" if t_val > 0 else "neutral"),
                }
            )

    return {
        "meta": {
            "plant": "ChkP 1810",
            "last_refresh": DATA_STORE["last_refresh"],
            "source_info": DATA_STORE["source_info"],
            "active_lop3_month": active_lop3_month,
            "has_lop3_data": has_lop3_data,
            "lop3_no_data_message": (
                f"No data available for {active_lop3_month}. Please update the data in KPI LOP3.xlsx and retry."
                if not has_lop3_data
                else ""
            ),
            "target_specified": bool(ototal_row["benchmark_val"] > 0),
            "target_message": "" if ototal_row["benchmark_val"] > 0 else "No target specified",
        },
        "active_filters": {
            "value_stream": vs,
            "year": year_filter_str,
            "month_year": month_year,
            "product": product_filter,
            "planner": planner_filter,
            "category": category_filter,
            "benchmark": benchmark,
        },
        "filter_options": {
            "value_streams": available_vs,
            "years": [str(y) for y in available_years],
            "month_years": available_month_years,
            "products": available_products,
            "planners": available_planners,
            "categories": available_categories,
            "benchmarks": ["Target"],
        },
        "summary_cards": summary_cards,
        "tree": {
            "tier1_root": tier1_root,
            "tier2a": tier2a_cards,
            "tier2b": tier2b_cards,
            "tier3": {
                "4W": branch_a,
                "2W": branch_b,
            },
        },
        "trends": trend_series,
        "tables": {
            "cf_main": main_table_records,
            "lop3": lop3_table_records,
        },
    }


# ==============================================================================
# 5. FLASK ROUTES & STATIC ASSET WHITELISTING
# ==============================================================================
BOSCH_SVG_FALLBACK = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 230 52" width="230" height="52">
  <g fill="none" fill-rule="evenodd">
    <circle cx="26" cy="26" r="22" stroke="#262626" stroke-width="3.2" fill="#ffffff"/>
    <path d="M17.5 14.5v23M34.5 14.5v23M17.5 26h17" stroke="#262626" stroke-width="3.2" stroke-linecap="round"/>
    <path d="M17.5 14.5A14.5 14.5 0 0 0 17.5 37.5M34.5 14.5A14.5 14.5 0 0 1 34.5 37.5" stroke="#262626" stroke-width="2.6"/>
    <text x="58" y="36" font-family="Bosch Sans, Arial Black, Arial, sans-serif" font-size="31" font-weight="900" letter-spacing="1.5" fill="#e20015">BOSCH</text>
  </g>
</svg>"""


@app.route("/")
def serve_index():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/<path:filename>")
def serve_whitelisted_asset(filename: str):
    clean_name = os.path.basename(filename)
    if clean_name not in WHITELISTED_ASSETS:
        return jsonify({"error": "Asset not whitelisted"}), 404

    full_path = os.path.join(BASE_DIR, clean_name)
    if os.path.isfile(full_path):
        return send_from_directory(BASE_DIR, clean_name)

    if clean_name == "bosch-logo.svg":
        png_path = os.path.join(BASE_DIR, "bosch-logo.png")
        if os.path.isfile(png_path):
            return send_from_directory(BASE_DIR, "bosch-logo.png")
        return Response(BOSCH_SVG_FALLBACK, mimetype="image/svg+xml")

    if clean_name == "favicon.ico":
        return Response(status=204)

    return jsonify({"error": "File not found"}), 404


@app.route("/login")
def login_route():
    session["user"] = {
        "name": "Manoj Bharambe",
        "preferred_username": "bhm3chk@bosch.com",
        "plant": "ChkP 1810",
    }
    return redirect(url_for("serve_index"))


@app.route("/getAToken")
def get_a_token():
    return redirect(url_for("serve_index"))


@app.route("/api/me")
@login_required
def api_me():
    return jsonify(
        {
            "authenticated": True,
            "auth_bypass_dev": AUTH_BYPASS_DEV,
            "user": session.get("user"),
        }
    )


@app.route("/api/dashboard", methods=["GET"])
@login_required
def api_dashboard():
    filters = {
        "value_stream": request.args.get("value_stream", "ESP-ABS"),
        "year": request.args.get("year", ""),
        "month_year": request.args.get("month_year", ""),
        "product": request.args.get("product", "ALL"),
        "planner": request.args.get("planner", "ALL"),
        "category": request.args.get("category", "ALL"),
        "benchmark": request.args.get("benchmark", ""),
    }
    payload = compute_dashboard_payload(filters)
    return jsonify(payload)


@app.route("/api/reload", methods=["POST"])
@login_required
def api_reload():
    reload_all_sheets()
    return jsonify(
        {
            "status": "ok",
            "last_refresh": DATA_STORE["last_refresh"],
            "source_info": DATA_STORE["source_info"],
        }
    )


# ==============================================================================
# 6. LOCALHOST WAITRESS SERVER LAUNCHER
# ==============================================================================
def open_local_browser():
    target_url = f"http://localhost:{PORT}"
    try:
        webbrowser.open(target_url)
    except Exception:
        pass


if __name__ == "__main__":
    url = f"http://localhost:{PORT}"
    print("=" * 82)
    print(" BOSCH VM/LOG-IN INVENTORY KPI TREE - PLANT: ChkP 1810")
    print(f" Server Bound Strictly To : http://{HOST}:{PORT}")
    print(f" Forecast Workbook        : {DATA_STORE['source_info'].get('inventory_file')}")
    print(f" Forecast Sheet Loaded    : {DATA_STORE['source_info'].get('sheets_used', {}).get('inventory_sheet')}")
    print(f" LOP3 Workbook            : {DATA_STORE['source_info'].get('lop3_file')}")
    print(f" LOP3 Sheet Loaded        : {DATA_STORE['source_info'].get('sheets_used', {}).get('lop3_sheet')}")
    print(f" Exact Planners Loaded    : {DATA_STORE['source_info'].get('lop3_planners')}")
    print("=" * 82)

    threading.Timer(1.0, open_local_browser).start()
    serve(app, host=HOST, port=PORT, threads=6)
