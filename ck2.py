"""
=============================================================================
INCIDENT REPORT GENERATOR
=============================================================================
CONFIGURATION — only 5 things to set:
  INPUT_FILE_PATH          : path to the source Excel workbook
  COL_INCIDENT_ID          : exact column name for the Incident ID
  COL_CLOSURE_DATE         : exact column name for the Closure Date
  COL_STATUS               : exact column name for the Status column
  COL_CLOSED_BY            : exact column name for who closed the incident

CREDENTIAL BREACH CONFIGURATION:
  CREDENTIAL_BREACH_SHEET  : exact sheet name for credential breach data
  COL_BREACH_EMAIL         : exact column name for the email address
  COL_BREACH_PASSWORD      : exact column name for the password

Everything else is automatic — date range is prompted at runtime.

Summary sheet contains:
  1. Module-wise unique closed incidents table (date range filtered)
  2. Pie chart — "Closed Incidents By Module" (date range filtered)
  3. Pie chart — "Overall Incident Status" (all data, unique IDs)
  4. User-wise closed incidents table + pie (date range filtered)
  5. Credential Breach password-strength pie chart
  6. Top 20 most-seen email addresses table

All pie chart labels are printed OUTSIDE the slice with leader lines —
  name, count and % shown in bold navy text, never overlapping the colours.
=============================================================================
"""

import os
import re
import sys
from datetime import datetime

import pandas as pd
import xlsxwriter
import warnings
warnings.filterwarnings("ignore")

# =============================================================================
# >>>  CONFIGURATION — only edit these lines  <<<
# =============================================================================

INPUT_FILE_PATH  = "incidents.xlsx"
OUTPUT_FOLDER    = "reports"              # folder where reports are saved (created if missing)
COL_INCIDENT_ID  = "Incident Id"
COL_CLOSURE_DATE = "Incident Closure on"
COL_STATUS       = "Status"
COL_CLOSED_BY    = "Incident Closed By"

# ----------  Credential Breach sheet  ----------------------------------------
CREDENTIAL_BREACH_SHEET = "Credential Breaches"   # exact sheet name
COL_BREACH_EMAIL        = "Email"                  # exact column name for email
COL_BREACH_PASSWORD     = "Password"               # exact column name for password
# =============================================================================

SUMMARY_SHEET_NAME = "Summary Dashboard"

# ---------------------------------------------------------------------------
# DATE PARSING
# ---------------------------------------------------------------------------

_ORDINAL_RE   = re.compile(r"(\d+)(st|nd|rd|th)\b", re.IGNORECASE)
_DATE_FORMATS = [
    "%d %b, %Y %I:%M:%S %p",
    "%d %B, %Y %I:%M:%S %p",
    "%d %b, %Y %H:%M:%S",
    "%d %B, %Y %H:%M:%S",
    "%d %b, %Y %I:%M %p",
    "%d %b, %Y",
    "%d %B, %Y",
    "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d-%B-%Y",
    "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d.%m.%Y",
    "%d %b %y", "%d %B %y",
]

def _strip_ordinals(text):
    return _ORDINAL_RE.sub(r"\1", str(text))

def _parse_date_series(series):
    def _clean(text):
        t = _strip_ordinals(str(text))
        t = re.sub(r"([A-Za-z]),", r"\1", t)
        return t.strip()

    cleaned = series.astype(str).apply(_clean)
    parsed  = pd.to_datetime(cleaned, infer_datetime_format=True,
                             dayfirst=True, errors="coerce")

    for fmt in _DATE_FORMATS:
        if not parsed.isna().any():
            break
        bad = parsed.isna()
        parsed[bad] = pd.to_datetime(cleaned[bad], format=fmt, errors="coerce")

    _EPOCH    = pd.Timestamp("1899-12-30")
    still_bad = parsed.isna()
    if still_bad.any():
        for idx in parsed.index[still_bad]:
            try:
                serial = float(str(series[idx]).strip())
                if 1 < serial < 2958466:
                    parsed[idx] = _EPOCH + pd.Timedelta(days=serial)
            except (ValueError, TypeError):
                pass

    return parsed


def _prompt_date(label):
    while True:
        raw     = input(f"  Enter {label} (e.g. 1 Jan 2024 / 01/01/2024 / 2024-01-01): ").strip()
        cleaned = _strip_ordinals(raw)
        for fmt in _DATE_FORMATS:
            try:
                return pd.Timestamp(datetime.strptime(cleaned, fmt))
            except ValueError:
                pass
        try:
            return pd.Timestamp(pd.to_datetime(cleaned, dayfirst=True))
        except Exception:
            print(f"  Could not understand '{raw}'. Please try again.")

# ---------------------------------------------------------------------------
# STEP 1 — LOAD
# ---------------------------------------------------------------------------

def load_all_sheets(path):
    if not os.path.exists(path):
        sys.exit(f"\nERROR: Input file not found -> {path}\n")
    print(f"\nLoading: {path}")

    str_sheets    = pd.read_excel(path, sheet_name=None, dtype=str)
    native_sheets = pd.read_excel(path, sheet_name=None)

    merged = {}
    for name, df_str in str_sheets.items():
        df = df_str.copy()
        if name in native_sheets and COL_CLOSURE_DATE in native_sheets[name].columns:
            df[COL_CLOSURE_DATE] = native_sheets[name][COL_CLOSURE_DATE].values
        merged[name] = df
    return merged

# ---------------------------------------------------------------------------
# STEP 2 — PROCESS
# ---------------------------------------------------------------------------

def process_sheet(df, start_dt, end_dt):
    df = df.copy()
    if COL_CLOSURE_DATE not in df.columns:
        return df, pd.DataFrame(columns=df.columns)

    df[COL_CLOSURE_DATE] = _parse_date_series(df[COL_CLOSURE_DATE])
    mask     = (df[COL_CLOSURE_DATE] >= start_dt) & (df[COL_CLOSURE_DATE] <= end_dt)
    filtered = df[mask].copy()
    if COL_INCIDENT_ID in filtered.columns:
        filtered = filtered.drop_duplicates(subset=[COL_INCIDENT_ID])
    return df, filtered

# ---------------------------------------------------------------------------
# STEP 3 — AGGREGATE
# ---------------------------------------------------------------------------

def aggregate(counts):
    module_names  = []
    module_counts = []
    for sheet_name, count in counts.items():
        if count > 0:
            module_names.append(sheet_name)
            module_counts.append(count)
    if not module_names:
        sys.exit("\nNo incidents found in the specified date range.\n")
    return module_names, module_counts

# ---------------------------------------------------------------------------
# STEP 3b — OVERALL STATUS BREAKDOWN
# ---------------------------------------------------------------------------

def _normalize_status(val):
    if pd.isna(val):
        return None
    s = str(val).strip().lower()
    if s.startswith("closed"):
        return "Closed"
    if s.startswith("open"):
        return "Open"
    if s.startswith("in progress") or s.startswith("inprogress"):
        return "In Progress"
    return None

def compute_status_breakdown(processed_raw, sheet_names):
    frames = []
    for name in sheet_names:
        df = processed_raw[name]
        if df.empty:
            continue
        cols_needed = [c for c in [COL_INCIDENT_ID, COL_STATUS] if c in df.columns]
        if cols_needed:
            frames.append(df[cols_needed].copy())

    if not frames:
        return [], []

    combined = pd.concat(frames, ignore_index=True)
    if COL_INCIDENT_ID in combined.columns:
        combined = combined.drop_duplicates(subset=[COL_INCIDENT_ID])

    if COL_STATUS not in combined.columns:
        print(f"  WARNING: Column '{COL_STATUS}' not found — status pie skipped.")
        return [], []

    combined["_status_norm"] = combined[COL_STATUS].apply(_normalize_status)
    combined = combined[combined["_status_norm"].notna()]
    counts   = combined["_status_norm"].value_counts()
    order    = ["Open", "In Progress", "Closed"]
    labels   = [s for s in order if s in counts.index]
    status_counts = [int(counts[s]) for s in labels]

    print(f"\n  Overall status breakdown (unique incidents, 3 buckets only):")
    for lbl, cnt in zip(labels, status_counts):
        print(f"    {lbl}: {cnt}")

    return labels, status_counts

# ---------------------------------------------------------------------------
# STEP 3c — USER-WISE BREAKDOWN
# ---------------------------------------------------------------------------

def compute_user_breakdown(filtered_raw, sheet_names):
    frames = []
    for name in sheet_names:
        df = filtered_raw.get(name, pd.DataFrame())
        if df.empty:
            continue
        cols = [c for c in [COL_INCIDENT_ID, COL_CLOSED_BY] if c in df.columns]
        if len(cols) == 2:
            frames.append(df[cols].copy())

    if not frames:
        print(f"  WARNING: Column '{COL_CLOSED_BY}' not found — user breakdown skipped.")
        return [], []

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.drop_duplicates(subset=[COL_INCIDENT_ID])
    combined = combined[
        ~combined[COL_CLOSED_BY].astype(str).str.strip().str.lower().isin(
            ["", "nan", "none", "n/a", "-"]
        )
    ]

    counts = (
        combined[COL_CLOSED_BY]
        .astype(str).str.strip()
        .value_counts()
        .sort_values(ascending=False)
    )

    user_names  = list(counts.index)
    user_counts = [int(c) for c in counts.values]

    print(f"\n  Closed-by breakdown (date range, unique incidents):")
    for u, c in zip(user_names, user_counts):
        print(f"    {u}: {c}")

    return user_names, user_counts

# ---------------------------------------------------------------------------
# STEP 3d — CREDENTIAL BREACH ANALYSIS  (NEW)
# ---------------------------------------------------------------------------

_STRONG_PASSWORD_RE = re.compile(
    r"^(?=.*[A-Z])(?=.*[a-z])(?=.*\d)(?=.*[^A-Za-z0-9]).+$"
)

_BLANK_VALUES = {"", "nan", "none", "n/a", "na", "-", "null"}


def _classify_password(val):
    """
    Returns one of three string labels:
      'No Password'      — cell is blank / N/A
      'Strong'           — meets policy: ≥1 upper, ≥1 lower, ≥1 digit, ≥1 symbol
      'Weak / No Policy' — non-blank but does not meet policy
    """
    if pd.isna(val):
        return "No Password"
    s = str(val).strip()
    if s.lower() in _BLANK_VALUES:
        return "No Password"
    if _STRONG_PASSWORD_RE.match(s):
        return "Strong"
    return "Weak / No Policy"


def compute_credential_breach_analysis(raw_sheets):
    """
    Reads the CREDENTIAL_BREACH_SHEET from the already-loaded raw_sheets dict.

    Returns:
      pwd_labels  : list[str]  — e.g. ['Strong', 'Weak / No Policy', 'No Password']
      pwd_counts  : list[int]
      top_emails  : list[tuple[str, int]]  — up to 20 (email, count) pairs, desc
    """
    if CREDENTIAL_BREACH_SHEET not in raw_sheets:
        print(f"\n  WARNING: Sheet '{CREDENTIAL_BREACH_SHEET}' not found — "
              f"credential breach analysis skipped.")
        return [], [], []

    df = raw_sheets[CREDENTIAL_BREACH_SHEET].copy()
    print(f"\n  Credential Breach sheet loaded — {len(df)} row(s).")

    # ── Password strength ────────────────────────────────────────────────────
    if COL_BREACH_PASSWORD not in df.columns:
        print(f"  WARNING: Column '{COL_BREACH_PASSWORD}' not found in "
              f"'{CREDENTIAL_BREACH_SHEET}' — password analysis skipped.")
        pwd_labels, pwd_counts = [], []
    else:
        df["_pwd_class"] = df[COL_BREACH_PASSWORD].apply(_classify_password)
        vc = df["_pwd_class"].value_counts()

        # Fixed display order
        order      = ["Strong", "Weak / No Policy", "No Password"]
        pwd_labels = [lbl for lbl in order if lbl in vc.index]
        pwd_counts = [int(vc[lbl]) for lbl in pwd_labels]

        print(f"\n  Password strength breakdown:")
        for lbl, cnt in zip(pwd_labels, pwd_counts):
            print(f"    {lbl}: {cnt}")

    # ── Top 20 emails ────────────────────────────────────────────────────────
    if COL_BREACH_EMAIL not in df.columns:
        print(f"  WARNING: Column '{COL_BREACH_EMAIL}' not found in "
              f"'{CREDENTIAL_BREACH_SHEET}' — top-email table skipped.")
        top_emails = []
    else:
        email_series = (
            df[COL_BREACH_EMAIL]
            .astype(str).str.strip().str.lower()
        )
        # Drop blanks
        email_series = email_series[~email_series.isin(_BLANK_VALUES)]
        top_emails = (
            email_series.value_counts()
            .head(20)
            .reset_index()
            .rename(columns={"index": "email", COL_BREACH_EMAIL: "count"})
            .values.tolist()          # list of [email, count]
        )
        # Normalise — pandas ≥2.0 changes value_counts column names
        if top_emails and len(top_emails[0]) == 2:
            top_emails = [(str(row[0]), int(row[1])) for row in top_emails]

        print(f"\n  Top {len(top_emails)} email(s) by occurrence:")
        for email, cnt in top_emails:
            print(f"    {email}: {cnt}")

    return pwd_labels, pwd_counts, top_emails

# ---------------------------------------------------------------------------
# STEP 4 — BUILD OUTPUT WORKBOOK
# ---------------------------------------------------------------------------

def build_workbook(module_names, module_counts, status_labels, status_counts,
                   user_names, user_counts,
                   pwd_labels, pwd_counts, top_emails,
                   filtered_raw, counts, sheet_names,
                   start_dt, end_dt, output_path):

    wb = xlsxwriter.Workbook(output_path)

    # ── FORMATS ──────────────────────────────────────────────────────────────
    NAVY    = "#1F3864"
    MIDBLUE = "#2E75B6"
    LTBLUE  = "#D6E4F7"
    ALT     = "#EFF5FB"
    WHITE   = "#FFFFFF"
    GREEN   = "#70AD47"   # used for credential breach section accent

    f_banner = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 15,
        "font_color": WHITE, "bg_color": NAVY,
        "align": "center", "valign": "vcenter",
    })
    f_section = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 13,
        "font_color": NAVY, "valign": "vcenter",
    })
    f_section_green = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 13,
        "font_color": GREEN, "valign": "vcenter",
    })
    f_col_hdr = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 11,
        "font_color": WHITE, "bg_color": NAVY,
        "align": "center", "valign": "vcenter", "border": 1,
    })
    f_col_hdr_green = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 11,
        "font_color": WHITE, "bg_color": GREEN,
        "align": "center", "valign": "vcenter", "border": 1,
    })
    f_num = wb.add_format({
        "font_name": "Arial", "font_size": 10,
        "align": "center", "valign": "vcenter", "border": 1,
    })
    f_lft = wb.add_format({
        "font_name": "Arial", "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1,
    })
    f_num_alt = wb.add_format({
        "font_name": "Arial", "font_size": 10,
        "align": "center", "valign": "vcenter", "border": 1, "bg_color": ALT,
    })
    f_lft_alt = wb.add_format({
        "font_name": "Arial", "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1, "bg_color": ALT,
    })
    f_total = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 11,
        "font_color": NAVY, "bg_color": LTBLUE,
        "align": "center", "valign": "vcenter", "border": 2,
    })
    f_data_hdr = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 10,
        "font_color": WHITE, "bg_color": MIDBLUE,
        "align": "center", "valign": "vcenter", "border": 1,
    })
    f_cell = wb.add_format({
        "font_name": "Arial", "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1,
    })
    f_cell_alt = wb.add_format({
        "font_name": "Arial", "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1, "bg_color": ALT,
    })
    f_date = wb.add_format({
        "font_name": "Arial", "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1,
        "num_format": "dd mmm yyyy",
    })
    f_date_alt = wb.add_format({
        "font_name": "Arial", "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1, "bg_color": ALT,
        "num_format": "dd mmm yyyy",
    })

    # ── SUMMARY SHEET ─────────────────────────────────────────────────────────
    n           = len(module_names)
    HDR_ROW     = 3
    DATA_START  = 4
    TOTAL_ROW   = DATA_START + n
    sum_formula = f"=SUM(C{DATA_START + 1}:C{DATA_START + n})"

    sw = wb.add_worksheet(SUMMARY_SHEET_NAME)

    sw.set_row(0, 42)
    sw.merge_range(
        0, 0, 0, 15,
        f"Incident Summary Dashboard  |  "
        f"{start_dt.strftime('%d %b %Y')}  ->  {end_dt.strftime('%d %b %Y')}",
        f_banner,
    )

    sw.set_row(2, 24)
    sw.write(2, 0, "Module-wise Unique Closed Incidents", f_section)

    sw.set_row(HDR_ROW, 22)
    sw.write(HDR_ROW, 0, "#",                       f_col_hdr)
    sw.write(HDR_ROW, 1, "Module Name",              f_col_hdr)
    sw.write(HDR_ROW, 2, "Unique Closed Incidents",  f_col_hdr)

    for i in range(n):
        r   = DATA_START + i
        alt = (i % 2 == 1)
        sw.set_row(r, 18)
        sw.write(r, 0, i + 1,            f_num_alt if alt else f_num)
        sw.write(r, 1, module_names[i],  f_lft_alt if alt else f_lft)
        sw.write(r, 2, module_counts[i], f_num_alt if alt else f_num)

    sw.set_row(TOTAL_ROW, 22)
    sw.merge_range(TOTAL_ROW, 0, TOTAL_ROW, 1, "TOTAL", f_total)
    sw.write_formula(TOTAL_ROW, 2, sum_formula, f_total)

    sw.set_column(0, 0, 6)
    sw.set_column(1, 1, max(max(len(m) for m in module_names) + 6, 24))
    sw.set_column(2, 2, 28)
    sw.freeze_panes(HDR_ROW, 0)

    # ── PIE COLOUR PALETTES ──────────────────────────────────────────────────
    _PALETTE = [
        "#4472C4", "#ED7D31", "#70AD47", "#FFC000", "#5B9BD5",
        "#A9D18E", "#FF7C80", "#9E480E", "#7030A0", "#636363",
        "#255E91", "#43682B", "#C00000", "#997300", "#7B5EA7",
    ]
    # Credential breach password pie: fixed semantic colours
    _PWD_PALETTE = {
        "Strong":           "#70AD47",   # green
        "Weak / No Policy": "#ED7D31",   # orange
        "No Password":      "#A6A6A6",   # grey
    }

    def _pie_points(count):
        return [{"fill": {"color": _PALETTE[i % len(_PALETTE)]}}
                for i in range(count)]

    def _pie_points_named(labels, palette_dict):
        return [{"fill": {"color": palette_dict.get(lbl, "#4472C4")}}
                for lbl in labels]

    def _write_legend(ws, start_row, start_col, labels, values,
                      palette=None):
        """
        Colour-swatch legend: [■] | Name | Count | Percentage
        palette: optional dict {label: colour}; falls back to _PALETTE index order.
        """
        total    = sum(values) or 1
        hdr_fmt  = wb.add_format({
            "bold": True, "font_name": "Arial", "font_size": 9,
            "font_color": WHITE, "bg_color": NAVY,
            "align": "center", "valign": "vcenter", "border": 1,
        })
        ws.write(start_row, start_col,     "",        hdr_fmt)
        ws.write(start_row, start_col + 1, "Name",    hdr_fmt)
        ws.write(start_row, start_col + 2, "Count",   hdr_fmt)
        ws.write(start_row, start_col + 3, "% Share", hdr_fmt)

        for i, (lbl, val) in enumerate(zip(labels, values)):
            r      = start_row + 1 + i
            colour = (palette.get(lbl) if palette
                      else _PALETTE[i % len(_PALETTE)])
            alt    = (i % 2 == 1)
            bg     = ALT if alt else WHITE

            swatch_fmt = wb.add_format({"bg_color": colour, "border": 1})
            name_fmt   = wb.add_format({
                "font_name": "Arial", "font_size": 9, "border": 1,
                "align": "left", "valign": "vcenter", "bg_color": bg,
            })
            val_fmt    = wb.add_format({
                "font_name": "Arial", "font_size": 9, "border": 1,
                "align": "center", "valign": "vcenter", "bg_color": bg,
            })
            pct_fmt    = wb.add_format({
                "font_name": "Arial", "font_size": 9, "border": 1,
                "align": "center", "valign": "vcenter", "bg_color": bg,
                "num_format": "0.0%",
            })

            ws.write(r, start_col,     "",          swatch_fmt)
            ws.write(r, start_col + 1, lbl,         name_fmt)
            ws.write(r, start_col + 2, val,         val_fmt)
            ws.write(r, start_col + 3, val / total, pct_fmt)

        ws.set_column(start_col,     start_col,     3)
        ws.set_column(start_col + 1, start_col + 1,
                      max((len(str(l)) for l in labels), default=12) + 4)
        ws.set_column(start_col + 2, start_col + 2, 10)
        ws.set_column(start_col + 3, start_col + 3, 10)

    CHART_ROW = TOTAL_ROW + 2
    PIE_ROWS  = 19

    # ── PIE 1 — Closed Incidents By Module ───────────────────────────────────
    pie1 = wb.add_chart({"type": "pie"})
    pie1.add_series({
        "name":       "Closed Incidents By Module",
        "categories": [SUMMARY_SHEET_NAME, DATA_START, 1, DATA_START + n - 1, 1],
        "values":     [SUMMARY_SHEET_NAME, DATA_START, 2, DATA_START + n - 1, 2],
        "points":     _pie_points(n),
    })
    pie1.set_title({"name": "Closed Incidents By Module"})
    pie1.set_legend({"none": True})
    pie1.set_style(10)
    pie1.set_size({"width": 420, "height": 360})
    sw.insert_chart(CHART_ROW, 0, pie1)

    L1_ROW = CHART_ROW + PIE_ROWS
    _write_legend(sw, L1_ROW, 0, module_names, module_counts)

    # ── PIE 2 — Overall Incident Status ──────────────────────────────────────
    ns = len(status_labels)
    if ns > 0:
        STATUS_SHEET = "Status Data"
        ss = wb.add_worksheet(STATUS_SHEET)
        for i, (lbl, cnt) in enumerate(zip(status_labels, status_counts)):
            ss.write(i, 0, lbl)
            ss.write(i, 1, cnt)

        pie2 = wb.add_chart({"type": "pie"})
        pie2.add_series({
            "name":       "Overall Incident Status",
            "categories": [STATUS_SHEET, 0, 0, ns - 1, 0],
            "values":     [STATUS_SHEET, 0, 1, ns - 1, 1],
            "points":     _pie_points(ns),
        })
        pie2.set_title({"name": "Overall Incident Status (All Modules, Unique IDs)"})
        pie2.set_legend({"none": True})
        pie2.set_style(10)
        pie2.set_size({"width": 420, "height": 360})
        sw.insert_chart(CHART_ROW, 7, pie2)

        _write_legend(sw, L1_ROW, 7, status_labels, status_counts)

    # ── USER TABLE + PIE 3 ────────────────────────────────────────────────────
    nu = len(user_names)
    if nu > 0:
        legend_rows      = max(n, ns if ns > 0 else 0) + 3
        USER_TABLE_START = L1_ROW + legend_rows + 2
        USER_HDR_ROW     = USER_TABLE_START + 1
        USER_DATA_START  = USER_HDR_ROW + 1
        USER_TOTAL_ROW   = USER_DATA_START + nu

        sw.set_row(USER_TABLE_START, 24)
        sw.write(USER_TABLE_START, 0,
                 "Incidents Closed By — User Wise (Date Range)",
                 f_section)

        sw.set_row(USER_HDR_ROW, 22)
        sw.write(USER_HDR_ROW, 0, "#",                       f_col_hdr)
        sw.write(USER_HDR_ROW, 1, "Closed By",               f_col_hdr)
        sw.write(USER_HDR_ROW, 2, "Unique Incidents Closed",  f_col_hdr)

        for i in range(nu):
            r   = USER_DATA_START + i
            alt = (i % 2 == 1)
            sw.set_row(r, 18)
            sw.write(r, 0, i + 1,          f_num_alt if alt else f_num)
            sw.write(r, 1, user_names[i],  f_lft_alt if alt else f_lft)
            sw.write(r, 2, user_counts[i], f_num_alt if alt else f_num)

        sw.set_row(USER_TOTAL_ROW, 22)
        sw.merge_range(USER_TOTAL_ROW, 0, USER_TOTAL_ROW, 1, "TOTAL", f_total)
        sw.write_formula(
            USER_TOTAL_ROW, 2,
            f"=SUM(C{USER_DATA_START + 1}:C{USER_DATA_START + nu})",
            f_total,
        )

        USER_SHEET = "User Data"
        us = wb.add_worksheet(USER_SHEET)
        for i, (uname, ucnt) in enumerate(zip(user_names, user_counts)):
            us.write(i, 0, uname)
            us.write(i, 1, ucnt)

        pie3 = wb.add_chart({"type": "pie"})
        pie3.add_series({
            "name":       "Incidents Closed By User",
            "categories": [USER_SHEET, 0, 0, nu - 1, 0],
            "values":     [USER_SHEET, 0, 1, nu - 1, 1],
            "points":     _pie_points(nu),
        })
        pie3.set_title({"name": "Incidents Closed By User (Date Range)"})
        pie3.set_legend({"none": True})
        pie3.set_style(10)
        pie3.set_size({"width": 420, "height": 360})
        sw.insert_chart(USER_TOTAL_ROW + 2, 0, pie3)

        _write_legend(sw, USER_TOTAL_ROW + 2 + PIE_ROWS,
                      0, user_names, user_counts)

        # Anchor row for credential breach section is below user section
        CRED_ANCHOR_ROW = USER_TOTAL_ROW + 2 + PIE_ROWS + nu + 5
    else:
        # No user data — credential breach goes below the legend tables
        CRED_ANCHOR_ROW = L1_ROW + max(n, ns if ns > 0 else 0) + 5

    # ── CREDENTIAL BREACH SECTION  (NEW) ─────────────────────────────────────
    np_ = len(pwd_labels)          # number of password buckets (0-3)
    ne  = len(top_emails)          # number of top-email rows (0-20)

    if np_ > 0 or ne > 0:
        # ── Separator / section banner ────────────────────────────────────────
        sw.set_row(CRED_ANCHOR_ROW, 36)
        sw.merge_range(
            CRED_ANCHOR_ROW, 0, CRED_ANCHOR_ROW, 15,
            f"Credential Breach Analysis  —  {CREDENTIAL_BREACH_SHEET}",
            wb.add_format({
                "bold": True, "font_name": "Arial", "font_size": 13,
                "font_color": WHITE, "bg_color": GREEN,
                "align": "center", "valign": "vcenter",
            }),
        )

        # ── Password strength table ───────────────────────────────────────────
        if np_ > 0:
            PWD_TBL_ROW = CRED_ANCHOR_ROW + 2

            sw.set_row(PWD_TBL_ROW, 24)
            sw.write(PWD_TBL_ROW, 0,
                     "Password Strength vs. Policy",
                     f_section_green)

            PWD_HDR = PWD_TBL_ROW + 1
            PWD_DATA_START = PWD_HDR + 1
            PWD_TOTAL_ROW  = PWD_DATA_START + np_

            sw.set_row(PWD_HDR, 22)
            sw.write(PWD_HDR, 0, "#",               f_col_hdr_green)
            sw.write(PWD_HDR, 1, "Category",        f_col_hdr_green)
            sw.write(PWD_HDR, 2, "Count",           f_col_hdr_green)

            for i in range(np_):
                r   = PWD_DATA_START + i
                alt = (i % 2 == 1)
                sw.set_row(r, 18)
                sw.write(r, 0, i + 1,            f_num_alt if alt else f_num)
                sw.write(r, 1, pwd_labels[i],    f_lft_alt if alt else f_lft)
                sw.write(r, 2, pwd_counts[i],    f_num_alt if alt else f_num)

            sw.set_row(PWD_TOTAL_ROW, 22)
            sw.merge_range(PWD_TOTAL_ROW, 0, PWD_TOTAL_ROW, 1, "TOTAL", f_total)
            sw.write_formula(
                PWD_TOTAL_ROW, 2,
                f"=SUM(C{PWD_DATA_START + 1}:C{PWD_DATA_START + np_})",
                f_total,
            )

            # Write password data to a helper sheet for chart reference
            PWD_SHEET = "Pwd Strength Data"
            ps = wb.add_worksheet(PWD_SHEET)
            for i, (lbl, cnt) in enumerate(zip(pwd_labels, pwd_counts)):
                ps.write(i, 0, lbl)
                ps.write(i, 1, cnt)

            pie4 = wb.add_chart({"type": "pie"})
            pie4.add_series({
                "name":       "Password Strength",
                "categories": [PWD_SHEET, 0, 0, np_ - 1, 0],
                "values":     [PWD_SHEET, 0, 1, np_ - 1, 1],
                "points":     _pie_points_named(pwd_labels, _PWD_PALETTE),
            })
            pie4.set_title({"name": "Credential Breach — Password Strength vs. Policy"})
            pie4.set_legend({"none": True})
            pie4.set_style(10)
            pie4.set_size({"width": 420, "height": 360})
            sw.insert_chart(PWD_TOTAL_ROW + 2, 0, pie4)

            # Legend for pie4
            PWD_LEGEND_ROW = PWD_TOTAL_ROW + 2 + PIE_ROWS
            _write_legend(sw, PWD_LEGEND_ROW, 0,
                          pwd_labels, pwd_counts, palette=_PWD_PALETTE)
        else:
            PWD_LEGEND_ROW = CRED_ANCHOR_ROW + 2

        # ── Top 20 emails table ───────────────────────────────────────────────
        if ne > 0:
            # Place email table to the right of the pie chart (column 7)
            EMAIL_TBL_ROW = CRED_ANCHOR_ROW + 2

            sw.set_row(EMAIL_TBL_ROW, 24)
            sw.write(EMAIL_TBL_ROW, 7,
                     f"Top {ne} Email Addresses by Breach Count",
                     f_section_green)

            EMAIL_HDR       = EMAIL_TBL_ROW + 1
            EMAIL_DATA_START = EMAIL_HDR + 1

            sw.set_row(EMAIL_HDR, 22)
            sw.write(EMAIL_HDR, 7, "#",              f_col_hdr_green)
            sw.write(EMAIL_HDR, 8, "Email Address",  f_col_hdr_green)
            sw.write(EMAIL_HDR, 9, "Breach Count",   f_col_hdr_green)

            # Auto-fit email column width
            max_email_len = max((len(str(e)) for e, _ in top_emails), default=20)
            sw.set_column(7, 7, 6)
            sw.set_column(8, 8, max(max_email_len + 4, 32))
            sw.set_column(9, 9, 16)

            for i, (email, cnt) in enumerate(top_emails):
                r   = EMAIL_DATA_START + i
                alt = (i % 2 == 1)
                sw.set_row(r, 18)
                sw.write(r, 7, i + 1,  f_num_alt if alt else f_num)
                sw.write(r, 8, email,  f_lft_alt if alt else f_lft)
                sw.write(r, 9, cnt,    f_num_alt if alt else f_num)

    # ── DATA SHEETS — only modules with incidents in range ────────────────────
    for name in sheet_names:
        if counts.get(name, 0) == 0:
            continue

        df      = filtered_raw[name]
        ws_name = name[:31]
        dw      = wb.add_worksheet(ws_name)

        if df.empty:
            continue

        headers      = list(df.columns)
        n_cols       = len(headers)
        n_rows       = len(df)
        date_col_idx = headers.index(COL_CLOSURE_DATE) if COL_CLOSURE_DATE in headers else -1

        dw.set_default_row(16)
        dw.set_row(0, 20)
        for ci, h in enumerate(headers):
            dw.write(0, ci, h, f_data_hdr)
            dw.set_column(ci, ci, max(len(str(h)) + 4, 14))

        data_values = df.values

        for ri in range(n_rows):
            excel_row = ri + 1
            alt = (excel_row % 2 == 0)
            for ci in range(n_cols):
                val         = data_values[ri, ci]
                is_date_col = (ci == date_col_idx)
                try:
                    is_null = pd.isna(val)
                except (TypeError, ValueError):
                    is_null = False

                if is_null:
                    dw.write_blank(excel_row, ci, None,
                                   f_date_alt if (is_date_col and alt) else
                                   f_date     if is_date_col else
                                   f_cell_alt if alt else f_cell)
                elif isinstance(val, pd.Timestamp):
                    dw.write_datetime(excel_row, ci, val.to_pydatetime(),
                                      f_date_alt if alt else f_date)
                else:
                    dw.write(excel_row, ci, val,
                             f_cell_alt if alt else f_cell)

        dw.freeze_panes(1, 0)
        dw.autofilter(0, 0, n_rows, n_cols - 1)

    wb.close()

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    print("\n" + "=" * 60)
    print("  INCIDENT REPORT GENERATOR")
    print("=" * 60)
    print("\nEnter the date range for filtering closed incidents.")

    start_dt = _prompt_date("START date (inclusive)")
    end_dt   = _prompt_date("END   date (inclusive)")

    if end_dt < start_dt:
        start_dt, end_dt = end_dt, start_dt
        print("  (Dates swapped — start was after end.)")

    end_dt = end_dt.replace(hour=23, minute=59, second=59)
    print(f"\n  Range: {start_dt.strftime('%d %b %Y')} -> {end_dt.strftime('%d %b %Y')}\n")

    # Step 1 — load all sheets
    raw_sheets  = load_all_sheets(INPUT_FILE_PATH)
    sheet_names = list(raw_sheets.keys())

    # Identify incident sheets (everything except the credential breach sheet)
    incident_sheet_names = [s for s in sheet_names if s != CREDENTIAL_BREACH_SHEET]
    print(f"Sheets found ({len(sheet_names)}): {sheet_names}")
    print(f"Incident sheets ({len(incident_sheet_names)}): {incident_sheet_names}\n")

    # Step 2 — process each incident sheet
    processed_raw = {}
    filtered_raw  = {}
    counts        = {}

    print("\n  DIAGNOSTIC — checking for closed incidents with missing closure dates:")
    for name in incident_sheet_names:
        df = raw_sheets.get(name, pd.DataFrame())
        if df.empty or COL_INCIDENT_ID not in df.columns:
            continue
        has_status  = COL_STATUS in df.columns
        has_cl_date = COL_CLOSURE_DATE in df.columns
        if not has_status or not has_cl_date:
            continue
        closed_mask = df[COL_STATUS].astype(str).str.lower().str.startswith("closed")
        closed_df   = df[closed_mask]
        blank_date  = closed_df[
            closed_df[COL_CLOSURE_DATE].astype(str).str.strip().isin(
                ["", "nan", "NaT", "None", "N/A", "-"]
            )
        ]
        if not blank_date.empty:
            print(f"    [{name}]  {len(blank_date)} closed incident(s) have NO closure date "
                  f"— these will never appear in module tally!")
        else:
            print(f"    [{name}]  all closed incidents have a closure date  ✓")
    print()

    for name in incident_sheet_names:
        raw_df, filtered_df = process_sheet(raw_sheets[name], start_dt, end_dt)
        processed_raw[name] = raw_df
        filtered_raw[name]  = filtered_df
        counts[name]        = len(filtered_df)
        status = (f"{counts[name]} unique incident(s) in range"
                  if counts[name] else "no incidents in range")
        print(f"  [{name}]  total rows = {len(raw_df)}  |  {status}")

        if COL_CLOSURE_DATE in raw_df.columns:
            orig_strings = raw_sheets[name][COL_CLOSURE_DATE].astype(str).str.strip()
            nat_rows     = raw_df[raw_df[COL_CLOSURE_DATE].isna()]
            truly_bad    = nat_rows[
                orig_strings.loc[nat_rows.index].str.lower().isin(
                    ["", "nan", "none", "nat", "n/a", "-"]
                ) == False
            ]
            if not truly_bad.empty:
                samples = orig_strings.loc[truly_bad.index].unique()[:5]
                print(f"    ⚠  {len(truly_bad)} row(s) had unparseable closure dates "
                      f"— excluded from count!")
                print(f"    ⚠  Sample date strings that failed: {list(samples)}")

    # Step 3 — aggregate incident modules
    module_names, module_counts = aggregate(counts)
    grand_total = sum(module_counts)
    skipped     = len(incident_sheet_names) - len(module_names)

    print(f"\n  Modules with incidents : {len(module_names)}")
    if skipped:
        print(f"  Modules excluded (zero): {skipped}  (not written to output)")
    print(f"  Grand total            : {grand_total}")

    # Step 3b — overall status breakdown
    status_labels, status_counts = compute_status_breakdown(
        processed_raw, incident_sheet_names
    )

    # Step 3c — user-wise breakdown
    user_names, user_counts = compute_user_breakdown(
        filtered_raw, incident_sheet_names
    )

    # Step 3d — credential breach analysis  (NEW)
    print(f"\n{'=' * 60}")
    print(f"  CREDENTIAL BREACH ANALYSIS")
    print(f"{'=' * 60}")
    pwd_labels, pwd_counts, top_emails = compute_credential_breach_analysis(raw_sheets)

    # Build output path
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    date_range_str = (
        f"{start_dt.strftime('%d %b %Y')} - {end_dt.strftime('%d %b %Y')}"
    )
    output_path = os.path.join(
        OUTPUT_FOLDER,
        f"CloudSek Incident Review - {date_range_str}.xlsx"
    )

    # Step 4 — write workbook
    build_workbook(
        module_names, module_counts,
        status_labels, status_counts,
        user_names, user_counts,
        pwd_labels, pwd_counts, top_emails,
        filtered_raw, counts, incident_sheet_names,
        start_dt, end_dt,
        output_path,
    )

    print(f"\n  Report saved -> {output_path}")
    print("Done.\n")


if __name__ == "__main__":
    main()
