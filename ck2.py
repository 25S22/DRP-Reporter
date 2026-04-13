"""
=============================================================================
INCIDENT REPORT GENERATOR
=============================================================================
CONFIGURATION — only 5 things to set:
  INPUT_FILE_PATH  : path to the source Excel workbook
  COL_INCIDENT_ID  : exact column name for the Incident ID
  COL_CLOSURE_DATE : exact column name for the Closure Date
  COL_STATUS       : exact column name for the Status column
  COL_CLOSED_BY    : exact column name for who closed the incident
  COL_COMMENT      : exact column name for the comment/notes column

Everything else is automatic — date range is prompted at runtime.

Summary sheet contains:
  1. Module-wise unique closed incidents table (date range filtered)
  2. Pie chart — "Closed Incidents By Module" (date range filtered)
  3. Pie chart — "Overall Incident Status" (all data, unique IDs)
  4. User-wise closed incidents table + pie (date range filtered)

All pie chart labels are printed OUTSIDE the slice with leader lines —
  name, count and % shown in bold navy text, never overlapping the colours.

Email extraction:
  Unique email addresses are harvested from COL_COMMENT for all
  "closed - resolved" incidents (all data, not date-range limited)
  and written to an "Emails - Closed Resolved" sheet.
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
COL_COMMENT      = "Comment"             # column to mine for email addresses

# =============================================================================
# INTERNALS
# =============================================================================

SUMMARY_SHEET_NAME = "Summary Dashboard"

# ---------------------------------------------------------------------------
# DATE PARSING
# ---------------------------------------------------------------------------

_ORDINAL_RE   = re.compile(r"(\d+)(st|nd|rd|th)\b", re.IGNORECASE)

# Formats listed BOTH with and WITHOUT the comma that can appear after the
# month abbreviation (e.g. "06 Apr, 2026" and "06 Apr 2026" are both tried).
# Lower-case am/pm is normalised to upper-case before matching (see _clean).
_DATE_FORMATS = [
    # ── with comma after month name ──────────────────────────────────────────
    "%d %b, %Y %I:%M:%S %p",   # 18 Feb, 2026 06:11:54 PM
    "%d %B, %Y %I:%M:%S %p",
    "%d %b, %Y %H:%M:%S",      # 18 Feb, 2026 18:11:54
    "%d %B, %Y %H:%M:%S",
    "%d %b, %Y %I:%M %p",      # 18 Feb, 2026 06:11 PM
    "%d %B, %Y %I:%M %p",
    "%d %b, %Y",                # 18 Feb, 2026
    "%d %B, %Y",
    # ── WITHOUT comma after month name ───────────────────────────────────────
    # These are used when _clean() strips the comma (or the source never had one)
    "%d %b %Y %I:%M:%S %p",    # 18 Feb 2026 06:11:54 PM  ← was MISSING — key fix
    "%d %B %Y %I:%M:%S %p",
    "%d %b %Y %H:%M:%S",       # 18 Feb 2026 18:11:54
    "%d %B %Y %H:%M:%S",
    "%d %b %Y %I:%M %p",       # 18 Feb 2026 06:11 PM
    "%d %B %Y %I:%M %p",
    "%d %b %Y", "%d %B %Y",
    "%d-%b-%Y", "%d-%B-%Y",
    # ── numeric variants ─────────────────────────────────────────────────────
    "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d.%m.%Y",
    "%d %b %y", "%d %B %y",
]

def _strip_ordinals(text):
    return _ORDINAL_RE.sub(r"\1", str(text))

def _clean(text):
    """
    Normalise a raw date string so it can be matched by strptime:
      1. Remove ordinal suffixes  (1st -> 1, 2nd -> 2 …)
      2. Uppercase am/pm          (pm -> PM, am -> AM)
         %p is locale-dependent and fails on lowercase on many platforms.
      3. Strip the comma that sometimes follows a month abbreviation
         ("Apr," -> "Apr") so the WITHOUT-comma formats below can match.
    """
    t = _strip_ordinals(str(text))
    # Uppercase am/pm — must come BEFORE comma removal to avoid regex conflict
    t = re.sub(r'\b(am|pm)\b', lambda m: m.group(0).upper(), t, flags=re.IGNORECASE)
    # Remove comma after letter (handles "Apr,", "February," etc.)
    t = re.sub(r'([A-Za-z]),', r'\1', t)
    return t.strip()

def _parse_date_series(series):
    """
    Robustly parse a Series into datetime.
    Handles ordinals, comma-after-month, lowercase am/pm,
    standard formats, and Excel numeric serials.
    """
    cleaned = series.astype(str).apply(_clean)

    # Pass 1: pandas inference on cleaned strings
    parsed = pd.to_datetime(cleaned, infer_datetime_format=True,
                            dayfirst=True, errors="coerce")

    # Pass 2: explicit format list — tries every format for still-bad rows
    for fmt in _DATE_FORMATS:
        if not parsed.isna().any():
            break
        bad = parsed.isna()
        parsed[bad] = pd.to_datetime(cleaned[bad], format=fmt, errors="coerce")

    # Pass 3: Excel numeric serial fallback
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
        cleaned = _clean(raw)
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

    # Load twice:
    #   str version    — keeps IDs, Status, Comment, all text columns exactly as-is
    #   native version — lets pandas parse the date column properly
    # Then splice: replace only the date column with the natively-parsed version.
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
# STEP 3b — OVERALL STATUS BREAKDOWN  (all data, unique incident IDs)
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

    order         = ["Open", "In Progress", "Closed"]
    labels        = [s for s in order if s in counts.index]
    status_counts = [int(counts[s]) for s in labels]

    print(f"\n  Overall status breakdown (unique incidents, 3 buckets only):")
    for lbl, cnt in zip(labels, status_counts):
        print(f"    {lbl}: {cnt}")

    return labels, status_counts

# ---------------------------------------------------------------------------
# STEP 3c — USER-WISE BREAKDOWN  (date range filtered, unique incident IDs)
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
# STEP 3d — EMAIL EXTRACTION  (Status = "closed - resolved", all data)
# ---------------------------------------------------------------------------

_EMAIL_RE = re.compile(
    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}",
    re.IGNORECASE,
)

def extract_emails_closed_resolved(processed_raw, sheet_names):
    """
    Scan COL_COMMENT for every row whose COL_STATUS is exactly
    'closed - resolved' (case-insensitive) across ALL sheets.
    Deduplicates on COL_INCIDENT_ID first, then collects all emails,
    returning a sorted list of unique addresses.
    """
    if not COL_COMMENT:
        print("  INFO: COL_COMMENT not configured — email extraction skipped.")
        return []

    frames = []
    for name in sheet_names:
        df = processed_raw.get(name, pd.DataFrame())
        if df.empty:
            continue
        needed = [c for c in [COL_INCIDENT_ID, COL_STATUS, COL_COMMENT]
                  if c in df.columns]
        if COL_STATUS not in needed or COL_COMMENT not in needed:
            continue
        frames.append(df[needed].copy())

    if not frames:
        print(f"  WARNING: Required columns not found — email extraction skipped.")
        return []

    combined = pd.concat(frames, ignore_index=True)

    # Deduplicate on incident ID so a multi-sheet incident isn't counted twice
    if COL_INCIDENT_ID in combined.columns:
        combined = combined.drop_duplicates(subset=[COL_INCIDENT_ID])

    # Filter for "closed - resolved" (case-insensitive, strip whitespace)
    mask     = (
        combined[COL_STATUS]
        .astype(str).str.strip().str.lower()
        == "closed - resolved"
    )
    cr_df    = combined[mask]

    print(f"\n  Rows with status 'closed - resolved': {len(cr_df)}")

    if cr_df.empty or COL_COMMENT not in cr_df.columns:
        print("  No matching rows found for email extraction.")
        return []

    emails = set()
    for cell_value in cr_df[COL_COMMENT].dropna():
        for match in _EMAIL_RE.findall(str(cell_value)):
            emails.add(match.lower().strip())

    unique_emails = sorted(emails)
    print(f"  Unique email addresses found: {len(unique_emails)}")
    return unique_emails

# ---------------------------------------------------------------------------
# STEP 4 — BUILD OUTPUT WORKBOOK
# ---------------------------------------------------------------------------

def build_workbook(module_names, module_counts, status_labels, status_counts,
                   user_names, user_counts, unique_emails,
                   filtered_raw, counts, sheet_names,
                   start_dt, end_dt, output_path):

    wb = xlsxwriter.Workbook(output_path)

    # ── FORMATS ──────────────────────────────────────────────────────────────
    NAVY    = "#1F3864"
    MIDBLUE = "#2E75B6"
    LTBLUE  = "#D6E4F7"
    ALT     = "#EFF5FB"
    WHITE   = "#FFFFFF"

    f_banner = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 15,
        "font_color": WHITE, "bg_color": NAVY,
        "align": "center", "valign": "vcenter",
    })
    f_section = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 13,
        "font_color": NAVY, "valign": "vcenter",
    })
    f_col_hdr = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 11,
        "font_color": WHITE, "bg_color": NAVY,
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

    # Banner
    sw.set_row(0, 42)
    sw.merge_range(
        0, 0, 0, 15,
        f"Incident Summary Dashboard  |  "
        f"{start_dt.strftime('%d %b %Y')}  ->  {end_dt.strftime('%d %b %Y')}",
        f_banner,
    )

    # Section title
    sw.set_row(2, 24)
    sw.write(2, 0, "Module-wise Unique Closed Incidents", f_section)

    # Column headers
    sw.set_row(HDR_ROW, 22)
    sw.write(HDR_ROW, 0, "#",                       f_col_hdr)
    sw.write(HDR_ROW, 1, "Module Name",              f_col_hdr)
    sw.write(HDR_ROW, 2, "Unique Closed Incidents",  f_col_hdr)

    # Data rows
    for i in range(n):
        r   = DATA_START + i
        alt = (i % 2 == 1)
        sw.set_row(r, 18)
        sw.write(r, 0, i + 1,            f_num_alt if alt else f_num)
        sw.write(r, 1, module_names[i],  f_lft_alt if alt else f_lft)
        sw.write(r, 2, module_counts[i], f_num_alt if alt else f_num)

    # Total row
    sw.set_row(TOTAL_ROW, 22)
    sw.merge_range(TOTAL_ROW, 0, TOTAL_ROW, 1, "TOTAL", f_total)
    sw.write_formula(TOTAL_ROW, 2, sum_formula, f_total)

    # Column widths
    sw.set_column(0, 0, 6)
    sw.set_column(1, 1, max(max(len(m) for m in module_names) + 6, 24))
    sw.set_column(2, 2, 28)
    sw.freeze_panes(HDR_ROW, 0)

    # ── PIE COLOUR PALETTE ───────────────────────────────────────────────────
    _PALETTE = [
        "#4472C4", "#ED7D31", "#70AD47", "#FFC000", "#5B9BD5",
        "#A9D18E", "#FF7C80", "#9E480E", "#7030A0", "#636363",
        "#255E91", "#43682B", "#C00000", "#997300", "#7B5EA7",
    ]

    def _pie_points(count):
        return [{"fill": {"color": _PALETTE[i % len(_PALETTE)]}}
                for i in range(count)]

    def _write_legend(ws, start_row, start_col, labels, values):
        total = sum(values) or 1
        hdr_fmt = wb.add_format({
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
            colour = _PALETTE[i % len(_PALETTE)]
            alt    = (i % 2 == 1)
            bg     = ALT if alt else WHITE

            swatch_fmt = wb.add_format({"bg_color": colour, "border": 1})
            name_fmt   = wb.add_format({
                "font_name": "Arial", "font_size": 9, "border": 1,
                "align": "left", "valign": "vcenter", "bg_color": bg,
            })
            val_fmt = wb.add_format({
                "font_name": "Arial", "font_size": 9, "border": 1,
                "align": "center", "valign": "vcenter", "bg_color": bg,
            })
            pct_fmt = wb.add_format({
                "font_name": "Arial", "font_size": 9, "border": 1,
                "align": "center", "valign": "vcenter", "bg_color": bg,
                "num_format": "0.0%",
            })

            ws.write(r, start_col,     "",         swatch_fmt)
            ws.write(r, start_col + 1, lbl,        name_fmt)
            ws.write(r, start_col + 2, val,        val_fmt)
            ws.write(r, start_col + 3, val / total, pct_fmt)

        ws.set_column(start_col,     start_col,     3)
        ws.set_column(start_col + 1, start_col + 1,
                      max((len(str(l)) for l in labels), default=12) + 4)
        ws.set_column(start_col + 2, start_col + 2, 10)
        ws.set_column(start_col + 3, start_col + 3, 10)

    # Chart anchor row (0-indexed)
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
                 "Incidents Closed By — User Wise (Date Range)", f_section)

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

    # ── EMAIL SHEET — unique emails from "closed - resolved" comments ─────────
    if unique_emails:
        f_email_hdr = wb.add_format({
            "bold": True, "font_name": "Arial", "font_size": 11,
            "font_color": WHITE, "bg_color": NAVY,
            "align": "center", "valign": "vcenter", "border": 1,
        })
        f_email_banner = wb.add_format({
            "bold": True, "font_name": "Arial", "font_size": 13,
            "font_color": WHITE, "bg_color": NAVY,
            "align": "center", "valign": "vcenter",
        })
        f_email_cell = wb.add_format({
            "font_name": "Arial", "font_size": 10,
            "align": "left", "valign": "vcenter", "border": 1,
        })
        f_email_cell_alt = wb.add_format({
            "font_name": "Arial", "font_size": 10,
            "align": "left", "valign": "vcenter", "border": 1, "bg_color": ALT,
        })
        f_email_num = wb.add_format({
            "font_name": "Arial", "font_size": 10,
            "align": "center", "valign": "vcenter", "border": 1,
        })
        f_email_num_alt = wb.add_format({
            "font_name": "Arial", "font_size": 10,
            "align": "center", "valign": "vcenter", "border": 1, "bg_color": ALT,
        })

        ew = wb.add_worksheet("Emails - Closed Resolved")

        # Banner
        ew.set_row(0, 36)
        ew.merge_range(
            0, 0, 0, 2,
            f"Unique Emails — Status: Closed - Resolved  |  {len(unique_emails)} addresses",
            f_email_banner,
        )

        # Column headers
        ew.set_row(2, 22)
        ew.write(2, 0, "#",             f_email_hdr)
        ew.write(2, 1, "Email Address", f_email_hdr)

        ew.set_column(0, 0, 6)
        ew.set_column(1, 1, max(len(e) for e in unique_emails) + 6)
        ew.freeze_panes(3, 0)

        for i, email in enumerate(unique_emails):
            r   = 3 + i
            alt = (i % 2 == 1)
            ew.set_row(r, 16)
            ew.write(r, 0, i + 1, f_email_num_alt if alt else f_email_num)
            ew.write(r, 1, email,  f_email_cell_alt if alt else f_email_cell)

    # ── DATA SHEETS — only modules with incidents in range ───────────────────
    for name in sheet_names:
        if counts[name] == 0:
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

    # Step 1 — load
    raw_sheets  = load_all_sheets(INPUT_FILE_PATH)
    sheet_names = list(raw_sheets.keys())
    print(f"Sheets found ({len(sheet_names)}): {sheet_names}\n")

    # ── DIAGNOSTIC ─────────────────────────────────────────────────────────
    print("\n  DIAGNOSTIC — checking for closed incidents with missing closure dates:")
    for name in sheet_names:
        df = raw_sheets.get(name, pd.DataFrame())
        if df.empty or COL_INCIDENT_ID not in df.columns:
            continue
        if COL_STATUS not in df.columns or COL_CLOSURE_DATE not in df.columns:
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

    # Step 2 — process each sheet
    processed_raw = {}
    filtered_raw  = {}
    counts        = {}

    for name in sheet_names:
        raw_df, filtered_df = process_sheet(raw_sheets[name], start_dt, end_dt)
        processed_raw[name] = raw_df
        filtered_raw[name]  = filtered_df
        counts[name]        = len(filtered_df)
        status = (f"{counts[name]} unique incident(s) in range"
                  if counts[name] else "no incidents in range")
        print(f"  [{name}]  total rows = {len(raw_df)}  |  {status}")

        # Diagnostic: unparseable closure dates
        if COL_CLOSURE_DATE in raw_df.columns:
            orig_strings = raw_sheets[name][COL_CLOSURE_DATE].astype(str).str.strip()
            nat_rows     = raw_df[raw_df[COL_CLOSURE_DATE].isna()]
            truly_bad    = nat_rows[
                ~orig_strings.loc[nat_rows.index].str.lower().isin(
                    ["", "nan", "none", "nat", "n/a", "-"]
                )
            ]
            if not truly_bad.empty:
                samples = orig_strings.loc[truly_bad.index].unique()[:5]
                print(f"    ⚠  {len(truly_bad)} row(s) had unparseable closure dates — excluded!")
                print(f"    ⚠  Sample date strings that failed: {list(samples)}")

    # Step 3 — aggregate
    module_names, module_counts = aggregate(counts)
    grand_total = sum(module_counts)
    skipped     = len(sheet_names) - len(module_names)

    print(f"\n  Modules with incidents : {len(module_names)}")
    if skipped:
        print(f"  Modules excluded (zero): {skipped}  (not written to output)")
    print(f"  Grand total            : {grand_total}")

    # Step 3b — overall status breakdown
    status_labels, status_counts = compute_status_breakdown(processed_raw, sheet_names)

    # Step 3c — user-wise breakdown
    user_names, user_counts = compute_user_breakdown(filtered_raw, sheet_names)

    # Step 3d — email extraction from closed-resolved comments
    unique_emails = extract_emails_closed_resolved(processed_raw, sheet_names)

    # Build output path
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    date_range_str = (
        f"{start_dt.strftime('%d %b %Y')} - {end_dt.strftime('%d %b %Y')}"
    )
    output_path = os.path.join(
        OUTPUT_FOLDER,
        f"CloudSek Incident Review - {date_range_str}.xlsx"
    )

    # Step 4 — write output
    build_workbook(
        module_names, module_counts, status_labels, status_counts,
        user_names, user_counts, unique_emails,
        filtered_raw, counts, sheet_names, start_dt, end_dt,
        output_path,
    )

    print(f"\n  Report saved -> {output_path}")
    print("Done.\n")


if __name__ == "__main__":
    main()
