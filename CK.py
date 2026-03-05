"""
=============================================================================
INCIDENT REPORT GENERATOR
=============================================================================
CONFIGURATION — only 3 things to set:
  INPUT_FILE_PATH  : path to the source Excel workbook
  COL_INCIDENT_ID  : exact column name for the Incident ID
  COL_CLOSURE_DATE : exact column name for the Closure Date

Everything else is automatic — date range is prompted at runtime.
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
# >>>  CONFIGURATION — only edit these three lines  <<<
# =============================================================================

INPUT_FILE_PATH  = "incidents.xlsx"
COL_INCIDENT_ID  = "Incident Id"
COL_CLOSURE_DATE = "Incident Closure on"

# =============================================================================
# INTERNALS
# =============================================================================

OUTPUT_FILE_PATH   = "incidents_report.xlsx"
SUMMARY_SHEET_NAME = "Summary Dashboard"

# ---------------------------------------------------------------------------
# DATE PARSING
# ---------------------------------------------------------------------------

_ORDINAL_RE   = re.compile(r"(\d+)(st|nd|rd|th)\b", re.IGNORECASE)
_DATE_FORMATS = [
    "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d-%B-%Y",
    "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d.%m.%Y",
    "%d %b %y", "%d %B %y",
]

def _strip_ordinals(text):
    return _ORDINAL_RE.sub(r"\1", str(text))

def _parse_date_series(series):
    cleaned = series.astype(str).apply(_strip_ordinals)
    parsed  = pd.to_datetime(cleaned, infer_datetime_format=True,
                             dayfirst=True, errors="coerce")
    for fmt in _DATE_FORMATS:
        if not parsed.isna().any():
            break
        bad = parsed.isna()
        parsed[bad] = pd.to_datetime(cleaned[bad], format=fmt, errors="coerce")
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
    return pd.read_excel(path, sheet_name=None, dtype=str)

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
    """
    counts is {sheet_name: unique_count}.
    Sheet names ARE the module names.
    Only include modules that have at least 1 incident.
    Returns two plain lists — no DataFrame column name issues downstream.
    """
    module_names = []
    module_counts = []
    for sheet_name, count in counts.items():
        if count > 0:
            module_names.append(sheet_name)
            module_counts.append(count)
    if not module_names:
        sys.exit("\nNo incidents found in the specified date range.\n")
    return module_names, module_counts

# ---------------------------------------------------------------------------
# STEP 4 — BUILD OUTPUT WORKBOOK
# ---------------------------------------------------------------------------

def build_workbook(module_names, module_counts, processed_raw,
                   sheet_names, start_dt, end_dt):

    wb = xlsxwriter.Workbook(OUTPUT_FILE_PATH)

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
    #
    # Row layout (0-indexed):
    #   0        : banner
    #   1        : spacer
    #   2        : section title
    #   3        : column headers
    #   4 .. 4+n-1 : data rows  (one per active module)
    #   4+n      : TOTAL row
    #
    n           = len(module_names)   # only modules with incidents
    HDR_ROW     = 3                   # 0-indexed header row
    DATA_START  = 4                   # 0-indexed first data row
    TOTAL_ROW   = DATA_START + n      # 0-indexed total row

    # Excel formula uses 1-indexed rows
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

    # Data rows — iterate plain lists, zero ambiguity
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

    # ── BAR CHART ─────────────────────────────────────────────────────────────
    bar = wb.add_chart({"type": "column"})
    bar.add_series({
        "name":       "Incidents",
        "categories": [SUMMARY_SHEET_NAME, DATA_START, 1, DATA_START + n - 1, 1],
        "values":     [SUMMARY_SHEET_NAME, DATA_START, 2, DATA_START + n - 1, 2],
        "gap":        80,
    })
    bar.set_title({"name": "Unique Closed Incidents by Module"})
    bar.set_x_axis({"name": "Module"})
    bar.set_y_axis({"name": "Count"})
    bar.set_style(10)
    bar_w = min(max(n * 90, 400), 720)
    bar.set_size({"width": bar_w, "height": 300})
    sw.insert_chart(TOTAL_ROW + 2, 0, bar)

    # ── PIE CHART ─────────────────────────────────────────────────────────────
    pie = wb.add_chart({"type": "pie"})
    pie.add_series({
        "name":       "Incidents",
        "categories": [SUMMARY_SHEET_NAME, DATA_START, 1, DATA_START + n - 1, 1],
        "values":     [SUMMARY_SHEET_NAME, DATA_START, 2, DATA_START + n - 1, 2],
    })
    pie.set_title({"name": "Incident Share by Module"})
    pie.set_style(10)
    pie.set_size({"width": 420, "height": 300})
    pie_col = round(bar_w / 64) + 1
    sw.insert_chart(TOTAL_ROW + 2, pie_col, pie)

    # ── DATA SHEETS ───────────────────────────────────────────────────────────
    for name in sheet_names:
        df       = processed_raw[name]
        ws_name  = name[:31]          # Excel sheet name max 31 chars
        dw       = wb.add_worksheet(ws_name)

        if df.empty:
            continue

        headers = list(df.columns)
        n_cols  = len(headers)
        n_rows  = len(df)

        # Header
        dw.set_row(0, 20)
        for ci, h in enumerate(headers):
            dw.write(0, ci, h, f_data_hdr)
            dw.set_column(ci, ci, max(len(str(h)) + 4, 14))

        # Data — iterate with iterrows() so column access by name is safe
        for ri, (_, row) in enumerate(df.iterrows(), 1):
            dw.set_row(ri, 16)
            alt = (ri % 2 == 0)
            for ci, col_name in enumerate(headers):
                val          = row[col_name]
                is_date_col  = (col_name == COL_CLOSURE_DATE)
                try:
                    is_null = pd.isna(val)
                except (TypeError, ValueError):
                    is_null = False

                if is_null:
                    dw.write_blank(ri, ci, None,
                                   f_date_alt if (is_date_col and alt) else
                                   f_date     if is_date_col else
                                   f_cell_alt if alt else f_cell)
                elif isinstance(val, pd.Timestamp):
                    dw.write_datetime(ri, ci, val.to_pydatetime(),
                                      f_date_alt if alt else f_date)
                else:
                    dw.write(ri, ci, val,
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

    print(f"\n  Range: {start_dt.strftime('%d %b %Y')} -> {end_dt.strftime('%d %b %Y')}\n")

    # Step 1 — load
    raw_sheets  = load_all_sheets(INPUT_FILE_PATH)
    sheet_names = list(raw_sheets.keys())
    print(f"Sheets found ({len(sheet_names)}): {sheet_names}\n")

    # Step 2 — process each sheet
    processed_raw = {}
    counts        = {}     # {sheet_name: unique_incident_count}
    for name in sheet_names:
        raw_df, filtered_df = process_sheet(raw_sheets[name], start_dt, end_dt)
        processed_raw[name] = raw_df
        counts[name]        = len(filtered_df)
        status = (f"{counts[name]} unique incident(s) in range"
                  if counts[name] else "no incidents in range")
        print(f"  [{name}]  total rows = {len(raw_df)}  |  {status}")

    # Step 3 — aggregate into plain lists (no DataFrame column-name ambiguity)
    module_names, module_counts = aggregate(counts)
    grand_total = sum(module_counts)
    skipped     = len(sheet_names) - len(module_names)

    print(f"\n  Modules with incidents : {len(module_names)}")
    if skipped:
        print(f"  Modules skipped (zero) : {skipped}  (sheets still written)")
    print(f"  Grand total            : {grand_total}")

    # Step 4 — write output
    build_workbook(module_names, module_counts,
                   processed_raw, sheet_names, start_dt, end_dt)

    print(f"\n  Report saved -> {OUTPUT_FILE_PATH}")
    print("Done.\n")


if __name__ == "__main__":
    main()
