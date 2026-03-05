"""
=============================================================================
INCIDENT REPORT GENERATOR
=============================================================================
CONFIGURATION — only 3 things to set:
  INPUT_FILE_PATH  : path to the source Excel workbook
  COL_INCIDENT_ID  : exact column name for the Incident ID
  COL_CLOSURE_DATE : exact column name for the Closure Date

Everything else is automatic:
  - Date range is prompted interactively at runtime
  - Any number of sheets / modules is handled (1 to 20+)
  - Only modules with incidents in the range appear in charts/summary
  - Sheet names are used directly as Module names
  - Date formats like "2 Mar 2026", "2nd March 2026", "01/03/2026" all work

Dependencies:  pip install pandas xlsxwriter openpyxl
  (openpyxl is used only to READ the input file; xlsxwriter handles all output)
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
    rows = [
        {"Module Name": k, "Unique Closed Incidents": v}
        for k, v in counts.items() if v > 0
    ]
    if not rows:
        sys.exit("\nNo incidents found in the specified date range.\n")
    return pd.DataFrame(rows)

# ---------------------------------------------------------------------------
# STEP 4 — BUILD OUTPUT WORKBOOK  (xlsxwriter — write-only, reliable charts)
# ---------------------------------------------------------------------------

def build_workbook(summary_df, processed_raw, sheet_names, start_dt, end_dt):

    wb = xlsxwriter.Workbook(OUTPUT_FILE_PATH)

    # ── FORMATS ──────────────────────────────────────────────────────────────
    NAVY    = "#1F3864"
    MIDBLUE = "#2E75B6"
    LTBLUE  = "#D6E4F7"
    ALT     = "#EFF5FB"
    WHITE   = "#FFFFFF"

    def _fmt(d):
        return wb.add_format(d)

    f_banner = _fmt({
        "bold": True, "font_name": "Arial", "font_size": 15,
        "font_color": WHITE, "bg_color": NAVY,
        "align": "center", "valign": "vcenter",
    })
    f_section = _fmt({
        "bold": True, "font_name": "Arial", "font_size": 13,
        "font_color": NAVY, "valign": "vcenter",
    })
    f_col_hdr = _fmt({
        "bold": True, "font_name": "Arial", "font_size": 11,
        "font_color": WHITE, "bg_color": NAVY,
        "align": "center", "valign": "vcenter", "border": 1,
    })
    f_num = _fmt({
        "font_name": "Arial", "font_size": 10,
        "align": "center", "valign": "vcenter", "border": 1,
    })
    f_lft = _fmt({
        "font_name": "Arial", "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1,
    })
    f_num_alt = _fmt({
        "font_name": "Arial", "font_size": 10,
        "align": "center", "valign": "vcenter", "border": 1, "bg_color": ALT,
    })
    f_lft_alt = _fmt({
        "font_name": "Arial", "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1, "bg_color": ALT,
    })
    f_total_merge = _fmt({
        "bold": True, "font_name": "Arial", "font_size": 11,
        "font_color": NAVY, "bg_color": LTBLUE,
        "align": "center", "valign": "vcenter", "border": 2,
    })
    f_total_val = _fmt({
        "bold": True, "font_name": "Arial", "font_size": 11,
        "font_color": NAVY, "bg_color": LTBLUE,
        "align": "center", "valign": "vcenter", "border": 2,
    })
    f_data_hdr = _fmt({
        "bold": True, "font_name": "Arial", "font_size": 10,
        "font_color": WHITE, "bg_color": MIDBLUE,
        "align": "center", "valign": "vcenter", "border": 1,
    })
    f_cell = _fmt({
        "font_name": "Arial", "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1,
    })
    f_cell_alt = _fmt({
        "font_name": "Arial", "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1, "bg_color": ALT,
    })
    f_date = _fmt({
        "font_name": "Arial", "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1,
        "num_format": "dd mmm yyyy",
    })
    f_date_alt = _fmt({
        "font_name": "Arial", "font_size": 10,
        "align": "left", "valign": "vcenter", "border": 1, "bg_color": ALT,
        "num_format": "dd mmm yyyy",
    })

    # ── SUMMARY SHEET ─────────────────────────────────────────────────────────
    #
    # xlsxwriter uses 0-indexed (row, col) everywhere EXCEPT formula strings
    # which use standard Excel A1 notation (1-indexed).
    #
    # Layout (0-indexed rows):
    #   row 0  : banner
    #   row 1  : (empty spacer)
    #   row 2  : section title
    #   row 3  : column headers  (#, Module Name, Unique Closed Incidents)
    #   row 4  : first data row
    #   ...
    #   row 4+n-1 : last data row
    #   row 4+n   : TOTAL row
    #
    n            = len(summary_df)
    DATA_ROW_0   = 4          # 0-indexed first data row
    total_row_0  = DATA_ROW_0 + n   # 0-indexed total row

    # Excel 1-indexed equivalents for formula only
    excel_data_start = DATA_ROW_0 + 1        # = 5
    excel_data_end   = DATA_ROW_0 + n        # = 4+n
    sum_formula      = f"=SUM(C{excel_data_start}:C{excel_data_end})"

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
    sw.set_row(3, 22)
    sw.write(3, 0, "#",                      f_col_hdr)
    sw.write(3, 1, "Module Name",             f_col_hdr)
    sw.write(3, 2, "Unique Closed Incidents", f_col_hdr)

    # Data rows
    for i, row in enumerate(summary_df.itertuples(index=False)):
        r   = DATA_ROW_0 + i
        d   = row._asdict()
        alt = (i % 2 == 1)
        sw.set_row(r, 18)
        sw.write(r, 0, i + 1,                         f_num_alt if alt else f_num)
        sw.write(r, 1, d["Module Name"],               f_lft_alt if alt else f_lft)
        sw.write(r, 2, d["Unique Closed Incidents"],   f_num_alt if alt else f_num)

    # Total row
    sw.set_row(total_row_0, 22)
    sw.merge_range(total_row_0, 0, total_row_0, 1, "TOTAL", f_total_merge)
    sw.write_formula(total_row_0, 2, sum_formula, f_total_val)

    # Column widths
    sw.set_column(0, 0, 6)
    sw.set_column(1, 1, max(
        summary_df["Module Name"].astype(str).str.len().max() + 6, 24
    ))
    sw.set_column(2, 2, 28)

    # Freeze pane below banner + spacer + section title
    sw.freeze_panes(3, 0)

    # ── BAR CHART ────────────────────────────────────────────────────────────
    #
    # xlsxwriter chart series references:
    #   [sheet_name, first_row, first_col, last_row, last_col]  — all 0-indexed
    #
    bar = wb.add_chart({"type": "column"})
    bar.add_series({
        "name":       "Incidents",
        "categories": [SUMMARY_SHEET_NAME,
                       DATA_ROW_0, 1, DATA_ROW_0 + n - 1, 1],
        "values":     [SUMMARY_SHEET_NAME,
                       DATA_ROW_0, 2, DATA_ROW_0 + n - 1, 2],
        "gap":        100,
    })
    bar.set_title({"name": "Unique Closed Incidents by Module"})
    bar.set_x_axis({"name": "Module"})
    bar.set_y_axis({"name": "Count"})
    bar.set_style(10)
    bar_px_w = min(max(n * 90, 400), 720)
    bar.set_size({"width": bar_px_w, "height": 300})
    sw.insert_chart(total_row_0 + 2, 0, bar)

    # ── PIE CHART ────────────────────────────────────────────────────────────
    pie = wb.add_chart({"type": "pie"})
    pie.add_series({
        "name":       "Incidents",
        "categories": [SUMMARY_SHEET_NAME,
                       DATA_ROW_0, 1, DATA_ROW_0 + n - 1, 1],
        "values":     [SUMMARY_SHEET_NAME,
                       DATA_ROW_0, 2, DATA_ROW_0 + n - 1, 2],
    })
    pie.set_title({"name": "Incident Share by Module"})
    pie.set_style(10)
    pie.set_size({"width": 420, "height": 300})
    # Offset pie to the right of bar — 64px ≈ 1 default Excel column
    pie_col_offset = round(bar_px_w / 64) + 1
    sw.insert_chart(total_row_0 + 2, pie_col_offset, pie)

    # ── DATA SHEETS ──────────────────────────────────────────────────────────
    for name in sheet_names:
        df = processed_raw[name]
        # Excel sheet names max 31 chars
        safe_name = name[:31]
        dw = wb.add_worksheet(safe_name)

        if df.empty:
            continue

        headers = list(df.columns)
        n_cols  = len(headers)
        n_rows  = len(df)

        # Header row (0-indexed row 0)
        dw.set_row(0, 20)
        for ci, h in enumerate(headers):
            dw.write(0, ci, h, f_data_hdr)
            dw.set_column(ci, ci, max(len(str(h)) + 4, 14))

        # Data rows
        for ri, row_data in enumerate(df.itertuples(index=False), 1):
            dw.set_row(ri, 16)
            alt = (ri % 2 == 0)
            for ci, val in enumerate(row_data):
                # Coerce NaT / NaN to None
                try:
                    if pd.isna(val):
                        val = None
                except (TypeError, ValueError):
                    pass

                is_date_col = (headers[ci] == COL_CLOSURE_DATE)

                if val is None:
                    fmt = f_date_alt if (is_date_col and alt) else \
                          f_date    if is_date_col else \
                          f_cell_alt if alt else f_cell
                    dw.write_blank(ri, ci, None, fmt)
                elif isinstance(val, pd.Timestamp):
                    fmt = f_date_alt if alt else f_date
                    dw.write_datetime(ri, ci, val.to_pydatetime(), fmt)
                else:
                    fmt = f_cell_alt if alt else f_cell
                    dw.write(ri, ci, val, fmt)

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

    # Step 1
    raw_sheets  = load_all_sheets(INPUT_FILE_PATH)
    sheet_names = list(raw_sheets.keys())
    print(f"Sheets found ({len(sheet_names)}): {sheet_names}\n")

    # Step 2
    processed_raw = {}
    counts        = {}
    for name in sheet_names:
        raw_df, filtered_df = process_sheet(raw_sheets[name], start_dt, end_dt)
        processed_raw[name] = raw_df
        counts[name]        = len(filtered_df)
        status = (f"{counts[name]} unique incident(s) in range"
                  if counts[name] else "no incidents in range")
        print(f"  [{name}]  total rows = {len(raw_df)}  |  {status}")

    # Step 3
    summary_df  = aggregate(counts)
    grand_total = summary_df["Unique Closed Incidents"].sum()
    active      = len(summary_df)
    skipped     = len(sheet_names) - active
    print(f"\n  Modules with incidents : {active}")
    if skipped:
        print(f"  Modules skipped (zero) : {skipped}  (sheets still written to output)")
    print(f"  Grand total            : {grand_total}")

    # Step 4
    build_workbook(summary_df, processed_raw, sheet_names, start_dt, end_dt)

    print(f"\n  Report saved -> {OUTPUT_FILE_PATH}")
    print("Done.\n")


if __name__ == "__main__":
    main()
