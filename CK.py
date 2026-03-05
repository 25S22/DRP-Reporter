"""
=============================================================================
INCIDENT REPORT GENERATOR
=============================================================================
CONFIGURATION (only 3 things to set):
  INPUT_FILE_PATH  — path to the source Excel workbook
  COL_INCIDENT_ID  — exact column name for the Incident ID
  COL_CLOSURE_DATE — exact column name for the Closure Date

Everything else — date range (prompted at runtime), number of modules,
chart layout, colours — is handled automatically.

Date formats like "2 Mar 2026", "2nd March 2026", "03/02/2026", "2026-03-02"
are all parsed correctly.
=============================================================================
"""

import os
import re
import sys
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter

import warnings
warnings.filterwarnings("ignore")

# =============================================================================
# >>>  CONFIGURATION — only edit these three lines  <<<
# =============================================================================

INPUT_FILE_PATH  = "incidents.xlsx"
COL_INCIDENT_ID  = "Incident Id"
COL_CLOSURE_DATE = "Incident Closure on"

# =============================================================================
# INTERNALS — do not edit below this line
# =============================================================================

OUTPUT_FILE_PATH   = "incidents_report.xlsx"
SUMMARY_SHEET_NAME = "Summary Dashboard"

# ---------------------------------------------------------------------------
# STYLES
# ---------------------------------------------------------------------------

def _border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def _thick_border():
    m = Side(style="medium")
    return Border(left=m, right=m, top=m, bottom=m)

NAVY       = "1F3864"
MID_BLUE   = "2E75B6"
LIGHT_BLUE = "D6E4F7"
ALT_BLUE   = "EFF5FB"
WHITE      = "FFFFFF"

HEADER_FILL = PatternFill("solid", start_color=NAVY)
SUBHDR_FILL = PatternFill("solid", start_color=MID_BLUE)
TOTAL_FILL  = PatternFill("solid", start_color=LIGHT_BLUE)
ALT_FILL    = PatternFill("solid", start_color=ALT_BLUE)
WHITE_FILL  = PatternFill("solid", start_color=WHITE)

H_FONT   = Font(name="Arial", bold=True, color=WHITE,    size=11)
D_FONT   = Font(name="Arial",            color="000000", size=10)
T_FONT   = Font(name="Arial", bold=True, color=NAVY,     size=11)
BIG_FONT = Font(name="Arial", bold=True, color=WHITE,    size=15)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

# ---------------------------------------------------------------------------
# DATE PARSING
# ---------------------------------------------------------------------------

_ORDINAL_RE = re.compile(r"(\d+)(st|nd|rd|th)\b", re.IGNORECASE)

def _strip_ordinals(text):
    return _ORDINAL_RE.sub(r"\1", str(text))

_DATE_FORMATS = [
    "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d-%B-%Y",
    "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d.%m.%Y",
    "%d %b %y", "%d %B %y",
]

def _parse_date_series(series):
    """
    Robustly parses a Series of mixed date strings.
    Strips ordinal suffixes first, then tries multiple formats.
    """
    cleaned = series.astype(str).apply(_strip_ordinals)
    parsed  = pd.to_datetime(cleaned, infer_datetime_format=True,
                             dayfirst=True, errors="coerce")
    for fmt in _DATE_FORMATS:
        still_bad = parsed.isna()
        if not still_bad.any():
            break
        parsed[still_bad] = pd.to_datetime(
            cleaned[still_bad], format=fmt, errors="coerce"
        )
    return parsed


def _prompt_date(label):
    """Interactively prompt for a date in any recognisable format."""
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
# STEP 3 — AGGREGATE  (only modules with ≥ 1 incident enter the charts)
# ---------------------------------------------------------------------------

def aggregate(counts):
    rows = [
        {"Module Name": k, "Unique Closed Incidents": v}
        for k, v in counts.items()
        if v > 0
    ]
    if not rows:
        sys.exit("\nNo incidents found in the specified date range across any sheet.\n")
    return pd.DataFrame(rows)

# ---------------------------------------------------------------------------
# STEP 4a — SUMMARY SHEET
# ---------------------------------------------------------------------------

def _write_summary_table(ws, summary_df, start_row):
    """Writes the styled table; returns (header_row, total_row)."""
    ws.row_dimensions[start_row].height = 24
    c = ws.cell(row=start_row, column=1,
                value="Module-wise Unique Closed Incidents")
    c.font      = Font(name="Arial", bold=True, size=13, color=NAVY)
    c.alignment = LEFT
    ws.merge_cells(start_row=start_row, start_column=1,
                   end_row=start_row,   end_column=3)

    header_row = start_row + 1
    ws.row_dimensions[header_row].height = 22
    for col, text in enumerate(["#", "Module Name", "Unique Closed Incidents"], 1):
        cell = ws.cell(row=header_row, column=col, value=text)
        cell.font      = H_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = CENTER
        cell.border    = _border()

    n = len(summary_df)
    for i, row in enumerate(summary_df.itertuples(index=False), 1):
        r    = header_row + i
        fill = ALT_FILL if i % 2 == 0 else WHITE_FILL
        ws.row_dimensions[r].height = 18
        d    = row._asdict()

        for col_idx, val in enumerate(
            [i, d["Module Name"], d["Unique Closed Incidents"]], 1
        ):
            cell = ws.cell(row=r, column=col_idx, value=val)
            cell.font      = D_FONT
            cell.fill      = fill
            cell.border    = _border()
            cell.alignment = LEFT if col_idx == 2 else CENTER

    total_row = header_row + n + 1
    ws.row_dimensions[total_row].height = 22

    for col in [1, 2, 3]:
        cell = ws.cell(row=total_row, column=col)
        cell.fill      = TOTAL_FILL
        cell.border    = _thick_border()
        cell.font      = T_FONT
        cell.alignment = CENTER

    ws.cell(row=total_row, column=1, value="TOTAL")
    ws.merge_cells(start_row=total_row, start_column=1,
                   end_row=total_row,   end_column=2)
    ws.cell(row=total_row, column=3,
            value=f"=SUM(C{header_row+1}:C{total_row-1})")

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = max(
        summary_df["Module Name"].astype(str).str.len().max() + 6, 24
    )
    ws.column_dimensions["C"].width = 28

    return header_row, total_row


def _build_bar_chart(ws, n_modules, header_row):
    """
    Bar chart — uses openpyxl's built-in style theming for colours.
    No per-DataPoint solidFill calls (those break across openpyxl versions).
    Style 10 gives distinct colours for up to ~12 series; cycles gracefully beyond.
    """
    bar = BarChart()
    bar.type         = "col"
    bar.grouping     = "clustered"
    bar.title        = "Unique Closed Incidents by Module"
    bar.y_axis.title = "Count"
    bar.x_axis.title = "Module"
    bar.style        = 10                                  # rich colour theme
    bar.width        = min(max(n_modules * 3.2, 20), 36)  # scales with module count
    bar.height       = 14

    data = Reference(ws, min_col=3, min_row=header_row,
                     max_col=3, max_row=header_row + n_modules)
    cats = Reference(ws, min_col=2, min_row=header_row + 1,
                     max_row=header_row + n_modules)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)

    # Give every bar its own colour by splitting into one series per data point.
    # This is the only version-safe way to get distinct bar colours in openpyxl.
    bar.series.clear()
    for idx in range(n_modules):
        data_single = Reference(ws, min_col=3, min_row=header_row + 1 + idx,
                                max_col=3, max_row=header_row + 1 + idx)
        cat_single  = Reference(ws, min_col=2, min_row=header_row + 1 + idx,
                                max_row=header_row + 1 + idx)
        bar.add_data(data_single, titles_from_data=False)
        bar.set_categories(cats)   # shared axis labels
        bar.series[idx].title.v = summary_df_global.iloc[idx]["Module Name"]

    return bar


def _build_pie_chart(ws, n_modules, header_row):
    """
    Pie chart — openpyxl natively colours each slice when there is one
    series with multiple categories, so no DataPoint manipulation needed.
    """
    pie = PieChart()
    pie.title  = "Incident Share by Module"
    pie.style  = 10
    pie.width  = 20
    pie.height = 14

    data = Reference(ws, min_col=3, min_row=header_row,
                     max_col=3, max_row=header_row + n_modules)
    cats = Reference(ws, min_col=2, min_row=header_row + 1,
                     max_row=header_row + n_modules)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(cats)

    return pie


# module-count-aware helper used inside _build_bar_chart
summary_df_global = None


def build_summary_sheet(wb, summary_df, start_dt, end_dt):
    global summary_df_global
    summary_df_global = summary_df

    ws = wb.create_sheet(SUMMARY_SHEET_NAME, 0)

    date_range_str = (
        f"{start_dt.strftime('%d %b %Y')}  ->  {end_dt.strftime('%d %b %Y')}"
    )

    # Banner
    ws.row_dimensions[1].height = 42
    banner = ws.cell(row=1, column=1,
                     value=f"Incident Summary Dashboard  |  {date_range_str}")
    banner.font      = BIG_FONT
    banner.fill      = HEADER_FILL
    banner.alignment = CENTER
    ws.merge_cells("A1:P1")

    header_row, total_row = _write_summary_table(ws, summary_df, start_row=3)

    n         = len(summary_df)
    chart_row = total_row + 2

    bar = _build_bar_chart(ws, n, header_row)
    ws.add_chart(bar, f"A{chart_row}")

    # Pie starts just after the bar (bar.width / ~7.5 cols per unit)
    pie_col_idx = round(bar.width / 7.5) + 2
    pie_anchor  = f"{get_column_letter(pie_col_idx)}{chart_row}"
    pie = _build_pie_chart(ws, n, header_row)
    ws.add_chart(pie, pie_anchor)

    ws.freeze_panes = "A4"

# ---------------------------------------------------------------------------
# STEP 4b — DATA SHEETS
# ---------------------------------------------------------------------------

def _write_data_sheet(ws, df):
    """
    Writes ALL rows with styled headers + AutoFilter dropdown arrows.
    AutoFilter ref covers the whole table so the user can filter any column.
    (Custom date-column pre-filtering is skipped — it uses internal openpyxl
    filter classes that break silently on some versions.)
    """
    if df.empty:
        return

    headers = list(df.columns)
    n_cols  = len(headers)

    ws.row_dimensions[1].height = 20
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font      = Font(name="Arial", bold=True, color=WHITE, size=10)
        cell.fill      = SUBHDR_FILL
        cell.alignment = CENTER
        cell.border    = _border()
        ws.column_dimensions[get_column_letter(col_idx)].width = max(
            len(str(h)) + 4, 14
        )

    for row_idx, row_data in enumerate(df.itertuples(index=False), 2):
        fill = ALT_FILL if row_idx % 2 == 0 else WHITE_FILL
        ws.row_dimensions[row_idx].height = 16
        for col_idx, val in enumerate(row_data, 1):
            try:
                if pd.isna(val):
                    val = None
            except (TypeError, ValueError):
                pass
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font      = D_FONT
            cell.fill      = fill
            cell.alignment = LEFT
            cell.border    = _border()

    ws.freeze_panes = "A2"
    last_col = get_column_letter(n_cols)
    ws.auto_filter.ref = f"A1:{last_col}{len(df) + 1}"

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    print("\n" + "=" * 60)
    print("  INCIDENT REPORT GENERATOR")
    print("=" * 60)
    print("\nPlease enter the date range for filtering closed incidents.")
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

    # Step 2 — process
    processed_raw = {}
    counts        = {}

    for name in sheet_names:
        raw_df, filtered_df = process_sheet(raw_sheets[name], start_dt, end_dt)
        processed_raw[name] = raw_df
        counts[name]        = len(filtered_df)
        status = (
            f"{counts[name]} unique incident(s) in range"
            if counts[name] else "no incidents in range"
        )
        print(f"  [{name}]  total rows = {len(raw_df)}  |  {status}")

    # Step 3 — aggregate
    summary_df  = aggregate(counts)
    grand_total = summary_df["Unique Closed Incidents"].sum()
    active      = len(summary_df)
    skipped     = len(sheet_names) - active

    print(f"\n  Modules with incidents  : {active}")
    if skipped:
        print(f"  Modules with zero hits  : {skipped} (data sheets written, excluded from charts)")
    print(f"  Grand total             : {grand_total}")

    # Step 4 — build fresh output workbook
    wb = Workbook()
    wb.remove(wb.active)   # remove the default blank sheet

    build_summary_sheet(wb, summary_df, start_dt, end_dt)

    for name in sheet_names:
        ws = wb.create_sheet(title=name)
        _write_data_sheet(ws, processed_raw[name])

    wb.save(OUTPUT_FILE_PATH)
    print(f"\n  Report saved -> {OUTPUT_FILE_PATH}")
    print("Done.\n")


if __name__ == "__main__":
    main()
