"""
=============================================================================
INCIDENT REPORT GENERATOR
=============================================================================
CONFIGURATION — only 4 things to set:
  INPUT_FILE_PATH  : path to the source Excel workbook
  COL_INCIDENT_ID  : exact column name for the Incident ID
  COL_CLOSURE_DATE : exact column name for the Closure Date
  COL_STATUS       : exact column name for the Status column

Everything else is automatic — date range is prompted at runtime.

Summary sheet contains:
  1. Module-wise unique closed incidents table (date range filtered)
  2. Pie chart — "Closed Incidents By Module" (date range filtered)
  3. Pie chart — "Overall Incident Status" (all data, all modules,
     unique by Incident ID; Closed variants grouped as one slice)
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
# >>>  CONFIGURATION — only edit these four lines  <<<
# =============================================================================

INPUT_FILE_PATH  = "incidents.xlsx"
COL_INCIDENT_ID  = "Incident Id"
COL_CLOSURE_DATE = "Incident Closure on"
COL_STATUS       = "Status"

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

    # Diagnostic — expose any rows where date failed to parse
    nat_mask = df[COL_CLOSURE_DATE].isna()
    if nat_mask.any():
        # Only flag rows that had a non-blank original value (blank dates are expected)
        pass  # printed per-sheet in main for visibility

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
# STEP 3b — OVERALL STATUS BREAKDOWN  (all data, unique incident IDs)
# ---------------------------------------------------------------------------

def _normalize_status(val):
    """
    Maps raw status to exactly one of three buckets, or None to discard.
      - anything starting with 'closed'                    -> 'Closed'
      - anything starting with 'open'                      -> 'Open'
      - anything starting with 'in progress'/'inprogress'  -> 'In Progress'
      - everything else (Reopen, Unknown, blank, etc.)     -> None (excluded)
    """
    if pd.isna(val):
        return None
    s = str(val).strip().lower()
    if s.startswith("closed"):
        return "Closed"
    if s.startswith("open"):
        return "Open"
    if s.startswith("in progress") or s.startswith("inprogress"):
        return "In Progress"
    return None   # discard Reopen, Unknown, or anything else

def compute_status_breakdown(processed_raw, sheet_names):
    """
    Combine every sheet, deduplicate globally on COL_INCIDENT_ID,
    normalise status, discard anything outside the three buckets,
    and return plain lists in fixed display order: Open, In Progress, Closed.
    """
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

    # Drop everything that didn't map to one of the three buckets
    combined = combined[combined["_status_norm"].notna()]

    counts = combined["_status_norm"].value_counts()

    # Fixed order — only include buckets that actually have data
    order  = ["Open", "In Progress", "Closed"]
    labels = [s for s in order if s in counts.index]
    status_counts = [int(counts[s]) for s in labels]

    print(f"\n  Overall status breakdown (unique incidents, 3 buckets only):")
    for lbl, cnt in zip(labels, status_counts):
        print(f"    {lbl}: {cnt}")

    return labels, status_counts

# ---------------------------------------------------------------------------
# STEP 4 — BUILD OUTPUT WORKBOOK
# ---------------------------------------------------------------------------

def build_workbook(module_names, module_counts, status_labels, status_counts,
                   filtered_raw, counts, sheet_names, start_dt, end_dt):

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

    # ── PIE 1 — Closed Incidents By Module (date range filtered) ─────────────
    pie1 = wb.add_chart({"type": "pie"})
    pie1.add_series({
        "name":       "Closed Incidents By Module",
        "categories": [SUMMARY_SHEET_NAME, DATA_START, 1, DATA_START + n - 1, 1],
        "values":     [SUMMARY_SHEET_NAME, DATA_START, 2, DATA_START + n - 1, 2],
        "data_labels": {
            "category":   True,
            "value":      True,
            "percentage": True,
            "separator":  "\n",
            "font":       {"bold": True, "size": 10, "color": "#FFFFFF"},
            "border":     {"none": True},
            "fill":       {"none": True},
        },
    })
    pie1.set_title({"name": "Closed Incidents By Module"})
    pie1.set_legend({"position": "bottom"})
    pie1.set_style(10)
    pie1.set_size({"width": 480, "height": 360})
    sw.insert_chart(TOTAL_ROW + 2, 0, pie1)

    # ── PIE 2 — Overall Incident Status (all data, unique IDs) ───────────────
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
            "data_labels": {
                "category":   True,
                "value":      True,
                "percentage": True,
                "separator":  "\n",
                "font":       {"bold": True, "size": 10, "color": "#FFFFFF"},
                "border":     {"none": True},
                "fill":       {"none": True},
            },
        })
        pie2.set_title({"name": "Overall Incident Status (All Modules, Unique IDs)"})
        pie2.set_legend({"position": "bottom"})
        pie2.set_style(10)
        pie2.set_size({"width": 480, "height": 360})
        sw.insert_chart(TOTAL_ROW + 2, 7, pie2)

    # ── DATA SHEETS — only modules with incidents, only counted rows ─────────
    # filtered_raw contains only the deduped rows that were counted.
    # Sheets with zero incidents in range are excluded entirely.
    for name in sheet_names:
        if counts[name] == 0:
            continue                    # skip modules with no incidents in range

        df      = filtered_raw[name]   # only the counted rows
        ws_name = name[:31]
        dw      = wb.add_worksheet(ws_name)

        if df.empty:
            continue

        headers      = list(df.columns)
        n_cols       = len(headers)
        n_rows       = len(df)
        date_col_idx = headers.index(COL_CLOSURE_DATE) if COL_CLOSURE_DATE in headers else -1

        # Default row height set once — no per-row XML tag written
        dw.set_default_row(16)

        # Header row
        dw.set_row(0, 20)
        for ci, h in enumerate(headers):
            dw.write(0, ci, h, f_data_hdr)
            dw.set_column(ci, ci, max(len(str(h)) + 4, 14))

        # Bulk value extraction — one numpy/Python pass over the whole frame
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

    print(f"\n  Range: {start_dt.strftime('%d %b %Y')} -> {end_dt.strftime('%d %b %Y')}\n")

    # Step 1 — load
    raw_sheets  = load_all_sheets(INPUT_FILE_PATH)
    sheet_names = list(raw_sheets.keys())
    print(f"Sheets found ({len(sheet_names)}): {sheet_names}\n")

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

        # ── Diagnostic: show rows where closure date failed to parse ──────────
        if COL_CLOSURE_DATE in raw_df.columns:
            # Re-read original strings before parsing to show what failed
            orig_strings = raw_sheets[name][COL_CLOSURE_DATE].astype(str).str.strip()
            nat_rows     = raw_df[raw_df[COL_CLOSURE_DATE].isna()]
            # Only flag rows that had a non-empty value originally
            truly_bad    = nat_rows[
                orig_strings.loc[nat_rows.index].str.lower().isin(
                    ["", "nan", "none", "nat", "n/a", "-"]
                ) == False
            ]
            if not truly_bad.empty:
                samples = orig_strings.loc[truly_bad.index].unique()[:5]
                print(f"    ⚠  {len(truly_bad)} row(s) had unparseable closure dates "
                      f"— excluded from count!")
                print(f"    ⚠  Sample date strings that failed: "
                      f"{list(samples)}")

    # Step 3 — aggregate into plain lists (no DataFrame column-name ambiguity)
    module_names, module_counts = aggregate(counts)
    grand_total = sum(module_counts)
    skipped     = len(sheet_names) - len(module_names)

    print(f"\n  Modules with incidents : {len(module_names)}")
    if skipped:
        print(f"  Modules excluded (zero): {skipped}  (not written to output)")
    print(f"  Grand total            : {grand_total}")

    # Step 3b — overall status breakdown across all modules, all data
    status_labels, status_counts = compute_status_breakdown(processed_raw, sheet_names)

    # Step 4 — write output
    build_workbook(module_names, module_counts, status_labels, status_counts,
                   filtered_raw, counts, sheet_names, start_dt, end_dt)

    print(f"\n  Report saved -> {OUTPUT_FILE_PATH}")
    print("Done.\n")


if __name__ == "__main__":
    main()
