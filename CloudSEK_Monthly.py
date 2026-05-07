"""
=============================================================================
INCIDENT REPORT GENERATOR
=============================================================================
CONFIGURATION — only 6 things to set:
  INPUT_FILE_PATH   : path to the source Excel workbook
  COL_INCIDENT_ID   : exact column name for the Incident ID
  COL_CREATION_DATE : exact column name for the Incident Creation Date
  COL_CLOSURE_DATE  : exact column name for the Closure Date
  COL_STATUS        : exact column name for the Status column
  COL_CLOSED_BY     : exact column name for who closed the incident
  COL_COMMENT       : exact column name for the comment/notes column

SUMMARY DASHBOARD LAYOUT (monthly stakeholder deck)
  ┌──────────────────────────────────────────────────────────────────────┐
  │  BANNER — report title + date range                                  │
  ├────────────┬────────────┬──────────────────┬────────────────────────┤
  │  TOTAL     │  TOTAL     │  RESOLUTION      │  STILL PENDING         │
  │  RAISED    │  CLOSED    │  RATE  (RAG)     │  (RAG — lower=better)  │
  ├────────────┴────────────┴──────────────────┴────────────────────────┤
  │ ① Module table — Closed in range          (navy)                    │
  │ ② Module table — Created in range         (green)                   │
  │ ③ Union table + status mini-table         (purple)                  │
  ├────────────────────┬───────────────────┬───────────────────────────┤
  │ Horizontal bar     │ Donut (RAG)       │ Clustered bar             │
  │ Closed per module  │ Open/InProg/Close │ Created vs Closed         │
  ├────────────────────┴───────────────────┴───────────────────────────┤
  │ User table + horizontal bar chart                                   │
  └─────────────────────────────────────────────────────────────────────┘

CHART DESIGN DECISIONS
  • Horizontal bar  — lengths are trivially comparable; module names fit
                      naturally as row labels; sorted so worst module is
                      immediately visible at the top
  • Donut (RAG)     — three-bucket health signal with semantic red/amber/
                      green; total shown in title; far less ambiguous than
                      a pie with arbitrary colours
  • Clustered bar   — the only chart that shows whether a module is
                      ACCUMULATING a backlog (created > closed) or
                      CLEARING it — impossible to see in any pie chart
  • Hidden _ChartData sheet — all chart cell-refs point here so they
                      never break when table row numbers shift

All filtering/aggregation logic is unchanged from the previous version.
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

INPUT_FILE_PATH   = "incidents.xlsx"
OUTPUT_FOLDER     = "reports"
COL_INCIDENT_ID   = "Incident Id"
COL_CREATION_DATE = "Created On"
COL_CLOSURE_DATE  = "Incident Closure on"
COL_STATUS        = "Status"
COL_CLOSED_BY     = "Incident Closed By"
COL_COMMENT       = "Comment"

# =============================================================================
# INTERNALS
# =============================================================================

SUMMARY_SHEET_NAME = "Summary Dashboard"

_STATUS_RAG = {
    "Open":        "#C00000",
    "In Progress": "#FFC000",
    "Closed":      "#70AD47",
}

_PALETTE = [
    "#4472C4", "#ED7D31", "#70AD47", "#FFC000", "#5B9BD5",
    "#A9D18E", "#FF7C80", "#9E480E", "#7030A0", "#636363",
    "#255E91", "#43682B", "#C00000", "#997300", "#7B5EA7",
]

# ---------------------------------------------------------------------------
# DATE PARSING  (unchanged)
# ---------------------------------------------------------------------------

_ORDINAL_RE = re.compile(r"(\d+)(st|nd|rd|th)\b", re.IGNORECASE)

_DATE_FORMATS = [
    "%d %b, %Y %I:%M:%S %p", "%d %B, %Y %I:%M:%S %p",
    "%d %b, %Y %H:%M:%S",    "%d %B, %Y %H:%M:%S",
    "%d %b, %Y %I:%M %p",    "%d %B, %Y %I:%M %p",
    "%d %b, %Y",              "%d %B, %Y",
    "%d %b %Y %I:%M:%S %p",  "%d %B %Y %I:%M:%S %p",
    "%d %b %Y %H:%M:%S",     "%d %B %Y %H:%M:%S",
    "%d %b %Y %I:%M %p",     "%d %B %Y %I:%M %p",
    "%d %b %Y", "%d %B %Y",
    "%d-%b-%Y", "%d-%B-%Y",
    "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d.%m.%Y",
    "%d %b %y", "%d %B %y",
]

def _strip_ordinals(text):
    return _ORDINAL_RE.sub(r"\1", str(text))

def _clean(text):
    t = _strip_ordinals(str(text))
    t = re.sub(r'\b(am|pm)\b', lambda m: m.group(0).upper(), t, flags=re.IGNORECASE)
    t = re.sub(r'([A-Za-z]),', r'\1', t)
    return t.strip()

def _parse_date_series(series):
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
# STEP 1 — LOAD  (unchanged)
# ---------------------------------------------------------------------------

def load_all_sheets(path):
    if not os.path.exists(path):
        sys.exit(f"\nERROR: Input file not found -> {path}\n")
    print(f"\nLoading: {path}")
    str_sheets    = pd.read_excel(path, sheet_name=None, dtype=str)
    native_sheets = pd.read_excel(path, sheet_name=None)
    merged = {}
    for name, df_str in str_sheets.items():
        df     = df_str.copy()
        native = native_sheets.get(name, pd.DataFrame())
        for col in [COL_CLOSURE_DATE, COL_CREATION_DATE]:
            if col and col in native.columns:
                df[col] = native[col].values
        merged[name] = df
    return merged

# ---------------------------------------------------------------------------
# STEP 2 — PROCESS  (unchanged)
# ---------------------------------------------------------------------------

def process_sheet(df, start_dt, end_dt):
    df = df.copy()
    if COL_CLOSURE_DATE in df.columns:
        df[COL_CLOSURE_DATE] = _parse_date_series(df[COL_CLOSURE_DATE])
        mask_closure = (df[COL_CLOSURE_DATE] >= start_dt) & (df[COL_CLOSURE_DATE] <= end_dt)
    else:
        mask_closure = pd.Series(False, index=df.index)

    if COL_CREATION_DATE and COL_CREATION_DATE in df.columns:
        df[COL_CREATION_DATE] = _parse_date_series(df[COL_CREATION_DATE])
        mask_creation = (df[COL_CREATION_DATE] >= start_dt) & (df[COL_CREATION_DATE] <= end_dt)
    else:
        mask_creation = pd.Series(False, index=df.index)

    mask_union = mask_closure | mask_creation

    def _dedup(sub_df):
        if COL_INCIDENT_ID in sub_df.columns:
            return sub_df.drop_duplicates(subset=[COL_INCIDENT_ID])
        return sub_df

    return (
        df,
        _dedup(df[mask_closure].copy()),
        _dedup(df[mask_creation].copy()),
        _dedup(df[mask_union].copy()),
    )

# ---------------------------------------------------------------------------
# STEP 3 — AGGREGATE  (unchanged)
# ---------------------------------------------------------------------------

def aggregate(counts):
    names  = [n for n, c in counts.items() if c > 0]
    cnts   = [c for c in counts.values()   if c > 0]
    if not names:
        sys.exit("\nNo incidents found in the specified date range.\n")
    return names, cnts

# ---------------------------------------------------------------------------
# STEP 3b — OVERALL STATUS BREAKDOWN  (unchanged)
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
        cols = [c for c in [COL_INCIDENT_ID, COL_STATUS] if c in df.columns]
        if cols:
            frames.append(df[cols].copy())
    if not frames:
        return [], []
    combined = pd.concat(frames, ignore_index=True)
    if COL_INCIDENT_ID in combined.columns:
        combined = combined.drop_duplicates(subset=[COL_INCIDENT_ID])
    if COL_STATUS not in combined.columns:
        print(f"  WARNING: Column '{COL_STATUS}' not found — status chart skipped.")
        return [], []
    combined["_s"] = combined[COL_STATUS].apply(_normalize_status)
    combined = combined[combined["_s"].notna()]
    vc       = combined["_s"].value_counts()
    labels   = [s for s in ["Open", "In Progress", "Closed"] if s in vc.index]
    counts   = [int(vc[s]) for s in labels]
    print(f"\n  Overall status breakdown:")
    for l, c in zip(labels, counts):
        print(f"    {l}: {c}")
    return labels, counts

# ---------------------------------------------------------------------------
# STEP 3c — USER-WISE BREAKDOWN  (unchanged)
# ---------------------------------------------------------------------------

def compute_user_breakdown(filtered_closure_raw, sheet_names):
    frames = []
    for name in sheet_names:
        df = filtered_closure_raw.get(name, pd.DataFrame())
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
        ~combined[COL_CLOSED_BY].astype(str).str.strip().str.lower()
         .isin(["", "nan", "none", "n/a", "-"])
    ]
    vc = (combined[COL_CLOSED_BY].astype(str).str.strip()
          .value_counts().sort_values(ascending=False))
    names  = list(vc.index)
    counts = [int(c) for c in vc.values]
    print(f"\n  Closed-by breakdown:")
    for u, c in zip(names, counts):
        print(f"    {u}: {c}")
    return names, counts

# ---------------------------------------------------------------------------
# STEP 3d — UNION STATUS BREAKDOWN  (unchanged)
# ---------------------------------------------------------------------------

def compute_union_status_breakdown(filtered_union_raw, sheet_names):
    frames = []
    for name in sheet_names:
        df = filtered_union_raw.get(name, pd.DataFrame())
        if df.empty:
            continue
        cols = [c for c in [COL_INCIDENT_ID, COL_STATUS] if c in df.columns]
        if len(cols) == 2:
            frames.append(df[cols].copy())
    if not frames:
        return [], []
    combined = pd.concat(frames, ignore_index=True)
    if COL_INCIDENT_ID in combined.columns:
        combined = combined.drop_duplicates(subset=[COL_INCIDENT_ID])
    if COL_STATUS not in combined.columns:
        return [], []
    combined["_s"] = combined[COL_STATUS].apply(_normalize_status)
    combined = combined[combined["_s"].notna()]
    vc       = combined["_s"].value_counts()
    labels   = [s for s in ["Open", "In Progress", "Closed"] if s in vc.index]
    counts   = [int(vc[s]) for s in labels]
    total    = sum(counts)
    print(f"\n  Union status breakdown ({total} unique incidents):")
    for l, c in zip(labels, counts):
        print(f"    {l}: {c}")
    return labels, counts

# ---------------------------------------------------------------------------
# STEP 3e — EMAIL EXTRACTION  (unchanged)
# ---------------------------------------------------------------------------

_EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}",
                        re.IGNORECASE)

def extract_emails_closed_resolved(processed_raw, sheet_names):
    if not COL_COMMENT:
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
        return []
    combined = pd.concat(frames, ignore_index=True)
    if COL_INCIDENT_ID in combined.columns:
        combined = combined.drop_duplicates(subset=[COL_INCIDENT_ID])
    mask  = combined[COL_STATUS].astype(str).str.strip().str.lower() == "closed - resolved"
    cr_df = combined[mask]
    print(f"\n  Rows with status 'closed - resolved': {len(cr_df)}")
    if cr_df.empty:
        return []
    emails = set()
    for val in cr_df[COL_COMMENT].dropna():
        for m in _EMAIL_RE.findall(str(val)):
            emails.add(m.lower().strip())
    unique_emails = sorted(emails)
    print(f"  Unique emails found: {len(unique_emails)}")
    return unique_emails

# ---------------------------------------------------------------------------
# STEP 4 — BUILD OUTPUT WORKBOOK
# ---------------------------------------------------------------------------

def build_workbook(
    module_names, module_counts,
    created_module_names, created_module_counts,
    union_module_names,   union_module_counts,
    union_status_labels,  union_status_counts,
    status_labels,        status_counts,
    user_names,           user_counts,
    unique_emails,
    filtered_closure_raw, filtered_creation_raw, filtered_union_raw,
    counts_closure, counts_union,
    sheet_names,
    start_dt, end_dt,
    output_path,
):
    wb = xlsxwriter.Workbook(output_path)

    # ── COLOUR TOKENS ────────────────────────────────────────────────────────
    NAVY    = "#1F3864"
    MIDBLUE = "#2E75B6"
    LTBLUE  = "#D6E4F7"
    D_GREEN = "#375623"
    LTGREEN = "#E2EFDA"
    D_PURP  = "#4B2878"
    LTPURP  = "#EAE0F1"
    ALT     = "#EFF5FB"
    WHITE   = "#FFFFFF"

    # ── KPI CALCULATIONS ─────────────────────────────────────────────────────
    total_raised  = sum(created_module_counts) if created_module_counts else 0
    total_closed  = sum(module_counts)
    total_union   = sum(union_module_counts)   if union_module_counts   else 0
    _denom        = total_raised if total_raised > 0 else (total_union or 1)
    res_rate      = total_closed / _denom * 100

    still_open    = next((c for l, c in zip(union_status_labels, union_status_counts)
                          if l == "Open"), 0)
    in_progress   = next((c for l, c in zip(union_status_labels, union_status_counts)
                          if l == "In Progress"), 0)
    still_pending = still_open + in_progress
    pend_pct      = still_pending / _denom * 100 if _denom else 0

    def _rag(value, g_thresh, a_thresh, higher_is_better=True):
        if higher_is_better:
            if value >= g_thresh: return D_GREEN, LTGREEN
            if value >= a_thresh: return "#7F6000", "#FFF2CC"
            return "#9C0006", "#FFC7CE"
        else:
            if value <= g_thresh: return D_GREEN, LTGREEN
            if value <= a_thresh: return "#7F6000", "#FFF2CC"
            return "#9C0006", "#FFC7CE"

    rate_hdr, rate_bg = _rag(res_rate,   80, 60, higher_is_better=True)
    pend_hdr, pend_bg = _rag(pend_pct,   10, 25, higher_is_better=False)

    # ── FORMAT FACTORY ───────────────────────────────────────────────────────
    def _f(**kw):
        base = {"font_name": "Arial", "font_size": 10, "valign": "vcenter"}
        base.update(kw)
        return wb.add_format(base)

    f_banner    = _f(bold=True, font_size=15, font_color=WHITE,
                     bg_color=NAVY, align="center")
    f_section   = _f(bold=True, font_size=13, font_color=NAVY)
    f_num       = _f(align="center", border=1)
    f_lft       = _f(align="left",   border=1)
    f_num_alt   = _f(align="center", border=1, bg_color=ALT)
    f_lft_alt   = _f(align="left",   border=1, bg_color=ALT)
    f_data_hdr  = _f(bold=True, font_color=WHITE, bg_color=MIDBLUE,
                     align="center", border=1)
    f_cell      = _f(align="left",  border=1)
    f_cell_alt  = _f(align="left",  border=1, bg_color=ALT)
    f_date      = _f(align="left",  border=1, num_format="dd mmm yyyy")
    f_date_alt  = _f(align="left",  border=1, bg_color=ALT,
                     num_format="dd mmm yyyy")

    def _hdr(bg):
        return _f(bold=True, font_size=11, font_color=WHITE,
                  bg_color=bg, align="center", border=1)
    def _tot(fg, bg):
        return _f(bold=True, font_size=11, font_color=fg,
                  bg_color=bg, align="center", border=2)

    f_hdr_navy  = _hdr(NAVY);    f_tot_navy  = _tot(NAVY,    LTBLUE)
    f_hdr_green = _hdr(D_GREEN); f_tot_green = _tot(D_GREEN, LTGREEN)
    f_hdr_purp  = _hdr(D_PURP);  f_tot_purp  = _tot(D_PURP,  LTPURP)

    # ── HIDDEN CHART-DATA SHEET ───────────────────────────────────────────────
    # Decouples chart cell-refs from visible table row positions.
    # All charts reference this sheet; tables can shift freely.
    CDSHEET = "_ChartData"
    cd      = wb.add_worksheet(CDSHEET)
    cd.hide()

    # Build a unified module list sorted by closed-count desc
    all_mods    = list(dict.fromkeys(
        list(module_names) + list(created_module_names) + list(union_module_names)
    ))
    closed_map  = dict(zip(module_names,         module_counts))
    created_map = dict(zip(created_module_names, created_module_counts))
    mods_sorted = sorted(all_mods, key=lambda m: closed_map.get(m, 0), reverse=True)
    NM          = len(mods_sorted)

    # Cols A-C  (0-2):  module | closed | created
    for i, m in enumerate(mods_sorted):
        cd.write(i, 0, m)
        cd.write(i, 1, closed_map.get(m,  0))
        cd.write(i, 2, created_map.get(m, 0))

    # Cols E-F  (4-5):  overall status label | count
    NS = len(status_labels)
    for i, (lbl, cnt) in enumerate(zip(status_labels, status_counts)):
        cd.write(i, 4, lbl)
        cd.write(i, 5, cnt)

    # Cols H-I  (7-8):  union status label | count
    for i, (lbl, cnt) in enumerate(zip(union_status_labels, union_status_counts)):
        cd.write(i, 7, lbl)
        cd.write(i, 8, cnt)

    # Cols K-L  (10-11): user name | count
    NU = len(user_names)
    for i, (u, c) in enumerate(zip(user_names, user_counts)):
        cd.write(i, 10, u)
        cd.write(i, 11, c)

    # ── SUMMARY SHEET ─────────────────────────────────────────────────────────
    sw = wb.add_worksheet(SUMMARY_SHEET_NAME)
    sw.set_column(0, 0,  6)
    sw.set_column(1, 1, 52)
    sw.set_column(2, 2, 28)

    # Row 0 — banner
    sw.set_row(0, 42)
    sw.merge_range(
        0, 0, 0, 14,
        f"Incident Summary Dashboard  |  "
        f"{start_dt.strftime('%d %b %Y')}  to  {end_dt.strftime('%d %b %Y')}",
        f_banner,
    )

    # ── KPI SCORECARD (rows 2-4) ──────────────────────────────────────────────
    KPI_TOP = 2
    BOX_W   = 3    # columns per box

    kpi = [
        ("TOTAL RAISED",     str(total_raised) if total_raised else "N/A",
         "Incidents created in range",    NAVY,     LTBLUE),
        ("TOTAL CLOSED",     str(total_closed),
         "Closure date in range",         D_GREEN,  LTGREEN),
        ("RESOLUTION RATE",  f"{res_rate:.1f}%" if total_raised else "N/A",
         "Closed / Raised in period",     rate_hdr, rate_bg),
        ("STILL PENDING",    str(still_pending),
         "Open + In Progress (union)",    pend_hdr, pend_bg),
    ]

    sw.set_row(KPI_TOP,     18)   # label
    sw.set_row(KPI_TOP + 1, 40)   # big number
    sw.set_row(KPI_TOP + 2, 16)   # subtitle

    for bi, (label, value, subtitle, hdr_c, bg_c) in enumerate(kpi):
        c1 = bi * BOX_W;  c2 = c1 + BOX_W - 1
        sw.merge_range(KPI_TOP,     c1, KPI_TOP,     c2, label,
                       _f(bold=True, font_size=9, font_color=WHITE,
                          bg_color=hdr_c, align="center", border=1))
        sw.merge_range(KPI_TOP + 1, c1, KPI_TOP + 1, c2, value,
                       _f(bold=True, font_size=22, font_color=hdr_c,
                          bg_color=bg_c, align="center", border=1))
        sw.merge_range(KPI_TOP + 2, c1, KPI_TOP + 2, c2, subtitle,
                       _f(font_size=8, font_color="#595959", italic=True,
                          bg_color=bg_c, align="center", border=1))

    cursor = KPI_TOP + 4   # start of tables

    # ── TABLE WRITER ─────────────────────────────────────────────────────────
    def _write_table(top, names, cnts, hdr_fmt, tot_fmt, label):
        n = len(names)
        sw.set_row(top, 24)
        sw.write(top, 0, label, f_section)
        HDR = top + 1;  DS = HDR + 1;  TR = DS + n
        sw.set_row(HDR, 22)
        for ci, h in enumerate(["#", "Module Name", "Unique Incidents"]):
            sw.write(HDR, ci, h, hdr_fmt)
        for i in range(n):
            r   = DS + i;  alt = (i % 2 == 1)
            sw.set_row(r, 18)
            sw.write(r, 0, i + 1,    f_num_alt if alt else f_num)
            sw.write(r, 1, names[i], f_lft_alt if alt else f_lft)
            sw.write(r, 2, cnts[i],  f_num_alt if alt else f_num)
        sw.set_row(TR, 22)
        sw.merge_range(TR, 0, TR, 1, "TOTAL", tot_fmt)
        sw.write_formula(TR, 2, f"=SUM(C{DS+1}:C{DS+n})", tot_fmt)
        return TR + 2

    # ① closed
    cursor = _write_table(cursor, module_names, module_counts,
                          f_hdr_navy, f_tot_navy,
                          "① Closed in Date Range — Module-wise (Closure Date filter)")
    # ② created
    if created_module_names:
        cursor = _write_table(cursor, created_module_names, created_module_counts,
                              f_hdr_green, f_tot_green,
                              "② Created in Date Range — Module-wise (Creation Date filter)")
    else:
        sw.write(cursor, 0,
                 "② No creation-date column configured or no incidents created in range.",
                 f_section)
        cursor += 2

    # ③ union
    if union_module_names:
        cursor = _write_table(cursor, union_module_names, union_module_counts,
                              f_hdr_purp, f_tot_purp,
                              "③ UNION — Created OR Closed in Range (complete period picture)")

    # ③ union status mini-table
    if union_status_labels:
        sw.set_row(cursor, 20)
        sw.write(cursor, 0,
                 "③ Union — Current Status of All Incidents in the Period", f_section)
        cursor += 1
        for ci, h in enumerate(["Status", "Count", "% Share"]):
            sw.write(cursor, ci, h, f_hdr_purp)
        cursor += 1
        total_u = sum(union_status_counts) or 1
        for lbl, cnt in zip(union_status_labels, union_status_counts):
            bg = _STATUS_RAG.get(lbl, ALT)
            sw.write(cursor, 0, lbl,
                     _f(bold=True, align="left",   border=1, bg_color=bg))
            sw.write(cursor, 1, cnt,
                     _f(align="center", border=1, bg_color=bg))
            sw.write(cursor, 2, cnt / total_u,
                     _f(align="center", border=1, bg_color=bg, num_format="0.0%"))
            cursor += 1
        cursor += 1

    sw.freeze_panes(KPI_TOP + 3, 0)

    # ── CHART SECTION ─────────────────────────────────────────────────────────
    cursor += 1
    sw.set_row(cursor, 22)
    sw.write(cursor, 0, "Visual Summary", f_section)
    cursor += 1
    CHART_ROW = cursor

    # shared chart styling helper
    def _style(chart):
        chart.set_plotarea({"border": {"none": True}})
        chart.set_chartarea({"border": {"color": "#D9D9D9"},
                             "fill":   {"color": WHITE}})
        chart.set_style(2)

    # ── CHART 1: Horizontal bar — Closed per module ───────────────────────────
    # Sorted descending so the most-impacted module is at the top.
    # Single professional-blue fill; data labels inside bars.
    if NM > 0:
        bar1 = wb.add_chart({"type": "bar"})
        bar1.add_series({
            "name":       "Closed Incidents",
            "categories": [CDSHEET, 0, 0, NM - 1, 0],
            "values":     [CDSHEET, 0, 1, NM - 1, 1],
            "fill":       {"color": "#4472C4"},
            "data_labels": {
                "value":    True,
                "position": "inside_end",
                "font":     {"bold": True, "size": 9, "color": WHITE},
            },
        })
        bar1.set_title({
            "name":      f"Closed Incidents by Module  (total: {total_closed})",
            "name_font": {"bold": True, "size": 11, "color": NAVY},
        })
        bar1.set_legend({"none": True})
        bar1.set_x_axis({"num_font":       {"size": 9},
                          "major_gridlines": {"visible": False}})
        bar1.set_y_axis({"num_font":       {"size": 9, "bold": True}})
        _style(bar1)
        bar1.set_size({"width": 480, "height": max(280, NM * 28 + 100)})
        sw.insert_chart(CHART_ROW, 0, bar1, {"x_offset": 5, "y_offset": 5})

    # ── CHART 2: Donut — Overall status with RAG colours ─────────────────────
    # Red=Open / Amber=In Progress / Green=Closed.
    # Total shown in title so stakeholders never have to add slices.
    if NS > 0:
        donut = wb.add_chart({"type": "doughnut"})
        donut.add_series({
            "name":       "Status",
            "categories": [CDSHEET, 0, 4, NS - 1, 4],
            "values":     [CDSHEET, 0, 5, NS - 1, 5],
            "points": [
                {"fill": {"color": _STATUS_RAG.get(l, "#4472C4")}}
                for l in status_labels
            ],
            "data_labels": {
                "percentage": True,
                "category":   True,
                "value":      True,
                "separator":  "\n",
                "font":       {"bold": True, "size": 9},
            },
        })
        donut.set_title({
            "name": (
                f"Overall Incident Status\n"
                f"Total: {sum(status_counts)} unique incidents (all data)"
            ),
            "name_font": {"bold": True, "size": 10, "color": NAVY},
        })
        donut.set_legend({"none": True})
        _style(donut)
        donut.set_size({"width": 380, "height": 380})
        sw.insert_chart(CHART_ROW, 8, donut, {"x_offset": 5, "y_offset": 5})

    # ── CHART 3: Clustered horizontal bar — Created vs Closed per module ─────
    # The only chart that shows whether a module is accumulating a backlog
    # (created bar > closed bar) or clearing carry-over work.
    has_creation_data = any(created_map.get(m, 0) > 0 for m in mods_sorted)
    if NM > 0 and has_creation_data:
        bar2 = wb.add_chart({"type": "bar", "subtype": "clustered"})
        bar2.add_series({
            "name":       "Closed in Range",
            "categories": [CDSHEET, 0, 0, NM - 1, 0],
            "values":     [CDSHEET, 0, 1, NM - 1, 1],
            "fill":       {"color": "#4472C4"},
            "data_labels": {"value": True, "font": {"size": 8}},
        })
        bar2.add_series({
            "name":       "Created in Range",
            "categories": [CDSHEET, 0, 0, NM - 1, 0],
            "values":     [CDSHEET, 0, 2, NM - 1, 2],
            "fill":       {"color": "#70AD47"},
            "data_labels": {"value": True, "font": {"size": 8}},
        })
        bar2.set_title({
            "name":      "Created vs Closed per Module",
            "name_font": {"bold": True, "size": 11, "color": NAVY},
        })
        bar2.set_legend({"position": "bottom"})
        bar2.set_x_axis({"num_font":       {"size": 9},
                          "major_gridlines": {"visible": False}})
        bar2.set_y_axis({"num_font":       {"size": 9, "bold": True}})
        _style(bar2)
        bar2.set_size({"width": 480, "height": max(300, NM * 40 + 120)})
        sw.insert_chart(CHART_ROW, 15, bar2, {"x_offset": 5, "y_offset": 5})

    # Advance cursor past chart area
    chart_rows = max(NM * 2 + 6, 20)
    cursor     = CHART_ROW + chart_rows

    # ── USER TABLE + HORIZONTAL BAR ───────────────────────────────────────────
    if NU > 0:
        cursor += 1
        sw.set_row(cursor, 24)
        sw.write(cursor, 0,
                 "Incidents Closed By — User Wise (Closure Date Range)", f_section)
        UHDR = cursor + 1;  UDS = UHDR + 1;  UTR = UDS + NU
        sw.set_row(UHDR, 22)
        for ci, h in enumerate(["#", "Closed By", "Unique Incidents Closed"]):
            sw.write(UHDR, ci, h, f_hdr_navy)
        for i in range(NU):
            r   = UDS + i;  alt = (i % 2 == 1)
            sw.set_row(r, 18)
            sw.write(r, 0, i + 1,         f_num_alt if alt else f_num)
            sw.write(r, 1, user_names[i], f_lft_alt if alt else f_lft)
            sw.write(r, 2, user_counts[i], f_num_alt if alt else f_num)
        sw.set_row(UTR, 22)
        sw.merge_range(UTR, 0, UTR, 1, "TOTAL", f_tot_navy)
        sw.write_formula(UTR, 2, f"=SUM(C{UDS+1}:C{UDS+NU})", f_tot_navy)

        bar3 = wb.add_chart({"type": "bar"})
        bar3.add_series({
            "name":       "Incidents Closed",
            "categories": [CDSHEET, 0, 10, NU - 1, 10],
            "values":     [CDSHEET, 0, 11, NU - 1, 11],
            "fill":       {"color": "#5B9BD5"},
            "data_labels": {
                "value":    True,
                "position": "inside_end",
                "font":     {"bold": True, "size": 9, "color": WHITE},
            },
        })
        bar3.set_title({
            "name":      f"Incidents Closed by User  (total: {sum(user_counts)})",
            "name_font": {"bold": True, "size": 11, "color": NAVY},
        })
        bar3.set_legend({"none": True})
        bar3.set_x_axis({"num_font":       {"size": 9},
                          "major_gridlines": {"visible": False}})
        bar3.set_y_axis({"num_font":       {"size": 9, "bold": True}})
        _style(bar3)
        bar3.set_size({"width": 480, "height": max(280, NU * 28 + 100)})
        sw.insert_chart(UTR + 2, 0, bar3, {"x_offset": 5, "y_offset": 5})

    # ── EMAIL SHEET ───────────────────────────────────────────────────────────
    if unique_emails:
        ew = wb.add_worksheet("Emails - Closed Resolved")
        ew.set_row(0, 36)
        ew.merge_range(
            0, 0, 0, 2,
            f"Unique Emails — Closed - Resolved  |  {len(unique_emails)} addresses",
            _f(bold=True, font_size=13, font_color=WHITE,
               bg_color=NAVY, align="center"),
        )
        eh = _f(bold=True, font_size=11, font_color=WHITE,
                bg_color=NAVY, align="center", border=1)
        ew.set_row(2, 22)
        ew.write(2, 0, "#",             eh)
        ew.write(2, 1, "Email Address", eh)
        ew.set_column(0, 0, 6)
        ew.set_column(1, 1, max(len(e) for e in unique_emails) + 6)
        ew.freeze_panes(3, 0)
        for i, email in enumerate(unique_emails):
            r   = 3 + i;  alt = (i % 2 == 1)
            bg  = ALT if alt else WHITE
            ew.set_row(r, 16)
            ew.write(r, 0, i + 1, _f(align="center", border=1, bg_color=bg))
            ew.write(r, 1, email,  _f(align="left",   border=1, bg_color=bg))

    # ── DATA SHEETS ───────────────────────────────────────────────────────────
    date_cols = {c for c in [COL_CREATION_DATE, COL_CLOSURE_DATE] if c}

    def _write_data_sheet(ws_name, df):
        if df is None or df.empty:
            return
        dw      = wb.add_worksheet(ws_name[:31])
        headers = list(df.columns)
        nc      = len(headers)
        nr      = len(df)
        di      = {i for i, h in enumerate(headers) if h in date_cols}

        dw.set_default_row(16);  dw.set_row(0, 20)
        for ci, h in enumerate(headers):
            dw.write(0, ci, h, f_data_hdr)
            dw.set_column(ci, ci, max(len(str(h)) + 4, 14))

        vals = df.values
        for ri in range(nr):
            er  = ri + 1;  alt = (er % 2 == 0)
            for ci in range(nc):
                val = vals[ri, ci];  isd = (ci in di)
                try:    nil = pd.isna(val)
                except: nil = False
                if nil:
                    dw.write_blank(er, ci, None,
                                   f_date_alt if (isd and alt) else
                                   f_date     if isd else
                                   f_cell_alt if alt else f_cell)
                elif isinstance(val, pd.Timestamp):
                    dw.write_datetime(er, ci, val.to_pydatetime(),
                                      f_date_alt if alt else f_date)
                else:
                    dw.write(er, ci, val, f_cell_alt if alt else f_cell)

        dw.freeze_panes(1, 0)
        dw.autofilter(0, 0, nr, nc - 1)

    for name in sheet_names:
        short = name[:25]
        if counts_closure.get(name, 0) > 0:
            _write_data_sheet(f"C - {short}", filtered_closure_raw.get(name))
        fc = filtered_creation_raw.get(name)
        if fc is not None and not fc.empty:
            _write_data_sheet(f"R - {short}", fc)
        if counts_union.get(name, 0) > 0:
            _write_data_sheet(f"U - {short}", filtered_union_raw.get(name))

    wb.close()

# ---------------------------------------------------------------------------
# MAIN  (unchanged)
# ---------------------------------------------------------------------------

def main():
    print("\n" + "=" * 60)
    print("  INCIDENT REPORT GENERATOR")
    print("=" * 60)
    print("\nEnter the date range for filtering incidents.")

    start_dt = _prompt_date("START date (inclusive)")
    end_dt   = _prompt_date("END   date (inclusive)")

    if end_dt < start_dt:
        start_dt, end_dt = end_dt, start_dt
        print("  (Dates swapped — start was after end.)")

    end_dt = end_dt.replace(hour=23, minute=59, second=59)
    print(f"\n  Range: {start_dt.strftime('%d %b %Y')} -> {end_dt.strftime('%d %b %Y')}\n")

    raw_sheets  = load_all_sheets(INPUT_FILE_PATH)
    sheet_names = list(raw_sheets.keys())
    print(f"Sheets found ({len(sheet_names)}): {sheet_names}\n")

    print("\n  DIAGNOSTIC — closed incidents without closure dates:")
    for name in sheet_names:
        df = raw_sheets.get(name, pd.DataFrame())
        if df.empty or COL_INCIDENT_ID not in df.columns:
            continue
        if COL_STATUS not in df.columns or COL_CLOSURE_DATE not in df.columns:
            continue
        closed_mask = df[COL_STATUS].astype(str).str.lower().str.startswith("closed")
        blank_date  = df[closed_mask][
            df[closed_mask][COL_CLOSURE_DATE].astype(str).str.strip()
            .isin(["", "nan", "NaT", "None", "N/A", "-"])
        ]
        if not blank_date.empty:
            print(f"    [{name}]  {len(blank_date)} closed incident(s) have NO closure date!")
        else:
            print(f"    [{name}]  all closed incidents have a closure date  ✓")
    print()

    processed_raw         = {}
    filtered_closure_raw  = {}
    filtered_creation_raw = {}
    filtered_union_raw    = {}
    counts_closure        = {}
    counts_creation       = {}
    counts_union          = {}

    for name in sheet_names:
        raw_df, fc, fcr, fu = process_sheet(raw_sheets[name], start_dt, end_dt)
        processed_raw[name]         = raw_df
        filtered_closure_raw[name]  = fc
        filtered_creation_raw[name] = fcr
        filtered_union_raw[name]    = fu
        counts_closure[name]        = len(fc)
        counts_creation[name]       = len(fcr)
        counts_union[name]          = len(fu)

        print(f"  [{name}]  "
              f"closed_in_range={counts_closure[name]}  "
              f"created_in_range={counts_creation[name]}  "
              f"union={counts_union[name]}")

        if COL_CLOSURE_DATE in raw_df.columns:
            orig   = raw_sheets[name][COL_CLOSURE_DATE].astype(str).str.strip()
            bad_df = raw_df[raw_df[COL_CLOSURE_DATE].isna()]
            truly  = bad_df[
                ~orig.loc[bad_df.index].str.lower()
                 .isin(["", "nan", "none", "nat", "n/a", "-"])
            ]
            if not truly.empty:
                print(f"    WARNING: {len(truly)} rows had unparseable closure dates!")
                print(f"    Samples: {list(orig.loc[truly.index].unique()[:5])}")

    module_names, module_counts = aggregate(counts_closure)

    if COL_CREATION_DATE and any(c > 0 for c in counts_creation.values()):
        pairs = [(n, c) for n, c in counts_creation.items() if c > 0]
        created_module_names  = [p[0] for p in pairs]
        created_module_counts = [p[1] for p in pairs]
    else:
        created_module_names, created_module_counts = [], []

    if any(c > 0 for c in counts_union.values()):
        pairs = [(n, c) for n, c in counts_union.items() if c > 0]
        union_module_names  = [p[0] for p in pairs]
        union_module_counts = [p[1] for p in pairs]
    else:
        union_module_names, union_module_counts = [], []

    print(f"\n  Closed in range  : {len(module_names)} modules / "
          f"{sum(module_counts)} incidents")
    print(f"  Created in range : {len(created_module_names)} modules / "
          f"{sum(created_module_counts)} incidents")
    print(f"  Union            : {len(union_module_names)} modules / "
          f"{sum(union_module_counts)} incidents")

    status_labels,       status_counts       = compute_status_breakdown(
        processed_raw, sheet_names)
    user_names,          user_counts         = compute_user_breakdown(
        filtered_closure_raw, sheet_names)
    union_status_labels, union_status_counts = compute_union_status_breakdown(
        filtered_union_raw, sheet_names)
    unique_emails                            = extract_emails_closed_resolved(
        processed_raw, sheet_names)

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    date_range_str = (f"{start_dt.strftime('%d %b %Y')} - "
                      f"{end_dt.strftime('%d %b %Y')}")
    output_path    = os.path.join(OUTPUT_FOLDER,
                                  f"Incident Review - {date_range_str}.xlsx")

    build_workbook(
        module_names, module_counts,
        created_module_names, created_module_counts,
        union_module_names,   union_module_counts,
        union_status_labels,  union_status_counts,
        status_labels,        status_counts,
        user_names,           user_counts,
        unique_emails,
        filtered_closure_raw, filtered_creation_raw, filtered_union_raw,
        counts_closure, counts_union,
        sheet_names,
        start_dt, end_dt,
        output_path,
    )

    print(f"\n  Report saved -> {output_path}")
    print("Done.\n")


if __name__ == "__main__":
    main()
