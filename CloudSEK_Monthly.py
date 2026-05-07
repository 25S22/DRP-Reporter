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
  │  TOTAL IN  │  CLOSED    │  IN PROGRESS     │  OPEN / PENDING        │
  │  PERIOD    │  (RAG)     │  (RAG)           │  (RAG — lower=better)  │
  ├────────────┴────────────┴──────────────────┴────────────────────────┤
  │ ③ Union table + status mini-table         (purple)                  │
  ├────────────────────────────────────────────────────────────────────-┤
  │ Stacked horizontal bar — per module union count, split by status    │
  │   ■ Closed (green)  ■ In Progress (amber)  ■ Open (red)            │
  ├────────────────────────────────────────────────────────────────────-┤
  │ User table + horizontal bar chart                                   │
  └──────────────────────────────────────────────────────────────────────┘

CHART DESIGN DECISIONS
  • Stacked horizontal bar — total bar = union count per module (all
    incidents touched in the period). The three coloured segments give
    an at-a-glance RAG health signal: a wide red or amber band signals
    a module still carrying open work. Sorted so the most-impacted
    module sits at the top. Data labels on each segment prevent readers
    having to estimate.
  • Donut (RAG)     — three-bucket health signal with semantic colours;
                      total shown in title; far less ambiguous than a pie.
  • Hidden _ChartData sheet — all chart cell-refs point here so they
                      never break when table row numbers shift.

All filtering/aggregation logic is unchanged from the previous version.
Only U (union) data sheets are written; C and R sheets are omitted.
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
# DATE PARSING
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
        df     = df_str.copy()
        native = native_sheets.get(name, pd.DataFrame())
        for col in [COL_CLOSURE_DATE, COL_CREATION_DATE]:
            if col and col in native.columns:
                df[col] = native[col].values
        merged[name] = df
    return merged

# ---------------------------------------------------------------------------
# STEP 2 — PROCESS
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
# STEP 3 — AGGREGATE
# ---------------------------------------------------------------------------

def aggregate(counts):
    names  = [n for n, c in counts.items() if c > 0]
    cnts   = [c for c in counts.values()   if c > 0]
    if not names:
        sys.exit("\nNo incidents found in the specified date range.\n")
    return names, cnts

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
# STEP 3c — USER-WISE BREAKDOWN
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
# STEP 3d — UNION STATUS BREAKDOWN
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
# STEP 3e — PER-MODULE STATUS BREAKDOWN (union incidents)
# ---------------------------------------------------------------------------

def compute_union_module_status_breakdown(filtered_union_raw, sheet_names):
    """
    For each module (sheet) returns a dict of status counts across union rows.
    Result: {module_name: {'Closed': n, 'In Progress': n, 'Open': n}}
    """
    result = {}
    for name in sheet_names:
        df = filtered_union_raw.get(name, pd.DataFrame())
        breakdown = {"Closed": 0, "In Progress": 0, "Open": 0}
        if not df.empty and COL_STATUS in df.columns:
            for val in df[COL_STATUS]:
                s = _normalize_status(val)
                if s in breakdown:
                    breakdown[s] += 1
        if any(v > 0 for v in breakdown.values()):
            result[name] = breakdown
    return result

# ---------------------------------------------------------------------------
# STEP 3f — EMAIL EXTRACTION
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
    union_module_names,        union_module_counts,
    union_module_status,       # {module: {Closed, In Progress, Open}}
    union_status_labels,       union_status_counts,
    status_labels,             status_counts,
    user_names,                user_counts,
    unique_emails,
    filtered_union_raw,
    counts_union,
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

    # Status palette (RAG)
    CLR_CLOSED  = "#70AD47"   # green
    CLR_INPROG  = "#FFC000"   # amber
    CLR_OPEN    = "#C00000"   # red
    CLR_CLOSED_D = "#375623"  # dark green text
    CLR_INPROG_D = "#7F6000"  # dark amber text
    CLR_OPEN_D   = "#9C0006"  # dark red text

    # ── KPI CALCULATIONS ─────────────────────────────────────────────────────
    total_union   = sum(union_module_counts) if union_module_counts else 0

    closed_cnt    = next((c for l, c in zip(union_status_labels, union_status_counts)
                          if l == "Closed"), 0)
    inprog_cnt    = next((c for l, c in zip(union_status_labels, union_status_counts)
                          if l == "In Progress"), 0)
    open_cnt      = next((c for l, c in zip(union_status_labels, union_status_counts)
                          if l == "Open"), 0)

    _denom        = total_union or 1
    res_rate      = closed_cnt / _denom * 100
    pend_pct      = (open_cnt + inprog_cnt) / _denom * 100

    def _rag(value, g_thresh, a_thresh, higher_is_better=True):
        if higher_is_better:
            if value >= g_thresh: return D_GREEN, LTGREEN, CLR_CLOSED
            if value >= a_thresh: return "#7F6000", "#FFF2CC", CLR_INPROG
            return "#9C0006",  "#FFC7CE", CLR_OPEN
        else:
            if value <= g_thresh: return D_GREEN, LTGREEN, CLR_CLOSED
            if value <= a_thresh: return "#7F6000", "#FFF2CC", CLR_INPROG
            return "#9C0006",  "#FFC7CE", CLR_OPEN

    rate_hdr, rate_bg, rate_bar = _rag(res_rate, 80, 60, higher_is_better=True)
    pend_hdr, pend_bg, pend_bar = _rag(pend_pct, 10, 25, higher_is_better=False)

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

    f_hdr_purp  = _hdr(D_PURP);  f_tot_purp  = _tot(D_PURP,  LTPURP)
    f_hdr_navy  = _hdr(NAVY);    f_tot_navy  = _tot(NAVY,    LTBLUE)

    # ── HIDDEN CHART-DATA SHEET ───────────────────────────────────────────────
    # Sorted by total union count descending (worst module first).
    # Layout:
    #   A  = module name
    #   B  = Closed count     (green bar segment)
    #   C  = In Progress count (amber bar segment)
    #   D  = Open count       (red bar segment)
    #   F  = overall status label (for donut)
    #   G  = overall status count
    #   I  = union status label
    #   J  = union status count
    #   L  = user name
    #   M  = user count

    CDSHEET = "_ChartData"
    cd      = wb.add_worksheet(CDSHEET)
    cd.hide()

    # Sort modules by union count descending
    mods_sorted = sorted(
        union_module_names,
        key=lambda m: union_module_counts[list(union_module_names).index(m)],
        reverse=True,
    )
    NM = len(mods_sorted)

    union_count_map  = dict(zip(union_module_names, union_module_counts))

    for i, m in enumerate(mods_sorted):
        sb  = union_module_status.get(m, {"Closed": 0, "In Progress": 0, "Open": 0})
        cd.write(i, 0, m)                      # A: module
        cd.write(i, 1, sb.get("Closed", 0))    # B: closed
        cd.write(i, 2, sb.get("In Progress", 0))  # C: in progress
        cd.write(i, 3, sb.get("Open", 0))      # D: open

    # Col F-G: overall status (for donut on all data)
    NS = len(status_labels)
    for i, (lbl, cnt) in enumerate(zip(status_labels, status_counts)):
        cd.write(i, 5, lbl)
        cd.write(i, 6, cnt)

    # Col I-J: union status
    for i, (lbl, cnt) in enumerate(zip(union_status_labels, union_status_counts)):
        cd.write(i, 8, lbl)
        cd.write(i, 9, cnt)

    # Col L-M: user
    NU = len(user_names)
    for i, (u, c) in enumerate(zip(user_names, user_counts)):
        cd.write(i, 11, u)
        cd.write(i, 12, c)

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
    # Four boxes: Total in Period | Closed | In Progress | Open/Pending
    KPI_TOP = 2
    BOX_W   = 3    # columns per box

    kpi = [
        ("TOTAL IN PERIOD",   str(total_union),
         "Unique incidents (union)",         NAVY,     LTBLUE),
        ("CLOSED",            str(closed_cnt),
         f"{res_rate:.1f}% resolution rate", rate_hdr, rate_bg),
        ("IN PROGRESS",       str(inprog_cnt),
         "Still being worked",               CLR_INPROG_D, "#FFF2CC"),
        ("OPEN / PENDING",    str(open_cnt),
         f"{pend_pct:.1f}% of period total", pend_hdr, pend_bg),
    ]

    sw.set_row(KPI_TOP,     18)
    sw.set_row(KPI_TOP + 1, 40)
    sw.set_row(KPI_TOP + 2, 16)

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

    cursor = KPI_TOP + 4

    # ── UNION TABLE (③) ───────────────────────────────────────────────────────
    if union_module_names:
        sw.set_row(cursor, 24)
        sw.write(cursor, 0,
                 "Incidents in Period — Module-wise  (Created OR Closed in Range)",
                 f_section)
        UHDR = cursor + 1;  UDS = UHDR + 1;  UTR = UDS + NM
        sw.set_row(UHDR, 22)
        for ci, h in enumerate(["#", "Module Name", "Unique Incidents"]):
            sw.write(UHDR, ci, h, f_hdr_purp)
        for i, m in enumerate(mods_sorted):
            r   = UDS + i;  alt = (i % 2 == 1)
            sw.set_row(r, 18)
            sw.write(r, 0, i + 1, f_num_alt if alt else f_num)
            sw.write(r, 1, m,     f_lft_alt if alt else f_lft)
            sw.write(r, 2, union_count_map.get(m, 0),
                     f_num_alt if alt else f_num)
        sw.set_row(UTR, 22)
        sw.merge_range(UTR, 0, UTR, 1, "TOTAL", f_tot_purp)
        sw.write_formula(UTR, 2, f"=SUM(C{UDS+1}:C{UDS+NM})", f_tot_purp)
        cursor = UTR + 2

    # ── UNION STATUS MINI-TABLE ───────────────────────────────────────────────
    if union_status_labels:
        sw.set_row(cursor, 20)
        sw.write(cursor, 0,
                 "Current Status Breakdown — All Incidents in Period", f_section)
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

    def _style(chart):
        chart.set_plotarea({"border": {"none": True}})
        chart.set_chartarea({"border": {"color": "#D9D9D9"},
                             "fill":   {"color": WHITE}})
        chart.set_style(2)

    # ── CHART 1: Stacked horizontal bar — union per module, split by status ──
    # Each module bar = total union incidents.
    # Segments: Closed (green) | In Progress (amber) | Open (red)
    # Sorted worst→best so problem modules are at top.
    if NM > 0:
        bar_h = max(360, NM * 38 + 120)
        bar_w = 700

        stacked = wb.add_chart({"type": "bar", "subtype": "stacked"})

        # Series 1: Closed — green (bottom/left segment, most positive)
        stacked.add_series({
            "name":       "Closed",
            "categories": [CDSHEET, 0, 0, NM - 1, 0],   # col A: module
            "values":     [CDSHEET, 0, 1, NM - 1, 1],   # col B: closed
            "fill":       {"color": CLR_CLOSED},
            "border":     {"color": WHITE, "width": 0.75},
            "data_labels": {
                "value":    True,
                "position": "inside_end",
                "font":     {"bold": True, "size": 9, "color": WHITE},
            },
        })
        # Series 2: In Progress — amber
        stacked.add_series({
            "name":       "In Progress",
            "categories": [CDSHEET, 0, 0, NM - 1, 0],
            "values":     [CDSHEET, 0, 2, NM - 1, 2],   # col C: in progress
            "fill":       {"color": CLR_INPROG},
            "border":     {"color": WHITE, "width": 0.75},
            "data_labels": {
                "value":    True,
                "position": "inside_end",
                "font":     {"bold": True, "size": 9, "color": "#3A3A3A"},
            },
        })
        # Series 3: Open — red (signals unresolved risk)
        stacked.add_series({
            "name":       "Open",
            "categories": [CDSHEET, 0, 0, NM - 1, 0],
            "values":     [CDSHEET, 0, 3, NM - 1, 3],   # col D: open
            "fill":       {"color": CLR_OPEN},
            "border":     {"color": WHITE, "width": 0.75},
            "data_labels": {
                "value":    True,
                "position": "inside_end",
                "font":     {"bold": True, "size": 9, "color": WHITE},
            },
        })

        stacked.set_title({
            "name": (
                f"Incidents by Module — Period Total: {total_union}  "
                f"│  ■ Closed  ■ In Progress  ■ Open"
            ),
            "name_font": {"bold": True, "size": 12, "color": NAVY},
        })
        stacked.set_legend({
            "position":  "bottom",
            "font":      {"bold": True, "size": 10},
        })
        stacked.set_x_axis({
            "num_font":        {"size": 9},
            "major_gridlines": {"visible": True,
                                "line":    {"color": "#E0E0E0", "width": 0.5}},
            "minor_gridlines": {"visible": False},
        })
        stacked.set_y_axis({
            "num_font": {"size": 10, "bold": True},
            "line":     {"none": True},
        })
        _style(stacked)
        stacked.set_size({"width": bar_w, "height": bar_h})
        sw.insert_chart(CHART_ROW, 0, stacked, {"x_offset": 5, "y_offset": 5})

    # ── CHART 2: Donut — union status (period incidents only) ────────────────
    NS_U = len(union_status_labels)
    if NS_U > 0:
        donut = wb.add_chart({"type": "doughnut"})
        donut.add_series({
            "name":       "Status",
            "categories": [CDSHEET, 0, 8, NS_U - 1, 8],
            "values":     [CDSHEET, 0, 9, NS_U - 1, 9],
            "points": [
                {"fill": {"color": _STATUS_RAG.get(l, "#4472C4")}}
                for l in union_status_labels
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
                f"Status Mix — Period Incidents\n"
                f"Total: {total_union} unique"
            ),
            "name_font": {"bold": True, "size": 10, "color": NAVY},
        })
        donut.set_legend({"none": True})
        _style(donut)
        donut.set_size({"width": 340, "height": 340})
        # Place to the right of the stacked bar
        sw.insert_chart(CHART_ROW, 12, donut, {"x_offset": 5, "y_offset": 5})

    chart_rows = max(NM * 3 + 6, 22)
    cursor     = CHART_ROW + chart_rows

    # ── USER TABLE + HORIZONTAL BAR ───────────────────────────────────────────
    if NU > 0:
        cursor += 1
        sw.set_row(cursor, 24)
        sw.write(cursor, 0,
                 "Incidents Closed By — User Wise (Closure Date Range)", f_section)
        UHDR2 = cursor + 1;  UDS2 = UHDR2 + 1;  UTR2 = UDS2 + NU
        sw.set_row(UHDR2, 22)
        for ci, h in enumerate(["#", "Closed By", "Unique Incidents Closed"]):
            sw.write(UHDR2, ci, h, f_hdr_navy)
        for i in range(NU):
            r   = UDS2 + i;  alt = (i % 2 == 1)
            sw.set_row(r, 18)
            sw.write(r, 0, i + 1,          f_num_alt if alt else f_num)
            sw.write(r, 1, user_names[i],  f_lft_alt if alt else f_lft)
            sw.write(r, 2, user_counts[i], f_num_alt if alt else f_num)
        sw.set_row(UTR2, 22)
        sw.merge_range(UTR2, 0, UTR2, 1, "TOTAL", f_tot_navy)
        sw.write_formula(UTR2, 2, f"=SUM(C{UDS2+1}:C{UDS2+NU})", f_tot_navy)

        bar3 = wb.add_chart({"type": "bar"})
        bar3.add_series({
            "name":       "Incidents Closed",
            "categories": [CDSHEET, 0, 11, NU - 1, 11],
            "values":     [CDSHEET, 0, 12, NU - 1, 12],
            "fill":       {"color": MIDBLUE},
            "border":     {"color": WHITE, "width": 0.75},
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
                          "major_gridlines": {"visible": True,
                                              "line":    {"color": "#E0E0E0",
                                                          "width": 0.5}}})
        bar3.set_y_axis({"num_font": {"size": 10, "bold": True},
                          "line":     {"none": True}})
        _style(bar3)
        bar3.set_size({"width": 520, "height": max(300, NU * 30 + 120)})
        sw.insert_chart(UTR2 + 2, 0, bar3, {"x_offset": 5, "y_offset": 5})

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

    # ── U (UNION) DATA SHEETS ONLY ────────────────────────────────────────────
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

    # Write only U sheets (union: created OR closed in range)
    for name in sheet_names:
        if counts_union.get(name, 0) > 0:
            _write_data_sheet(f"U - {name[:25]}", filtered_union_raw.get(name))

    wb.close()

# ---------------------------------------------------------------------------
# MAIN
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

    if not any(c > 0 for c in counts_union.values()):
        sys.exit("\nNo incidents found in the specified date range.\n")

    # Union module list
    pairs               = [(n, c) for n, c in counts_union.items() if c > 0]
    union_module_names  = [p[0] for p in pairs]
    union_module_counts = [p[1] for p in pairs]

    print(f"\n  Union (period total): {len(union_module_names)} modules / "
          f"{sum(union_module_counts)} incidents")

    status_labels,          status_counts          = compute_status_breakdown(
        processed_raw, sheet_names)
    user_names,             user_counts            = compute_user_breakdown(
        filtered_closure_raw, sheet_names)
    union_status_labels,    union_status_counts    = compute_union_status_breakdown(
        filtered_union_raw, sheet_names)
    union_module_status                            = compute_union_module_status_breakdown(
        filtered_union_raw, sheet_names)
    unique_emails                                  = extract_emails_closed_resolved(
        processed_raw, sheet_names)

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    date_range_str = (f"{start_dt.strftime('%d %b %Y')} - "
                      f"{end_dt.strftime('%d %b %Y')}")
    output_path    = os.path.join(OUTPUT_FOLDER,
                                  f"Incident Review - {date_range_str}.xlsx")

    build_workbook(
        union_module_names,     union_module_counts,
        union_module_status,
        union_status_labels,    union_status_counts,
        status_labels,          status_counts,
        user_names,             user_counts,
        unique_emails,
        filtered_union_raw,
        counts_union,
        sheet_names,
        start_dt, end_dt,
        output_path,
    )

    print(f"\n  Report saved -> {output_path}")
    print("Done.\n")


if __name__ == "__main__":
    main()
