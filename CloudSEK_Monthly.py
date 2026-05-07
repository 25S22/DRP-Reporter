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
    "Open":        "#EF4444",
    "In Progress": "#F59E0B",
    "Closed":      "#22C55E",
}

_PALETTE = [
    "#2563EB", "#F59E0B", "#22C55E", "#EF4444", "#8B5CF6",
    "#06B6D4", "#EC4899", "#A855F7", "#14B8A6", "#64748B",
    "#1D4ED8", "#0F766E", "#B91C1C", "#C2410C", "#4C1D95",
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
def _progress_bar(pct, width=15):
    """Returns a Unicode block-char progress strip  e.g. '██████░░░░░░░░░  40%'"""
    filled = round(min(max(float(pct), 0.0), 100.0) / 100.0 * width)
    return "█" * filled + "░" * (width - filled) + f"   {pct:.0f}%"


def _add_insights_panel(sw, wb, cursor,
                         mods_sorted, union_count_map, union_module_status,
                         res_rate, user_names, user_counts,
                         NAVY, WHITE):
    """
    Writes a 4-box insight strip starting at `cursor`.
    Each box: coloured icon header | bold metric | italic subtitle | thin accent.
    Returns the next free row after the panel.
    """
 
    def _f(**kw):
        base = {"font_name": "Calibri", "font_size": 10, "valign": "vcenter"}
        base.update(kw)
        return wb.add_format(base)
 
    # ── Compute four insight values ───────────────────────────────────────────
    most_mod = mods_sorted[0] if mods_sorted else "—"
    most_cnt = union_count_map.get(most_mod, 0)
 
    if   res_rate >= 80: rate_h, rate_b, rate_s = "#065F46", "#D1FAE5", "✓  ON TARGET"
    elif res_rate >= 60: rate_h, rate_b, rate_s = "#B45309", "#FEF3C7", "~  NEEDS ATTENTION"
    else:                rate_h, rate_b, rate_s = "#B91C1C", "#FEE2E2", "✗  BELOW TARGET"
 
    risk_mod = "—"; risk_cnt = 0
    for m in mods_sorted:
        n = union_module_status.get(m, {}).get("Open", 0)
        if n > risk_cnt:
            risk_cnt = n; risk_mod = m
 
    top_u = user_names[0]  if user_names  else "—"
    top_c = user_counts[0] if user_counts else 0
 
    boxes = [
        {
            "icon": "◆  MOST IMPACTED MODULE",
            "val":  most_mod[:24],
            "sub":  f"{most_cnt:,} incidents in period",
            "h": "#2563EB", "b": "#E0E7FF", "t": "#111827",
        },
    ]
 
    # ── Section label ─────────────────────────────────────────────────────────
    sw.set_row(cursor, 16)
    sw.merge_range(cursor, 0, cursor, 14,
                   "  KEY INSIGHTS  —  AUTO-GENERATED FROM PERIOD DATA",
                   _f(bold=True, font_size=9, font_color="#374151", italic=True,
                      bg_color="#F9FAFB", align="left",
                      left=5, left_color="#2563EB", top=1, bottom=1,
                      top_color="#E5E7EB", bottom_color="#E5E7EB"))
    cursor += 1
 
    BOX_W = 15  # spans columns 0-14
    sw.set_row(cursor,     20)   # icon header row
    sw.set_row(cursor + 1, 44)   # big metric row
    sw.set_row(cursor + 2, 16)   # subtitle row
    sw.set_row(cursor + 3,  6)   # thin accent strip
 
    for bi, bx in enumerate(boxes):
        c1 = bi * BOX_W;  c2 = c1 + BOX_W - 1
 
        # Icon / header band
        sw.merge_range(cursor, c1, cursor, c2, bx["icon"],
                       _f(bold=True, font_size=9, font_color=WHITE,
                          bg_color=bx["h"], align="center",
                          border=1, top=2, top_color=bx["h"]))
 
        # Large metric value
        sw.merge_range(cursor + 1, c1, cursor + 1, c2, bx["val"],
                       _f(bold=True, font_size=20, font_color=bx["t"],
                          bg_color=bx["b"], align="center", border=1))
 
        # Italic subtitle
        sw.merge_range(cursor + 2, c1, cursor + 2, c2, bx["sub"],
                       _f(font_size=9, italic=True, font_color="#4B5563",
                          bg_color=bx["b"], align="center", border=1))
 
        # Thin bottom accent strip (same colour as header)
        sw.merge_range(cursor + 3, c1, cursor + 3, c2, "",
                       _f(bg_color=bx["h"]))
 
    return cursor + 5   # 4 content rows + 1 blank spacer row

 
def build_workbook(
    union_module_names,        union_module_counts,
    union_module_status,
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
    import xlsxwriter
    import pandas as pd
    from datetime import datetime
 
    wb = xlsxwriter.Workbook(output_path)
 
    # ── COLOUR TOKENS ────────────────────────────────────────────────────────
    NAVY     = "#111827"
    MIDBLUE  = "#2563EB"
    LTBLUE   = "#DBEAFE"
    D_GREEN  = "#065F46"
    LTGREEN  = "#D1FAE5"
    D_PURP   = "#6D28D9"
    LTPURP   = "#EDE9FE"
    ALT      = "#F3F4F6"
    WHITE    = "#FFFFFF"
    OFFWHITE = "#F9FAFB"
    SUBTLE   = "#4B5563"
 
    CLR_CLOSED   = "#22C55E"
    CLR_INPROG   = "#F59E0B"
    CLR_OPEN     = "#EF4444"
    CLR_CLOSED_D = "#15803D"
    CLR_INPROG_D = "#B45309"
    CLR_OPEN_D   = "#B91C1C"
 
    # ── KPI CALCULATIONS ─────────────────────────────────────────────────────
    total_union  = sum(union_module_counts) if union_module_counts else 0
    closed_cnt   = next((c for l, c in zip(union_status_labels, union_status_counts)
                         if l == "Closed"),      0)
    inprog_cnt   = next((c for l, c in zip(union_status_labels, union_status_counts)
                         if l == "In Progress"), 0)
    open_cnt     = next((c for l, c in zip(union_status_labels, union_status_counts)
                         if l == "Open"),        0)
    _d         = total_union or 1
    res_rate   = closed_cnt  / _d * 100
    pend_pct   = (open_cnt + inprog_cnt) / _d * 100
 
    def _rag_colors(value, g, a, higher=True):
        if higher:
            if value >= g: return D_GREEN,     LTGREEN,  CLR_CLOSED
            if value >= a: return CLR_INPROG_D, "#FEF3C7", CLR_INPROG
            return CLR_OPEN_D, "#FEE2E2", CLR_OPEN
        else:
            if value <= g: return D_GREEN,     LTGREEN,  CLR_CLOSED
            if value <= a: return CLR_INPROG_D, "#FEF3C7", CLR_INPROG
            return CLR_OPEN_D, "#FEE2E2", CLR_OPEN
 
    rate_hdr, rate_bg, _ = _rag_colors(res_rate, 80, 60, higher=True)
    pend_hdr, pend_bg, _ = _rag_colors(pend_pct, 10, 25, higher=False)
 
    # ── FORMAT FACTORY ───────────────────────────────────────────────────────
    def _f(**kw):
        base = {"font_name": "Calibri", "font_size": 10, "valign": "vcenter"}
        base.update(kw)
        return wb.add_format(base)
 
    border_grey = "#E5E7EB"
    f_num       = _f(align="center", border=1, border_color=border_grey,
                     font_color=NAVY)
    f_lft       = _f(align="left",   border=1, border_color=border_grey,
                     font_color=NAVY)
    f_num_alt   = _f(align="center", border=1, border_color=border_grey,
                     bg_color=ALT, font_color=NAVY)
    f_lft_alt   = _f(align="left",   border=1, border_color=border_grey,
                     bg_color=ALT, font_color=NAVY)
    f_data_hdr  = _f(bold=True, font_color=WHITE, bg_color=MIDBLUE,
                     align="center", border=1, border_color=MIDBLUE)
    f_cell      = _f(align="left",  border=1, border_color=border_grey,
                     font_color=NAVY)
    f_cell_alt  = _f(align="left",  border=1, border_color=border_grey,
                     bg_color=ALT, font_color=NAVY)
    f_date      = _f(align="left",  border=1, border_color=border_grey,
                     num_format="dd mmm yyyy", font_color=NAVY)
    f_date_alt  = _f(align="left",  border=1, border_color=border_grey,
                     bg_color=ALT, num_format="dd mmm yyyy", font_color=NAVY)
 
    # Section heading — left accent bar effect
    f_section = _f(bold=True, font_size=12, font_color=NAVY,
                   left=5, left_color=MIDBLUE, bg_color=OFFWHITE)
 
    def _hdr(bg):
        return _f(bold=True, font_size=11, font_color=WHITE,
                  bg_color=bg, align="center", border=1, border_color=bg)
    def _tot(fg, bg):
        return _f(bold=True, font_size=11, font_color=fg,
                  bg_color=bg, align="center", border=1, border_color=border_grey)
 
    f_hdr_purp = _hdr(D_PURP);  f_tot_purp = _tot(D_PURP, LTPURP)
    f_hdr_navy = _hdr(NAVY);    f_tot_navy = _tot(NAVY,   LTBLUE)
 
    # ── HIDDEN CHART-DATA SHEET ───────────────────────────────────────────────
    # Sorted by union count descending — worst module at top of every chart.
    # Layout:
    #   A  module name          F  overall status label    K  user name
    #   B  Closed count         G  overall status count    L  user count
    #   C  In Progress count    I  union status label
    #   D  Open count           J  union status count
 
    CDSHEET = "_ChartData"
    cd = wb.add_worksheet(CDSHEET)
    cd.hide()
 
    union_count_map = dict(zip(union_module_names, union_module_counts))
    mods_sorted = sorted(
        union_module_names,
        key=lambda m: union_count_map.get(m, 0),
        reverse=True,
    )
    NM = len(mods_sorted)
 
    for i, m in enumerate(mods_sorted):
        sb = union_module_status.get(m, {"Closed": 0, "In Progress": 0, "Open": 0})
        cd.write(i, 0, m)
        cd.write(i, 1, sb.get("Closed", 0))
        cd.write(i, 2, sb.get("In Progress", 0))
        cd.write(i, 3, sb.get("Open", 0))
 
    NS = len(status_labels)
    for i, (l, c) in enumerate(zip(status_labels, status_counts)):
        cd.write(i, 5, l);  cd.write(i, 6, c)
 
    NS_U = len(union_status_labels)
    for i, (l, c) in enumerate(zip(union_status_labels, union_status_counts)):
        cd.write(i, 8, l);  cd.write(i, 9, c)
 
    NU = len(user_names)
    for i, (u, c) in enumerate(zip(user_names, user_counts)):
        cd.write(i, 11, u);  cd.write(i, 12, c)
 
    # ── SUMMARY SHEET ─────────────────────────────────────────────────────────
    sw = wb.add_worksheet(SUMMARY_SHEET_NAME)
    sw.set_zoom(90)
    sw.set_column(0,  0,   6)
    sw.set_column(1,  1,  52)
    sw.set_column(2,  2,  20)
    sw.set_column(3, 14,   9)
 
    # ── BANNER (2 rows) ───────────────────────────────────────────────────────
    #   Row 0: bold dark title
    #   Row 1: lighter subtitle with period and generated date
 
    sw.set_row(0, 40)
    sw.merge_range(0, 0, 0, 14,
                   "  INCIDENT SUMMARY DASHBOARD",
                   _f(bold=True, font_size=20, font_color=WHITE,
                      bg_color=NAVY, align="left", valign="vcenter"))
 
    sw.set_row(1, 20)
    sw.merge_range(1, 0, 1, 14,
                   f"  Reporting Period:  "
                   f"{start_dt.strftime('%d %b %Y')}  ─  {end_dt.strftime('%d %b %Y')}"
                   f"     │     Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}",
                   _f(font_size=10, font_color="#E0E7FF", italic=True,
                      bg_color=MIDBLUE, align="left", valign="vcenter"))
 
    sw.set_row(2, 8)    # breathing gap below banner
 
    # ── KPI SCORECARD (rows 3-7) ──────────────────────────────────────────────
    #   Row 3: small label header
    #   Row 4: large number
    #   Row 5: subtitle + progress bar
    #   Row 6: ultra-thin accent strip
    #   Row 7: breathing gap
 
    KPI_TOP = 3
    BOX_W   = 3
 
    kpi = [
        ("TOTAL IN PERIOD",   str(total_union),
         _progress_bar(100, 15),
         NAVY, LTBLUE),
        ("CLOSED",            str(closed_cnt),
         _progress_bar(res_rate),
         rate_hdr, rate_bg),
        ("IN PROGRESS",       str(inprog_cnt),
         _progress_bar(inprog_cnt / _d * 100),
         CLR_INPROG_D, "#FEF3C7"),
        ("OPEN / PENDING",    str(open_cnt),
         _progress_bar(pend_pct),
         pend_hdr, pend_bg),
    ]
 
    sw.set_row(KPI_TOP,     16)   # label row
    sw.set_row(KPI_TOP + 1, 46)   # BIG number row
    sw.set_row(KPI_TOP + 2, 18)   # progress bar row
    sw.set_row(KPI_TOP + 3,  4)   # thin accent strip
    sw.set_row(KPI_TOP + 4,  8)   # gap
 
    for bi, (label, value, progbar, hdr_c, bg_c) in enumerate(kpi):
        c1 = bi * BOX_W;  c2 = c1 + BOX_W - 1
 
        # Label header
        sw.merge_range(KPI_TOP, c1, KPI_TOP, c2, label,
                       _f(bold=True, font_size=10, font_color=WHITE,
                          bg_color=hdr_c, align="center", border=1,
                          border_color=hdr_c))
 
        # Large number
        sw.merge_range(KPI_TOP + 1, c1, KPI_TOP + 1, c2, value,
                       _f(bold=True, font_size=28, font_color=NAVY,
                          bg_color=bg_c, align="center",
                          left=1, right=1, top=0, bottom=0,
                          left_color="#E5E7EB", right_color="#E5E7EB"))
 
        # Progress bar row
        sw.merge_range(KPI_TOP + 2, c1, KPI_TOP + 2, c2, progbar,
                       _f(font_size=8, font_color=hdr_c, bold=True,
                          bg_color=bg_c, align="center",
                          left=1, right=1, top=0, bottom=0,
                          left_color="#E5E7EB", right_color="#E5E7EB",
                          font_name="Courier New"))   # monospace for block chars
 
        # Thin accent strip  (solid colour = visual "bottom border")
        sw.merge_range(KPI_TOP + 3, c1, KPI_TOP + 3, c2, "",
                       _f(bg_color=hdr_c))
 
    # Freeze everything above the insights panel
    sw.freeze_panes(KPI_TOP + 5, 0)
 
    cursor = KPI_TOP + 5   # row 8 — start of scrollable content
 
    # ── THIN SECTION DIVIDER ──────────────────────────────────────────────────
    def _divider(row, color=MIDBLUE):
        sw.set_row(row, 3)
        sw.merge_range(row, 0, row, 14, "", _f(bg_color=color))
 
    # ── INSIGHTS PANEL ────────────────────────────────────────────────────────
    cursor = _add_insights_panel(
        sw, wb, cursor,
        mods_sorted, union_count_map, union_module_status,
        res_rate, user_names, user_counts,
        NAVY, WHITE,
    )
 
    _divider(cursor);  cursor += 1
 
    # ── MODULE TABLE (union) ──────────────────────────────────────────────────
    if union_module_names:
        sw.set_row(cursor, 26)
        sw.write(cursor, 0,
                 "  Incidents in Period  —  Module-wise  (Created OR Closed in Range)",
                 f_section)
        HDR = cursor + 1;  DS = HDR + 1;  TR = DS + NM
 
        sw.set_row(HDR, 24)
        for ci, h in enumerate(["#", "Module Name", "Unique Incidents"]):
            sw.write(HDR, ci, h, f_hdr_purp)
 
        for i, m in enumerate(mods_sorted):
            r = DS + i;  alt = (i % 2 == 1)
            sw.set_row(r, 18)
            sw.write(r, 0, i + 1,                f_num_alt if alt else f_num)
            sw.write(r, 1, m,                    f_lft_alt if alt else f_lft)
            sw.write(r, 2, union_count_map[m],   f_num_alt if alt else f_num)
 
        # ── Data bar on count column ──────────────────────────────────────────
        sw.conditional_format(DS, 2, TR - 1, 2, {
            "type":             "data_bar",
            "data_bar_2010":    True,
            "bar_color":        "#60A5FA",
            "bar_border_color": MIDBLUE,
            "bar_solid":        True,
            "min_type":         "num",
            "min_value":        0,
            "bar_direction":    "left",
        })
 
        sw.set_row(TR, 24)
        sw.merge_range(TR, 0, TR, 1, "TOTAL", f_tot_purp)
        sw.write_formula(TR, 2, f"=SUM(C{DS+1}:C{DS+NM})", f_tot_purp)
        cursor = TR + 2
 
    # ── UNION STATUS MINI-TABLE ───────────────────────────────────────────────
    _STATUS_RAG = {"Open": "#EF4444", "In Progress": "#F59E0B", "Closed": "#22C55E"}
 
    if union_status_labels:
        sw.set_row(cursor, 26)
        sw.write(cursor, 0,
                 "  Current Status Breakdown  —  All Incidents in Period",
                 f_section)
        cursor += 1
        for ci, h in enumerate(["Status", "Count", "% Share"]):
            sw.write(cursor, ci, h, f_hdr_purp)
        cursor += 1
        total_u = sum(union_status_counts) or 1
        for lbl, cnt in zip(union_status_labels, union_status_counts):
            bg = _STATUS_RAG.get(lbl, ALT)
            sw.write(cursor, 0, lbl,
                     _f(bold=True, align="left",   border=1, bg_color=bg,
                        font_color=WHITE))
            sw.write(cursor, 1, cnt,
                     _f(bold=True, align="center", border=1, bg_color=bg,
                        font_color=WHITE))
            sw.write(cursor, 2, cnt / total_u,
                     _f(align="center", border=1, bg_color=bg,
                        num_format="0.0%", font_color=WHITE))
            cursor += 1
        cursor += 1
 
    _divider(cursor);  cursor += 1
 
    # ── CHART SECTION ─────────────────────────────────────────────────────────
    sw.set_row(cursor, 22)
    sw.write(cursor, 0, "  Visual Summary", f_section)
    cursor += 1
    CHART_ROW = cursor
 
    # Shared chart polish helper
    def _style(chart):
        chart.set_plotarea({
            "border": {"none": True},
            "fill":   {"color": OFFWHITE},
        })
        chart.set_chartarea({
            "border": {"color": "#E5E7EB", "width": 0.75},
            "fill":   {"color": WHITE},
        })
        chart.set_style(2)
 
    # ── CHART 1 — Stacked horizontal bar: Union per module split by status ────
    #
    # The biggest single visual upgrade is set_gap(50):
    #   default gap = 150 % of bar width → bars look thin and "airy"
    #   gap = 50 → bars are 3× thicker — immediately more impactful
    #
    # Segment order: Closed (green, left) | In Progress (amber) | Open (red)
    # Each segment carries a bold white data label with the exact count.
    # Modules sorted worst → best so the most-impacted row is at the top.
 
    if NM > 0:
        bar_h = max(380, NM * 42 + 130)
        bar_w = 680
 
        stacked = wb.add_chart({"type": "bar", "subtype": "stacked"})
        stacked.set_gap(50)         # ← FAT bars — single biggest visual impact
 
        # Series: Closed (green)
        stacked.add_series({
            "name":       "Closed",
            "categories": [CDSHEET, 0, 0, NM - 1, 0],
            "values":     [CDSHEET, 0, 1, NM - 1, 1],
            "fill":       {"color": CLR_CLOSED},
            "border":     {"color": WHITE, "width": 1.0},
            "data_labels": {
                "value":    True,
                "position": "inside_end",
                "font":     {"bold": True, "size": 9, "color": WHITE,
                             "name": "Calibri"},
            },
        })
        # Series: In Progress (amber)
        stacked.add_series({
            "name":       "In Progress",
            "categories": [CDSHEET, 0, 0, NM - 1, 0],
            "values":     [CDSHEET, 0, 2, NM - 1, 2],
            "fill":       {"color": CLR_INPROG},
            "border":     {"color": WHITE, "width": 1.0},
            "data_labels": {
                "value":    True,
                "position": "inside_end",
                "font":     {"bold": True, "size": 9, "color": NAVY,
                             "name": "Calibri"},
            },
        })
        # Series: Open (red — signals unresolved risk)
        stacked.add_series({
            "name":       "Open",
            "categories": [CDSHEET, 0, 0, NM - 1, 0],
            "values":     [CDSHEET, 0, 3, NM - 1, 3],
            "fill":       {"color": CLR_OPEN},
            "border":     {"color": WHITE, "width": 1.0},
            "data_labels": {
                "value":    True,
                "position": "inside_end",
                "font":     {"bold": True, "size": 9, "color": WHITE,
                             "name": "Calibri"},
            },
        })
 
        stacked.set_title({
            "name":      f"Incidents by Module  ·  Period Total: {total_union:,}",
            "name_font": {"bold": True, "size": 12, "color": NAVY,
                          "name": "Calibri"},
        })
        stacked.set_legend({
            "position": "bottom",
            "font":     {"bold": True, "size": 10, "color": SUBTLE,
                         "name": "Calibri"},
            "border":   {"color": "#E5E7EB"},
            "fill":     {"color": OFFWHITE},
        })
        stacked.set_x_axis({
            "num_font":        {"size": 9, "color": SUBTLE},
            "major_gridlines": {
                "visible": True,
                "line":    {"color": "#E5E7EB", "width": 0.6,
                            "dash_type": "dash"},
            },
            "line":            {"color": "#E5E7EB"},
            "num_format":      "0",
            "min":             0,
        })
        stacked.set_y_axis({
            "num_font":        {"size": 10, "bold": True, "color": NAVY},
            "line":            {"none": True},
            "major_tick_mark": "none",
            "major_gridlines": {"visible": False},
        })
        _style(stacked)
        stacked.set_size({"width": bar_w, "height": bar_h})
        sw.insert_chart(CHART_ROW, 0, stacked, {"x_offset": 5, "y_offset": 5})
 
    # ── CHART 2 — Donut: status mix for union incidents ───────────────────────
    #
    # RAG colours (red/amber/green) encode health at a glance.
    # Category + value + percentage on each slice — no legend needed.
    # Total shown in title so stakeholders never need to sum slices.
 
    if NS_U > 0:
        donut = wb.add_chart({"type": "doughnut"})
        donut.add_series({
            "name":       "Status",
            "categories": [CDSHEET, 0, 8, NS_U - 1, 8],
            "values":     [CDSHEET, 0, 9, NS_U - 1, 9],
            "points": [
                {"fill": {"color": _STATUS_RAG.get(l, MIDBLUE)}}
                for l in union_status_labels
            ],
            "data_labels": {
                "percentage": True,
                "category":   True,
                "value":      True,
                "separator":  "\n",
                "font":       {"bold": True, "size": 9, "name": "Calibri",
                               "color": NAVY},
            },
        })
        donut.set_title({
            "name": (
                f"Period Status Mix\n"
                f"Total: {total_union:,} unique incidents"
            ),
            "name_font": {"bold": True, "size": 10, "color": NAVY,
                          "name": "Calibri"},
        })
        donut.set_legend({"none": True})
        _style(donut)
        donut.set_size({"width": 340, "height": 350})
        # Positioned to the right of the stacked bar
        sw.insert_chart(CHART_ROW, 12, donut, {"x_offset": 5, "y_offset": 5})
 
    chart_rows = max(NM * 3 + 8, 24)
    cursor     = CHART_ROW + chart_rows
 
    _divider(cursor);  cursor += 1
 
    # ── USER TABLE + HORIZONTAL BAR ───────────────────────────────────────────
    if NU > 0:
        sw.set_row(cursor, 26)
        sw.write(cursor, 0,
                 "  Incidents Closed By  —  User Wise  (Closure Date Range)",
                 f_section)
        UHDR = cursor + 1;  UDS = UHDR + 1;  UTR = UDS + NU
 
        sw.set_row(UHDR, 24)
        for ci, h in enumerate(["#", "Closed By", "Incidents Closed"]):
            sw.write(UHDR, ci, h, f_hdr_navy)
 
        for i in range(NU):
            r = UDS + i;  alt = (i % 2 == 1)
            sw.set_row(r, 18)
            sw.write(r, 0, i + 1,          f_num_alt if alt else f_num)
            sw.write(r, 1, user_names[i],  f_lft_alt if alt else f_lft)
            sw.write(r, 2, user_counts[i], f_num_alt if alt else f_num)
 
        # Data bar on user count column too
        sw.conditional_format(UDS, 2, UTR - 1, 2, {
            "type":             "data_bar",
            "data_bar_2010":    True,
            "bar_color":        "#60A5FA",
            "bar_border_color": MIDBLUE,
            "bar_solid":        True,
            "min_type":         "num",
            "min_value":        0,
            "bar_direction":    "left",
        })
 
        sw.set_row(UTR, 24)
        sw.merge_range(UTR, 0, UTR, 1, "TOTAL", f_tot_navy)
        sw.write_formula(UTR, 2, f"=SUM(C{UDS+1}:C{UDS+NU})", f_tot_navy)
 
        # User bar chart — consistent with stacked bar palette
        bar3 = wb.add_chart({"type": "bar"})
        bar3.set_gap(55)
 
        bar3.add_series({
            "name":       "Incidents Closed",
            "categories": [CDSHEET, 0, 11, NU - 1, 11],
            "values":     [CDSHEET, 0, 12, NU - 1, 12],
            "fill":       {"color": MIDBLUE},
            "border":     {"color": WHITE, "width": 0.75},
            "data_labels": {
                "value":    True,
                "position": "inside_end",
                "font":     {"bold": True, "size": 9, "color": WHITE,
                             "name": "Calibri"},
            },
        })
        bar3.set_title({
            "name":      f"Incidents Closed by User  ·  Total: {sum(user_counts):,}",
            "name_font": {"bold": True, "size": 11, "color": NAVY,
                          "name": "Calibri"},
        })
        bar3.set_legend({"none": True})
        bar3.set_x_axis({
            "num_font":        {"size": 9, "color": SUBTLE},
            "major_gridlines": {
                "visible": True,
                "line":    {"color": "#E5E7EB", "width": 0.6,
                            "dash_type": "dash"},
            },
            "line":       {"color": "#E5E7EB"},
            "num_format": "0",
            "min":        0,
        })
        bar3.set_y_axis({
            "num_font":        {"size": 10, "bold": True, "color": NAVY},
            "line":            {"none": True},
            "major_tick_mark": "none",
            "major_gridlines": {"visible": False},
        })
        _style(bar3)
        bar3.set_size({"width": 520, "height": max(300, NU * 32 + 120)})
        sw.insert_chart(UTR + 2, 0, bar3, {"x_offset": 5, "y_offset": 5})
 
    # ── EMAIL SHEET ───────────────────────────────────────────────────────────
    if unique_emails:
        ew = wb.add_worksheet("Emails - Closed Resolved")
        ew.set_row(0, 38)
        ew.merge_range(
            0, 0, 0, 2,
            f"Unique Emails — Closed Resolved  │  {len(unique_emails):,} addresses",
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
            r  = 3 + i;  alt = (i % 2 == 1)
            bg = ALT if alt else WHITE
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
 
        dw.set_default_row(16)
        dw.set_row(0, 22)
        for ci, h in enumerate(headers):
            dw.write(0, ci, h, f_data_hdr)
            dw.set_column(ci, ci, max(len(str(h)) + 4, 14))
 
        vals = df.values
        for ri in range(nr):
            er = ri + 1;  alt = (er % 2 == 0)
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
