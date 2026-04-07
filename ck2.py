"""
=============================================================================
INCIDENT REPORT GENERATOR  (with Credential Breach Analysis)
=============================================================================
CONFIGURATION — only edit the lines in the CONFIGURATION block below.

  INPUT_FILE_PATH          : path to the source Excel workbook
  COL_INCIDENT_ID          : exact column name for the Incident ID
  COL_CLOSURE_DATE         : exact column name for the Closure Date
  COL_STATUS               : exact column name for the Status column
  COL_CLOSED_BY            : exact column name for who closed the incident

  CREDENTIAL_BREACH_SHEET  : exact sheet name for credential breach data
  COL_BREACH_EMAIL         : exact column name for the email address
  COL_BREACH_PASSWORD      : exact column name for the password

Summary Dashboard contains:
  1.  Module-wise unique closed incidents table  (date-range filtered)
  2.  Pie — Closed Incidents By Module           (date-range filtered)
  3.  Pie — Overall Incident Status              (all data, unique IDs)
  4.  User-wise closed incidents table + pie     (date-range filtered)
  ── Credential Breach section (reads raw sheet, not date-filtered) ──
  5.  Password Strength vs. Policy table + pie
  6.  Top 20 most-seen email addresses table
  7.  Domain breakdown table + pie               (new insight)
  8.  Breach severity summary (strong/weak/none) bar context

Core incident counting logic is identical to v1 — numbers will match exactly.
=============================================================================
"""

import os
import re
import sys
from collections import Counter
from datetime import datetime

import pandas as pd
import xlsxwriter
import warnings
warnings.filterwarnings("ignore")

# =============================================================================
# >>>  CONFIGURATION — only edit these lines  <<<
# =============================================================================

INPUT_FILE_PATH  = "incidents.xlsx"
OUTPUT_FOLDER    = "reports"

COL_INCIDENT_ID  = "Incident Id"
COL_CLOSURE_DATE = "Incident Closure on"
COL_STATUS       = "Status"
COL_CLOSED_BY    = "Incident Closed By"

CREDENTIAL_BREACH_SHEET = "Credential Breaches"
COL_BREACH_EMAIL        = "Email"
COL_BREACH_PASSWORD     = "Password"

# =============================================================================
# INTERNALS
# =============================================================================

SUMMARY_SHEET_NAME = "Summary Dashboard"

_ORDINAL_RE = re.compile(r"(\d+)(st|nd|rd|th)\b", re.IGNORECASE)
_DATE_FORMATS = [
    "%d %b, %Y %I:%M:%S %p", "%d %B, %Y %I:%M:%S %p",
    "%d %b, %Y %H:%M:%S",    "%d %B, %Y %H:%M:%S",
    "%d %b, %Y %I:%M %p",    "%d %b, %Y", "%d %B, %Y",
    "%d %b %Y", "%d %B %Y",  "%d-%b-%Y",  "%d-%B-%Y",
    "%d/%m/%Y", "%m/%d/%Y",  "%Y-%m-%d",  "%d.%m.%Y",
    "%d %b %y", "%d %B %y",
]

_BLANK_VALUES = {"", "nan", "none", "n/a", "na", "-", "null", "nat"}

# ---------------------------------------------------------------------------
# DATE PARSING  (unchanged from v1)
# ---------------------------------------------------------------------------

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
# STEP 1 — LOAD  (unchanged from v1)
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
# STEP 2 — PROCESS  (unchanged from v1)
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
# STEP 3 — AGGREGATE  (unchanged from v1)
# ---------------------------------------------------------------------------

def aggregate(counts):
    module_names, module_counts = [], []
    for sheet_name, count in counts.items():
        if count > 0:
            module_names.append(sheet_name)
            module_counts.append(count)
    if not module_names:
        sys.exit("\nNo incidents found in the specified date range.\n")
    return module_names, module_counts

# ---------------------------------------------------------------------------
# STEP 3b — OVERALL STATUS BREAKDOWN  (unchanged from v1)
# ---------------------------------------------------------------------------

def _normalize_status(val):
    if pd.isna(val):
        return None
    s = str(val).strip().lower()
    if s.startswith("closed"):       return "Closed"
    if s.startswith("open"):         return "Open"
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
        print(f"  WARNING: Column '{COL_STATUS}' not found — status pie skipped.")
        return [], []

    combined["_status_norm"] = combined[COL_STATUS].apply(_normalize_status)
    combined = combined[combined["_status_norm"].notna()]
    vc       = combined["_status_norm"].value_counts()
    order    = ["Open", "In Progress", "Closed"]
    labels   = [s for s in order if s in vc.index]
    counts   = [int(vc[s]) for s in labels]

    print(f"\n  Overall status breakdown (unique incidents, 3 buckets only):")
    for lbl, cnt in zip(labels, counts):
        print(f"    {lbl}: {cnt}")

    return labels, counts

# ---------------------------------------------------------------------------
# STEP 3c — USER-WISE BREAKDOWN  (unchanged from v1)
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
        ~combined[COL_CLOSED_BY].astype(str).str.strip().str.lower().isin(_BLANK_VALUES)
    ]

    vc = (
        combined[COL_CLOSED_BY]
        .astype(str).str.strip()
        .value_counts()
        .sort_values(ascending=False)
    )

    user_names  = list(vc.index)
    user_counts = [int(c) for c in vc.values]

    print(f"\n  Closed-by breakdown (date range, unique incidents):")
    for u, c in zip(user_names, user_counts):
        print(f"    {u}: {c}")

    return user_names, user_counts

# ---------------------------------------------------------------------------
# STEP 3d — CREDENTIAL BREACH ANALYSIS
#   Reads the raw credential breach sheet directly (no date filtering).
#   Returns:
#     pwd_labels, pwd_counts   — password strength buckets
#     top_emails               — list of (email_str, count_int), up to 20
#     domain_labels            — top email domains
#     domain_counts
# ---------------------------------------------------------------------------

_STRONG_PASSWORD_RE = re.compile(
    r"^(?=.*[A-Z])(?=.*[a-z])(?=.*\d)(?=.*[^A-Za-z0-9]).+$"
)


def _classify_password(val):
    if pd.isna(val):
        return "No Password"
    s = str(val).strip()
    if s.lower() in _BLANK_VALUES:
        return "No Password"
    return "Strong" if _STRONG_PASSWORD_RE.match(s) else "Weak / No Policy"


def compute_credential_breach_analysis(raw_sheets):
    """
    Reads CREDENTIAL_BREACH_SHEET from raw_sheets (all rows, no date filter).
    Returns:
      pwd_labels  : list[str]
      pwd_counts  : list[int]
      top_emails  : list[tuple[str, int]]  — up to 20, descending
      domain_labels : list[str]
      domain_counts : list[int]            — top 10 domains
    """
    empty = [], [], [], [], []

    if CREDENTIAL_BREACH_SHEET not in raw_sheets:
        print(f"\n  WARNING: Sheet '{CREDENTIAL_BREACH_SHEET}' not found — "
              f"credential breach analysis skipped.")
        return empty

    df = raw_sheets[CREDENTIAL_BREACH_SHEET].copy()
    print(f"\n  Credential Breach sheet loaded — {len(df)} row(s).")

    # ── Password strength ─────────────────────────────────────────────────────
    if COL_BREACH_PASSWORD not in df.columns:
        print(f"  WARNING: Column '{COL_BREACH_PASSWORD}' not found — "
              f"password analysis skipped.")
        pwd_labels, pwd_counts = [], []
    else:
        df["_pwd_class"] = df[COL_BREACH_PASSWORD].apply(_classify_password)
        vc = df["_pwd_class"].value_counts()
        order      = ["Strong", "Weak / No Policy", "No Password"]
        pwd_labels = [lbl for lbl in order if lbl in vc.index]
        pwd_counts = [int(vc[lbl]) for lbl in pwd_labels]

        print(f"\n  Password strength breakdown:")
        for lbl, cnt in zip(pwd_labels, pwd_counts):
            print(f"    {lbl}: {cnt}")

    # ── Top 20 emails — pandas-version-safe approach ──────────────────────────
    if COL_BREACH_EMAIL not in df.columns:
        print(f"  WARNING: Column '{COL_BREACH_EMAIL}' not found — "
              f"top-email table skipped.")
        top_emails    = []
        domain_labels = []
        domain_counts = []
    else:
        email_series = (
            df[COL_BREACH_EMAIL]
            .astype(str).str.strip().str.lower()
        )
        # Drop blanks
        email_series = email_series[~email_series.isin(_BLANK_VALUES)]

        # Top 20 emails — use Series directly to avoid pandas 2.x rename issues
        email_vc   = email_series.value_counts()
        top_emails = [(str(email), int(cnt))
                      for email, cnt in email_vc.head(20).items()]

        print(f"\n  Top {len(top_emails)} email(s) by occurrence:")
        for email, cnt in top_emails:
            print(f"    {email}: {cnt}")

        # ── Domain breakdown (new insight) ────────────────────────────────────
        # Extract domain from each email (right of @), count occurrences
        domains = email_series[email_series.str.contains("@", na=False)]
        domains = domains.str.split("@").str[-1]   # take part after last @
        domain_vc     = domains.value_counts()
        domain_labels = [str(d) for d in domain_vc.head(10).index]
        domain_counts = [int(c) for c in domain_vc.head(10).values]

        print(f"\n  Top {len(domain_labels)} email domain(s):")
        for d, c in zip(domain_labels, domain_counts):
            print(f"    {d}: {c}")

    return pwd_labels, pwd_counts, top_emails, domain_labels, domain_counts

# ---------------------------------------------------------------------------
# STEP 4 — BUILD OUTPUT WORKBOOK
# ---------------------------------------------------------------------------

def build_workbook(module_names, module_counts,
                   status_labels, status_counts,
                   user_names, user_counts,
                   pwd_labels, pwd_counts, top_emails,
                   domain_labels, domain_counts,
                   filtered_raw, counts, sheet_names,
                   start_dt, end_dt, output_path):

    wb = xlsxwriter.Workbook(output_path)

    # ── COLOUR CONSTANTS ─────────────────────────────────────────────────────
    NAVY    = "#1F3864"
    MIDBLUE = "#2E75B6"
    LTBLUE  = "#D6E4F7"
    ALT     = "#EFF5FB"
    WHITE   = "#FFFFFF"
    GREEN   = "#375623"    # dark green for credential breach section
    LGREEN  = "#70AD47"    # light green accent

    # ── FORMAT FACTORY ───────────────────────────────────────────────────────
    def _fmt(**kw):
        base = {"font_name": "Arial", "font_size": 10,
                "valign": "vcenter", "border": 1}
        base.update(kw)
        return wb.add_format(base)

    f_banner = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 15,
        "font_color": WHITE, "bg_color": NAVY,
        "align": "center", "valign": "vcenter",
    })
    f_banner_green = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 13,
        "font_color": WHITE, "bg_color": GREEN,
        "align": "center", "valign": "vcenter",
    })
    f_section = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 13,
        "font_color": NAVY, "valign": "vcenter",
    })
    f_section_green = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 13,
        "font_color": LGREEN, "valign": "vcenter",
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
    f_num     = _fmt(align="center")
    f_lft     = _fmt(align="left")
    f_num_alt = _fmt(align="center", bg_color=ALT)
    f_lft_alt = _fmt(align="left",   bg_color=ALT)
    f_total   = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 11,
        "font_color": NAVY, "bg_color": LTBLUE,
        "align": "center", "valign": "vcenter", "border": 2,
    })
    f_data_hdr = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 10,
        "font_color": WHITE, "bg_color": MIDBLUE,
        "align": "center", "valign": "vcenter", "border": 1,
    })
    f_cell     = _fmt(align="left")
    f_cell_alt = _fmt(align="left", bg_color=ALT)
    f_date     = _fmt(align="left", num_format="dd mmm yyyy")
    f_date_alt = _fmt(align="left", bg_color=ALT, num_format="dd mmm yyyy")

    # ── PIE COLOUR PALETTES ──────────────────────────────────────────────────
    _PALETTE = [
        "#4472C4", "#ED7D31", "#70AD47", "#FFC000", "#5B9BD5",
        "#A9D18E", "#FF7C80", "#9E480E", "#7030A0", "#636363",
        "#255E91", "#43682B", "#C00000", "#997300", "#7B5EA7",
    ]
    _PWD_PALETTE = {
        "Strong":           "#70AD47",
        "Weak / No Policy": "#ED7D31",
        "No Password":      "#A6A6A6",
    }

    def _pie_points(count):
        return [{"fill": {"color": _PALETTE[i % len(_PALETTE)]}}
                for i in range(count)]

    def _pie_points_named(labels, palette):
        return [{"fill": {"color": palette.get(lbl, "#4472C4")}}
                for lbl in labels]

    # ── LEGEND HELPER ────────────────────────────────────────────────────────
    def _write_legend(ws, start_row, start_col, labels, values, palette=None):
        """Colour-swatch legend: [■] | Name | Count | % Share"""
        total   = sum(values) or 1
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
            colour = (palette.get(lbl) if palette else _PALETTE[i % len(_PALETTE)])
            alt    = (i % 2 == 1)
            bg     = ALT if alt else WHITE

            ws.write(r, start_col,     "", wb.add_format({"bg_color": colour, "border": 1}))
            ws.write(r, start_col + 1, lbl, wb.add_format({
                "font_name": "Arial", "font_size": 9, "border": 1,
                "align": "left", "valign": "vcenter", "bg_color": bg}))
            ws.write(r, start_col + 2, val, wb.add_format({
                "font_name": "Arial", "font_size": 9, "border": 1,
                "align": "center", "valign": "vcenter", "bg_color": bg}))
            ws.write(r, start_col + 3, val / total, wb.add_format({
                "font_name": "Arial", "font_size": 9, "border": 1,
                "align": "center", "valign": "vcenter", "bg_color": bg,
                "num_format": "0.0%"}))

        ws.set_column(start_col,     start_col,     3)
        ws.set_column(start_col + 1, start_col + 1,
                      max((len(str(l)) for l in labels), default=12) + 4)
        ws.set_column(start_col + 2, start_col + 2, 10)
        ws.set_column(start_col + 3, start_col + 3, 10)

        # Return total rows consumed (header + data rows)
        return 1 + len(labels)

    def _make_pie(title, categories_ref, values_ref, points, size=None):
        """Create a styled pie chart object."""
        chart = wb.add_chart({"type": "pie"})
        chart.add_series({
            "name":       title,
            "categories": categories_ref,
            "values":     values_ref,
            "points":     points,
        })
        chart.set_title({"name": title})
        chart.set_legend({"none": True})
        chart.set_style(10)
        chart.set_size(size or {"width": 420, "height": 360})
        return chart

    PIE_ROWS = 19   # approximate row height a 360px chart occupies at default row height

    # =========================================================================
    # SUMMARY SHEET
    # =========================================================================
    n  = len(module_names)
    ns = len(status_labels)
    nu = len(user_names)
    np_ = len(pwd_labels)
    nd  = len(domain_labels)
    ne  = len(top_emails)

    HDR_ROW    = 3
    DATA_START = 4
    TOTAL_ROW  = DATA_START + n

    sw = wb.add_worksheet(SUMMARY_SHEET_NAME)

    # Banner
    sw.set_row(0, 42)
    sw.merge_range(
        0, 0, 0, 15,
        f"Incident Summary Dashboard  |  "
        f"{start_dt.strftime('%d %b %Y')}  ->  {end_dt.strftime('%d %b %Y')}",
        f_banner,
    )

    # ── Section 1: Module-wise table ─────────────────────────────────────────
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
    sw.write_formula(TOTAL_ROW, 2,
                     f"=SUM(C{DATA_START + 1}:C{DATA_START + n})", f_total)

    sw.set_column(0, 0, 6)
    sw.set_column(1, 1, max(max(len(m) for m in module_names) + 6, 24))
    sw.set_column(2, 2, 28)
    sw.freeze_panes(HDR_ROW, 0)

    # ── Charts row anchor ─────────────────────────────────────────────────────
    CHART_ROW = TOTAL_ROW + 2

    # Pie 1 — Closed By Module
    pie1 = _make_pie(
        "Closed Incidents By Module",
        [SUMMARY_SHEET_NAME, DATA_START, 1, DATA_START + n - 1, 1],
        [SUMMARY_SHEET_NAME, DATA_START, 2, DATA_START + n - 1, 2],
        _pie_points(n),
    )
    sw.insert_chart(CHART_ROW, 0, pie1)

    # Legend 1 (below pie1)
    L1_ROW      = CHART_ROW + PIE_ROWS
    L1_rows_used = _write_legend(sw, L1_ROW, 0, module_names, module_counts)

    # Pie 2 — Overall Incident Status
    if ns > 0:
        STATUS_SHEET = "Status Data"
        ss = wb.add_worksheet(STATUS_SHEET)
        for i, (lbl, cnt) in enumerate(zip(status_labels, status_counts)):
            ss.write(i, 0, lbl)
            ss.write(i, 1, cnt)

        pie2 = _make_pie(
            "Overall Incident Status (All Modules, Unique IDs)",
            [STATUS_SHEET, 0, 0, ns - 1, 0],
            [STATUS_SHEET, 0, 1, ns - 1, 1],
            _pie_points(ns),
        )
        sw.insert_chart(CHART_ROW, 7, pie2)
        L2_rows_used = _write_legend(sw, L1_ROW, 7, status_labels, status_counts)
    else:
        L2_rows_used = 0

    # ── Section 4: User-wise table + pie ─────────────────────────────────────
    # Position below the taller of the two legend tables
    legend_rows_max  = max(L1_rows_used, L2_rows_used)
    USER_TABLE_START = L1_ROW + legend_rows_max + 2
    USER_HDR_ROW     = USER_TABLE_START + 1
    USER_DATA_START  = USER_HDR_ROW + 1
    USER_TOTAL_ROW   = USER_DATA_START + nu       # row after last user data row

    if nu > 0:
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

        # User data helper sheet for pie
        USER_SHEET = "User Data"
        us = wb.add_worksheet(USER_SHEET)
        for i, (uname, ucnt) in enumerate(zip(user_names, user_counts)):
            us.write(i, 0, uname)
            us.write(i, 1, ucnt)

        pie3 = _make_pie(
            "Incidents Closed By User (Date Range)",
            [USER_SHEET, 0, 0, nu - 1, 0],
            [USER_SHEET, 0, 1, nu - 1, 1],
            _pie_points(nu),
        )
        USER_PIE_ROW = USER_TOTAL_ROW + 2
        sw.insert_chart(USER_PIE_ROW, 0, pie3)

        USER_LEGEND_ROW  = USER_PIE_ROW + PIE_ROWS
        user_legend_used = _write_legend(sw, USER_LEGEND_ROW, 0,
                                         user_names, user_counts)

        # Credential breach section anchors below user legend
        CRED_ANCHOR_ROW = USER_LEGEND_ROW + user_legend_used + 3
    else:
        # No user data — credential breach goes below legend tables
        CRED_ANCHOR_ROW = L1_ROW + legend_rows_max + 3

    # =========================================================================
    # CREDENTIAL BREACH SECTION
    # =========================================================================
    has_cred = (np_ > 0 or ne > 0 or nd > 0)

    if has_cred:
        # Section banner
        sw.set_row(CRED_ANCHOR_ROW, 36)
        sw.merge_range(
            CRED_ANCHOR_ROW, 0, CRED_ANCHOR_ROW, 15,
            f"Credential Breach Analysis  —  {CREDENTIAL_BREACH_SHEET}",
            f_banner_green,
        )

        # ── Password strength table ───────────────────────────────────────────
        if np_ > 0:
            PWD_TBL_ROW    = CRED_ANCHOR_ROW + 2
            PWD_HDR_ROW    = PWD_TBL_ROW + 1
            PWD_DATA_START = PWD_HDR_ROW + 1
            PWD_TOTAL_ROW  = PWD_DATA_START + np_

            sw.set_row(PWD_TBL_ROW, 24)
            sw.write(PWD_TBL_ROW, 0, "Password Strength vs. Policy", f_section_green)

            sw.set_row(PWD_HDR_ROW, 22)
            sw.write(PWD_HDR_ROW, 0, "#",          f_col_hdr_green)
            sw.write(PWD_HDR_ROW, 1, "Category",   f_col_hdr_green)
            sw.write(PWD_HDR_ROW, 2, "Count",      f_col_hdr_green)

            for i in range(np_):
                r   = PWD_DATA_START + i
                alt = (i % 2 == 1)
                sw.set_row(r, 18)
                sw.write(r, 0, i + 1,           f_num_alt if alt else f_num)
                sw.write(r, 1, pwd_labels[i],   f_lft_alt if alt else f_lft)
                sw.write(r, 2, pwd_counts[i],   f_num_alt if alt else f_num)

            sw.set_row(PWD_TOTAL_ROW, 22)
            sw.merge_range(PWD_TOTAL_ROW, 0, PWD_TOTAL_ROW, 1, "TOTAL", f_total)
            sw.write_formula(
                PWD_TOTAL_ROW, 2,
                f"=SUM(C{PWD_DATA_START + 1}:C{PWD_DATA_START + np_})",
                f_total,
            )

            # Helper sheet for pie4
            PWD_SHEET = "Pwd Strength Data"
            ps = wb.add_worksheet(PWD_SHEET)
            for i, (lbl, cnt) in enumerate(zip(pwd_labels, pwd_counts)):
                ps.write(i, 0, lbl)
                ps.write(i, 1, cnt)

            pie4 = _make_pie(
                "Credential Breach — Password Strength vs. Policy",
                [PWD_SHEET, 0, 0, np_ - 1, 0],
                [PWD_SHEET, 0, 1, np_ - 1, 1],
                _pie_points_named(pwd_labels, _PWD_PALETTE),
            )
            PWD_PIE_ROW = PWD_TOTAL_ROW + 2
            sw.insert_chart(PWD_PIE_ROW, 0, pie4)

            PWD_LEGEND_ROW  = PWD_PIE_ROW + PIE_ROWS
            pwd_legend_used = _write_legend(sw, PWD_LEGEND_ROW, 0,
                                            pwd_labels, pwd_counts,
                                            palette=_PWD_PALETTE)

            # Domain pie anchors to right of password pie
            DOMAIN_PIE_COL_OFFSET = 7
        else:
            PWD_TBL_ROW     = CRED_ANCHOR_ROW + 2
            PWD_PIE_ROW     = PWD_TBL_ROW
            PWD_LEGEND_ROW  = PWD_TBL_ROW
            pwd_legend_used = 0
            DOMAIN_PIE_COL_OFFSET = 7

        # ── Domain breakdown table + pie (new insight) ────────────────────────
        if nd > 0:
            DOM_TBL_ROW    = CRED_ANCHOR_ROW + 2
            DOM_HDR_ROW    = DOM_TBL_ROW + 1
            DOM_DATA_START = DOM_HDR_ROW + 1

            sw.set_row(DOM_TBL_ROW, 24)
            sw.write(DOM_TBL_ROW, DOMAIN_PIE_COL_OFFSET,
                     f"Top {nd} Email Domains by Breach Count", f_section_green)

            sw.set_row(DOM_HDR_ROW, 22)
            sw.write(DOM_HDR_ROW, DOMAIN_PIE_COL_OFFSET,     "#",           f_col_hdr_green)
            sw.write(DOM_HDR_ROW, DOMAIN_PIE_COL_OFFSET + 1, "Domain",      f_col_hdr_green)
            sw.write(DOM_HDR_ROW, DOMAIN_PIE_COL_OFFSET + 2, "Breach Count", f_col_hdr_green)

            max_dom_len = max((len(d) for d in domain_labels), default=20)
            sw.set_column(DOMAIN_PIE_COL_OFFSET,     DOMAIN_PIE_COL_OFFSET,     6)
            sw.set_column(DOMAIN_PIE_COL_OFFSET + 1, DOMAIN_PIE_COL_OFFSET + 1, max(max_dom_len + 4, 28))
            sw.set_column(DOMAIN_PIE_COL_OFFSET + 2, DOMAIN_PIE_COL_OFFSET + 2, 16)

            for i, (dom, cnt) in enumerate(zip(domain_labels, domain_counts)):
                r   = DOM_DATA_START + i
                alt = (i % 2 == 1)
                sw.set_row(r, 18)
                sw.write(r, DOMAIN_PIE_COL_OFFSET,     i + 1, f_num_alt if alt else f_num)
                sw.write(r, DOMAIN_PIE_COL_OFFSET + 1, dom,   f_lft_alt if alt else f_lft)
                sw.write(r, DOMAIN_PIE_COL_OFFSET + 2, cnt,   f_num_alt if alt else f_num)

            DOM_TOTAL_ROW = DOM_DATA_START + nd
            sw.set_row(DOM_TOTAL_ROW, 22)
            sw.merge_range(DOM_TOTAL_ROW, DOMAIN_PIE_COL_OFFSET,
                           DOM_TOTAL_ROW, DOMAIN_PIE_COL_OFFSET + 1, "TOTAL", f_total)
            sw.write_formula(
                DOM_TOTAL_ROW, DOMAIN_PIE_COL_OFFSET + 2,
                f"=SUM({chr(65 + DOMAIN_PIE_COL_OFFSET + 2)}{DOM_DATA_START + 1}"
                f":{chr(65 + DOMAIN_PIE_COL_OFFSET + 2)}{DOM_DATA_START + nd})",
                f_total,
            )

            DOM_SHEET = "Domain Data"
            ds = wb.add_worksheet(DOM_SHEET)
            for i, (dom, cnt) in enumerate(zip(domain_labels, domain_counts)):
                ds.write(i, 0, dom)
                ds.write(i, 1, cnt)

            pie5 = _make_pie(
                "Credential Breach — Top Email Domains",
                [DOM_SHEET, 0, 0, nd - 1, 0],
                [DOM_SHEET, 0, 1, nd - 1, 1],
                _pie_points(nd),
            )
            DOM_PIE_ROW = DOM_TOTAL_ROW + 2
            sw.insert_chart(DOM_PIE_ROW, DOMAIN_PIE_COL_OFFSET, pie5)

            DOM_LEGEND_ROW = DOM_PIE_ROW + PIE_ROWS
            _write_legend(sw, DOM_LEGEND_ROW, DOMAIN_PIE_COL_OFFSET,
                          domain_labels, domain_counts)

        # ── Top 20 email addresses table ──────────────────────────────────────
        if ne > 0:
            # Place below the pwd + domain section
            # Anchor row = max bottom of pwd legend vs domain legend
            if np_ > 0:
                pwd_bottom = PWD_LEGEND_ROW + pwd_legend_used + 2
            else:
                pwd_bottom = CRED_ANCHOR_ROW + 4

            if nd > 0:
                dom_bottom = DOM_LEGEND_ROW + nd + 1 + 3
            else:
                dom_bottom = CRED_ANCHOR_ROW + 4

            EMAIL_TBL_ROW    = max(pwd_bottom, dom_bottom)
            EMAIL_HDR_ROW    = EMAIL_TBL_ROW + 1
            EMAIL_DATA_START = EMAIL_HDR_ROW + 1

            sw.set_row(EMAIL_TBL_ROW, 24)
            sw.write(EMAIL_TBL_ROW, 0,
                     f"Top {ne} Email Addresses by Breach Count", f_section_green)

            sw.set_row(EMAIL_HDR_ROW, 22)
            sw.write(EMAIL_HDR_ROW, 0, "#",              f_col_hdr_green)
            sw.write(EMAIL_HDR_ROW, 1, "Email Address",  f_col_hdr_green)
            sw.write(EMAIL_HDR_ROW, 2, "Breach Count",   f_col_hdr_green)

            max_email_len = max((len(str(e)) for e, _ in top_emails), default=20)
            sw.set_column(0, 0, max(sw.col_size_changed(0), 6))
            sw.set_column(1, 1, max(max_email_len + 4, 36))
            sw.set_column(2, 2, max(16, 16))

            for i, (email, cnt) in enumerate(top_emails):
                r   = EMAIL_DATA_START + i
                alt = (i % 2 == 1)
                sw.set_row(r, 18)
                sw.write(r, 0, i + 1,  f_num_alt if alt else f_num)
                sw.write(r, 1, email,  f_lft_alt if alt else f_lft)
                sw.write(r, 2, cnt,    f_num_alt if alt else f_num)

    # =========================================================================
    # DATA SHEETS — one per incident module that has data in range
    # =========================================================================
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
        date_col_idx = (headers.index(COL_CLOSURE_DATE)
                        if COL_CLOSURE_DATE in headers else -1)

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
    print("  INCIDENT REPORT GENERATOR  (with Credential Breach Analysis)")
    print("=" * 60)
    print("\nEnter the date range for filtering closed incidents.")

    start_dt = _prompt_date("START date (inclusive)")
    end_dt   = _prompt_date("END   date (inclusive)")

    if end_dt < start_dt:
        start_dt, end_dt = end_dt, start_dt
        print("  (Dates swapped — start was after end.)")

    end_dt = end_dt.replace(hour=23, minute=59, second=59)
    print(f"\n  Range: {start_dt.strftime('%d %b %Y')} -> {end_dt.strftime('%d %b %Y')}\n")

    # Step 1 — load all sheets once
    raw_sheets  = load_all_sheets(INPUT_FILE_PATH)
    sheet_names = list(raw_sheets.keys())

    # Separate incident sheets from the credential breach sheet
    incident_sheet_names = [s for s in sheet_names if s != CREDENTIAL_BREACH_SHEET]

    print(f"Sheets found ({len(sheet_names)}): {sheet_names}")
    print(f"Incident sheets ({len(incident_sheet_names)}): {incident_sheet_names}\n")

    # Step 2 — diagnostic: closed incidents without closure dates
    print("  DIAGNOSTIC — checking for closed incidents with missing closure dates:")
    for name in incident_sheet_names:
        df = raw_sheets.get(name, pd.DataFrame())
        if df.empty or COL_INCIDENT_ID not in df.columns:
            continue
        if COL_STATUS not in df.columns or COL_CLOSURE_DATE not in df.columns:
            continue
        closed_df  = df[df[COL_STATUS].astype(str).str.lower().str.startswith("closed")]
        blank_date = closed_df[
            closed_df[COL_CLOSURE_DATE].astype(str).str.strip().isin(
                ["", "nan", "NaT", "None", "N/A", "-"]
            )
        ]
        if blank_date.empty:
            print(f"    [{name}]  all closed incidents have a closure date  ✓")
        else:
            print(f"    [{name}]  {len(blank_date)} closed incident(s) have NO closure date "
                  f"— will not appear in module tally!")
    print()

    # Step 3 — process each incident sheet
    processed_raw = {}
    filtered_raw  = {}
    counts        = {}

    for name in incident_sheet_names:
        raw_df, filtered_df = process_sheet(raw_sheets[name], start_dt, end_dt)
        processed_raw[name] = raw_df
        filtered_raw[name]  = filtered_df
        counts[name]        = len(filtered_df)
        status_str = (f"{counts[name]} unique incident(s) in range"
                      if counts[name] else "no incidents in range")
        print(f"  [{name}]  total rows = {len(raw_df)}  |  {status_str}")

        # Warn about unparseable dates
        if COL_CLOSURE_DATE in raw_df.columns:
            orig_strings = raw_sheets[name][COL_CLOSURE_DATE].astype(str).str.strip()
            nat_rows     = raw_df[raw_df[COL_CLOSURE_DATE].isna()]
            truly_bad    = nat_rows[
                ~orig_strings.loc[nat_rows.index].str.lower().isin(_BLANK_VALUES)
            ]
            if not truly_bad.empty:
                samples = orig_strings.loc[truly_bad.index].unique()[:5]
                print(f"    ⚠  {len(truly_bad)} row(s) had unparseable closure dates "
                      f"— excluded from count!")
                print(f"    ⚠  Sample date strings that failed: {list(samples)}")

    # Aggregate
    module_names, module_counts = aggregate(counts)
    grand_total = sum(module_counts)
    skipped     = len(incident_sheet_names) - len(module_names)

    print(f"\n  Modules with incidents : {len(module_names)}")
    if skipped:
        print(f"  Modules excluded (zero): {skipped}  (not written to output)")
    print(f"  Grand total            : {grand_total}")

    # Status breakdown (all incident data, unique IDs)
    status_labels, status_counts = compute_status_breakdown(
        processed_raw, incident_sheet_names
    )

    # User-wise breakdown (date-range filtered)
    user_names, user_counts = compute_user_breakdown(
        filtered_raw, incident_sheet_names
    )

    # Credential breach analysis (reads raw sheet, no date filter)
    print(f"\n{'=' * 60}")
    print(f"  CREDENTIAL BREACH ANALYSIS")
    print(f"{'=' * 60}")
    pwd_labels, pwd_counts, top_emails, domain_labels, domain_counts = \
        compute_credential_breach_analysis(raw_sheets)

    # Build output path
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    date_range_str = (
        f"{start_dt.strftime('%d %b %Y')} - {end_dt.strftime('%d %b %Y')}"
    )
    output_path = os.path.join(
        OUTPUT_FOLDER,
        f"CloudSek Incident Review - {date_range_str}.xlsx"
    )

    # Write workbook
    build_workbook(
        module_names, module_counts,
        status_labels, status_counts,
        user_names, user_counts,
        pwd_labels, pwd_counts, top_emails,
        domain_labels, domain_counts,
        filtered_raw, counts, incident_sheet_names,
        start_dt, end_dt,
        output_path,
    )

    print(f"\n  Report saved -> {output_path}")
    print("Done.\n")


if __name__ == "__main__":
    main()
