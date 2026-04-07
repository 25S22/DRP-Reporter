"""
=============================================================================
CREDENTIAL BREACH ANALYSIS (Standalone with Optional Date Filter)
=============================================================================
CONFIGURATION — only edit the lines in the CONFIGURATION block below.
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

SUMMARY_SHEET_NAME = "Credential Dashboard"
_BLANK_VALUES = {"", "nan", "none", "n/a", "na", "-", "null", "nat"}

# Password Policy: Uppercase, Lowercase, Number, Special Character AND min length 9
_STRONG_PASSWORD_RE = re.compile(
    r"^(?=.*[A-Z])(?=.*[a-z])(?=.*\d)(?=.*[^A-Za-z0-9]).{9,}$"
)

_ORDINAL_RE = re.compile(r"(\d+)(st|nd|rd|th)\b", re.IGNORECASE)
_DATE_FORMATS = [
    "%d %b, %Y %I:%M:%S %p", "%d %B, %Y %I:%M:%S %p",
    "%d %b, %Y %H:%M:%S",    "%d %B, %Y %H:%M:%S",
    "%d %b, %Y %I:%M %p",    "%d %b, %Y", "%d %B, %Y",
    "%d %b %Y", "%d %B %Y",  "%d-%b-%Y",  "%d-%B-%Y",
    "%d/%m/%Y", "%m/%d/%Y",  "%Y-%m-%d",  "%d.%m.%Y",
    "%d %b %y", "%d %B %y",
]

# ---------------------------------------------------------------------------
# DATE PARSING HELPER FUNCTIONS
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
        raw = input(f"  Enter {label} (e.g. 1 Jan 2024 / YYYY-MM-DD) [Leave blank to skip]: ").strip()
        if not raw:
            return None
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
    
    # Load all sheets as strings to preserve exact formatting
    return pd.read_excel(path, sheet_name=None, dtype=str)

# ---------------------------------------------------------------------------
# STEP 2 — CREDENTIAL BREACH ANALYSIS
# ---------------------------------------------------------------------------

def _classify_password(val):
    if pd.isna(val):
        return "No Password"
    s = str(val).strip()
    if s.lower() in _BLANK_VALUES:
        return "No Password"
    return "Strong" if _STRONG_PASSWORD_RE.match(s) else "Weak / No Policy"

def compute_credential_breach_analysis(raw_sheets, start_dt, end_dt):
    empty = [], [], [], [], []

    if CREDENTIAL_BREACH_SHEET not in raw_sheets:
        print(f"\n  WARNING: Sheet '{CREDENTIAL_BREACH_SHEET}' not found — "
              f"credential breach analysis skipped.")
        return empty

    df = raw_sheets[CREDENTIAL_BREACH_SHEET].copy()
    print(f"\n  Credential Breach sheet loaded — {len(df)} row(s).")

    # ── Optional Date Filtering (Auto-detects date column) ────────────────────
    if start_dt and end_dt:
        # Look for a likely date column since it's not strictly in the config
        possible_date_cols = ["Date", COL_CLOSURE_DATE, "Breach Date", "Timestamp", "Created On"]
        target_date_col = None
        
        for col in possible_date_cols:
            if col in df.columns:
                target_date_col = col
                break
                
        if not target_date_col:
            print("  WARNING: Could not find a recognizable date column in the Credential sheet.")
            print("  Cannot filter by date. Processing entire sheet.")
        else:
            print(f"  Filtering dates using column: '{target_date_col}'")
            df[target_date_col] = _parse_date_series(df[target_date_col])
            mask = (df[target_date_col] >= start_dt) & (df[target_date_col] <= end_dt)
            df = df[mask].copy()
            print(f"  Date filter applied. {len(df)} row(s) remain in range.")
            
            if df.empty:
                print("  WARNING: No records found in the specified date range.")
                return empty

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

    # ── Top 20 emails ─────────────────────────────────────────────────────────
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
        email_series = email_series[~email_series.isin(_BLANK_VALUES)]

        email_vc   = email_series.value_counts()
        top_emails = [(str(email), int(cnt))
                      for email, cnt in email_vc.head(20).items()]

        print(f"\n  Top {len(top_emails)} email(s) by occurrence:")
        for email, cnt in top_emails:
            print(f"    {email}: {cnt}")

        # ── Domain breakdown ──────────────────────────────────────────────────
        domains = email_series[email_series.str.contains("@", na=False)]
        domains = domains.str.split("@").str[-1]   
        domain_vc     = domains.value_counts()
        domain_labels = [str(d) for d in domain_vc.head(10).index]
        domain_counts = [int(c) for c in domain_vc.head(10).values]

        print(f"\n  Top {len(domain_labels)} email domain(s):")
        for d, c in zip(domain_labels, domain_counts):
            print(f"    {d}: {c}")

    return pwd_labels, pwd_counts, top_emails, domain_labels, domain_counts

# ---------------------------------------------------------------------------
# STEP 3 — BUILD OUTPUT WORKBOOK
# ---------------------------------------------------------------------------

def build_workbook(pwd_labels, pwd_counts, top_emails,
                   domain_labels, domain_counts, output_path, title_suffix=""):

    wb = xlsxwriter.Workbook(output_path)

    # ── COLOUR CONSTANTS ─────────────────────────────────────────────────────
    NAVY    = "#1F3864"
    LTBLUE  = "#D6E4F7"
    ALT     = "#EFF5FB"
    WHITE   = "#FFFFFF"
    GREEN   = "#375623"
    LGREEN  = "#70AD47"

    # ── FORMAT FACTORY ───────────────────────────────────────────────────────
    def _fmt(**kw):
        base = {"font_name": "Arial", "font_size": 10,
                "valign": "vcenter", "border": 1}
        base.update(kw)
        return wb.add_format(base)

    f_banner_green = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 15,
        "font_color": WHITE, "bg_color": GREEN,
        "align": "center", "valign": "vcenter",
    })
    f_section_green = wb.add_format({
        "bold": True, "font_name": "Arial", "font_size": 13,
        "font_color": LGREEN, "valign": "vcenter",
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

    # ── PIE COLOUR PALETTES ──────────────────────────────────────────────────
    _PALETTE = [
        "#4472C4", "#ED7D31", "#70AD47", "#FFC000", "#5B9BD5",
        "#A9D18E", "#FF7C80", "#9E480E", "#7030A0", "#636363",
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

        return 1 + len(labels)

    def _make_pie(title, categories_ref, values_ref, points, size=None):
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

    PIE_ROWS = 19
    sw = wb.add_worksheet(SUMMARY_SHEET_NAME)

    np_ = len(pwd_labels)
    nd  = len(domain_labels)
    ne  = len(top_emails)
    
    CRED_ANCHOR_ROW = 1 

    has_cred = (np_ > 0 or ne > 0 or nd > 0)

    if has_cred:
        sw.set_row(CRED_ANCHOR_ROW, 42)
        sw.merge_range(
            CRED_ANCHOR_ROW, 0, CRED_ANCHOR_ROW, 15,
            f"Credential Breach Analysis {title_suffix} —  {CREDENTIAL_BREACH_SHEET}",
            f_banner_green,
        )

        # ── Password strength table ───────────────────────────────────────────
        if np_ > 0:
            PWD_TBL_ROW    = CRED_ANCHOR_ROW + 3
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
            DOMAIN_PIE_COL_OFFSET = 7
        else:
            PWD_TBL_ROW     = CRED_ANCHOR_ROW + 3
            PWD_LEGEND_ROW  = PWD_TBL_ROW
            pwd_legend_used = 0
            DOMAIN_PIE_COL_OFFSET = 7

        # ── Domain breakdown table + pie ──────────────────────────────────────
        if nd > 0:
            DOM_TBL_ROW    = CRED_ANCHOR_ROW + 3
            DOM_HDR_ROW    = DOM_TBL_ROW + 1
            DOM_DATA_START = DOM_HDR_ROW + 1

            sw.set_row(DOM_TBL_ROW, 24)
            sw.write(DOM_TBL_ROW, DOMAIN_PIE_COL_OFFSET,
                     f"Top {nd} Email Domains by Breach Count", f_section_green)

            sw.set_row(DOM_HDR_ROW, 22)
            sw.write(DOM_HDR_ROW, DOMAIN_PIE_COL_OFFSET,     "#",            f_col_hdr_green)
            sw.write(DOM_HDR_ROW, DOMAIN_PIE_COL_OFFSET + 1, "Domain",       f_col_hdr_green)
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
            sw.write(EMAIL_HDR_ROW, 0, "#",               f_col_hdr_green)
            sw.write(EMAIL_HDR_ROW, 1, "Email Address",   f_col_hdr_green)
            sw.write(EMAIL_HDR_ROW, 2, "Breach Count",    f_col_hdr_green)

            max_email_len = max((len(str(e)) for e, _ in top_emails), default=20)
            
            # ---> THE FIX IS APPLIED ON THIS LINE <---
            sw.set_column(0, 0, 6)
            
            sw.set_column(1, 1, max(max_email_len + 4, 36))
            sw.set_column(2, 2, max(16, 16))

            for i, (email, cnt) in enumerate(top_emails):
                r   = EMAIL_DATA_START + i
                alt = (i % 2 == 1)
                sw.set_row(r, 18)
                sw.write(r, 0, i + 1,  f_num_alt if alt else f_num)
                sw.write(r, 1, email,  f_lft_alt if alt else f_lft)
                sw.write(r, 2, cnt,    f_num_alt if alt else f_num)

    wb.close()

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    print("\n" + "=" * 60)
    print("  CREDENTIAL BREACH ANALYSIS")
    print("=" * 60)

    print("\nSelect a date range for filtering. To process the ENTIRE sheet, just press ENTER.")
    start_dt = _prompt_date("START date (inclusive)")
    
    if start_dt:
        end_dt = _prompt_date("END   date (inclusive)")
        if not end_dt:
            end_dt = pd.Timestamp.now()
            print(f"  No end date provided. Defaulting to today: {end_dt.strftime('%d %b %Y')}")
            
        if end_dt < start_dt:
            start_dt, end_dt = end_dt, start_dt
            print("  (Dates swapped — start was after end.)")
            
        end_dt = end_dt.replace(hour=23, minute=59, second=59)
        title_suffix = f"({start_dt.strftime('%d %b %Y')} - {end_dt.strftime('%d %b %Y')})"
    else:
        end_dt = None
        title_suffix = "(All Dates)"
        print("  Processing entire sheet...")

    # Step 1 — load sheets
    raw_sheets = load_all_sheets(INPUT_FILE_PATH)
    
    # Step 2 — compute metrics
    pwd_labels, pwd_counts, top_emails, domain_labels, domain_counts = \
        compute_credential_breach_analysis(raw_sheets, start_dt, end_dt)

    if not pwd_labels and not top_emails and not domain_labels:
        print("\n  No data found or sheet empty. Exiting.")
        return

    # Step 3 — write output
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    if start_dt and end_dt:
        file_suffix = f"{start_dt.strftime('%Y%m%d')}_to_{end_dt.strftime('%Y%m%d')}"
    else:
        file_suffix = "All_Dates"
        
    output_path = os.path.join(OUTPUT_FOLDER, f"Standalone Credential Review - {file_suffix}.xlsx")

    build_workbook(
        pwd_labels, pwd_counts, 
        top_emails, 
        domain_labels, domain_counts,
        output_path,
        title_suffix
    )

    print(f"\n  Report saved -> {output_path}")
    print("Done.\n")

if __name__ == "__main__":
    main()
