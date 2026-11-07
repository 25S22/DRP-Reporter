#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
===============================================================================
 INCIDENT REPORT GENERATOR  ·  v3.2  (Excel dashboard always · PowerPoint on request)
===============================================================================

WHAT CHANGED vs v3.1
-------------------------------------------------------------------------------
 1. POWERPOINT IS NOW OPT-IN
      The Excel dashboard is always produced. The one-slide PowerPoint is only
      produced when you say yes:
        · interactive run  -> after the dates you are asked "Create the
                              PowerPoint as well? (y/N)"  (Enter = no)
        · command line     -> add  --pptx  to say yes  (--no-pptx forces no)
        · scheduler / no keyboard attached -> no PowerPoint unless --pptx

 2. "ONE EXTRA INCIDENT" HUNT
      The total is now built from one frame and its arithmetic is printed and
      shown on the dashboard:   created in period + closed-from-earlier = total.
      Likely causes of a stray +1 are now handled or made visible:
        · footer / total / stray rows that carry a date but NO incident id are
          excluded and listed on the Data Quality tab (DROP_ROWS_WITHOUT_ID)
        · an id present in two modules is flagged on the dashboard
        · NEW "Count Audit" tab lists every single row behind the total, why it
          was counted, and whether the same id sits in another module - so the
          extra row can be spotted in seconds
      Also fixed: Data Quality "Excel Row" for unparseable dates pointed at the
      combined index instead of the real row in the source sheet.

 3. CLEARER EXCEL DASHBOARD
      "Incidents in period" now spells out its own arithmetic
      (X created + Y closed from earlier), a one-line definitions strip sits
      under the KPI cards, and only the title rows are frozen (the old freeze
      took a third of a laptop screen).

 4. POWERPOINT REDESIGN  (all charts / tables are native + editable)
      · line chart (created vs closed, rolling months) instead of bars
      · coloured incident breakdown table by module (same colours as Excel)
      · coloured status breakdown with doughnut
      · bottom-left panel left completely blank for the monthly highlights
        (click it and type / paste)
      · the "attention required" / ageing / data-quality panel is removed;
        nothing on the slide is a warning

 5. STATUS MATCHING SAFETY FIX
      Keywords now match at the START of a word, so "Unassigned" is no longer
      read as "assigned" (-> In Progress), "Incomplete" no longer as "complete"
      (-> Closed), and "Unresolved"/"Not closed" are recognised as Open.

REQUIRED PACKAGES
-------------------------------------------------------------------------------
     pip install pandas xlsxwriter python-pptx python-dateutil

OPTIONAL
-------------------------------------------------------------------------------
     pip install dateparser      (natural-language / non-English dates only)

USAGE
-------------------------------------------------------------------------------
     python incident_report_generator.py                    (prompts for all)
     python incident_report_generator.py -i in.xlsx -s "1 Jan 2024" -e "31 Jan 2024"
     python incident_report_generator.py ... --pptx         (also build the slide)
===============================================================================
"""

from __future__ import annotations

import argparse
import math
import os
import re
import sys
from datetime import datetime

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name
from pandas.api.types import is_datetime64_any_dtype

import warnings
warnings.filterwarnings("ignore")

try:
    from dateutil import parser as _du_parser
    _HAS_DATEUTIL = True
except ImportError:                                    # pragma: no cover
    _HAS_DATEUTIL = False

try:
    import dateparser as _dateparser
    _HAS_DATEPARSER = True
except ImportError:                                    # pragma: no cover
    _HAS_DATEPARSER = False


# =============================================================================
# >>>  CONFIGURATION — only edit this block  <<<
# =============================================================================

INPUT_FILE_PATH   = "incidents.xlsx"
OUTPUT_FOLDER     = "reports"

COL_INCIDENT_ID   = "Incident Id"
COL_CREATION_DATE = "Created On"
COL_CLOSURE_DATE  = "Incident Closure on"
COL_STATUS        = "Status"
COL_CLOSED_BY     = "Incident Closed By"
COL_COMMENT       = "Comment"
COL_EVENT_COUNT   = "Event Count"      # "" to skip the Total Events KPI

# --- behaviour switches -------------------------------------------------------
DAYFIRST              = True    # True -> 05/01/2024 is 5 Jan. False -> 5 May.
DEDUPE_ACROSS_MODULES = False   # False: an ID may legitimately exist in two
                                # modules and is counted in each (module totals
                                # then equal the grand total). True: one global
                                # ID wins; duplicates are dropped and listed on
                                # the Data Quality sheet.
DROP_ROWS_WITHOUT_ID  = True    # rows with no incident id (footers, totals,
                                # stray notes) are not incidents: exclude them
                                # from every count and list them on Data Quality
EMAIL_SCOPE           = "period"  # "period" (in-range only) or "all"
STATUS_FOR_EMAILS     = "closed - resolved"
MAX_DQ_ROWS           = 500     # cap rows written to the Data Quality sheet

# --- outputs ------------------------------------------------------------------
PPTX_DEFAULT          = False   # used only when nobody can be asked (scheduler)
TREND_MONTHS          = 6       # rolling months ending in the report month
TOP_N_MODULES         = 4       # modules listed individually on the PPT table
AGED_DAYS             = 90      # Excel only: open/in-progress ageing threshold
PPT_NOTES_TITLE       = "MONTHLY HIGHLIGHTS"   # title above the blank panel
PPT_FONT              = "Calibri"

# --- status bucketing ---------------------------------------------------------
# Exact matches win first (compare lowercased + trimmed), then keyword rules in
# order. Keywords match from the START of a word (so "assigned" does not match
# "unassigned"). Anything unmatched lands in "Other" and is still counted.
STATUS_OVERRIDES: dict[str, str] = {
    # "closed - not an incident": "Other",
}

STATUS_RULES: list[tuple[str, list[str]]] = [
    ("Closed", [
        "closed", "resolved", "complete", "done", "fixed", "cancel",
        "rejected", "duplicate", "withdrawn", "no action",
    ]),
    ("In Progress", [
        "in progress", "inprogress", "in-progress", "wip", "work in progress",
        "working", "assigned", "investigat", "under review", "in review",
        "pending vendor", "pending customer", "pending user", "awaiting",
        "on hold", "onhold", "escalat", "reopen", "re-open", "in analysis",
    ]),
    ("Open", [
        "open", "new", "unassigned", "raised", "logged", "backlog",
        "pending", "queued", "to do", "todo", "not started", "active",
    ]),
]

STATUS_ORDER = ["Closed", "In Progress", "Open", "Other"]


# =============================================================================
# INTERNALS — palette
# =============================================================================

SUMMARY_SHEET_NAME = "Summary Dashboard"
DQ_SHEET_NAME      = "Data Quality"
AUDIT_SHEET_NAME   = "Count Audit"
CDSHEET            = "_ChartData"

# One neutral slate ramp for structure, ONE accent (indigo) for anything
# structural, and the RAG colours reserved strictly for status meaning.
INK       = "#0F172A"      # banner / darkest text
SLATE     = "#334155"
MUTED     = "#64748B"      # captions, axis text
LINE      = "#E2E8F0"      # hairline borders
LINE_SOFT = "#F1F5F9"      # zebra rows
CANVAS    = "#F8FAFC"      # page background bands
WHITE     = "#FFFFFF"
ACCENT    = "#4F46E5"      # indigo - section rules, totals, data bars
ACCENT_L  = "#EEF2FF"
ACCENT_D  = "#3730A3"

GREEN, GREEN_L, GREEN_D = "#16A34A", "#DCFCE7", "#15803D"
AMBER, AMBER_L, AMBER_D = "#D97706", "#FEF3C7", "#B45309"
RED,   RED_L,   RED_D   = "#DC2626", "#FEE2E2", "#B91C1C"
GREY,  GREY_L,  GREY_D  = "#94A3B8", "#F1F5F9", "#475569"

_STATUS_RAG  = {"Closed": GREEN,   "In Progress": AMBER,
                "Open": RED,       "Other": GREY}
_STATUS_TINT = {"Closed": GREEN_L, "In Progress": AMBER_L,
                "Open": RED_L,     "Other": GREY_L}
_STATUS_DARK = {"Closed": GREEN_D, "In Progress": AMBER_D,
                "Open": RED_D,     "Other": GREY_D}

# --- dashboard grid ----------------------------------------------------------
GUTTER   = 0          # narrow spacer column on the left
C0       = 1          # first content column
CN       = 20         # last content column
COL_W    = 8.5        # ~64 px per column
COL_PX   = 64
GUT_W    = 2

# =============================================================================
# COLUMN NAME RESOLUTION
# =============================================================================

def _norm_col(name) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(name).lower())


def _configured_columns() -> list[str]:
    return [c for c in (COL_INCIDENT_ID, COL_CREATION_DATE, COL_CLOSURE_DATE,
                        COL_STATUS, COL_CLOSED_BY, COL_COMMENT,
                        COL_EVENT_COUNT) if c]


def resolve_columns(df: pd.DataFrame, sheet: str, issues: list) -> pd.DataFrame:
    """Rename near-miss headers onto the configured names (case/space/punct)."""
    lookup = {_norm_col(c): c for c in _configured_columns()}
    rename, seen = {}, set()
    for col in df.columns:
        key = _norm_col(col)
        target = lookup.get(key)
        if target and col != target and target not in df.columns and target not in seen:
            rename[col] = target
            seen.add(target)
    if rename:
        for old, new in rename.items():
            issues.append({"Type": "Header auto-matched", "Module": sheet,
                           "Column": new, "Row": "", "Detail": f"'{old}' -> '{new}'"})
        df = df.rename(columns=rename)
    return df


# =============================================================================
# DATE ENGINE
# =============================================================================

_ORDINAL_RE  = re.compile(r"(\d+)\s*(st|nd|rd|th)\b", re.IGNORECASE)
_MULTISPACE  = re.compile(r"\s+")
_TRAIL_TZ_RE = re.compile(
    r"\s*(?:\(?\b(?:ist|utc|gmt|est|pst|cst|cet|bst|z|hrs|hours|hrs\.)\b\)?)+\s*$",
    re.IGNORECASE)
_PAREN_TZ_RE = re.compile(r"\s*\([A-Za-z][A-Za-z \.\+\-0-9]*\)\s*$")

_NULL_TOKENS = {
    "", "nan", "nat", "none", "null", "n/a", "na", "-", "--", "---", "?",
    "tbd", "not closed", "not applicable", "notclosed", "0", "nil", "blank",
    "#n/a", "#value!", "#ref!", "not available", "<na>",
}

_EPOCH_1900 = pd.Timestamp("1899-12-30")


def _build_formats(dayfirst: bool) -> list[str]:
    dmy = ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y"]
    mdy = ["%m/%d/%Y", "%m-%d-%Y", "%m.%d.%Y", "%m/%d/%y", "%m-%d-%y"]
    numeric = (dmy + mdy) if dayfirst else (mdy + dmy)

    named, base = [], ["%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d-%B-%Y",
                       "%b %d %Y", "%B %d %Y", "%d %b %y", "%d %B %y"]
    times = ["", " %I:%M:%S %p", " %H:%M:%S", " %I:%M %p", " %H:%M",
             " %H:%M:%S.%f"]
    for b in base:
        for t in times:
            named.append(b + t)

    iso = ["%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
           "%Y-%m-%d %H:%M:%S.%f", "%Y/%m/%d", "%Y/%m/%d %H:%M:%S",
           "%Y%m%d", "%Y-%m-%dT%H:%M:%S"]

    out = []
    for n in numeric:
        for t in times:
            out.append(n + t)
    return iso + named + out


_DATE_FORMATS = _build_formats(DAYFIRST)


def _clean_date_text(value) -> str:
    t = str(value)
    t = t.replace("\u00a0", " ").replace("\u200b", "")
    t = re.sub(r"[\u2010-\u2015\u2212]", "-", t)
    t = _ORDINAL_RE.sub(r"\1", t)
    t = re.sub(r"^(?:\d{4}-\d{2}-\d{2})T", lambda m: m.group(0)[:-1] + " ", t)
    t = re.sub(r"(?i)\ba\.?\s*m\.?\b", "AM", t)
    t = re.sub(r"(?i)\bp\.?\s*m\.?\b", "PM", t)
    t = _PAREN_TZ_RE.sub("", t)
    t = _TRAIL_TZ_RE.sub("", t)
    # strip a UTC offset ONLY when it follows a clock time (never the "-2026" of 07-10-2026)
    t = re.sub(r"(\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:\s*[AaPp][Mm])?)\s*[+-]\d{2}:?\d{2}$", r"\1", t)
    t = re.sub(r"(?<=\d)Z$", "", t)
    t = t.replace(",", " ")
    t = _MULTISPACE.sub(" ", t)
    return t.strip()


def _looks_like_date_text(t: str) -> bool:
    if len(t) < 5 or not re.search(r"\d", t):
        return False
    return bool(re.search(r"\d{1,4}[/\-. ]\d{1,2}", t) or
                re.search(r"[A-Za-z]{3}", t))


def _plausible(parsed: pd.Series) -> pd.Series:
    """True where a parsed date is real and in a sane range (guards against '07-10' -> year 1)."""
    ok = parsed.notna()
    if ok.any():
        ok = ok & parsed.dt.year.between(1990, 2100)
    return ok


def parse_date_series(series: pd.Series, *, module: str, column: str,
                      issues: list, src_rows: pd.Series | None = None) -> pd.Series:
    """Six-stage cascade. Returns tz-naive datetime64; logs true failures.

    src_rows (optional) maps the series index to the real Excel row number in
    the source sheet, so the Data Quality tab points at the right row.
    """
    if is_datetime64_any_dtype(series):
        out = pd.to_datetime(series, errors="coerce")
        try:
            out = out.dt.tz_localize(None)
        except (TypeError, AttributeError):
            pass
        return out

    out = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

    # stage 0 — genuine datetime objects
    mask = series.map(lambda v: isinstance(v, (datetime, pd.Timestamp)))
    if mask.any():
        out.loc[mask] = pd.to_datetime(series[mask], errors="coerce")

    # stage 1 — numeric serials / epochs
    todo = out.isna()
    if todo.any():
        num = pd.to_numeric(series.where(todo), errors="coerce")
        ok = num.notna() & todo
        if ok.any():
            v = num[ok]
            ser = v[(v > 1) & (v < 80000)]
            if len(ser):
                out.loc[ser.index] = _EPOCH_1900 + pd.to_timedelta(ser, unit="D")
            eps = v[(v >= 1e9) & (v < 4.1e9)]
            if len(eps):
                out.loc[eps.index] = pd.to_datetime(eps.astype("int64"), unit="s")
            epm = v[(v >= 1e12) & (v < 4.1e12)]
            if len(epm):
                out.loc[epm.index] = pd.to_datetime(epm.astype("int64"), unit="ms")

    # remaining -> text
    todo = out.isna()
    if not todo.any():
        return out
    txt = series[todo].astype(str).map(_clean_date_text)
    blank = txt.str.strip().str.lower().isin(_NULL_TOKENS)
    txt = txt[~blank]                       # blanks are legitimately empty
    if txt.empty:
        return out

    # stage 2 — exact formats, vectorised
    for fmt in _DATE_FORMATS:
        if txt.empty:
            break
        p = pd.to_datetime(txt, format=fmt, errors="coerce")
        good = _plausible(p)
        if good.any():
            out.loc[txt.index[good]] = p[good]
            txt = txt[~good]

    # stage 3 — pandas flexible
    if not txt.empty:
        p = pd.to_datetime(txt, errors="coerce", dayfirst=DAYFIRST)
        good = _plausible(p)
        if good.any():
            out.loc[txt.index[good]] = p[good]
            txt = txt[~good]

    # stage 4 — dateutil fuzzy
    if not txt.empty and _HAS_DATEUTIL:
        for idx, t in txt.items():
            if not _looks_like_date_text(t):
                continue
            try:
                d = _du_parser.parse(t, dayfirst=DAYFIRST, fuzzy=True)
                if 1990 <= d.year <= 2100:
                    out.loc[idx] = pd.Timestamp(d).tz_localize(None) \
                        if d.tzinfo else pd.Timestamp(d)
            except (ValueError, OverflowError, TypeError):
                pass
        txt = txt[out.loc[txt.index].isna()]

    # stage 5 — dateparser (optional)
    if not txt.empty and _HAS_DATEPARSER:
        settings = {"DATE_ORDER": "DMY" if DAYFIRST else "MDY",
                    "RETURN_AS_TIMEZONE_AWARE": False}
        for idx, t in txt.items():
            try:
                d = _dateparser.parse(t, settings=settings)
                if d and 1990 <= d.year <= 2100:
                    out.loc[idx] = pd.Timestamp(d)
            except Exception:
                pass
        txt = txt[out.loc[txt.index].isna()]

    # whatever is left is a genuine data-quality problem
    for idx, _ in txt.items():
        if src_rows is not None and idx in src_rows.index:
            row_ref = int(src_rows.loc[idx])
        else:
            row_ref = int(idx) + 2 if isinstance(idx, int) else str(idx)
        issues.append({
            "Type":   "Unparseable date",
            "Module": module,
            "Column": column,
            "Row":    row_ref,
            "Detail": str(series.loc[idx])[:200],
        })
    return out


def parse_single_date(raw: str):
    t = _clean_date_text(raw)
    for fmt in _DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(t, fmt))
        except ValueError:
            continue
    try:
        return pd.Timestamp(pd.to_datetime(t, dayfirst=DAYFIRST))
    except Exception:
        pass
    if _HAS_DATEUTIL:
        try:
            return pd.Timestamp(_du_parser.parse(t, dayfirst=DAYFIRST, fuzzy=True))
        except Exception:
            pass
    return None


def prompt_date(label: str, default: pd.Timestamp | None = None):
    hint = f" [{default:%d %b %Y}]" if default is not None else ""
    while True:
        raw = input(f"  {label}{hint}: ").strip()
        if not raw and default is not None:
            return default
        got = parse_single_date(raw)
        if got is not None:
            return got
        print(f"    Could not read '{raw}'. Try 1 Jan 2024 / 01-01-2024 / 2024-01-01.")


# =============================================================================
# STATUS BUCKETING
# =============================================================================

_unmapped_statuses: dict[str, int] = {}

# "not resolved", "unresolved", "non-closed", "not fixed" ... -> still open work
_NEGATED_CLOSED_RE = re.compile(
    r"(?<![a-z0-9])(?:not|non|un|never)[\s\-]*(?:yet[\s\-]*)?"
    r"(?:closed|resolved|fixed|complete|completed|done)")
_kw_cache: dict[str, re.Pattern] = {}


def _kw_regex(kw: str) -> re.Pattern:
    """Keyword anchored at the START of a word ('assigned' != 'unassigned')."""
    rx = _kw_cache.get(kw)
    if rx is None:
        rx = _kw_cache[kw] = re.compile(r"(?<![a-z0-9])" + re.escape(kw))
    return rx


def normalize_status(value) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "Other"
    try:
        if pd.isna(value):
            return "Other"
    except (TypeError, ValueError):
        pass
    s = _MULTISPACE.sub(" ", str(value).strip().lower())
    if not s or s in _NULL_TOKENS:
        return "Other"
    if s in STATUS_OVERRIDES:
        return STATUS_OVERRIDES[s]
    if _NEGATED_CLOSED_RE.search(s):
        return "Open"
    for bucket, keywords in STATUS_RULES:
        for kw in keywords:
            if _kw_regex(kw).search(s):
                return bucket
    _unmapped_statuses[str(value).strip()] = _unmapped_statuses.get(
        str(value).strip(), 0) + 1
    return "Other"



# =============================================================================
# LOAD  ->  ONE TIDY MASTER FRAME
# =============================================================================

def load_master(path: str, issues: list):
    if not os.path.exists(path):
        sys.exit(f"\nERROR: input file not found -> {path}\n")

    print(f"\nLoading: {path}")
    sheets = pd.read_excel(path, sheet_name=None, dtype=object)
    if not sheets:
        sys.exit("\nERROR: workbook contains no sheets.\n")

    frames, order = [], []
    for name, df in sheets.items():
        if df is None or df.empty:
            print(f"  [{name}] empty - skipped")
            continue
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        df = resolve_columns(df, name, issues)
        df = df.dropna(how="all")
        if df.empty:
            continue
        df["_Module"] = name
        df["_SrcRow"] = df.index + 2          # Excel row number, header = row 1
        frames.append(df)
        order.append(name)

    if not frames:
        sys.exit("\nERROR: no usable rows found in the workbook.\n")

    master = pd.concat(frames, ignore_index=True, sort=False)
    print(f"  Sheets used ({len(order)}): {order}")
    print(f"  Total raw rows: {len(master):,}")
    return master, order


# =============================================================================
# PREPARE — ids, dates, status, dedupe
# =============================================================================

def prepare(master: pd.DataFrame, modules: list[str], issues: list):
    # ---- normalised id (first, so non-incident rows can be removed up front)
    if COL_INCIDENT_ID and COL_INCIDENT_ID in master.columns:
        ids = master[COL_INCIDENT_ID].astype(str).str.strip().str.upper()
        ids = ids.where(~ids.isin(["", "NAN", "NONE", "NAT", "<NA>", "NULL"]), None)
        master["_Id"] = ids
    else:
        master["_Id"] = None
        print(f"  WARNING: column '{COL_INCIDENT_ID}' not found - dedup disabled.")

    # ---- rows with no incident id are not incidents (footers, totals, notes)
    if DROP_ROWS_WITHOUT_ID and COL_INCIDENT_ID in master.columns:
        no_id = master["_Id"].isna()
        if no_id.any():
            for _, row in master[no_id].head(MAX_DQ_ROWS).iterrows():
                bits = [f"{c}={str(row.get(c))[:40]}"
                        for c in (COL_CREATION_DATE, COL_STATUS)
                        if c and c in master.columns and pd.notna(row.get(c))]
                issues.append({"Type": "Row without incident ID (excluded)",
                               "Module": row["_Module"], "Column": COL_INCIDENT_ID,
                               "Row": row["_SrcRow"],
                               "Detail": " | ".join(bits) or "(no other values)"})
            print(f"  NOTE: {int(no_id.sum())} row(s) have no '{COL_INCIDENT_ID}' and were "
                  f"excluded from all counts (see the Data Quality tab).")
            master = master[~no_id].copy()

    # ---- dates
    for col, tag in ((COL_CREATION_DATE, "creation"), (COL_CLOSURE_DATE, "closure")):
        target = f"_{tag}"
        if col and col in master.columns:
            parts = []
            for mod, grp in master.groupby("_Module", sort=False):
                parts.append(parse_date_series(grp[col], module=mod, column=col,
                                               issues=issues, src_rows=grp["_SrcRow"]))
            master[target] = pd.concat(parts).reindex(master.index)
        else:
            master[target] = pd.NaT
            if col:
                print(f"  WARNING: column '{col}' not found in any sheet.")

    # ---- status
    if COL_STATUS and COL_STATUS in master.columns:
        master["_Status"] = master[COL_STATUS].map(normalize_status)
        blanks = master[COL_STATUS].astype(str).str.strip().str.lower().isin(_NULL_TOKENS)
        for _, row in master[blanks].head(MAX_DQ_ROWS).iterrows():
            issues.append({"Type": "Blank status", "Module": row["_Module"],
                           "Column": COL_STATUS, "Row": row["_SrcRow"],
                           "Detail": "counted under 'Other'"})
    else:
        master["_Status"] = "Other"
        print(f"  WARNING: column '{COL_STATUS}' not found - all rows -> 'Other'.")

    # ---- cross-module duplicates (diagnostic, always reported)
    has_id = master["_Id"].notna()
    if has_id.any():
        spread = (master[has_id].groupby("_Id")["_Module"].nunique())
        multi = spread[spread > 1]
        for _id in list(multi.index)[:MAX_DQ_ROWS]:
            mods = sorted(master.loc[master["_Id"] == _id, "_Module"].unique())
            issues.append({"Type": "ID in multiple modules", "Module": ", ".join(mods),
                           "Column": COL_INCIDENT_ID, "Row": "",
                           "Detail": f"{_id} appears in {len(mods)} modules"})
        if len(multi):
            print(f"  NOTE: {len(multi)} incident id(s) appear in more than one module "
                  f"(DEDUPE_ACROSS_MODULES={DEDUPE_ACROSS_MODULES}).")

    # ---- closed but no closure date
    closed_mask = master["_Status"].eq("Closed") & master["_closure"].isna()
    for _, row in master[closed_mask].head(MAX_DQ_ROWS).iterrows():
        issues.append({"Type": "Closed without closure date", "Module": row["_Module"],
                       "Column": COL_CLOSURE_DATE, "Row": row["_SrcRow"],
                       "Detail": str(row.get(COL_STATUS, ""))[:120]})
    if closed_mask.any():
        print(f"  NOTE: {int(closed_mask.sum())} closed incident(s) have no closure date "
              f"- they cannot appear in closure-date metrics.")
    return master


def dedupe(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or df["_Id"].isna().all():
        return df
    keys = ["_Id"] if DEDUPE_ACROSS_MODULES else ["_Module", "_Id"]
    with_id = df[df["_Id"].notna()].drop_duplicates(subset=keys, keep="first")
    without = df[df["_Id"].isna()]
    return pd.concat([with_id, without]).sort_index()


def _period_frames(master: pd.DataFrame, start_dt, end_dt):
    """ONE place that decides who is 'in the period'.

    union   = created in period OR closed in period (de-duplicated once)
    created = the rows of `union` created in the period
    closed  = the rows of `union` closed in the period
    Because created/closed are slices of union, the arithmetic always ties:
        len(union) == len(created) + (closed in period but created earlier)
    """
    in_created = master["_creation"].between(start_dt, end_dt)
    in_closed  = master["_closure"].between(start_dt, end_dt)
    union   = dedupe(master[in_created | in_closed].copy())
    created = union[union["_creation"].between(start_dt, end_dt)]
    closed  = union[union["_closure"].between(start_dt, end_dt)]
    return union, created, closed


def _mttr(closed: pd.DataFrame):
    """(mean, median, n_excluded) days to resolve; only rows with ordered dates."""
    if closed is None or closed.empty:
        return None, None, 0
    valid = closed.dropna(subset=["_creation", "_closure"])
    if valid.empty:
        return None, None, 0
    days = (valid["_closure"] - valid["_creation"]).dt.total_seconds() / 86400.0
    ok = days[days >= 0]
    bad = len(valid) - len(ok)
    if ok.empty:
        return None, None, bad
    return float(ok.mean()), float(ok.median()), bad


def _period_snapshot(master: pd.DataFrame, start_dt, end_dt) -> dict:
    """Side-effect-free summary used to compare the period with its predecessor."""
    union, created, closed = _period_frames(master, start_dt, end_dt)
    total = len(union)
    n_closed_status = int((union["_Status"] == "Closed").sum()) if total else 0
    mean, med, _ = _mttr(closed)
    return {
        "total": total, "n_created": len(created), "n_closed": len(closed),
        "n_closed_status": n_closed_status,
        "resolution_rate": (n_closed_status / total * 100.0) if total else None,
        "net_backlog_change": len(created) - len(closed),
        "mttr_mean": mean, "mttr_med": med,
    }


def _prior_period_window(start_dt, end_dt):
    """Previous full calendar month, or the same-length period immediately before this one."""
    start = pd.Timestamp(start_dt).normalize()
    end = pd.Timestamp(end_dt).normalize()
    is_month = start.day == 1 and end == (start + pd.offsets.MonthEnd(0)).normalize()
    if is_month:
        prev_start = start - pd.offsets.MonthBegin(1)
        label = "prior month"
    else:
        span = (end + pd.Timedelta(days=1)) - start
        prev_start = start - span
        label = "prior period"
    prev_end = start - pd.Timedelta(seconds=1)
    return prev_start, prev_end, label


def _aging_snapshot(master: pd.DataFrame, end_dt) -> dict:
    """Age currently open/in-progress records at the reporting period end."""
    all_rows = dedupe(master.copy())
    active = all_rows[all_rows["_Status"].isin(["Open", "In Progress"])].copy()
    cutoff = pd.Timestamp(end_dt)
    # Exclude records created after the reporting cutoff; retain missing-date rows for QA.
    active = active[active["_creation"].isna() | active["_creation"].le(cutoff)].copy()
    no_date = int(active["_creation"].isna().sum())
    age = (cutoff.normalize() - active["_creation"].dt.normalize()).dt.days
    age = age.where(age >= 0)
    active["Age (days)"] = age

    def _age_bucket(value):
        if pd.isna(value):
            return "No creation date"
        if value <= 7:
            return "0–7 days"
        if value <= 30:
            return "8–30 days"
        if value <= AGED_DAYS:
            return f"31–{AGED_DAYS} days"
        return f">{AGED_DAYS} days"

    active["Age Bucket"] = active["Age (days)"].map(_age_bucket)
    active["_age_sort"] = active["Age (days)"].fillna(-1)
    active = active.sort_values("_age_sort", ascending=False, kind="mergesort").drop(columns=["_age_sort"])
    # Present both source fields and normalized/parsed audit fields on the detail tab.
    active = active.rename(columns={
        "_Module": "Module (worksheet)", "_SrcRow": "Source Excel Row", "_creation": "Parsed Created Date",
        "_closure": "Parsed Closure Date", "_Status": "Normalized Status", "_Id": "Normalized Incident ID",
    })
    known_ages = active["Age (days)"].dropna()
    buckets = {
        "0–7 days": int((known_ages <= 7).sum()),
        "8–30 days": int(((known_ages > 7) & (known_ages <= 30)).sum()),
        f"31–{AGED_DAYS} days": int(((known_ages > 30) & (known_ages <= AGED_DAYS)).sum()),
        f">{AGED_DAYS} days": int((known_ages > AGED_DAYS).sum()),
        "No creation date": no_date,
    }
    oldest_age = int(known_ages.max()) if len(known_ages) else None
    return {
        "active_total": int(len(active)), "aged_over_threshold": buckets[f">{AGED_DAYS} days"],
        "oldest_age": oldest_age, "no_date": no_date, "buckets": buckets,
        "detail": active,
    }


# =============================================================================
# AGGREGATION — every number below comes from the same frames
# =============================================================================

def build_metrics(master: pd.DataFrame, modules: list[str],
                  start_dt, end_dt, issues: list) -> dict:
    union, created, closed = _period_frames(master, start_dt, end_dt)

    if union.empty:
        sys.exit("\nNo incidents found in the specified date range.\n")

    # module x status matrix (single source of truth for table AND chart)
    matrix = (union.groupby(["_Module", "_Status"]).size()
              .unstack(fill_value=0).reindex(columns=STATUS_ORDER, fill_value=0))
    matrix["Total"] = matrix.sum(axis=1)
    matrix = matrix.sort_values("Total", ascending=False)
    matrix = matrix[matrix["Total"] > 0]

    module_names  = list(matrix.index)
    module_totals = [int(v) for v in matrix["Total"]]
    module_status = {m: {s: int(matrix.loc[m, s]) for s in STATUS_ORDER}
                     for m in module_names}

    vc = union["_Status"].value_counts()
    status_labels = [s for s in STATUS_ORDER if int(vc.get(s, 0)) > 0]
    status_counts = [int(vc.get(s, 0)) for s in status_labels]

    # user breakdown — from closed-in-period rows
    user_names, user_counts = [], []
    if COL_CLOSED_BY and COL_CLOSED_BY in closed.columns and not closed.empty:
        u = closed[COL_CLOSED_BY].astype(str).str.strip()
        u = u[~u.str.lower().isin(_NULL_TOKENS)]
        if len(u):
            s = u.value_counts()
            user_names  = list(s.index)
            user_counts = [int(c) for c in s.values]
    elif COL_CLOSED_BY:
        print(f"  WARNING: column '{COL_CLOSED_BY}' not found - user table skipped.")

    # event count
    total_events = 0
    if COL_EVENT_COUNT and COL_EVENT_COUNT in closed.columns and not closed.empty:
        ev = pd.to_numeric(closed[COL_EVENT_COUNT], errors="coerce")
        total_events = int(ev.fillna(0).sum())
    elif COL_EVENT_COUNT:
        print(f"  INFO: column '{COL_EVENT_COUNT}' not found - Events KPI skipped.")

    # resolution time (only where both dates are valid and ordered)
    mttr_mean, mttr_med, bad = _mttr(closed)
    if bad:
        issues.append({"Type": "Closure before creation", "Module": "(various)",
                       "Column": COL_CLOSURE_DATE, "Row": "",
                       "Detail": f"{bad} row(s) excluded from MTTR"})

    # Rolling monthly trend ending in the report month (not limited to the selected period).
    last_month = end_dt.to_period("M")
    months = pd.period_range(last_month - (TREND_MONTHS - 1), last_month, freq="M")
    trend_labels, trend_created, trend_closed = [], [], []
    for month in months:
        month_start = month.start_time
        month_end = min(month.end_time, pd.Timestamp(end_dt))
        month_created = dedupe(master[master["_creation"].between(month_start, month_end)].copy())
        month_closed = dedupe(master[master["_closure"].between(month_start, month_end)].copy())
        trend_labels.append(month.strftime("%b %y"))
        trend_created.append(len(month_created))
        trend_closed.append(len(month_closed))

    # ALL-TIME status mix - every incident in the file, ignoring the date range
    alltime = dedupe(master.copy())
    avc = alltime["_Status"].value_counts()
    alltime_labels = [x for x in STATUS_ORDER if int(avc.get(x, 0)) > 0]
    alltime_counts = [int(avc.get(x, 0)) for x in alltime_labels]
    alltime_total  = sum(alltime_counts)
    alltime_backlog = int(avc.get("Open", 0)) + int(avc.get("In Progress", 0))

    # emails
    emails = extract_emails(master if EMAIL_SCOPE == "all" else union)

    # ids that sit in more than one module inside this period's union
    ids_in = union[union["_Id"].notna()]
    cross_module_dups = int((ids_in.groupby("_Id")["_Module"].nunique() > 1).sum()) \
        if len(ids_in) else 0

    total_union = int(matrix["Total"].sum())
    n_created = len(created)
    metrics = {
        "union": union, "closed": closed, "created": created,
        "module_names": module_names, "module_totals": module_totals,
        "module_status": module_status,
        "status_labels": status_labels, "status_counts": status_counts,
        "user_names": user_names, "user_counts": user_counts,
        "total_events": total_events,
        "mttr_mean": mttr_mean, "mttr_med": mttr_med,
        "alltime_labels": alltime_labels, "alltime_counts": alltime_counts,
        "alltime_total": alltime_total, "alltime_backlog": alltime_backlog,
        "trend_labels": trend_labels, "trend_created": trend_created,
        "trend_closed": trend_closed,
        "emails": emails, "total_union": total_union,
        "n_created": n_created, "n_closed": len(closed),
        "n_closed_from_earlier": total_union - n_created,
        "cross_module_dups": cross_module_dups,
        "n_closed_status": int(vc.get("Closed", 0)),
        "resolution_rate": (int(vc.get("Closed", 0)) / total_union * 100.0) if total_union else None,
        "net_backlog_change": n_created - len(closed),
    }
    reconcile(metrics)
    return metrics


_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


def extract_emails(df: pd.DataFrame) -> list[str]:
    if not COL_COMMENT or COL_COMMENT not in df.columns or COL_STATUS not in df.columns:
        return []
    m = df[COL_STATUS].astype(str).str.strip().str.lower() == STATUS_FOR_EMAILS
    sub = df[m]
    print(f"\n  Rows with status '{STATUS_FOR_EMAILS}' ({EMAIL_SCOPE} scope): {len(sub)}")
    if sub.empty:
        return []
    found = set()
    for val in sub[COL_COMMENT].dropna():
        for hit in _EMAIL_RE.findall(str(val)):
            found.add(hit.lower().strip(" .,;:"))
    out = sorted(found)
    print(f"  Unique emails: {len(out)}")
    return out


def reconcile(m: dict) -> None:
    """Hard guarantee that the dashboard's numbers agree with each other."""
    print("\n  " + "-" * 56)
    print("  RECONCILIATION")
    a = sum(m["module_totals"])
    b = sum(m["status_counts"])
    c = m["total_union"]
    d = len(m["union"])
    per_module_sum = sum(sum(v.values()) for v in m["module_status"].values())
    checks = [
        ("module table total      == period total", a, c),
        ("status breakdown total  == period total", b, c),
        ("stacked-bar segments    == period total", per_module_sum, c),
        ("de-duplicated row count == period total", d, c),
        ("created + closed-earlier== period total",
         m["n_created"] + m["n_closed_from_earlier"], c),
    ]
    ok = True
    for label, x, y in checks:
        flag = "PASS" if x == y else "FAIL"
        ok &= (x == y)
        print(f"    [{flag}] {label}   ({x:,} vs {y:,})")
    at = sum(m["alltime_counts"])
    print(f"    [{'PASS' if at == m['alltime_total'] else 'FAIL'}] "
          f"all-time status total  == all-time rows   "
          f"({at:,} vs {m['alltime_total']:,})")
    print(f"    [INFO] {c:,} incidents = {m['n_created']:,} created in period "
          f"+ {m['n_closed_from_earlier']:,} closed in period but created earlier")
    if m["cross_module_dups"]:
        print(f"    [INFO] {m['cross_module_dups']} incident id(s) sit in more than one module "
              f"and are counted once per module (see 'Count Audit' tab)")
    if m["user_counts"]:
        print(f"    [INFO] closed-by total {sum(m['user_counts']):,} of "
              f"{m['n_closed']:,} closed-in-period "
              f"(difference = rows with no '{COL_CLOSED_BY}' value)")
    print("  " + "-" * 56)
    if not ok:
        print("  !! Reconciliation failed - please report this input file.")



# =============================================================================
# WORKBOOK
# =============================================================================

def _rows_for(px: int, row_px: int = 20) -> int:
    return int(math.ceil(px / row_px)) + 2


def _days_text(value) -> str:
    if value is None or pd.isna(value):
        return "—"
    return f"{value:.1f} d" if value < 10 else f"{value:.0f} d"


def _period_label(start_dt, end_dt) -> str:
    start = pd.Timestamp(start_dt).normalize()
    end = pd.Timestamp(end_dt).normalize()
    if start.day == 1 and start.month == end.month and start.year == end.year and end.day == end.days_in_month:
        return start.strftime("%B %Y")
    return f"{start:%d %b %Y} – {end:%d %b %Y}"


def build_workbook(m: dict, modules: list[str], master: pd.DataFrame,
                   issues: list, start_dt, end_dt, output_path: str) -> None:
    wb = xlsxwriter.Workbook(output_path, {"nan_inf_to_errors": True})

    # Worksheet order matters: the dashboard must be the first, active tab.
    sw = wb.add_worksheet(SUMMARY_SHEET_NAME)
    cd = wb.add_worksheet(CDSHEET)
    cd.hide()

    def _f(**kw):
        base = {"font_name": "Calibri", "font_size": 10, "valign": "vcenter"}
        base.update(kw)
        return wb.add_format(base)

    # ------------------------------------------------------------- formats
    f_canvas   = _f(bg_color=CANVAS)
    f_banner   = _f(bold=True, font_size=19, font_color=WHITE, bg_color=INK,
                    align="left")
    f_accent   = _f(bg_color=ACCENT)
    f_meta     = _f(font_size=9, font_color=MUTED, bg_color=CANVAS, align="left")
    f_meta_r   = _f(font_size=9, font_color=MUTED, bg_color=CANVAS, align="right")
    f_section  = _f(bold=True, font_size=10, font_color=INK, bg_color=CANVAS,
                    align="left", left=5, left_color=ACCENT,
                    bottom=1, bottom_color=LINE)
    f_note     = _f(font_size=9, italic=True, font_color=MUTED, bg_color=CANVAS,
                    align="left")
    f_note_warn = _f(font_size=9, italic=True, font_color=AMBER_D, bg_color=CANVAS,
                     align="left")

    f_th       = _f(bold=True, font_size=10, font_color=WHITE, bg_color=INK,
                    align="center", border=1, border_color=INK)
    f_th_l     = _f(bold=True, font_size=10, font_color=WHITE, bg_color=INK,
                    align="left", border=1, border_color=INK)

    def _cell(align="center", alt=False, **kw):
        base = dict(align=align, border=1, border_color=LINE, font_color=SLATE)
        if alt:
            base["bg_color"] = LINE_SOFT
        base.update(kw)
        return _f(**base)

    f_c, f_c_alt = _cell(), _cell(alt=True)
    f_l, f_l_alt = _cell("left", font_color=INK), _cell("left", alt=True, font_color=INK)
    f_p, f_p_alt = _cell(num_format="0.0%"), _cell(alt=True, num_format="0.0%")
    f_b, f_b_alt = _cell(bold=True, font_color=INK), _cell(alt=True, bold=True, font_color=INK)

    f_tot      = _f(bold=True, font_size=10, font_color=ACCENT_D, bg_color=ACCENT_L,
                    align="center", border=1, border_color=LINE)
    f_tot_l    = _f(bold=True, font_size=10, font_color=ACCENT_D, bg_color=ACCENT_L,
                    align="left", border=1, border_color=LINE)
    f_tot_p    = _f(bold=True, font_size=10, font_color=ACCENT_D, bg_color=ACCENT_L,
                    align="center", border=1, border_color=LINE, num_format="0.0%")

    f_data_hdr = _f(bold=True, font_color=WHITE, bg_color=INK, align="center",
                    border=1, border_color=INK)
    f_cell     = _cell("left", font_color=INK)
    f_cell_alt = _cell("left", alt=True, font_color=INK)
    f_date     = _cell("left", font_color=INK, num_format="dd mmm yyyy hh:mm")
    f_date_alt = _cell("left", alt=True, font_color=INK,
                       num_format="dd mmm yyyy hh:mm")

    f_status_hdr = {st: _f(bold=True, font_size=10, font_color=WHITE,
                           bg_color=_STATUS_RAG[st], align="center",
                           border=1, border_color=_STATUS_RAG[st])
                    for st in STATUS_ORDER}
    f_status_cell = {st: _f(align="center", border=1, border_color=LINE,
                            bg_color=_STATUS_TINT[st], font_color=_STATUS_DARK[st],
                            bold=True) for st in STATUS_ORDER}

    def W(row, c1, c2, value, fmt):
        """Write, merging when the span is wider than one column."""
        if c2 > c1:
            sw.merge_range(row, c1, row, c2, value, fmt)
        else:
            sw.write(row, c1, value, fmt)

    def band(row, height, fmt=None):
        """Fill the full page width so the sheet reads as a designed page."""
        sw.set_row(row, height)
        sw.merge_range(row, GUTTER, row, CN + 1, "", fmt or f_canvas)

    def section(row, text, sub=""):
        sw.set_row(row, 24)
        sw.write_blank(row, GUTTER, None, f_canvas)
        W(row, C0, CN, f"  {text.upper()}" + (f"      {sub}" if sub else ""),
          f_section)
        sw.write_blank(row, CN + 1, None, f_canvas)

    # -------------------------------------------------------- chart data sheet
    mods   = m["module_names"]
    NM     = len(mods)
    st_lbl = m["status_labels"]
    at_lbl = m["alltime_labels"]

    mods_chart = list(reversed(mods))          # ascending -> biggest on top
    for i, mod in enumerate(mods_chart):
        cd.write_string(i, 0, str(mod))
        for j, st in enumerate(STATUS_ORDER, start=1):
            v = m["module_status"][mod].get(st, 0)
            if v:
                cd.write_number(i, j, v)
            else:
                cd.write_blank(i, j, None)     # blank -> no segment, no label

    for i, (l, c) in enumerate(zip(st_lbl, m["status_counts"])):
        cd.write_string(i, 6, l)
        cd.write_number(i, 7, c)
    NS = len(st_lbl)

    NU = len(m["user_names"])
    for i, (u, c) in enumerate(zip(reversed(m["user_names"]),
                                   reversed(m["user_counts"]))):
        cd.write_string(i, 9, str(u))
        cd.write_number(i, 10, c)

    NT = len(m["trend_labels"])
    for i in range(NT):
        cd.write_string(i, 12, m["trend_labels"][i])
        cd.write_number(i, 13, m["trend_created"][i])
        cd.write_number(i, 14, m["trend_closed"][i])

    for i, (l, c) in enumerate(zip(at_lbl, m["alltime_counts"])):
        cd.write_string(i, 16, l)
        cd.write_number(i, 17, c)
    NA = len(at_lbl)

    # ------------------------------------------------------------ page set-up
    sw.set_zoom(85)
    sw.hide_gridlines(2)
    sw.set_tab_color(ACCENT)
    sw.set_landscape()
    sw.set_paper(9)
    sw.fit_to_pages(1, 0)
    sw.set_margins(0.25, 0.25, 0.35, 0.35)
    sw.set_column(GUTTER, GUTTER, GUT_W)
    sw.set_column(C0, CN, COL_W)
    sw.set_column(CN + 1, CN + 1, GUT_W)

    total  = m["total_union"] or 1
    def _pick(labels, counts, name):
        return next((c for l, c in zip(labels, counts) if l == name), 0)
    closed = _pick(st_lbl, m["status_counts"], "Closed")
    inprog = _pick(st_lbl, m["status_counts"], "In Progress")
    openc  = _pick(st_lbl, m["status_counts"], "Open")
    other  = _pick(st_lbl, m["status_counts"], "Other")
    res_rate = closed / total * 100
    pend_pct = (openc + inprog) / total * 100

    # ---------------------------------------------------------------- banner
    sw.set_row(0, 46)
    W(0, GUTTER, CN + 1, "   INCIDENT SUMMARY DASHBOARD", f_banner)
    band(1, 5, f_accent)
    sw.set_row(2, 20)
    sw.write_blank(2, GUTTER, None, f_canvas)
    W(2, C0, 12,
      f"  {start_dt:%d %b %Y}  —  {end_dt:%d %b %Y}"
      f"      ·      {m['n_created']:,} created in period"
      f"      ·      {m['n_closed']:,} closed in period", f_meta)
    W(2, 13, CN, f"Generated {datetime.now():%d %b %Y, %H:%M}   ", f_meta_r)
    sw.write_blank(2, CN + 1, None, f_canvas)
    band(3, 10)

    # ------------------------------------------------------------- executive KPI cards
    KPI = 4
    tiles = _kpi_tiles(m)          # same tiles feed the PowerPoint
    sw.set_row(KPI, 5); sw.set_row(KPI + 1, 17); sw.set_row(KPI + 2, 42)
    sw.set_row(KPI + 3, 15); sw.set_row(KPI + 4, 14)
    for r in range(KPI, KPI + 5):
        sw.write_blank(r, GUTTER, None, f_canvas)
        sw.write_blank(r, CN + 1, None, f_canvas)
    step = (CN - C0 + 1) // len(tiles)
    for bi, (label, value, delta_text, caption, colour, delta_colour) in enumerate(tiles):
        c1 = C0 + bi * step
        c2 = CN if bi == len(tiles) - 1 else c1 + step - 1
        side = dict(left=1, right=1, left_color=LINE, right_color=LINE)
        W(KPI, c1, c2, "", _f(bg_color=colour))
        W(KPI + 1, c1, c2, f"  {label}",
          _f(bold=True, font_size=9, font_color=MUTED, bg_color=WHITE,
             align="left", **side))
        W(KPI + 2, c1, c2, value,
          _f(bold=True, font_size=27, font_color=INK, bg_color=WHITE,
             align="center", **side))
        W(KPI + 3, c1, c2, delta_text,
          _f(font_size=8, bold=True, font_color=delta_colour, bg_color=WHITE,
             align="center", **side))
        W(KPI + 4, c1, c2, f"  {caption}",
          _f(font_size=8, italic=True, font_color=MUTED, bg_color=WHITE,
             align="left", bottom=1, bottom_color=LINE, **side))

    # Freeze only the title rows - the KPI cards scroll away like everything else.
    sw.freeze_panes(3, 0)
    cursor = KPI + 5

    # one-line definitions so nobody has to guess how the headline number is built
    sw.set_row(cursor, 16)
    sw.write_blank(cursor, GUTTER, None, f_canvas)
    W(cursor, C0, CN,
      "  How to read this:  Incidents in period = created in the period + closed in the period "
      "but created earlier (each counted once).  Status = current status in the source file.",
      f_note)
    sw.write_blank(cursor, CN + 1, None, f_canvas)
    cursor += 1
    if m.get("cross_module_dups"):
        sw.set_row(cursor, 16)
        sw.write_blank(cursor, GUTTER, None, f_canvas)
        W(cursor, C0, CN,
          f"  Note: {m['cross_module_dups']} incident id(s) appear in more than one module and are "
          f"counted once per module. See the '{AUDIT_SHEET_NAME}' tab "
          f"(set DEDUPE_ACROSS_MODULES = True to count each id once overall).",
          f_note_warn)
        sw.write_blank(cursor, CN + 1, None, f_canvas)
        cursor += 1
    band(cursor, 12); cursor += 1

    # --------------------------------------------------------- executive insight cards
    aging = m.get("aging", {"aged_over_threshold": 0, "oldest_age": None})
    boxes = [("MOST IMPACTED MODULE", ACCENT, ACCENT_L,
              (str(mods[0])[:24] if mods else "—"),
              f"{(m['module_totals'][0] if mods else 0):,} incidents in period")]
    if COL_EVENT_COUNT:
        boxes.append(("TOTAL EVENTS CLOSED", GREEN, GREEN_L,
                      f"{m['total_events']:,}",
                      f"sum of '{COL_EVENT_COUNT}' for closures in period"))
    boxes.append(("OPEN BACKLOG · ALL TIME", RED, RED_L,
                  f"{m['alltime_backlog']:,}",
                  f"open + in progress across all {m['alltime_total']:,} incidents"))
    boxes.append((f"OPEN ITEMS > {AGED_DAYS} DAYS", RED if aging.get("aged_over_threshold", 0) else GREEN,
                  RED_L if aging.get("aged_over_threshold", 0) else GREEN_L,
                  f"{aging.get('aged_over_threshold', 0):,}",
                  f"oldest known item: {_days_text(aging.get('oldest_age'))}" if aging.get("oldest_age") is not None else "no dated open items"))
    boxes.append(("DATA QUALITY FLAGS", AMBER if m.get("quality_flag_count", 0) else GREEN,
                  AMBER_L if m.get("quality_flag_count", 0) else GREEN_L,
                  f"{m.get('quality_flag_count', 0):,}",
                  "excluding auto-matched headers"))

    section(cursor, "Key insights", "executive takeaways from period data")
    cursor += 1
    sw.set_row(cursor, 5); sw.set_row(cursor + 1, 16)
    sw.set_row(cursor + 2, 34); sw.set_row(cursor + 3, 16)
    for r in range(cursor, cursor + 4):
        sw.write_blank(r, GUTTER, None, f_canvas)
        sw.write_blank(r, CN + 1, None, f_canvas)
    width, rem = divmod(CN - C0 + 1, len(boxes))
    c1 = C0
    for bi, (title, colour, tint, big, sub) in enumerate(boxes):
        w = width + (1 if bi < rem else 0)
        c2 = c1 + w - 1
        side = dict(left=1, right=1, left_color=LINE, right_color=LINE)
        W(cursor, c1, c2, "", _f(bg_color=colour))
        W(cursor + 1, c1, c2, f"  {title}",
          _f(bold=True, font_size=8, font_color=colour, bg_color=tint,
             align="left", **side))
        W(cursor + 2, c1, c2, big,
          _f(bold=True, font_size=17, font_color=INK, bg_color=tint,
             align="center", **side))
        W(cursor + 3, c1, c2, f"  {sub}",
          _f(font_size=8, italic=True, font_color=SLATE, bg_color=tint,
             align="left", bottom=1, bottom_color=LINE, **side))
        c1 = c2 + 1
    cursor += 4
    band(cursor, 12); cursor += 1

    # --------------------------------------------------------- manual triage notes
    section(cursor, "Incidents triaged this month",
            "editable blank space for the monthly executive narrative")
    cursor += 1
    note_fmt = _f(bg_color=WHITE, border=1, border_color=LINE,
                   font_color=SLATE, align="left", valign="top", text_wrap=True)
    for note_row in range(cursor, cursor + 3):
        sw.set_row(note_row, 28)
        sw.write_blank(note_row, GUTTER, None, f_canvas)
        sw.merge_range(note_row, C0, note_row, CN, "", note_fmt)
        sw.write_blank(note_row, CN + 1, None, f_canvas)
    cursor += 3
    band(cursor, 12); cursor += 1

    # ------------------------------------------------------- module break-down
    MOD_COLS = [("#", 1, 1), ("Module Name", 2, 8), ("Total", 9, 10),
                ("Closed", 11, 12), ("In Progress", 13, 14), ("Open", 15, 16),
                ("Other", 17, 18), ("% of Period", 19, 20)]
    section(cursor, "Module breakdown",
            "incidents created OR closed inside the reporting period")
    HDR = cursor + 1
    DS  = HDR + 1
    TR  = DS + NM
    sw.set_row(HDR, 22)
    sw.write_blank(HDR, GUTTER, None, f_canvas)
    sw.write_blank(HDR, CN + 1, None, f_canvas)
    for name, c1, c2 in MOD_COLS:
        fmt = f_status_hdr[name] if name in STATUS_ORDER else (
            f_th_l if name == "Module Name" else f_th)
        W(HDR, c1, c2, name, fmt)

    for i, mod in enumerate(mods):
        r = DS + i
        alt = (i % 2 == 1)
        sb = m["module_status"][mod]
        sw.set_row(r, 19)
        sw.write_blank(r, GUTTER, None, f_canvas)
        sw.write_blank(r, CN + 1, None, f_canvas)
        W(r, 1, 1, i + 1, f_c_alt if alt else f_c)
        W(r, 2, 8, f"  {mod}", f_l_alt if alt else f_l)
        W(r, 9, 10, m["module_totals"][i], f_b_alt if alt else f_b)
        for j, st in enumerate(STATUS_ORDER):
            c1 = 11 + j * 2
            W(r, c1, c1 + 1, sb.get(st, 0), f_status_cell[st])
        W(r, 19, 20, m["module_totals"][i] / total, f_p_alt if alt else f_p)

    sw.conditional_format(DS, 9, TR - 1, 9, {
        "type": "data_bar", "data_bar_2010": True, "bar_color": "#A5B4FC",
        "bar_border_color": ACCENT, "bar_solid": True,
        "min_type": "num", "min_value": 0, "bar_direction": "left"})

    sw.set_row(TR, 22)
    sw.write_blank(TR, GUTTER, None, f_canvas)
    sw.write_blank(TR, CN + 1, None, f_canvas)
    W(TR, 1, 8, "  TOTAL", f_tot_l)
    L = xl_col_to_name(9)
    W(TR, 9, 10, "", f_tot)
    sw.write_formula(TR, 9, f"=SUM({L}{DS + 1}:{L}{DS + NM})", f_tot,
                     m["total_union"])
    for j, st in enumerate(STATUS_ORDER):
        c1 = 11 + j * 2
        L = xl_col_to_name(c1)
        val = sum(m["module_status"][mod].get(st, 0) for mod in mods)
        W(TR, c1, c1 + 1, "", f_tot)
        sw.write_formula(TR, c1, f"=SUM({L}{DS + 1}:{L}{DS + NM})", f_tot, val)
    L = xl_col_to_name(19)
    W(TR, 19, 20, "", f_tot_p)
    sw.write_formula(TR, 19, f"=SUM({L}{DS + 1}:{L}{DS + NM})", f_tot_p, 1.0)
    cursor = TR + 1
    band(cursor, 8); cursor += 1

    # ---------------------------------------------------- stacked module chart
    def _style(ch):
        ch.set_plotarea({"border": {"none": True}, "fill": {"color": WHITE}})
        ch.set_chartarea({"border": {"color": LINE, "width": 0.75},
                          "fill": {"color": WHITE}})
        ch.set_style(2)

    def _labels(values, colour):
        """Per-point labels with zero points deleted."""
        return {"value": True, "position": "center",
                "font": {"bold": True, "size": 11, "color": colour,
                         "name": "Calibri"},
                "custom": [({"delete": True} if not v else None) for v in values]}

    if NM:
        bar_h = max(440, NM * 56 + 170)
        stacked = wb.add_chart({"type": "bar", "subtype": "stacked"})
        for j, st in enumerate(STATUS_ORDER, start=1):
            vals = [m["module_status"][mod].get(st, 0) for mod in mods_chart]
            if not any(vals):
                continue
            stacked.add_series({
                "name": st,
                "categories": [CDSHEET, 0, 0, NM - 1, 0],
                "values": [CDSHEET, 0, j, NM - 1, j],
                "fill": {"color": _STATUS_RAG[st]},
                "border": {"color": WHITE, "width": 1.25},
                "gap": 32, "overlap": 100,
                "data_labels": _labels(vals, WHITE),
            })
        stacked.set_title({
            "name": f"Incidents by module, split by status  ·  {m['total_union']:,} in period",
            "name_font": {"bold": True, "size": 12, "color": INK, "name": "Calibri"}})
        stacked.set_legend({"position": "bottom",
                            "font": {"bold": True, "size": 10, "color": SLATE,
                                     "name": "Calibri"},
                            "border": {"none": True}, "fill": {"none": True}})
        stacked.set_x_axis({"num_font": {"size": 10, "color": MUTED},
                            "major_gridlines": {"visible": True,
                                                "line": {"color": LINE, "width": 0.75,
                                                         "dash_type": "dash"}},
                            "line": {"none": True}, "num_format": "0", "min": 0})
        stacked.set_y_axis({"num_font": {"size": 11, "bold": True, "color": INK},
                            "line": {"color": LINE}, "major_tick_mark": "none",
                            "major_gridlines": {"visible": False}})
        stacked.show_blanks_as("gap")
        _style(stacked)
        stacked.set_size({"width": 1215, "height": bar_h})
        sw.insert_chart(cursor, C0, stacked, {"x_offset": 2, "y_offset": 2})
        cursor += _rows_for(bar_h)
        band(cursor, 12); cursor += 1

    # ------------------------------------------- status mix: period + all time
    section(cursor, "Status mix",
            "left: the reporting period   ·   right: every incident in the file")
    cursor += 1
    mix_top = cursor

    def _mini_table(row, heading, labels, counts, grand):
        sw.set_row(row, 20)
        sw.write_blank(row, GUTTER, None, f_canvas)
        W(row, 1, 8, f"  {heading}",
          _f(bold=True, font_size=9, font_color=WHITE, bg_color=SLATE, align="left"))
        row += 1
        sw.set_row(row, 20)
        W(row, 1, 4, "Status", f_th_l)
        W(row, 5, 6, "Count", f_th)
        W(row, 7, 8, "% Share", f_th)
        row += 1
        first = row
        g = grand or 1
        for lbl, cnt in zip(labels, counts):
            sw.set_row(row, 19)
            W(row, 1, 4, f"  {lbl}",
              _f(bold=True, align="left", border=1, border_color=LINE,
                 bg_color=_STATUS_TINT[lbl], font_color=_STATUS_DARK[lbl]))
            W(row, 5, 6, cnt,
              _f(bold=True, align="center", border=1, border_color=LINE,
                 bg_color=_STATUS_TINT[lbl], font_color=_STATUS_DARK[lbl]))
            W(row, 7, 8, cnt / g,
              _f(align="center", border=1, border_color=LINE, num_format="0.0%",
                 bg_color=_STATUS_TINT[lbl], font_color=_STATUS_DARK[lbl]))
            row += 1
        sw.set_row(row, 21)
        W(row, 1, 4, "  TOTAL", f_tot_l)
        col = xl_col_to_name(5)
        W(row, 5, 6, "", f_tot)
        sw.write_formula(row, 5, f"=SUM({col}{first + 1}:{col}{row})", f_tot, grand)
        W(row, 7, 8, 1.0, f_tot_p)
        return row + 1

    r = _mini_table(mix_top, "IN PERIOD", st_lbl, m["status_counts"],
                    m["total_union"])
    r += 1
    r = _mini_table(r, "ALL TIME  ·  ENTIRE FILE", at_lbl, m["alltime_counts"],
                    m["alltime_total"])

    if NS:
        donut = wb.add_chart({"type": "doughnut"})
        donut.add_series({
            "name": "In period",
            "categories": [CDSHEET, 0, 6, NS - 1, 6],
            "values": [CDSHEET, 0, 7, NS - 1, 7],
            "points": [{"fill": {"color": _STATUS_RAG.get(l, ACCENT)},
                        "border": {"color": WHITE, "width": 2}} for l in st_lbl],
            "data_labels": {"percentage": True, "category": True, "value": True,
                            "separator": "\n",
                            "font": {"bold": True, "size": 9, "name": "Calibri",
                                     "color": INK}},
        })
        donut.set_hole_size(58)
        donut.set_title({"name": f"In period\n{m['total_union']:,} incidents",
                         "name_font": {"bold": True, "size": 11, "color": INK,
                                       "name": "Calibri"}})
        donut.set_legend({"none": True})
        _style(donut)
        donut.set_size({"width": 372, "height": 310})
        sw.insert_chart(mix_top, 9, donut, {"x_offset": 4, "y_offset": 0})

    if NA:
        pie = wb.add_chart({"type": "pie"})
        pie.add_series({
            "name": "All time",
            "categories": [CDSHEET, 0, 16, NA - 1, 16],
            "values": [CDSHEET, 0, 17, NA - 1, 17],
            "points": [{"fill": {"color": _STATUS_RAG.get(l, ACCENT)},
                        "border": {"color": WHITE, "width": 2}} for l in at_lbl],
            "data_labels": {"percentage": True, "category": True, "value": True,
                            "separator": "\n",
                            "font": {"bold": True, "size": 9, "name": "Calibri",
                                     "color": INK}},
        })
        pie.set_title({"name": f"All time · entire file\n{m['alltime_total']:,} incidents",
                       "name_font": {"bold": True, "size": 11, "color": INK,
                                     "name": "Calibri"}})
        pie.set_legend({"none": True})
        _style(pie)
        pie.set_size({"width": 372, "height": 310})
        sw.insert_chart(mix_top, 15, pie, {"x_offset": 4, "y_offset": 0})

    cursor = mix_top + max(r - mix_top, _rows_for(310))
    if other:
        sw.set_row(cursor, 16)
        sw.write_blank(cursor, GUTTER, None, f_canvas)
        W(cursor, C0, CN,
          f"  {other:,} incident(s) in the period carry a status outside the "
          f"Closed / In Progress / Open mapping and are grouped as 'Other'. "
          f"The Data Quality tab lists the exact values — add them to "
          f"STATUS_RULES to reclassify.", f_note)
        cursor += 1
    band(cursor, 12); cursor += 1

    # ------------------------------------------------------------ trend chart
    if NT >= 2:
        section(cursor, "Monthly trend", f"created vs closed · rolling {TREND_MONTHS} months through report end")
        cursor += 1
        trend = wb.add_chart({"type": "line"})
        for name, col, colour in (("Created", 13, ACCENT), ("Closed", 14, GREEN)):
            trend.add_series({
                "name": name,
                "categories": [CDSHEET, 0, 12, NT - 1, 12],
                "values": [CDSHEET, 0, col, NT - 1, col],
                "line": {"color": colour, "width": 2.5},
                "marker": {"type": "circle", "size": 7,
                           "fill": {"color": colour},
                           "border": {"color": WHITE, "width": 1.5}},
                "data_labels": {"value": True,
                                "font": {"bold": True, "size": 9, "color": colour,
                                         "name": "Calibri"}},
            })
        trend.set_title({"name": "", "none": True})
        trend.set_legend({"position": "bottom",
                          "font": {"bold": True, "size": 10, "color": SLATE},
                          "border": {"none": True}})
        trend.set_x_axis({"num_font": {"size": 10, "bold": True, "color": INK},
                          "line": {"color": LINE}, "major_tick_mark": "none"})
        trend.set_y_axis({"num_font": {"size": 10, "color": MUTED},
                          "major_gridlines": {"visible": True,
                                              "line": {"color": LINE, "width": 0.75,
                                                       "dash_type": "dash"}},
                          "line": {"none": True}, "min": 0, "num_format": "0"})
        _style(trend)
        trend.set_size({"width": 1215, "height": 300})
        sw.insert_chart(cursor, C0, trend, {"x_offset": 2, "y_offset": 2})
        cursor += _rows_for(300)
        band(cursor, 12); cursor += 1

    # ------------------------------------------------------ closed-by-user
    if NU:
        section(cursor, "Closed by user",
                "incidents whose closure date falls inside the period")
        UHDR = cursor + 1
        UDS  = UHDR + 1
        UTR  = UDS + NU
        sw.set_row(UHDR, 22)
        sw.write_blank(UHDR, GUTTER, None, f_canvas)
        W(UHDR, 1, 1, "#", f_th)
        W(UHDR, 2, 6, "Closed By", f_th_l)
        W(UHDR, 7, 8, "Incidents Closed", f_th)
        W(UHDR, 9, 10, "% Share", f_th)
        tot_u = sum(m["user_counts"]) or 1
        for i in range(NU):
            rr = UDS + i
            alt = (i % 2 == 1)
            sw.set_row(rr, 19)
            sw.write_blank(rr, GUTTER, None, f_canvas)
            W(rr, 1, 1, i + 1, f_c_alt if alt else f_c)
            W(rr, 2, 6, f"  {m['user_names'][i]}", f_l_alt if alt else f_l)
            W(rr, 7, 8, m["user_counts"][i], f_b_alt if alt else f_b)
            W(rr, 9, 10, m["user_counts"][i] / tot_u, f_p_alt if alt else f_p)
        sw.conditional_format(UDS, 7, UTR - 1, 7, {
            "type": "data_bar", "data_bar_2010": True, "bar_color": "#A5B4FC",
            "bar_border_color": ACCENT, "bar_solid": True,
            "min_type": "num", "min_value": 0, "bar_direction": "left"})
        sw.set_row(UTR, 22)
        sw.write_blank(UTR, GUTTER, None, f_canvas)
        W(UTR, 1, 6, "  TOTAL", f_tot_l)
        L = xl_col_to_name(7)
        W(UTR, 7, 8, "", f_tot)
        sw.write_formula(UTR, 7, f"=SUM({L}{UDS + 1}:{L}{UDS + NU})", f_tot, tot_u)
        W(UTR, 9, 10, 1.0, f_tot_p)

        ubar = wb.add_chart({"type": "bar"})
        ubar.add_series({
            "name": "Incidents Closed",
            "categories": [CDSHEET, 0, 9, NU - 1, 9],
            "values": [CDSHEET, 0, 10, NU - 1, 10],
            "fill": {"color": ACCENT},
            "border": {"color": WHITE, "width": 0.75},
            "gap": 45,
            "data_labels": {"value": True, "position": "inside_end",
                            "font": {"bold": True, "size": 10, "color": WHITE,
                                     "name": "Calibri"}},
        })
        ubar.set_title({"name": f"Closures by user  ·  {tot_u:,} in period",
                        "name_font": {"bold": True, "size": 11, "color": INK,
                                      "name": "Calibri"}})
        ubar.set_legend({"none": True})
        ubar.set_x_axis({"num_font": {"size": 10, "color": MUTED},
                         "major_gridlines": {"visible": True,
                                             "line": {"color": LINE, "width": 0.6,
                                                      "dash_type": "dash"}},
                         "line": {"none": True}, "num_format": "0", "min": 0})
        ubar.set_y_axis({"num_font": {"size": 10, "bold": True, "color": INK},
                         "line": {"color": LINE}, "major_tick_mark": "none",
                         "major_gridlines": {"visible": False}})
        _style(ubar)
        u_h = max(270, NU * 32 + 110)
        ubar.set_size({"width": 610, "height": u_h})
        sw.insert_chart(UHDR, 11, ubar, {"x_offset": 6, "y_offset": 0})
        cursor = UHDR + max(UTR - UHDR + 1, _rows_for(u_h))
        band(cursor, 12); cursor += 1

    # ---------------------------------------------------------------- footer
    sw.set_row(cursor, 18)
    sw.write_blank(cursor, GUTTER, None, f_canvas)
    W(cursor, C0, CN,
      f"  Every figure above is de-duplicated and reconciled against the same "
      f"source rows  ·  {m.get('quality_flag_count', len(issues)):,} data-quality flag(s) logged on the "
      f"'{DQ_SHEET_NAME}' tab  ·  row-by-row proof on the '{AUDIT_SHEET_NAME}' tab", f_note)
    sw.write_blank(cursor, CN + 1, None, f_canvas)

    # ---- emails sheet --------------------------------------------------------
    emails = m["emails"]
    if emails:
        ew = wb.add_worksheet("Emails - Closed Resolved")
        ew.hide_gridlines(2)
        ew.set_tab_color(GREEN_D)
        ew.set_row(0, 38)
        ew.merge_range(0, 0, 0, 2,
                       f"Unique Emails — {STATUS_FOR_EMAILS.title()}  │  {len(emails):,} addresses",
                       _f(bold=True, font_size=13, font_color=WHITE, bg_color=INK, align="center"))
        eh = _f(bold=True, font_size=11, font_color=WHITE, bg_color=INK,
                align="center", border=1)
        ew.set_row(2, 22)
        ew.write(2, 0, "#", eh)
        ew.write(2, 1, "Email Address", eh)
        ew.set_column(0, 0, 6)
        ew.set_column(1, 1, max(max(len(e) for e in emails) + 6, 30))
        ew.freeze_panes(3, 0)
        ew.autofilter(2, 0, 2 + len(emails), 1)
        for i, e in enumerate(emails):
            r = 3 + i
            bg = LINE_SOFT if i % 2 else WHITE
            ew.set_row(r, 16)
            ew.write(r, 0, i + 1, _f(align="center", border=1, border_color=LINE, bg_color=bg))
            ew.write(r, 1, e, _f(align="left", border=1, border_color=LINE, bg_color=bg))

    # ---- count audit sheet ---------------------------------------------------
    au = wb.add_worksheet(AUDIT_SHEET_NAME)
    au.hide_gridlines(2)
    au.set_tab_color(ACCENT)
    audit = m["union"].sort_values(["_Module", "_creation"], kind="mergesort").copy()
    a_cre = audit["_creation"].between(start_dt, end_dt)
    a_clo = audit["_closure"].between(start_dt, end_dt)
    id_mods = (master[master["_Id"].notna()].groupby("_Id")["_Module"]
               .agg(lambda s: sorted(set(s))).to_dict())
    au_cols = ["#", "Module", "Incident ID", "Created", "Closed",
               "Status (source)", "Bucket", "Why it is counted",
               "Also appears in module(s)"]
    au.set_row(0, 38)
    au.merge_range(0, 0, 0, len(au_cols) - 1,
                   f"Count Audit  │  {len(audit):,} incident(s)  =  {m['n_created']:,} created in period"
                   f"  +  {m['n_closed_from_earlier']:,} closed in period (created earlier)",
                   _f(bold=True, font_size=13, font_color=WHITE, bg_color=INK, align="center"))
    au.set_row(1, 18)
    au.merge_range(0 + 1, 0, 1, len(au_cols) - 1,
                   "  One line per incident behind the dashboard total. Amber = the same incident id "
                   "also exists in another module.",
                   _f(italic=True, font_size=9, font_color=MUTED, align="left"))
    au_hdr = _f(bold=True, font_size=10, font_color=WHITE, bg_color=INK,
                align="center", border=1, border_color=INK)
    au.set_row(2, 22)
    for ci, h in enumerate(au_cols):
        au.write(2, ci, h, au_hdr)
    for ci, w_ in enumerate([6, 24, 18, 20, 20, 24, 13, 40, 30]):
        au.set_column(ci, ci, w_)
    au.freeze_panes(3, 0)
    f_warn = _f(align="left", border=1, border_color=LINE, bg_color=AMBER_L,
                font_color=AMBER_D, bold=True)
    src_status = audit[COL_STATUS] if COL_STATUS in audit.columns else pd.Series("", index=audit.index)
    for i, ((idx, row), cin, cout) in enumerate(zip(audit.iterrows(), a_cre, a_clo)):
        r = 3 + i
        alt = (i % 2 == 1)
        base_l = f_cell_alt if alt else f_cell
        base_c = f_c_alt if alt else f_c
        base_d = f_date_alt if alt else f_date
        mod, _id = row["_Module"], row["_Id"]
        elsewhere = ", ".join(x for x in id_mods.get(_id, []) if x != mod) if _id else ""
        why = ("Created and closed in period" if (cin and cout) else
               "Created in period" if cin else "Closed in period (created earlier)")
        au.write_number(r, 0, i + 1, base_c)
        au.write(r, 1, str(mod), base_l)
        au.write(r, 2, str(row.get(COL_INCIDENT_ID, _id)) if _id else "(none)", base_l)
        for ci, key in ((3, "_creation"), (4, "_closure")):
            v = row[key]
            if pd.isna(v):
                au.write_blank(r, ci, None, base_d)
            else:
                au.write_datetime(r, ci, pd.Timestamp(v).to_pydatetime(), base_d)
        sv = src_status.loc[idx]
        au.write(r, 5, "" if pd.isna(sv) else str(sv), base_l)
        au.write(r, 6, row["_Status"], f_status_cell.get(row["_Status"], base_c))
        au.write(r, 7, why, base_l)
        au.write(r, 8, elsewhere, f_warn if elsewhere else base_l)
    if len(audit):
        au.autofilter(2, 0, 2 + len(audit), len(au_cols) - 1)

    # ---- data quality sheet --------------------------------------------------
    dq = wb.add_worksheet(DQ_SHEET_NAME)
    dq.hide_gridlines(2)
    dq.set_tab_color(RED if issues else GREEN)
    dq.set_row(0, 38)
    dq.merge_range(0, 0, 0, 4,
                   f"Data Quality Log  │  {len(issues):,} item(s) — fix these at source "
                   f"to sharpen next month's numbers",
                   _f(bold=True, font_size=13, font_color=WHITE, bg_color=INK, align="center"))
    dq_hdr = _f(bold=True, font_size=11, font_color=WHITE, bg_color=INK,
                align="center", border=1)
    cols = ["Type", "Module", "Column", "Excel Row", "Detail / Raw Value"]
    dq.set_row(2, 22)
    for ci, h in enumerate(cols):
        dq.write(2, ci, h, dq_hdr)
    dq.set_column(0, 0, 34); dq.set_column(1, 1, 22)
    dq.set_column(2, 2, 24); dq.set_column(3, 3, 11); dq.set_column(4, 4, 62)
    dq.freeze_panes(3, 0)

    if issues:
        shown = issues[:MAX_DQ_ROWS]
        for i, it in enumerate(shown):
            r = 3 + i
            bg = LINE_SOFT if i % 2 else WHITE
            cell = _f(align="left", border=1, border_color=LINE, bg_color=bg)
            ctr = _f(align="center", border=1, border_color=LINE, bg_color=bg)
            dq.write(r, 0, it["Type"], cell)
            dq.write(r, 1, str(it["Module"]), cell)
            dq.write(r, 2, str(it["Column"]), cell)
            dq.write(r, 3, it["Row"], ctr)
            dq.write(r, 4, str(it["Detail"]), cell)
        dq.autofilter(2, 0, 2 + len(shown), 4)
        if len(issues) > MAX_DQ_ROWS:
            dq.write(3 + len(shown), 0,
                     f"... {len(issues) - MAX_DQ_ROWS:,} more suppressed "
                     f"(raise MAX_DQ_ROWS to see all)",
                     _f(italic=True, font_color=MUTED))
    else:
        dq.merge_range(3, 0, 3, 4, "  ✓  No data quality issues detected in this period.",
                       _f(bold=True, font_size=11, font_color=GREEN_D, bg_color=GREEN_L,
                          align="left", border=1))

    if _unmapped_statuses:
        r0 = 3 + min(len(issues), MAX_DQ_ROWS) + 3
        dq.merge_range(r0, 0, r0, 4,
                       "  Unmapped status values (counted as 'Other') — add them to "
                       "STATUS_RULES / STATUS_OVERRIDES",
                       _f(bold=True, font_size=10, font_color=WHITE, bg_color=AMBER_D,
                          align="left"))
        for i, (raw, n) in enumerate(sorted(_unmapped_statuses.items(),
                                            key=lambda kv: -kv[1])):
            dq.write(r0 + 1 + i, 0, raw,
                     _f(align="left", border=1, border_color=LINE))
            dq.write(r0 + 1 + i, 1, n,
                     _f(align="center", border=1, border_color=LINE))

    # ---- open backlog detail -------------------------------------------------
    ob = wb.add_worksheet("Open Backlog")
    ob.hide_gridlines(2)
    ob.set_tab_color(RED if m.get("aging", {}).get("aged_over_threshold", 0) else GREEN)
    open_df = m.get("aging", {}).get("detail", pd.DataFrame()).copy()
    ob_headers = list(open_df.columns)
    last_col = max(len(ob_headers) - 1, 0)
    ob.set_row(0, 38)
    ob.merge_range(0, 0, 0, last_col,
                   f"Open Backlog at {pd.Timestamp(end_dt).strftime('%d %b %Y')}  |  "
                   f"{len(open_df):,} Open / In Progress item(s)",
                   _f(bold=True, font_size=13, font_color=WHITE, bg_color=INK, align="center"))
    ob_header = _f(bold=True, font_color=WHITE, bg_color=INK, align="center",
                   border=1, border_color=INK, text_wrap=True)
    for ci, h in enumerate(ob_headers):
        ob.write(2, ci, str(h), ob_header)
        width = max(len(str(h)) + 3, 13)
        if h in {COL_CREATION_DATE, COL_CLOSURE_DATE, "Parsed Created Date", "Parsed Closure Date"}:
            width = 22
        elif h in {"Comment", "Summary", "Description"}:
            width = 36
        elif h in {"Age Bucket", "Normalized Status", "Module"}:
            width = 20
        ob.set_column(ci, ci, min(width, 42))
    ob.freeze_panes(3, 0)
    if ob_headers and len(open_df):
        for ri, vals in enumerate(open_df.itertuples(index=False, name=None), start=3):
            alt = (ri % 2 == 0)
            for ci, v in enumerate(vals):
                fmt = _f(align="left", border=1, border_color=LINE,
                         bg_color=LINE_SOFT if alt else WHITE, font_color=INK)
                if ob_headers[ci] in {COL_CREATION_DATE, COL_CLOSURE_DATE, "Parsed Created Date", "Parsed Closure Date"}:
                    fmt = _f(align="left", border=1, border_color=LINE,
                             bg_color=LINE_SOFT if alt else WHITE, font_color=INK,
                             num_format="dd mmm yyyy hh:mm")
                try:
                    nil = pd.isna(v)
                except (TypeError, ValueError):
                    nil = False
                if nil:
                    ob.write_blank(ri, ci, None, fmt)
                elif isinstance(v, (pd.Timestamp, datetime)):
                    ob.write_datetime(ri, ci, pd.Timestamp(v).to_pydatetime(), fmt)
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    ob.write_number(ri, ci, v, fmt)
                else:
                    ob.write(ri, ci, str(v) if not isinstance(v, str) else v, fmt)
        ob.autofilter(2, 0, 2 + len(open_df), len(ob_headers) - 1)
        if "Age (days)" in ob_headers:
            age_col = ob_headers.index("Age (days)")
            ob.conditional_format(3, age_col, 2 + len(open_df), age_col, {
                "type": "cell", "criteria": ">", "value": AGED_DAYS,
                "format": _f(bg_color=RED_L, font_color=RED_D, bold=True)})
    elif ob_headers:
        ob.merge_range(3, 0, 3, last_col, "No Open or In Progress incidents at the reporting cutoff.",
                       _f(bold=True, font_size=11, font_color=GREEN_D, bg_color=GREEN_L, align="left"))
    else:
        ob.merge_range(2, 0, 2, 2, "No open backlog columns were available.",
                       _f(italic=True, font_color=MUTED))

    # ---- U (union) data sheets ----------------------------------------------
    drop = {"_Module", "_SrcRow", "_creation", "_closure", "_Status", "_Id"}
    date_cols = {c for c in (COL_CREATION_DATE, COL_CLOSURE_DATE) if c}
    union = m["union"]

    used_names = set()

    def _safe_name(base: str) -> str:
        name = re.sub(r"[\[\]:*?/\\]", "-", f"U - {base}")[:31]
        n = 1
        while name.lower() in used_names:
            suffix = f"~{n}"
            name = name[:31 - len(suffix)] + suffix
            n += 1
        used_names.add(name.lower())
        return name

    for mod in mods:
        sub = union[union["_Module"] == mod].copy()
        if sub.empty:
            continue
        # write parsed dates back so the data sheet shows real Excel dates
        for col, src in ((COL_CREATION_DATE, "_creation"), (COL_CLOSURE_DATE, "_closure")):
            if col and col in sub.columns:
                sub[col] = sub[src]
        sub = sub.drop(columns=[c for c in drop if c in sub.columns])
        sub = sub.dropna(axis=1, how="all")
        if sub.empty or not len(sub.columns):
            continue

        dw = wb.add_worksheet(_safe_name(str(mod)[:25]))
        dw.hide_gridlines(2)
        headers = list(sub.columns)
        dw.set_default_row(16)
        dw.set_row(0, 22)
        for ci, h in enumerate(headers):
            dw.write(0, ci, str(h), f_data_hdr)
            width = max(len(str(h)) + 4, 14)
            if h in date_cols:
                width = 22
            dw.set_column(ci, ci, min(width, 55))
        vals = sub.values
        di = {i for i, h in enumerate(headers) if h in date_cols}
        for ri in range(len(sub)):
            er = ri + 1
            alt = (er % 2 == 0)
            for ci in range(len(headers)):
                v = vals[ri, ci]
                isd = ci in di
                try:
                    nil = pd.isna(v)
                except (TypeError, ValueError):
                    nil = False
                if nil:
                    dw.write_blank(er, ci, None,
                                   (f_date_alt if alt else f_date) if isd
                                   else (f_cell_alt if alt else f_cell))
                elif isinstance(v, (pd.Timestamp, datetime)):
                    dw.write_datetime(er, ci, pd.Timestamp(v).to_pydatetime(),
                                      f_date_alt if alt else f_date)
                else:
                    dw.write(er, ci, v, f_cell_alt if alt else f_cell)
        dw.freeze_panes(1, 0)
        dw.autofilter(0, 0, len(sub), len(headers) - 1)

    sw.activate()
    sw.set_first_sheet()
    wb.close()



# =============================================================================
# KPI TILES — one definition, used by BOTH the Excel dashboard and the slide
# =============================================================================

def _kpi_tiles(m: dict) -> list[tuple]:
    """Return (label, value, delta_text, caption, accent_colour, delta_colour) x5."""
    prior = m.get("prior_period")
    prior_label = m.get("prior_label", "prior period")
    total = m["total_union"]
    res_rate = m.get("resolution_rate") or 0.0

    def count_delta(current, previous):
        if previous is None:
            return "No prior-period data"
        if previous == 0:
            return "New vs prior 0" if current else f"Flat vs {prior_label}"
        change = (current - previous) / previous * 100.0
        if abs(change) < 0.5:
            return f"▬ Flat vs {prior_label}"
        return f"{'▲' if change > 0 else '▼'} {abs(change):.0f}% vs {prior_label}"

    def point_delta(current, previous):
        if previous is None or current is None:
            return "No prior-period data", MUTED
        change = current - previous
        if abs(change) < 0.05:
            return f"Flat vs {prior_label}", MUTED
        return (f"{'▲' if change > 0 else '▼'} {abs(change):.1f} pts vs {prior_label}",
                GREEN_D if change > 0 else RED_D)

    net_change = int(m["n_created"] - m["n_closed"])
    res_delta, res_delta_col = point_delta(
        m.get("resolution_rate"), prior.get("resolution_rate") if prior else None)

    if m.get("mttr_med") is not None and prior and prior.get("mttr_med") is not None:
        prev = prior["mttr_med"]
        if not prev:
            mttr_delta, mttr_col = "Prior median was 0 days", MUTED
        else:
            pct = (m["mttr_med"] - prev) / prev * 100.0
            if abs(pct) < 0.5:
                mttr_delta, mttr_col = f"Flat vs {prior_label}", MUTED
            else:
                mttr_delta = f"{'▼' if pct < 0 else '▲'} {abs(pct):.0f}% vs {prior_label}"
                mttr_col = GREEN_D if pct < 0 else RED_D
    else:
        mttr_delta, mttr_col = "No prior-period data", MUTED

    prior_net_text = (f"Prior-period net: {prior['net_backlog_change']:+,}"
                      if prior else "No prior-period data")
    if net_change > 0:
        net_colour, net_caption = RED, "Backlog growing · created > closed"
    elif net_change < 0:
        net_colour, net_caption = GREEN, "Backlog shrinking · closed > created"
    else:
        net_colour, net_caption = GREY, "Backlog flat · created = closed"

    return [
        ("INCIDENTS IN PERIOD", f"{total:,}",
         count_delta(total, prior["total"] if prior else None),
         f"{m['n_created']:,} created + {m['n_closed_from_earlier']:,} closed from earlier",
         ACCENT, MUTED),
        ("INCIDENTS CLOSED", f"{m['n_closed']:,}",
         count_delta(m["n_closed"], prior["n_closed"] if prior else None),
         "Closure-date events in this period", GREEN, MUTED),
        ("RESOLUTION RATE", f"{res_rate:.0f}%", res_delta,
         "Closed status ÷ incidents in period",
         GREEN if res_rate >= 80 else AMBER if res_rate >= 60 else RED, res_delta_col),
        ("NET BACKLOG CHANGE", f"{net_change:+,}", prior_net_text, net_caption,
         net_colour, RED_D if net_change > 0 else GREEN_D if net_change < 0 else MUTED),
        ("MEDIAN TIME TO RESOLVE", _days_text(m.get("mttr_med")), mttr_delta,
         f"Mean {_days_text(m.get('mttr_mean'))} · closed in period, valid dates",
         ACCENT_D, mttr_col),
    ]


# =============================================================================
# ONE-SLIDE EXECUTIVE POWERPOINT  (built only on request)
# =============================================================================
# Design rules: native, editable charts/tables; status colours identical to the
# Excel dashboard; nothing on the slide is a warning; the bottom-left panel is
# deliberately empty so the monthly highlights can be typed or pasted straight in.

_PPT = None


def _load_pptx():
    """Import python-pptx lazily so Excel-only runs never need it."""
    global _PPT
    if _PPT is not None:
        return _PPT
    import types
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import (XL_CHART_TYPE, XL_LABEL_POSITION,
                                 XL_LEGEND_POSITION, XL_MARKER_STYLE, XL_TICK_MARK)
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Inches, Pt
    from lxml import etree
    _PPT = types.SimpleNamespace(
        Presentation=Presentation, CategoryChartData=CategoryChartData,
        RGBColor=RGBColor, XL_CHART_TYPE=XL_CHART_TYPE,
        XL_LABEL_POSITION=XL_LABEL_POSITION, XL_LEGEND_POSITION=XL_LEGEND_POSITION,
        XL_MARKER_STYLE=XL_MARKER_STYLE, XL_TICK_MARK=XL_TICK_MARK,
        MSO_SHAPE=MSO_SHAPE, MSO_ANCHOR=MSO_ANCHOR, MSO_AUTO_SIZE=MSO_AUTO_SIZE,
        PP_ALIGN=PP_ALIGN, qn=qn, Inches=Inches, Pt=Pt, etree=etree)
    return _PPT


def _rgb(value: str):
    return _PPT.RGBColor.from_string(value.lstrip("#"))


def _ppt_rect(slide, x, y, w, h, fill, line=None, radius=True):
    P = _PPT
    kind = P.MSO_SHAPE.ROUNDED_RECTANGLE if radius else P.MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, P.Inches(x), P.Inches(y), P.Inches(w), P.Inches(h))
    if radius:
        shape.adjustments[0] = 0.035
    shape.shadow.inherit = False
    style = shape._element.find(P.qn("p:style"))      # drop theme effects (shadows) for every viewer
    if style is not None:
        shape._element.remove(style)
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(fill)
    if line:
        shape.line.color.rgb = _rgb(line)
        shape.line.width = P.Pt(0.75)
    else:
        shape.line.fill.background()
    return shape


def _ppt_text(slide, x, y, w, h, text, size=10, color=SLATE, bold=False,
              align=None, valign=None, italic=False):
    P = _PPT
    shape = slide.shapes.add_textbox(P.Inches(x), P.Inches(y), P.Inches(w), P.Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = P.MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = P.Inches(0.02)
    tf.margin_top = tf.margin_bottom = P.Inches(0.01)
    tf.vertical_anchor = valign or P.MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align or P.PP_ALIGN.LEFT
    r = p.add_run()
    r.text = str(text)
    r.font.name = PPT_FONT
    r.font.size = P.Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = _rgb(color)
    return shape


def _ppt_cell(cell, text, size=8.5, bold=False, color=SLATE, fill=None,
              align=None, bottom_line=None):
    """Format one native-table cell (fill, font, alignment, optional hairline)."""
    P = _PPT
    cell.margin_left = cell.margin_right = P.Inches(0.08)
    cell.margin_top = cell.margin_bottom = P.Inches(0.02)
    cell.vertical_anchor = P.MSO_ANCHOR.MIDDLE
    p = cell.text_frame.paragraphs[0]
    p.alignment = align or P.PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(text)
    r.font.name = PPT_FONT
    r.font.size = P.Pt(size)
    r.font.bold = bold
    r.font.color.rgb = _rgb(color)
    if fill:
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(fill)
    else:
        cell.fill.background()
    if bottom_line:
        tcPr = cell._tc.get_or_add_tcPr()
        ln = P.etree.Element(P.qn("a:lnB"), w=str(int(0.75 * 12700)), cap="flat",
                             cmpd="sng", algn="ctr")
        sf = P.etree.SubElement(ln, P.qn("a:solidFill"))
        P.etree.SubElement(sf, P.qn("a:srgbClr"), val=bottom_line.lstrip("#"))
        P.etree.SubElement(ln, P.qn("a:prstDash"), val="solid")
        tcPr.insert(0, ln)          # borders must precede the fill element


def _ppt_panel(slide, x, y, w, h, title, accent=ACCENT):
    _ppt_rect(slide, x, y, w, h, WHITE, LINE)
    _ppt_rect(slide, x, y, w, 0.05, accent, radius=False)
    _ppt_text(slide, x + 0.18, y + 0.10, w - 0.36, 0.24, title, 9, INK, True)


def build_powerpoint(m: dict, start_dt, end_dt, output_path: str) -> bool:
    """Write the one-slide executive summary from the same metric bundle as Excel."""
    try:
        P = _load_pptx()
    except ImportError:
        print("  PowerPoint NOT created: python-pptx is not installed "
              "(pip install python-pptx). The Excel report is unaffected.")
        return False

    prs = P.Presentation()
    prs.slide_width = P.Inches(13.333)
    prs.slide_height = P.Inches(7.5)
    prs.core_properties.title = f"Incident Review — {_period_label(start_dt, end_dt)}"
    prs.core_properties.subject = "Monthly executive incident summary"
    prs.core_properties.author = "Incident Report Generator"
    slide = prs.slides.add_slide(prs.slide_layouts[6])         # blank layout
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(CANVAS)

    LM, CW = 0.44, 12.453                                      # left margin, content width
    total = int(m["total_union"])

    # ------------------------------------------------------------ header band
    _ppt_rect(slide, 0, 0, 13.333, 0.80, INK, radius=False)
    _ppt_rect(slide, 0, 0, 0.09, 0.80, ACCENT, radius=False)
    _ppt_text(slide, LM, 0.09, 6.2, 0.18, "MONTHLY INCIDENT REVIEW", 8.5, "A5B4FC", True)
    _ppt_text(slide, LM, 0.29, 8.6, 0.38,
              f"Incident performance — {_period_label(start_dt, end_dt)}", 21, WHITE, True)
    _ppt_text(slide, 9.2, 0.29, 3.69, 0.30, f"{start_dt:%d %b %Y} – {end_dt:%d %b %Y}",
              10.5, "CBD5E1", False, align=P.PP_ALIGN.RIGHT)

    # ------------------------------------------------------------- KPI cards
    tiles = _kpi_tiles(m)
    gap = 0.14
    card_w = (CW - gap * (len(tiles) - 1)) / len(tiles)
    card_y, card_h = 0.98, 1.04
    for i, (label, value, delta, caption, accent, delta_col) in enumerate(tiles):
        x = LM + i * (card_w + gap)
        _ppt_rect(slide, x, card_y, card_w, card_h, WHITE, LINE)
        _ppt_rect(slide, x, card_y, card_w, 0.055, accent, radius=False)
        _ppt_text(slide, x + 0.14, card_y + 0.11, card_w - 0.28, 0.17, label, 7.5, MUTED, True)
        _ppt_text(slide, x + 0.14, card_y + 0.28, card_w - 0.28, 0.38, value, 22, INK, True)
        _ppt_text(slide, x + 0.14, card_y + 0.66, card_w - 0.28, 0.15, delta, 7.5, delta_col, True)
        _ppt_text(slide, x + 0.14, card_y + 0.82, card_w - 0.28, 0.16, caption, 6.8, MUTED)

    # ---------------------------------------------- middle left: LINE chart
    mid_y, mid_h = 2.20, 2.40
    lx, lw = LM, 6.25
    _ppt_panel(slide, lx, mid_y, lw, mid_h,
               f"INCIDENT FLOW · CREATED VS CLOSED · LAST {len(m['trend_labels'])} MONTHS")
    labels = list(m["trend_labels"])
    created_v = [int(v) for v in m["trend_created"]]
    closed_v = [int(v) for v in m["trend_closed"]]
    cdata = P.CategoryChartData()
    cdata.categories = labels
    cdata.add_series("Created", created_v)
    cdata.add_series("Closed", closed_v)
    gf = slide.shapes.add_chart(P.XL_CHART_TYPE.LINE_MARKERS, P.Inches(lx + 0.08),
                                P.Inches(mid_y + 0.38), P.Inches(lw - 0.16),
                                P.Inches(mid_h - 0.44), cdata)
    chart = gf.chart
    chart.has_title = False
    chart.font.name = PPT_FONT
    chart.font.size = P.Pt(9)
    chart.font.color.rgb = _rgb(MUTED)
    chart.has_legend = True
    chart.legend.position = P.XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False
    chart.legend.font.size = P.Pt(9)
    chart.legend.font.color.rgb = _rgb(SLATE)
    top = max(created_v + closed_v + [1])
    va, ca = chart.value_axis, chart.category_axis
    va.minimum_scale = 0
    va.maximum_scale = math.ceil(top * 1.25)
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = _rgb(LINE)
    va.major_gridlines.format.line.width = P.Pt(0.5)
    va.format.line.fill.background()
    va.major_tick_mark = P.XL_TICK_MARK.NONE
    va.tick_labels.font.size = P.Pt(8.5)
    va.tick_labels.number_format = "0"
    va.tick_labels.number_format_is_linked = False
    ca.format.line.color.rgb = _rgb(LINE)
    ca.major_tick_mark = P.XL_TICK_MARK.NONE
    ca.tick_labels.font.size = P.Pt(9)
    ca.tick_labels.font.bold = True
    ca.tick_labels.font.color.rgb = _rgb(INK)
    for si, (ser, vals, colour, dark) in enumerate(
            zip(chart.plots[0].series, (created_v, closed_v), (ACCENT, GREEN), (ACCENT_D, GREEN_D))):
        ser.smooth = False
        ser.format.line.color.rgb = _rgb(colour)
        ser.format.line.width = P.Pt(2.5)
        ser.marker.style = P.XL_MARKER_STYLE.CIRCLE
        ser.marker.size = 8
        ser.marker.format.fill.solid()
        ser.marker.format.fill.fore_color.rgb = _rgb(colour)
        ser.marker.format.line.color.rgb = _rgb(WHITE)
        for pi, v in enumerate(vals):
            other = closed_v[pi] if si == 0 else created_v[pi]
            above = (v >= other) if si == 0 else (v > other)   # keeps labels apart
            dl = ser.points[pi].data_label
            dl.position = P.XL_LABEL_POSITION.ABOVE if above else P.XL_LABEL_POSITION.BELOW
            dl.font.size = P.Pt(9)
            dl.font.bold = True
            dl.font.color.rgb = _rgb(dark)

    # -------------------------- middle right: coloured breakdown by module
    rx, rw = LM + lw + 0.14, CW - lw - 0.14
    _ppt_panel(slide, rx, mid_y, rw, mid_h,
               f"INCIDENT BREAKDOWN BY MODULE · {total:,} IN PERIOD")
    names = m["module_names"]
    if len(names) > TOP_N_MODULES + 1:
        shown, rest = names[:TOP_N_MODULES], names[TOP_N_MODULES:]
    else:
        shown, rest = names, []
    rows = []
    for nme in shown:
        sb = m["module_status"][nme]
        rows.append((str(nme)[:26], sum(sb.values()), [sb.get(s, 0) for s in STATUS_ORDER]))
    if rest:
        sums = [sum(m["module_status"][n].get(s, 0) for n in rest) for s in STATUS_ORDER]
        rows.append((f"All other ({len(rest)} modules)", sum(sums), sums))
    grand = [sum(m["module_status"][n].get(s, 0) for n in names) for s in STATUS_ORDER]
    n_rows = len(rows) + 2
    row_h = 0.265
    tbl_gf = slide.shapes.add_table(n_rows, 6, P.Inches(rx + 0.12), P.Inches(mid_y + 0.43),
                                    P.Inches(rw - 0.24), P.Inches(row_h * n_rows))
    tbl = tbl_gf.table
    tbl.first_row = False
    tbl.horz_banding = False
    style_id = tbl_gf._element.graphic.graphicData.tbl.tblPr.find(P.qn("a:tableStyleId"))
    if style_id is not None:
        style_id.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"      # "No Style, No Grid"
    inner_w = rw - 0.24
    col_w = [0.335, 0.12, 0.14, 0.16, 0.12, 0.125]
    scale = inner_w / sum(col_w)
    for ci, cw_ in enumerate(col_w):
        tbl.columns[ci].width = P.Inches(cw_ * scale)
    for ri in range(n_rows):
        tbl.rows[ri].height = P.Inches(row_h)
    L, C = P.PP_ALIGN.LEFT, P.PP_ALIGN.CENTER
    _ppt_cell(tbl.cell(0, 0), "Module", 8.5, True, WHITE, INK, L)
    _ppt_cell(tbl.cell(0, 1), "Total", 8.5, True, WHITE, INK, C)
    for j, st in enumerate(STATUS_ORDER):
        _ppt_cell(tbl.cell(0, 2 + j), st, 8.5, True, WHITE, _STATUS_RAG[st], C)
    for ri, (nme, tot, vals) in enumerate(rows, start=1):
        _ppt_cell(tbl.cell(ri, 0), nme, 8.5, True, INK, WHITE, L, bottom_line=LINE)
        _ppt_cell(tbl.cell(ri, 1), f"{tot:,}", 9, True, INK, WHITE, C, bottom_line=LINE)
        for j, st in enumerate(STATUS_ORDER):
            _ppt_cell(tbl.cell(ri, 2 + j), f"{vals[j]:,}", 9, True,
                      _STATUS_DARK[st], _STATUS_TINT[st], C, bottom_line=WHITE)
    tr = n_rows - 1
    _ppt_cell(tbl.cell(tr, 0), "TOTAL", 8.5, True, ACCENT_D, ACCENT_L, L)
    _ppt_cell(tbl.cell(tr, 1), f"{total:,}", 9, True, ACCENT_D, ACCENT_L, C)
    for j, st in enumerate(STATUS_ORDER):
        _ppt_cell(tbl.cell(tr, 2 + j), f"{grand[j]:,}", 9, True, ACCENT_D, ACCENT_L, C)

    # ----------------------------- bottom left: BLANK panel for the highlights
    bot_y, bot_h = 4.78, 2.18
    bw = 7.25
    hl = slide.shapes.add_textbox(P.Inches(LM), P.Inches(bot_y), P.Inches(bw), P.Inches(bot_h))
    hl.name = "Monthly highlights (click and type or paste)"
    hl.fill.solid()
    hl.fill.fore_color.rgb = _rgb(WHITE)
    hl.line.color.rgb = _rgb(LINE)
    hl.line.width = P.Pt(0.75)
    tf = hl.text_frame
    tf.word_wrap = True
    tf.auto_size = P.MSO_AUTO_SIZE.NONE               # stay full size even while empty
    tf.vertical_anchor = P.MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = P.Inches(0.22)
    tf.margin_top, tf.margin_bottom = P.Inches(0.50), P.Inches(0.15)
    # typing style for whatever gets pasted in (dark slate, 12 pt, same font)
    epr = tf.paragraphs[0]._p.get_or_add_endParaRPr()
    epr.set("lang", "en-US")
    epr.set("sz", "1200")
    sf = P.etree.SubElement(epr, P.qn("a:solidFill"))
    P.etree.SubElement(sf, P.qn("a:srgbClr"), val=SLATE.lstrip("#"))
    P.etree.SubElement(epr, P.qn("a:latin"), typeface=PPT_FONT)
    _ppt_rect(slide, LM, bot_y, bw, 0.05, ACCENT, radius=False)
    _ppt_text(slide, LM + 0.18, bot_y + 0.10, bw - 0.36, 0.24, PPT_NOTES_TITLE, 9, INK, True)

    # ------------------------- bottom right: coloured status breakdown
    sx = LM + bw + 0.14
    sw_ = CW - bw - 0.14
    _ppt_panel(slide, sx, bot_y, sw_, bot_h, f"STATUS BREAKDOWN · {total:,} IN PERIOD")
    dcd = P.CategoryChartData()
    dcd.categories = list(m["status_labels"])
    dcd.add_series("Incidents", [int(v) for v in m["status_counts"]])
    d_x, d_y, d_w, d_h = sx + 0.10, bot_y + 0.40, 1.82, 1.70
    dgf = slide.shapes.add_chart(P.XL_CHART_TYPE.DOUGHNUT, P.Inches(d_x), P.Inches(d_y),
                                 P.Inches(d_w), P.Inches(d_h), dcd)
    dch = dgf.chart
    dch.has_title = False
    dch.has_legend = False
    dch.plots[0].has_data_labels = False
    dch.plots[0].vary_by_categories = True
    for pi, lbl in enumerate(m["status_labels"]):
        pt = dch.plots[0].series[0].points[pi]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = _rgb(_STATUS_RAG[lbl])
        pt.format.line.color.rgb = _rgb(WHITE)
        pt.format.line.width = P.Pt(1.5)
    for hs in dch._chartSpace.xpath(".//c:holeSize"):
        hs.set("val", "64")
    _ppt_text(slide, d_x + d_w / 2 - 0.45, d_y + d_h / 2 - 0.25, 0.9, 0.32,
              f"{total:,}", 16, INK, True, align=P.PP_ALIGN.CENTER)
    _ppt_text(slide, d_x + d_w / 2 - 0.45, d_y + d_h / 2 + 0.05, 0.9, 0.16,
              "incidents", 7, MUTED, align=P.PP_ALIGN.CENTER)

    lx0 = sx + 2.05
    lw0 = sw_ - 2.05 - 0.18
    ly = bot_y + 0.50
    for lbl, cnt in zip(m["status_labels"], m["status_counts"]):
        _ppt_rect(slide, lx0, ly + 0.07, 0.13, 0.13, _STATUS_RAG[lbl], radius=False)
        _ppt_text(slide, lx0 + 0.22, ly, lw0 - 1.30, 0.27, lbl, 9, SLATE, True)
        _ppt_text(slide, lx0 + lw0 - 1.08, ly, 0.55, 0.27, f"{cnt:,}", 10, _STATUS_DARK[lbl],
                  True, align=P.PP_ALIGN.RIGHT)
        _ppt_text(slide, lx0 + lw0 - 0.50, ly, 0.50, 0.27,
                  f"{cnt / (total or 1) * 100:.0f}%", 9, MUTED, False,
                  align=P.PP_ALIGN.RIGHT)
        ly += 0.31
    _ppt_rect(slide, lx0, ly + 0.02, lw0, 0.01, LINE, radius=False)
    _ppt_text(slide, lx0 + 0.22, ly + 0.07, lw0 - 1.30, 0.27, "Total", 9, ACCENT_D, True)
    _ppt_text(slide, lx0 + lw0 - 1.08, ly + 0.07, 0.55, 0.27, f"{total:,}", 10, ACCENT_D,
              True, align=P.PP_ALIGN.RIGHT)
    _ppt_text(slide, lx0 + lw0 - 0.50, ly + 0.07, 0.50, 0.27, "100%", 9, MUTED,
              False, align=P.PP_ALIGN.RIGHT)

    # ---------------------------------------------------------------- footer
    _ppt_text(slide, LM, 7.10, CW, 0.20,
              "Incidents in period = created in period + closed in period but created earlier "
              "(each counted once)  ·  Status = current status in the incident log  ·  "
              f"Generated {datetime.now():%d %b %Y}", 7, MUTED)

    prs.save(output_path)
    if len(prs.slides) != 1:
        raise AssertionError(f"Executive deck must contain exactly one slide; found {len(prs.slides)}")
    return True


# =============================================================================
# MAIN
# =============================================================================

def _cli():
    p = argparse.ArgumentParser(description="Incident report generator "
                                "(Excel always; PowerPoint only on request)")
    p.add_argument("-i", "--input", default=INPUT_FILE_PATH)
    p.add_argument("-o", "--output-folder", default=OUTPUT_FOLDER)
    p.add_argument("-s", "--start", default=None, help='e.g. "1 Jan 2024"')
    p.add_argument("-e", "--end", default=None, help='e.g. "31 Jan 2024"')
    g = p.add_mutually_exclusive_group()
    g.add_argument("--pptx", action="store_true", help="also create the one-slide PowerPoint")
    g.add_argument("--no-pptx", action="store_true", help="do not create the PowerPoint (and do not ask)")
    return p.parse_args()


def wants_powerpoint(args) -> bool:
    """Excel is always built. The slide only on an explicit yes."""
    if args.pptx:
        return True
    if args.no_pptx:
        return False
    if sys.stdin is not None and sys.stdin.isatty():
        ans = input("  Create the PowerPoint as well? (y/N): ").strip().lower()
        return ans in ("y", "yes")
    return PPTX_DEFAULT


def main():
    args = _cli()

    print("\n" + "=" * 62)
    print("  INCIDENT REPORT GENERATOR  v3.2")
    print("=" * 62)
    if not _HAS_DATEPARSER:
        print("  (optional) pip install dateparser  -> extra date-format coverage")

    issues: list[dict] = []
    master, modules = load_master(args.input, issues)
    master = prepare(master, modules, issues)

    lo = pd.concat([master["_creation"], master["_closure"]]).min()
    hi = pd.concat([master["_creation"], master["_closure"]]).max()
    if pd.isna(lo):
        sys.exit("\nERROR: no usable dates found in either date column.\n")
    print(f"\n  Dates available in file: {lo:%d %b %Y}  ->  {hi:%d %b %Y}")

    if args.start:
        start_dt = parse_single_date(args.start)
        if start_dt is None:
            sys.exit(f"Could not read --start '{args.start}'")
    else:
        start_dt = prompt_date("START date (inclusive)", lo.normalize())
    if args.end:
        end_dt = parse_single_date(args.end)
        if end_dt is None:
            sys.exit(f"Could not read --end '{args.end}'")
    else:
        end_dt = prompt_date("END   date (inclusive)", hi.normalize())

    if end_dt < start_dt:
        start_dt, end_dt = end_dt, start_dt
        print("  (dates swapped — start was after end)")
    start_dt = pd.Timestamp(start_dt).normalize()
    end_dt   = pd.Timestamp(end_dt).normalize() + pd.Timedelta(
        hours=23, minutes=59, seconds=59)
    print(f"\n  Range: {start_dt:%d %b %Y}  ->  {end_dt:%d %b %Y}\n")

    make_ppt = wants_powerpoint(args)

    metrics = build_metrics(master, modules, start_dt, end_dt, issues)
    metrics["current_period"] = _period_snapshot(master, start_dt, end_dt)
    if metrics["current_period"]["total"] != metrics["total_union"]:
        raise AssertionError(
            f"Current-period reconciliation failed: snapshot={metrics['current_period']['total']:,}, "
            f"dashboard={metrics['total_union']:,}"
        )
    prev_start, prev_end, prior_label = _prior_period_window(start_dt, end_dt)
    prior_summary = _period_snapshot(master, prev_start, prev_end)
    metrics["prior_period"] = prior_summary if prior_summary["total"] else None
    metrics["prior_label"] = prior_label
    metrics["aging"] = _aging_snapshot(master, end_dt)
    metrics["quality_flag_count"] = (
        sum(1 for issue in issues if issue.get("Type") != "Header auto-matched")
        + sum(_unmapped_statuses.values())
    )

    print(f"\n  Modules in period: {len(metrics['module_names'])}")
    for mod, tot in zip(metrics["module_names"], metrics["module_totals"]):
        sb = metrics["module_status"][mod]
        print(f"    {mod:<24} total={tot:<5} "
              f"closed={sb['Closed']:<4} inprog={sb['In Progress']:<4} "
              f"open={sb['Open']:<4} other={sb['Other']}")

    if metrics["mttr_mean"] is not None:
        print(f"\n  Avg resolution time: {metrics['mttr_mean']:.1f} days "
              f"(median {metrics['mttr_med']:.1f})")
    print(f"  All-time status mix ({metrics['alltime_total']:,} incidents): "
          + ", ".join(f"{l} {c:,}" for l, c in zip(metrics["alltime_labels"],
                                                   metrics["alltime_counts"])))

    if _unmapped_statuses:
        print("\n  Unmapped status values (bucketed as 'Other'):")
        for raw, n in sorted(_unmapped_statuses.items(), key=lambda kv: -kv[1])[:15]:
            print(f"    '{raw}'  x{n}")

    os.makedirs(args.output_folder, exist_ok=True)
    stem = f"Incident Review - {start_dt:%d %b %Y} - {end_dt:%d %b %Y}"
    out = os.path.join(args.output_folder, stem + ".xlsx")
    build_workbook(metrics, modules, master, issues, start_dt, end_dt, out)
    print(f"\n  Excel report saved -> {out}")

    if make_ppt:
        ppt_out = os.path.join(args.output_folder, stem + ".pptx")
        if build_powerpoint(metrics, start_dt, end_dt, ppt_out):
            print(f"  PowerPoint saved   -> {ppt_out} (1 slide)")
    else:
        print("  PowerPoint not requested (add --pptx or answer 'y' when asked).")

    print(f"\n  Data quality items logged: {len(issues):,}")
    print("Done.\n")


if __name__ == "__main__":
    main()
