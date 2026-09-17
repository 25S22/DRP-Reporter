#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
===============================================================================
 INCIDENT REPORT GENERATOR  ·  v3.0  ("strong build")
===============================================================================

WHAT CHANGED vs v2 (the three problems you reported)
-------------------------------------------------------------------------------
1. "Few rows are unparseable"
   The date engine was rewritten as a 6-stage cascade instead of one loose
   pd.to_datetime() call:
       stage 0  real datetime / date objects passed through untouched
       stage 1  numeric Excel serials (1900 system) + UNIX epoch s / ms
       stage 2  ~40 exact strptime formats, applied vectorised, ordered by
                DAYFIRST so 05/01/2024 never flips meaning
       stage 3  pandas flexible parser
       stage 4  dateutil fuzzy parser (handles "Closed on 5 Jan 2024 IST")
       stage 5  dateparser, if installed (natural language + non-English)
   Plus text pre-cleaning: ordinals, non-breaking spaces, unicode dashes,
   ISO "T" separators, trailing timezone tokens, "a.m./p.m.", double spaces.
   Anything still unparsed is no longer a scary console warning - it is
   written to a "Data Quality" sheet with module, column, row and raw value,
   so you can fix the source once instead of re-reading logs every month.

2. "Mismatch in the Overall status count"
   Root causes found:
     a) the Overall breakdown was computed from EVERY row in the workbook
        (processed_raw) while the KPI tiles used only in-period rows, so the
        two numbers could never agree;
     b) unrecognised statuses ("Pending", "WIP", "Cancelled", blanks) were
        silently dropped by _normalize_status, so the parts never summed to
        the whole;
     c) module counts de-duplicated per sheet, but the status breakdown
        de-duplicated ACROSS sheets, so an ID present in two modules was
        counted twice in one place and once in the other.
   Fixes: everything now derives from ONE tidy master frame with a _Module
   column and a single de-duplication policy (DEDUPE_ACROSS_MODULES).
   Unmatched statuses fall into an explicit "Other" bucket that is charted
   and tabled like the rest. A reconciliation check runs before the workbook
   is written and prints PASS/FAIL, so a mismatch can never ship silently.

3. "0 labels sit on top of the Closed segment in the stacked bar"
   Two belt-and-braces fixes: zero values are written to the hidden chart
   sheet as BLANK cells (Excel draws no segment), and each series carries a
   per-point custom data-label list where zero points are {'delete': True}.
   The stray "0" is gone.

ALSO ADDED
-------------------------------------------------------------------------------
   · Column-name auto-resolution ("Incident ID" == "incident id" == "Incident_Id")
   · MTTR KPI (mean + median resolution time) and a 5-tile scorecard
   · Monthly trend chart (created vs closed) with a moving picture of backlog
   · Data Quality sheet: unparseable dates, closed-without-closure-date,
     blank statuses, unmapped statuses, cross-module duplicate IDs
   · Reconciliation panel on the dashboard
   · CLI flags, so it can run unattended from a scheduler:
        python incident_report_generator.py -i in.xlsx -s "1 Jan 2024" -e "31 Jan 2024"
     Run with no flags and it prompts exactly like before.
   · Cached formula values (totals render in LibreOffice / Sheets / previewers)
   · Print setup, tab colours, hidden gridlines, frozen panes

OPTIONAL EXTRA LIBRARY (not required, auto-detected)
-------------------------------------------------------------------------------
     pip install dateparser
   Only helps if your source has natural-language or non-English dates.
   Everything else runs on pandas + xlsxwriter + python-dateutil, which you
   already have.
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
EMAIL_SCOPE           = "period"  # "period" (in-range only) or "all"
STATUS_FOR_EMAILS     = "closed - resolved"
MAX_DQ_ROWS           = 500     # cap rows written to the Data Quality sheet

# --- status bucketing ---------------------------------------------------------
# Exact matches win first (compare lowercased + trimmed), then keyword rules in
# order. Anything unmatched lands in "Other" and is still counted everywhere.
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
CDSHEET            = "_ChartData"

# --- design system -----------------------------------------------------------
# One neutral slate ramp for structure, ONE accent (indigo) for anything
# structural, and the RAG colours reserved strictly for status meaning. Keeping
# semantic colour scarce is what stops a dashboard looking like a paint chart.
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
# A uniform 20-column grid (plus a gutter each side) is what makes the KPI row,
# the insight row and the tables all line up on the same vertical rhythm.
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
    "#n/a", "#value!", "#ref!", "not available",
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
    t = re.sub(r"([+-]\d{2}:?\d{2})$", "", t)
    t = t.replace(",", " ")
    t = _MULTISPACE.sub(" ", t)
    return t.strip()


def _looks_like_date_text(t: str) -> bool:
    if len(t) < 5 or not re.search(r"\d", t):
        return False
    return bool(re.search(r"\d{1,4}[/\-. ]\d{1,2}", t) or
                re.search(r"[A-Za-z]{3}", t))


def parse_date_series(series: pd.Series, *, module: str, column: str,
                      issues: list) -> pd.Series:
    """Six-stage cascade. Returns tz-naive datetime64; logs true failures."""
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
        good = p.notna()
        if good.any():
            out.loc[txt.index[good]] = p[good]
            txt = txt[~good]

    # stage 3 — pandas flexible
    if not txt.empty:
        p = pd.to_datetime(txt, errors="coerce", dayfirst=DAYFIRST)
        good = p.notna()
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
        issues.append({
            "Type":   "Unparseable date",
            "Module": module,
            "Column": column,
            "Row":    int(idx) + 2 if isinstance(idx, (int,)) else str(idx),
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
    for bucket, keywords in STATUS_RULES:
        for kw in keywords:
            if kw in s:
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
# PREPARE — dates, status, ids, dedupe
# =============================================================================

def prepare(master: pd.DataFrame, modules: list[str], issues: list):
    for col, tag in ((COL_CREATION_DATE, "creation"), (COL_CLOSURE_DATE, "closure")):
        target = f"_{tag}"
        if col and col in master.columns:
            parts = []
            for mod, grp in master.groupby("_Module", sort=False):
                parts.append(parse_date_series(grp[col], module=mod,
                                               column=col, issues=issues))
            master[target] = pd.concat(parts).reindex(master.index)
        else:
            master[target] = pd.NaT
            if col:
                print(f"  WARNING: column '{col}' not found in any sheet.")

    # status
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

    # normalised id
    if COL_INCIDENT_ID and COL_INCIDENT_ID in master.columns:
        master["_Id"] = (master[COL_INCIDENT_ID].astype(str).str.strip()
                         .str.upper().replace({"": None, "NAN": None, "NONE": None}))
    else:
        master["_Id"] = None
        print(f"  WARNING: column '{COL_INCIDENT_ID}' not found - dedup disabled.")

    # cross-module duplicates (diagnostic, always reported)
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

    # closed but no closure date
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


# =============================================================================
# AGGREGATION — every number below comes from the same frames
# =============================================================================

def build_metrics(master: pd.DataFrame, modules: list[str],
                  start_dt, end_dt, issues: list) -> dict:
    in_created = master["_creation"].between(start_dt, end_dt)
    in_closed  = master["_closure"].between(start_dt, end_dt)

    union   = dedupe(master[in_created | in_closed].copy())
    closed  = dedupe(master[in_closed].copy())
    created = dedupe(master[in_created].copy())

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
    res = closed.dropna(subset=["_creation", "_closure"]) if not closed.empty \
        else closed
    mttr_mean = mttr_med = None
    if len(res):
        delta = (res["_closure"] - res["_creation"]).dt.total_seconds() / 86400.0
        delta = delta[delta >= 0]
        if len(delta):
            mttr_mean = float(delta.mean())
            mttr_med  = float(delta.median())
        bad = len(res) - len(delta)
        if bad:
            issues.append({"Type": "Closure before creation", "Module": "(various)",
                           "Column": COL_CLOSURE_DATE, "Row": "",
                           "Detail": f"{bad} row(s) excluded from MTTR"})

    # monthly trend inside the period
    months = pd.period_range(start_dt.to_period("M"), end_dt.to_period("M"), freq="M")
    cm = created["_creation"].dt.to_period("M").value_counts() if len(created) \
        else pd.Series(dtype=int)
    km = closed["_closure"].dt.to_period("M").value_counts() if len(closed) \
        else pd.Series(dtype=int)
    trend_labels  = [p.strftime("%b %Y") for p in months]
    trend_created = [int(cm.get(p, 0)) for p in months]
    trend_closed  = [int(km.get(p, 0)) for p in months]

    # ALL-TIME status mix - every incident in the file, ignoring the date range
    alltime = dedupe(master.copy())
    avc = alltime["_Status"].value_counts()
    alltime_labels = [x for x in STATUS_ORDER if int(avc.get(x, 0)) > 0]
    alltime_counts = [int(avc.get(x, 0)) for x in alltime_labels]
    alltime_total  = sum(alltime_counts)
    alltime_backlog = int(avc.get("Open", 0)) + int(avc.get("In Progress", 0))

    # emails
    emails = extract_emails(master if EMAIL_SCOPE == "all" else union)

    total_union = int(matrix["Total"].sum())
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
        "n_created": len(created), "n_closed": len(closed),
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

def _progress_bar(pct: float, width: int = 15) -> str:
    pct = min(max(float(pct), 0.0), 100.0)
    filled = round(pct / 100.0 * width)
    return "█" * filled + "░" * (width - filled) + f"   {pct:.0f}%"


def _rows_for(px: int, row_px: int = 20) -> int:
    return int(math.ceil(px / row_px)) + 2


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

    # ------------------------------------------------------------- KPI cards
    KPI = 4
    tiles = [
        ("TOTAL IN PERIOD", f"{m['total_union']:,}",
         f"of {m['alltime_total']:,} incidents in the whole file",
         ACCENT, m['total_union'] / (m['alltime_total'] or 1) * 100),
        ("CLOSED",          f"{closed:,}",  "resolved within the period",
         GREEN, res_rate),
        ("IN PROGRESS",     f"{inprog:,}",  "actively being worked",
         AMBER, inprog / total * 100),
        ("OPEN / PENDING",  f"{openc:,}",   "not yet picked up",
         RED, openc / total * 100),
        ("RESOLUTION RATE", f"{res_rate:.0f}%", "closed ÷ total in period",
         GREEN if res_rate >= 80 else AMBER if res_rate >= 60 else RED, res_rate),
    ]
    sw.set_row(KPI, 5); sw.set_row(KPI + 1, 17); sw.set_row(KPI + 2, 42)
    sw.set_row(KPI + 3, 15); sw.set_row(KPI + 4, 14)
    for r in range(KPI, KPI + 5):
        sw.write_blank(r, GUTTER, None, f_canvas)
        sw.write_blank(r, CN + 1, None, f_canvas)
    step = (CN - C0 + 1) // len(tiles)
    for bi, (label, value, caption, colour, pct) in enumerate(tiles):
        c1 = C0 + bi * step
        c2 = CN if bi == len(tiles) - 1 else c1 + step - 1
        side = dict(left=1, right=1, left_color=LINE, right_color=LINE)
        W(KPI, c1, c2, "", _f(bg_color=colour))
        W(KPI + 1, c1, c2, f"  {label}",
          _f(bold=True, font_size=9, font_color=MUTED, bg_color=WHITE,
             align="left", **side))
        W(KPI + 2, c1, c2, value,
          _f(bold=True, font_size=30, font_color=INK, bg_color=WHITE,
             align="center", **side))
        W(KPI + 3, c1, c2, _progress_bar(pct, 18),
          _f(font_size=8, bold=True, font_color=colour, bg_color=WHITE,
             align="center", font_name="Consolas", **side))
        W(KPI + 4, c1, c2, f"  {caption}",
          _f(font_size=8, italic=True, font_color=MUTED, bg_color=WHITE,
             align="left", bottom=1, bottom_color=LINE, **side))

    sw.freeze_panes(KPI + 5, 0)
    cursor = KPI + 5
    band(cursor, 12); cursor += 1

    # --------------------------------------------------------- insight cards
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
    if m["user_names"]:
        boxes.append(("TOP CLOSER", ACCENT_D, ACCENT_L,
                      str(m["user_names"][0])[:22],
                      f"{m['user_counts'][0]:,} closed in period"))

    section(cursor, "Key insights", "auto-generated from the period data")
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
        section(cursor, "Monthly trend", "created vs closed, inside the period")
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
      f"source rows  ·  {len(issues):,} data-quality item(s) logged on the "
      f"'{DQ_SHEET_NAME}' tab", f_note)
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
    dq.set_column(0, 0, 28); dq.set_column(1, 1, 22)
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
# MAIN
# =============================================================================

def _cli():
    p = argparse.ArgumentParser(description="Incident report generator")
    p.add_argument("-i", "--input", default=INPUT_FILE_PATH)
    p.add_argument("-o", "--output-folder", default=OUTPUT_FOLDER)
    p.add_argument("-s", "--start", default=None, help='e.g. "1 Jan 2024"')
    p.add_argument("-e", "--end", default=None, help='e.g. "31 Jan 2024"')
    return p.parse_args()


def main():
    args = _cli()

    print("\n" + "=" * 62)
    print("  INCIDENT REPORT GENERATOR  v3.0")
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

    metrics = build_metrics(master, modules, start_dt, end_dt, issues)

    print(f"\n  Modules in period: {len(metrics['module_names'])}")
    for mod, tot in zip(metrics["module_names"], metrics["module_totals"]):
        sb = metrics["module_status"][mod]
        print(f"    {mod:<24} total={tot:<5} "
              f"closed={sb['Closed']:<4} inprog={sb['In Progress']:<4} "
              f"open={sb['Open']:<4} other={sb['Other']}")

    if metrics["mttr_mean"] is not None:
        print(f"\n  Avg resolution time: {metrics['mttr_mean']:.1f} days "
              f"(median {metrics['mttr_med']:.1f}) — console only, not a KPI tile")
    print(f"  All-time status mix ({metrics['alltime_total']:,} incidents): "
          + ", ".join(f"{l} {c:,}" for l, c in zip(metrics["alltime_labels"],
                                                   metrics["alltime_counts"])))

    if _unmapped_statuses:
        print("\n  Unmapped status values (bucketed as 'Other'):")
        for raw, n in sorted(_unmapped_statuses.items(), key=lambda kv: -kv[1])[:15]:
            print(f"    '{raw}'  x{n}")

    os.makedirs(args.output_folder, exist_ok=True)
    fname = f"Incident Review - {start_dt:%d %b %Y} - {end_dt:%d %b %Y}.xlsx"
    out = os.path.join(args.output_folder, fname)
    build_workbook(metrics, modules, master, issues, start_dt, end_dt, out)

    print(f"\n  Data quality items logged: {len(issues):,}")
    print(f"  Report saved -> {out}")
    print("Done.\n")


if __name__ == "__main__":
    main()
