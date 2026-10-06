#!/usr/bin/env python3
"""
CloudSEK -> Outlook draft-reply automation (v2, aligned with api.cloudsek.com docs).

Scans an Outlook folder for CloudSEK incident alert emails, reads the Incident ID
(XVA-<number>) from the email's Summary block, checks the incident on the CloudSEK
API, and for incidents that are Closed* creates a Reply-All DRAFT containing the
latest comment. Nothing is ever sent.

CloudSEK endpoints used (base https://api.cloudsek.com, Bearer token):
  GET /v2/snapshot/incident?incident_display_id=XVA-123   -> current status
  GET /v2/changelog?entity_identifier=XVA-123             -> comment history

Documented limits: reads = 20 req/min per token; changelog = 100 req/hour per org.
The script therefore paces calls, caches per incident, and only hits the changelog
for incidents that are already closed.

Requirements (Windows, desktop Outlook signed in):
    pip install pywin32 requests pytz
"""

import argparse
import html
import json
import logging
import os
import re
import sys
import time
from datetime import datetime, timedelta

import pytz
import requests
import win32com.client

# --------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------
OUTLOOK_FOLDER_PATH = "Inbox/CloudSEK_Alerts"
LOOKBACK_DAYS = 3
LOCAL_TIMEZONE = "Asia/Kolkata"

API_BASE = os.getenv("CLOUDSEK_API_BASE", "https://api.cloudsek.com")
SNAPSHOT_PATH = "/v2/snapshot/incident"
CHANGELOG_PATH = "/v2/changelog"
CLOUDSEK_TOKEN = os.getenv("CLOUDSEK_TOKEN", "")

REQUEST_TIMEOUT = 30
MAX_RETRIES = 5
DEFAULT_RETRY_AFTER = 30
MIN_SECONDS_BETWEEN_CALLS = 3.1   # docs: pace reads ~3s apart (20/min limit)

STATE_FILE = "processed_incidents.json"

# Matches XVA-24388918 style IDs (docs: Display ID = XVA-<number>)
INCIDENT_ID_PATTERN = re.compile(r"\bXVA-[A-Za-z0-9]+\b")
# "Incident ID" label in the email's Summary block, then the XVA id within a short window
LABELLED_ID_PATTERN = re.compile(r"Incident\s*ID[^A-Za-z0-9]{0,40}?(XVA-[A-Za-z0-9]+)", re.IGNORECASE)

OL_FOLDER_INBOX = 6
OL_MAIL_CLASS = 43
PR_LAST_VERB_EXECUTED = "http://schemas.microsoft.com/mapi/proptag/0x10810003"
REPLIED_VERBS = {102, 103, 104}   # reply, reply-all, forward

ACCENT = "#9b59b6"
ACCENT_LIGHT = "#bb86fc"

DEBUG_DUMP = False   # set by --debug-dump

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    stream=sys.stdout,
)
log = logging.getLogger("cloudsek-drafts")

_last_call = 0.0


# --------------------------------------------------------------------------
# State
# --------------------------------------------------------------------------
def load_state():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as fh:
            return set(json.load(fh))
    except (FileNotFoundError, json.JSONDecodeError):
        return set()


def save_state(processed):
    with open(STATE_FILE, "w", encoding="utf-8") as fh:
        json.dump(sorted(processed), fh, indent=2)


# --------------------------------------------------------------------------
# Phase A: Outlook
# --------------------------------------------------------------------------
def get_target_folder(namespace, path):
    parts = [p for p in re.split(r"[\\/]", path) if p]
    if not parts:
        raise ValueError("OUTLOOK_FOLDER_PATH is empty")
    inbox = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    if parts[0].lower() == "inbox":
        folder, parts = inbox, parts[1:]
    else:
        folder = inbox.Parent
    for name in parts:
        folder = folder.Folders[name]
    return folder


def build_restricted_items(folder):
    tz = pytz.timezone(LOCAL_TIMEZONE)
    cutoff = datetime.now(tz) - timedelta(days=LOOKBACK_DAYS)
    cutoff_str = cutoff.strftime("%m/%d/%Y %I:%M %p")
    log.info("Lookback cutoff: %s (%s)", cutoff_str, LOCAL_TIMEZONE)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    return items.Restrict(f"[ReceivedTime] >= '{cutoff_str}'")


def already_replied(item):
    if (item.Subject or "").strip().upper().startswith("RE:"):
        return True
    try:
        if item.PropertyAccessor.GetProperty(PR_LAST_VERB_EXECUTED) in REPLIED_VERBS:
            return True
    except Exception:
        pass  # property absent on never-replied items
    return False


# --------------------------------------------------------------------------
# Phase B: extraction + API
# --------------------------------------------------------------------------
def extract_incident_id(body):
    """
    1) 'Incident ID ... XVA-123' label (Summary block of the alert email)
    2) first XVA-xxx after the word 'Summary'
    3) first XVA-xxx anywhere in the body
    """
    if not body:
        return None
    m = LABELLED_ID_PATTERN.search(body)
    if m:
        return m.group(1)
    s = re.search(r"summary", body, re.IGNORECASE)
    if s:
        m = INCIDENT_ID_PATTERN.search(body, s.end())
        if m:
            return m.group(0)
    m = INCIDENT_ID_PATTERN.search(body)
    return m.group(0) if m else None


def api_get(path, params):
    """Paced GET with Bearer auth; honours 429 Retry-After. Returns JSON or None."""
    global _last_call
    url = API_BASE.rstrip("/") + path
    headers = {"Authorization": f"Bearer {CLOUDSEK_TOKEN}", "Accept": "application/json"}

    for attempt in range(1, MAX_RETRIES + 1):
        gap = MIN_SECONDS_BETWEEN_CALLS - (time.monotonic() - _last_call)
        if gap > 0:
            time.sleep(gap)
        try:
            resp = requests.get(url, headers=headers, params=params, timeout=REQUEST_TIMEOUT)
            _last_call = time.monotonic()
        except requests.RequestException as exc:
            _last_call = time.monotonic()
            log.error("API request error (attempt %d/%d): %s", attempt, MAX_RETRIES, exc)
            time.sleep(min(2 ** attempt, 30))
            continue

        if resp.status_code == 429:
            ra = resp.headers.get("Retry-After", "")
            wait = int(float(ra)) if re.fullmatch(r"\d+(\.\d+)?", ra or "") else DEFAULT_RETRY_AFTER
            log.warning("429 RATE_LIMIT_EXCEEDED on %s, sleeping %ds (attempt %d/%d)",
                        path, wait, attempt, MAX_RETRIES)
            time.sleep(wait)
            continue

        if resp.status_code in (401, 403):
            log.error("Auth failure %d on %s. Check CLOUDSEK_TOKEN (Integrations -> Alerts API).",
                      resp.status_code, path)
            return None
        if resp.status_code >= 500:
            log.error("Server error %d on %s (attempt %d/%d)", resp.status_code, path, attempt, MAX_RETRIES)
            time.sleep(min(2 ** attempt, 30))
            continue
        if not resp.ok:
            log.error("API %d on %s: %s", resp.status_code, path, resp.text[:300])
            return None
        try:
            return resp.json()
        except ValueError:
            log.error("Non-JSON response from %s", path)
            return None

    log.error("Gave up on %s after %d attempts.", path, MAX_RETRIES)
    return None


def _rows(payload):
    if isinstance(payload, dict):
        data = payload.get("data")
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            return [data]
    if isinstance(payload, list):
        return payload
    return []


def _pick(d, keys):
    if not isinstance(d, dict):
        return None
    for k in keys:
        if d.get(k) not in (None, ""):
            return d[k]
    return None


def norm_status(s):
    """'ClosedResolved', 'closed-resolved', 'Closed Resolved' -> 'closedresolved'."""
    return re.sub(r"[^a-z]", "", str(s or "").lower())


def get_incident_status(incident_id):
    """Returns the incident's current status string, or None."""
    payload = api_get(SNAPSHOT_PATH, {"incident_display_id": incident_id})
    if payload is None:
        return None
    rows = _rows(payload)
    if DEBUG_DUMP:
        log.info("DEBUG snapshot raw for %s:\n%s", incident_id, json.dumps(payload, indent=2)[:3000])
    for row in rows:
        if row.get("incidentDisplayId") == incident_id or len(rows) == 1:
            st = _pick(row, ["status", "incidentStatus", "currentStatus"])
            if isinstance(st, dict):
                st = _pick(st, ["name", "value", "status"])
            return str(st) if st else None
    return None


def _sort_key(row):
    for k in ("sortableId", "actionedAt"):
        v = row.get(k)
        try:
            return int(v)
        except (TypeError, ValueError):
            continue
    return 0


def get_latest_comment(incident_id):
    """
    Latest comment on the incident, from /v2/changelog (entity_identifier=XVA-...).
    The docs list comment actions (added/edited/pinned/deleted) but the page does not
    publish the field that carries the comment text, so we look in currentState and
    top-level row keys. Use --debug-dump once to confirm the field name for your tenant.
    """
    payload = api_get(CHANGELOG_PATH, {"entity_identifier": incident_id, "limit": 200})
    if payload is None:
        return None
    rows = _rows(payload)
    if DEBUG_DUMP:
        log.info("DEBUG changelog raw for %s:\n%s", incident_id, json.dumps(payload, indent=2)[:6000])

    comment_rows = []
    for r in rows:
        action = str(r.get("actionType", "")).lower()
        if "comment" in action and not any(x in action for x in ("delete", "pin")):
            comment_rows.append(r)
    comment_rows.sort(key=_sort_key)

    text_keys = ["comment", "commentText", "comment_text", "text", "message", "content", "body"]
    for r in reversed(comment_rows):
        state = r.get("currentState")
        text = _pick(state, text_keys) or _pick(r, text_keys)
        if isinstance(text, dict):
            text = _pick(text, text_keys)
        if text:
            return str(text)
    return None


def is_closed(status):
    return norm_status(status).startswith("closed")


# --------------------------------------------------------------------------
# Phase C: draft
# --------------------------------------------------------------------------
def build_reply_html(comment):
    safe = html.escape(comment).replace("\r\n", "\n").replace("\n", "<br>")
    return f"""
<div style="font-family:Segoe UI, Calibri, Arial, sans-serif; font-size:11pt;">
  <p style="margin:0 0 12px 0;">
    Kindly find the comment added on the CloudSEK portal regarding this alert,
    and kindly close the incident.
  </p>
  <div style="border-left:4px solid {ACCENT}; border-top:1px solid {ACCENT};
              border-right:1px solid {ACCENT}; border-bottom:1px solid {ACCENT};
              padding:8px 14px; margin:12px 0;">
    <div style="color:{ACCENT_LIGHT}; font-weight:bold; font-size:9pt;
                letter-spacing:1px; text-transform:uppercase; margin-bottom:6px;">
      CloudSEK Comment
    </div>
    <div>{safe}</div>
  </div>
  <p style="margin:16px 0 0 0;">
    Regards,<br>
    CloudSek Automation - CyberDefence
  </p>
  <hr style="border:0; border-top:1px solid {ACCENT}; margin:16px 0;">
</div>
"""


def create_draft_reply(item, comment):
    reply = item.ReplyAll()
    block = build_reply_html(comment)
    original = reply.HTMLBody or ""
    tag = re.search(r"<body[^>]*>", original, re.IGNORECASE)
    reply.HTMLBody = (original[: tag.end()] + block + original[tag.end():]) if tag else block + original
    reply.Save()   # draft only; .Send() is never called


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    global DEBUG_DUMP
    ap = argparse.ArgumentParser()
    ap.add_argument("--debug-dump", action="store_true",
                    help="Log raw snapshot/changelog JSON for each checked incident (verify field names).")
    ap.add_argument("--test-id", help="Skip Outlook; just query this XVA ID and print what the script sees.")
    args = ap.parse_args()
    DEBUG_DUMP = args.debug_dump or bool(args.test_id)

    if not CLOUDSEK_TOKEN:
        log.critical("CLOUDSEK_TOKEN environment variable is not set.")
        return 1

    if args.test_id:
        status = get_incident_status(args.test_id)
        log.info("Status: %s | closed=%s", status, is_closed(status))
        if status and is_closed(status):
            log.info("Latest comment: %r", get_latest_comment(args.test_id))
        return 0

    try:
        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        folder = get_target_folder(namespace, OUTLOOK_FOLDER_PATH)
        items = build_restricted_items(folder)
        messages = [m for m in items]
    except Exception as exc:
        log.critical("Could not connect to Outlook / read folder: %s", exc)
        return 1
    log.info("Found %d item(s) in the lookback window.", len(messages))

    processed = load_state()
    stats = {"drafted": 0, "skipped": 0, "failed": 0}
    cache = {}   # incident_id -> (status, comment)

    for item in messages:
        try:
            if getattr(item, "Class", None) != OL_MAIL_CLASS:
                continue
            subject = item.Subject or "(no subject)"

            if already_replied(item):
                log.info("SKIP (already replied): %s", subject)
                stats["skipped"] += 1
                continue

            incident_id = extract_incident_id(item.Body)
            if not incident_id:
                log.info("SKIP (no XVA ID found): %s", subject)
                stats["skipped"] += 1
                continue
            if incident_id in processed:
                log.info("SKIP (draft already created): %s", incident_id)
                stats["skipped"] += 1
                continue

            if incident_id not in cache:
                status = get_incident_status(incident_id)
                comment = get_latest_comment(incident_id) if status and is_closed(status) else None
                cache[incident_id] = (status, comment)
            status, comment = cache[incident_id]

            if status is None:
                log.error("FAILED (no status from API): %s", incident_id)
                stats["failed"] += 1
                continue
            if not is_closed(status):
                log.info("SKIP (status '%s'): %s", status, incident_id)
                stats["skipped"] += 1
                continue
            if not comment:
                log.warning("SKIP (closed, but no comment found in changelog): %s", incident_id)
                stats["skipped"] += 1
                continue

            create_draft_reply(item, comment)
            processed.add(incident_id)
            save_state(processed)
            log.info("DRAFTED: %s (status %s) | %s", incident_id, status, subject)
            stats["drafted"] += 1

        except Exception as exc:
            log.error("FAILED on item: %s", exc)
            stats["failed"] += 1

    log.info("Done. drafted=%(drafted)d skipped=%(skipped)d failed=%(failed)d", stats)
    return 0 if stats["failed"] == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
