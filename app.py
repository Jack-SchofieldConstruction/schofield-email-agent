"""
Schofield Construction email agent.

Single-file Flask service:
  POST /scan     -> reads new mail from Hostinger IMAP, classifies via Claude,
                    writes enquiries + contacts to Supabase, returns JSON
                    in the shape the portal expects.
  GET  /health   -> sanity check
  GET  /         -> simple "hello" so the Render URL doesn't 404

Required environment variables (set in Render dashboard, NOT in code):
  IMAP_HOST       e.g. imap.hostinger.com
  IMAP_PORT       typically 993
  IMAP_USER       e.g. info@schofieldconstruction.site
  IMAP_PASSWORD   your email or app password
  ANTHROPIC_API_KEY
  SUPABASE_URL    e.g. https://xxxxx.supabase.co
  SUPABASE_KEY    the service_role key
  PORTAL_ORIGIN   e.g. https://schofieldgroundworks.co.uk (for CORS)
"""

import os
import json
import email
import imaplib
from email.header import decode_header
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime

import anthropic
import requests
from flask import Flask, request, jsonify
from flask_cors import CORS

# -----------------------------------------------------------------------------
# Config
# -----------------------------------------------------------------------------
IMAP_HOST        = os.environ.get("IMAP_HOST", "imap.hostinger.com")
IMAP_PORT        = int(os.environ.get("IMAP_PORT", "993"))
IMAP_USER        = os.environ["IMAP_USER"]
IMAP_PASSWORD    = os.environ["IMAP_PASSWORD"]
ANTHROPIC_KEY    = os.environ["ANTHROPIC_API_KEY"]
SUPABASE_URL     = os.environ["SUPABASE_URL"].rstrip("/")
SUPABASE_KEY     = os.environ["SUPABASE_KEY"]
PORTAL_ORIGIN    = os.environ.get("PORTAL_ORIGIN", "*")

MAX_EMAILS_PER_SCAN = 25  # cap to keep API costs predictable per scan
CLASSIFY_MODEL = "claude-haiku-4-5-20251001"

app = Flask(__name__)
CORS(app, origins=[PORTAL_ORIGIN] if PORTAL_ORIGIN != "*" else "*")
anthropic_client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)


# -----------------------------------------------------------------------------
# Supabase helpers (we use the REST API directly, no SDK needed)
# -----------------------------------------------------------------------------
def _sb_headers():
    return {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=representation",
    }


def sb_select(table, params=None):
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/{table}",
        headers=_sb_headers(),
        params=params or {},
        timeout=15,
    )
    r.raise_for_status()
    return r.json()


def sb_insert(table, row):
    # Drop any keys with None values to avoid unknown-column errors if the schema
    # doesn't yet have an optional column (e.g. before DB migration runs)
    clean_row = {k: v for k, v in row.items() if v is not None}
    r = requests.post(
        f"{SUPABASE_URL}/rest/v1/{table}",
        headers=_sb_headers(),
        json=clean_row,
        timeout=15,
    )
    if r.status_code == 409:  # unique violation — already exists, skip silently
        return None
    if r.status_code == 400 and "role" in r.text.lower():
        # Column doesn't exist yet — retry without it
        clean_row.pop("role", None)
        r = requests.post(
            f"{SUPABASE_URL}/rest/v1/{table}",
            headers=_sb_headers(),
            json=clean_row,
            timeout=15,
        )
        if r.status_code == 409:
            return None
    r.raise_for_status()
    data = r.json()
    return data[0] if isinstance(data, list) and data else data


def sb_update(table, match_col, match_val, updates):
    r = requests.patch(
        f"{SUPABASE_URL}/rest/v1/{table}",
        headers=_sb_headers(),
        params={match_col: f"eq.{match_val}"},
        json=updates,
        timeout=15,
    )
    r.raise_for_status()
    return r.json()


def get_last_scan():
    rows = sb_select("agent_state", {"key": "eq.last_email_scan", "select": "value"})
    if rows and rows[0].get("value"):
        return rows[0]["value"]
    return None


def set_last_scan(iso):
    sb_update("agent_state", "key", "last_email_scan",
              {"value": iso, "updated_at": datetime.now(timezone.utc).isoformat()})


def contact_exists(email_addr):
    if not email_addr:
        return False
    rows = sb_select("contacts", {"email": f"eq.{email_addr}", "select": "id"})
    return len(rows) > 0


def enquiry_exists(message_id):
    rows = sb_select("enquiries", {"outlook_message_id": f"eq.{message_id}", "select": "id"})
    return len(rows) > 0


# -----------------------------------------------------------------------------
# IMAP — fetch emails
# -----------------------------------------------------------------------------
def decode_value(raw):
    if not raw:
        return ""
    parts = decode_header(raw)
    out = []
    for text, enc in parts:
        if isinstance(text, bytes):
            try:
                out.append(text.decode(enc or "utf-8", errors="replace"))
            except LookupError:
                out.append(text.decode("utf-8", errors="replace"))
        else:
            out.append(text)
    return "".join(out)


def extract_body(msg):
    """Get a plain-text body, falling back to stripping HTML."""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition") or "")
            if ctype == "text/plain" and "attachment" not in disp:
                try:
                    return part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8", errors="replace"
                    )
                except Exception:
                    continue
        # fallback to HTML
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                try:
                    html = part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8", errors="replace"
                    )
                    return strip_html(html)
                except Exception:
                    continue
        return ""
    else:
        try:
            payload = msg.get_payload(decode=True)
            text = payload.decode(msg.get_content_charset() or "utf-8", errors="replace")
            if msg.get_content_type() == "text/html":
                text = strip_html(text)
            return text
        except Exception:
            return ""


def strip_html(html):
    import re
    text = re.sub(r"<style[\s\S]*?</style>", " ", html, flags=re.I)
    text = re.sub(r"<script[\s\S]*?</script>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def fetch_unread_emails():
    """Return list of dicts for all UNREAD emails in INBOX (capped at MAX_EMAILS_PER_SCAN, newest first).

    We don't mark them as read — the user's inbox read/unread state stays under their control.
    Deduplication against re-processing is handled by Message-ID in Supabase.
    """
    mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    mail.login(IMAP_USER, IMAP_PASSWORD)
    mail.select("INBOX")

    # UNSEEN = unread. Use PEEK in the fetch so we don't accidentally mark as read.
    status, data = mail.search(None, "UNSEEN")
    if status != "OK":
        mail.logout()
        return []

    ids = data[0].split()
    # Most recent first; cap how many we look at
    ids = list(reversed(ids))[:MAX_EMAILS_PER_SCAN]

    emails = []

    for msg_id in ids:
        # BODY.PEEK[] reads the message WITHOUT setting the \Seen flag
        status, msg_data = mail.fetch(msg_id, "(BODY.PEEK[])")
        if status != "OK":
            continue
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)

        received_raw = msg.get("Date")
        try:
            received_dt = parsedate_to_datetime(received_raw)
        except Exception:
            received_dt = datetime.now(timezone.utc)
        if received_dt.tzinfo is None:
            received_dt = received_dt.replace(tzinfo=timezone.utc)

        message_id_hdr = msg.get("Message-ID") or msg.get("Message-Id") or f"no-id-{msg_id.decode()}"
        from_hdr = decode_value(msg.get("From", ""))
        subject  = decode_value(msg.get("Subject", ""))
        body     = extract_body(msg)

        # Parse "Name <addr@example.com>"
        from_name, from_addr = email.utils.parseaddr(from_hdr)

        emails.append({
            "message_id":  message_id_hdr.strip(),
            "from_addr":   from_addr.lower().strip(),
            "from_name":   decode_value(from_name).strip(),
            "subject":     subject.strip(),
            "body":        body[:4000],   # cap body length for token costs
            "received":    received_dt.isoformat(),
        })

    mail.logout()
    return emails


# -----------------------------------------------------------------------------
# Claude classification
# -----------------------------------------------------------------------------
CLASSIFY_PROMPT_TEMPLATE = """You classify inbound business emails for Jack at Schofield Construction.

About Jack's situation:
- He is PRIMARILY looking for freelance/contract opportunities for himself: Site Manager, Project Manager, and Quantity Surveyor roles, plus day-rate cover work.
- He also receives some project enquiries for Schofield Construction (groundworks, civils, paving, fencing, small builds).
- Recruiter cold-emails about specific freelance/contract roles ARE valid enquiries — they're a primary lead source.

Return ONLY valid JSON (no markdown fences, no prose) matching this shape exactly:

{
  "is_enquiry": boolean,
  "type": "project" | "freelance" | null,
  "role": "site_manager" | "project_manager" | "quantity_surveyor" | "other" | null,
  "priority": "hot" | "warm" | "cold" | null,
  "estimated_value_gbp": number | null,
  "summary": string,
  "from_name": string,
  "company": string
}

Classification rules:

is_enquiry = TRUE when:
- A recruiter is offering Jack a specific contract/freelance role (Site Manager, Project Manager, QS, or similar construction freelance work). Even if generic/cold, treat as TRUE.
- Someone is asking about project work for Schofield Construction (quotes, builds, civils, groundworks).
- Someone is offering a day-rate or short-term contract opportunity.
- Anyone asking about Jack's availability for site management, project management, or QS work.

is_enquiry = FALSE when:
- Newsletters, marketing blasts, industry news, job-board digests (e.g. "10 new jobs matching your search").
- Receipts, invoices, billing, HMRC, banks, insurance.
- Automated notifications (LinkedIn updates, software notifications, calendar invites).
- Personal mail, internal admin, replies to existing threads from known contacts already in Jack's pipeline.
- Recruiter spam that is NOT for construction freelance roles (e.g. tech jobs, sales jobs, anything outside SM/PM/QS/site work).

type:
- "freelance" — contract or day-rate role for Jack personally (SM, PM, QS, site cover, etc.). This is the MAJORITY case.
- "project" — work for Schofield Construction (groundworks, builds, civils, etc.).

role (only relevant when type = "freelance"):
- "site_manager" — site management or SM role
- "project_manager" — PM, project lead, or similar
- "quantity_surveyor" — QS, commercial manager, estimating
- "other" — freelance but doesn't fit the above (e.g. day rate labourer cover)
- null — when type is "project"

priority:
- "hot" — named decision-maker, specific role with start date, explicit day rate or value, urgent timing, OR matches Jack's specialism (SM/PM/QS) closely.
- "warm" — relevant but generic/early-stage.
- "cold" — vague, exploratory, weak fit.

estimated_value_gbp:
- For freelance: estimate total contract value if duration mentioned (e.g. "£500/day x 12 weeks = ~£30,000"). Use UK day rates: SM £350-500, PM £400-600, QS £400-550.
- For projects: best estimate from scope.
- null if genuinely unknowable.

summary: one sentence, under 20 words, plain English. State the role/type and key detail (location, duration, rate if known).

from_name and company: extract from signature or sender. Empty string if unknowable.

EMAIL:
From: __FROM_FIELD__
Subject: __SUBJECT__

Body:
__BODY__
"""


def classify_email(em):
    prompt = (CLASSIFY_PROMPT_TEMPLATE
              .replace("__FROM_FIELD__", f"{em['from_name']} <{em['from_addr']}>")
              .replace("__SUBJECT__", em["subject"])
              .replace("__BODY__", em["body"] or "(empty body)"))
    try:
        resp = anthropic_client.messages.create(
            model=CLASSIFY_MODEL,
            max_tokens=400,
            messages=[{"role": "user", "content": prompt}],
        )
        text = resp.content[0].text.strip()
        # Strip any accidental code fences
        if text.startswith("```"):
            text = text.strip("`")
            if text.lower().startswith("json"):
                text = text[4:]
            text = text.strip()
        return json.loads(text)
    except Exception as e:
        app.logger.exception("Classification failed for %s: %s", em["message_id"], e)
        return None


# -----------------------------------------------------------------------------
# Routes
# -----------------------------------------------------------------------------
@app.route("/", methods=["GET"])
def index():
    return "Schofield email agent — POST /scan to run."


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"ok": True, "time": datetime.now(timezone.utc).isoformat()})


@app.route("/scan", methods=["POST", "OPTIONS"])
def scan():
    if request.method == "OPTIONS":
        return ("", 204)

    scan_started_at = datetime.now(timezone.utc).isoformat()

    try:
        emails = fetch_unread_emails()
    except Exception as e:
        app.logger.exception("IMAP fetch failed")
        return jsonify({"error": f"Could not fetch mail: {e}"}), 500

    new_enquiries = []
    new_contacts  = []

    for em in emails:
        if enquiry_exists(em["message_id"]):
            continue

        cls = classify_email(em)
        if not cls or not cls.get("is_enquiry"):
            continue

        role = cls.get("role")

        # Insert enquiry
        enq_row = {
            "outlook_message_id": em["message_id"],
            "received_date": em["received"],
            "from_name":  cls.get("from_name") or em["from_name"],
            "from_email": em["from_addr"],
            "company":    cls.get("company") or "",
            "subject":    em["subject"],
            "summary":    cls.get("summary") or "",
            "type":       cls.get("type"),
            "role":       role,
            "stage":      "lead",
            "estimated_value": cls.get("estimated_value_gbp") or 0,
            "priority":   cls.get("priority") or "warm",
        }
        inserted = sb_insert("enquiries", enq_row)
        if not inserted:
            continue  # dedupe

        # Portal-shape row for the immediate response
        new_enquiries.append({
            "date":     em["received"][:16].replace("T", " "),
            "from":     enq_row["from_name"],
            "company":  enq_row["company"],
            "subject":  enq_row["subject"],
            "type":     enq_row["type"] or "freelance",
            "role":     role,
            "stage":    enq_row["stage"],
            "value":    enq_row["estimated_value"],
            "priority": enq_row["priority"],
        })

        # Insert contact if new
        if em["from_addr"] and not contact_exists(em["from_addr"]):
            contact_row = {
                "name":    cls.get("from_name") or em["from_name"] or em["from_addr"],
                "company": cls.get("company") or "",
                "email":   em["from_addr"],
                "phone":   "",
            }
            c_inserted = sb_insert("contacts", contact_row)
            if c_inserted:
                new_contacts.append({
                    "name":    contact_row["name"],
                    "company": contact_row["company"],
                    "email":   contact_row["email"],
                    "phone":   contact_row["phone"],
                })

    # Update the cursor for the UI's "last scanned" display
    set_last_scan(scan_started_at)

    return jsonify({
        "enquiries":       new_enquiries,
        "contacts":        new_contacts,
        "scanCompletedAt": scan_started_at,
        "scannedCount":    len(emails),
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
