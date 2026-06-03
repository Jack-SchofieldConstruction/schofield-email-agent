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
from datetime import datetime, timezone, timedelta
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

MAX_EMAILS_PER_SCAN = 200  # cap to keep one scan bounded; dedup means re-scans are cheap
SCAN_DAYS_BACK = int(os.environ.get("SCAN_DAYS_BACK", "90"))
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


def classification_exists(message_id):
    rows = sb_select("classifications", {"outlook_message_id": f"eq.{message_id}", "select": "id"})
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


def fetch_recent_emails():
    """Fetch all emails from INBOX received in the last SCAN_DAYS_BACK days.

    Dedup happens later via Message-ID lookup against Supabase, so re-running this
    repeatedly is cheap — only never-seen emails get sent to Claude for classification.
    BODY.PEEK ensures we don't change the read/unread state.
    """
    mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    mail.login(IMAP_USER, IMAP_PASSWORD)
    mail.select("INBOX")

    # IMAP SINCE wants a date like "01-Jan-2026" — gives us all emails ON or AFTER that day.
    since_dt = datetime.now(timezone.utc) - timedelta(days=SCAN_DAYS_BACK)
    since_str = since_dt.strftime("%d-%b-%Y")
    status, data = mail.search(None, f'(SINCE "{since_str}")')
    if status != "OK":
        mail.logout()
        return []

    ids = data[0].split()
    # Newest first; cap how many we look at in a single scan
    ids = list(reversed(ids))[:MAX_EMAILS_PER_SCAN]

    emails = []
    for msg_id in ids:
        # PEEK so we don't mark as read
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

        emails.append({
            "message_id":  message_id_hdr.strip(),
            "from_addr":   "",      # filled in later if we proceed past dedup
            "from_name":   "",
            "from_hdr":    from_hdr,
            "subject":     subject.strip(),
            "msg":         msg,     # keep the parsed message so we only extract the body if needed
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
  "company": string,
  "reason": string
}

Classification rules:

is_enquiry = TRUE when:
- A recruiter is offering Jack a specific contract/freelance role (Site Manager, Project Manager, QS, or similar construction freelance work). Even if generic/cold, treat as TRUE.
- A job board alert (Indeed, Hampshire Jobs, LinkedIn, CV-Library, etc.) features a SPECIFIC role for SM, PM, QS, or similar construction freelance/contract work with identifiable details (employer, location, rate, or role title). These ARE legitimate leads even though they're automated — Jack actively looks at job boards for these.
- Remote, hybrid, or flexible-location SM/PM/QS roles — treat as INTERESTING regardless of geography.
- Someone is asking about project work for Schofield Construction (quotes, builds, civils, groundworks).
- Someone is offering a day-rate or short-term contract opportunity.
- Anyone asking about Jack's availability for site management, project management, or QS work.

is_enquiry = FALSE when:
- A job board email that is just a GENERIC weekly digest with no specific roles, OR contains only roles outside construction (tech, retail, healthcare, etc.).
- Marketing/newsletter content from tool vendors, software companies, industry publications.
- Receipts, invoices, billing, HMRC, banks, insurance, utilities.
- Automated system notifications (LinkedIn connection updates, calendar invites, password resets).
- Personal mail, internal admin, replies to existing threads from known contacts already in Jack's pipeline.
- Recruiter spam that is NOT for construction freelance roles (e.g. tech jobs, sales jobs, anything outside SM/PM/QS/site/construction).

Edge cases — favour TRUE when:
- The email lists 1-3 specific construction roles even if framed as "alert" or "digest" — extract the most relevant one in `summary`.
- The role mentions "freelance", "contract", "day rate", or specific £/day figures — these are strong positive signals.

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

reason: ALWAYS provide a one-sentence (under 25 words) explanation of WHY you classified this as an enquiry or not. Be specific — name what triggered the decision. Examples:
- "Recruiter offering 12-week SM contract in Reading at £450/day."
- "LinkedIn weekly digest of unrelated job suggestions, not a specific role offer."
- "Marketing newsletter from a tool vendor, no role or project on offer."
- "Receipt from Hostinger billing — automated, no opportunity."
- "Reply from existing contractor about ongoing project — internal thread."

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
    raw_text = ""
    try:
        resp = anthropic_client.messages.create(
            model=CLASSIFY_MODEL,
            max_tokens=500,
            messages=[{"role": "user", "content": prompt}],
        )
        raw_text = resp.content[0].text.strip()
        text = raw_text
        if text.startswith("```"):
            text = text.strip("`")
            if text.lower().startswith("json"):
                text = text[4:]
            text = text.strip()
        return json.loads(text), raw_text
    except json.JSONDecodeError as e:
        app.logger.warning("Could not parse Claude JSON for %s: %s — raw: %s",
                           em["message_id"], e, raw_text[:300])
        return None, f"[JSON_DECODE_ERROR] {e}\n---RAW---\n{raw_text}"
    except Exception as e:
        # Capture the actual exception so we can see what's going wrong (auth, rate limit, etc.)
        err_type = type(e).__name__
        err_msg = str(e)
        app.logger.exception("Classification failed for %s: %s: %s",
                             em["message_id"], err_type, err_msg)
        return None, f"[{err_type}] {err_msg}\n---RAW---\n{raw_text}"


# -----------------------------------------------------------------------------
# Routes
# -----------------------------------------------------------------------------
@app.route("/", methods=["GET"])
def index():
    return "Schofield email agent — POST /scan to run."


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"ok": True, "time": datetime.now(timezone.utc).isoformat()})


@app.route("/recent-rejections", methods=["GET"])
def recent_rejections():
    """Return the most recent emails classified as NOT enquiries, with Claude's reason.
       Optional query params:
         ?limit=50    (default 50)
         ?search=foo  filter subject/from
    """
    limit = int(request.args.get("limit", "50"))
    search = request.args.get("search", "").strip()

    params = {
        "is_enquiry": "eq.false",
        "select": "received_date,from_name,from_email,subject,reason,raw_response,created_at",
        "order": "created_at.desc",
        "limit": str(limit),
    }
    if search:
        # Postgrest 'or' filter — match search in subject OR from_name OR from_email
        params["or"] = f"(subject.ilike.*{search}*,from_name.ilike.*{search}*,from_email.ilike.*{search}*)"

    rows = sb_select("classifications", params)
    return jsonify({"count": len(rows), "rejections": rows})


@app.route("/recent-classifications", methods=["GET"])
def recent_classifications():
    """All recent classifications, enquiries AND rejections, for browsing.
       Optional ?limit=50
    """
    limit = int(request.args.get("limit", "50"))
    rows = sb_select("classifications", {
        "select": "received_date,from_name,from_email,subject,is_enquiry,type,role,priority,summary,reason,created_at",
        "order": "created_at.desc",
        "limit": str(limit),
    })
    return jsonify({"count": len(rows), "results": rows})


@app.route("/scan", methods=["POST", "OPTIONS"])
def scan():
    if request.method == "OPTIONS":
        return ("", 204)

    scan_started_at = datetime.now(timezone.utc).isoformat()

    try:
        emails = fetch_recent_emails()
    except Exception as e:
        app.logger.exception("IMAP fetch failed")
        return jsonify({"error": f"Could not fetch mail: {e}"}), 500

    new_enquiries = []
    new_contacts  = []
    skipped_dupes = 0
    classified    = 0
    rejections    = 0
    parse_errors  = 0

    for em in emails:
        # Cheap path: skip anything already in the DB BEFORE parsing body or calling Claude
        if enquiry_exists(em["message_id"]):
            skipped_dupes += 1
            continue
        if classification_exists(em["message_id"]):
            # Already classified once (as a non-enquiry) — don't re-spend tokens on it
            skipped_dupes += 1
            continue

        # Only now do the expensive body extraction and parsing
        msg = em.pop("msg")
        body = extract_body(msg)[:4000]
        from_name, from_addr = email.utils.parseaddr(em["from_hdr"])
        em["from_addr"] = from_addr.lower().strip()
        em["from_name"] = decode_value(from_name).strip()
        em["body"]      = body

        cls, raw_text = classify_email(em)
        classified += 1

        # Always log the classification result, enquiry or not
        log_row = {
            "outlook_message_id": em["message_id"],
            "received_date": em["received"],
            "from_name":  em["from_name"],
            "from_email": em["from_addr"],
            "subject":    em["subject"],
            "raw_response": raw_text[:2000] if raw_text else None,
        }
        if cls:
            log_row.update({
                "is_enquiry": bool(cls.get("is_enquiry")),
                "type":       cls.get("type"),
                "role":       cls.get("role"),
                "priority":   cls.get("priority"),
                "estimated_value": cls.get("estimated_value_gbp") or 0,
                "summary":    cls.get("summary") or "",
                "reason":     cls.get("reason") or "",
            })
        else:
            log_row.update({"is_enquiry": False, "reason": "JSON parse error — see raw_response"})
            parse_errors += 1
        sb_insert("classifications", log_row)

        if not cls or not cls.get("is_enquiry"):
            rejections += 1
            continue

        role = cls.get("role")

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
            continue

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

    set_last_scan(scan_started_at)

    return jsonify({
        "enquiries":         new_enquiries,
        "contacts":          new_contacts,
        "scanCompletedAt":   scan_started_at,
        "scannedCount":      len(emails),
        "newlyClassified":   classified,
        "skippedDuplicates": skipped_dupes,
        "rejections":        rejections,
        "parseErrors":       parse_errors,
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
