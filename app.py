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

MAX_EMAILS_PER_SCAN = int(os.environ.get("MAX_EMAILS_PER_SCAN", "20"))  # per HTTP call, keeps memory low
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
    """Fetch up to MAX_EMAILS_PER_SCAN UNSEEN-BY-AGENT emails from INBOX in the last SCAN_DAYS_BACK days.

    Walks newest-to-oldest. For each candidate, checks Supabase to see if it's already been
    classified, and only fetches the FULL body for emails that need processing.
    This keeps memory bounded — we never load 200 full email bodies at once.
    """
    mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    mail.login(IMAP_USER, IMAP_PASSWORD)
    mail.select("INBOX")

    since_dt = datetime.now(timezone.utc) - timedelta(days=SCAN_DAYS_BACK)
    since_str = since_dt.strftime("%d-%b-%Y")
    status, data = mail.search(None, f'(SINCE "{since_str}")')
    if status != "OK":
        mail.logout()
        return [], 0

    all_ids = list(reversed(data[0].split()))  # newest first
    emails = []
    seen_count = 0

    for msg_id in all_ids:
        if len(emails) >= MAX_EMAILS_PER_SCAN:
            break

        # Step 1: fetch JUST the Message-ID header to check dedup cheaply (no body, low memory)
        status, header_data = mail.fetch(msg_id, "(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)])")
        if status != "OK":
            continue
        header_blob = header_data[0][1].decode("utf-8", errors="replace")
        msg_id_match = ""
        for line in header_blob.splitlines():
            if line.lower().startswith("message-id:"):
                msg_id_match = line.split(":", 1)[1].strip()
                break
        if not msg_id_match:
            msg_id_match = f"no-id-{msg_id.decode()}"

        # Step 2: skip if already in DB (enquiry or classification)
        if enquiry_exists(msg_id_match) or classification_exists(msg_id_match):
            seen_count += 1
            continue

        # Step 3: full fetch only for emails we'll actually process
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

        from_hdr = decode_value(msg.get("From", ""))
        subject  = decode_value(msg.get("Subject", ""))
        body     = extract_body(msg)[:4000]
        from_name, from_addr = email.utils.parseaddr(from_hdr)

        emails.append({
            "message_id":  msg_id_match,
            "from_addr":   from_addr.lower().strip(),
            "from_name":   decode_value(from_name).strip(),
            "subject":     subject.strip(),
            "body":        body,
            "received":    received_dt.isoformat(),
        })

        # Free the raw msg object before next loop
        del msg, raw, msg_data

    mail.logout()
    return emails, seen_count


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


def rams_exists(message_id):
    rows = sb_select("rams_submissions", {"outlook_message_id": f"eq.{message_id}", "select": "id"})
    return len(rows) > 0


# Common Sent folder names in IMAP — Hostinger uses "INBOX.Sent", others vary.
SENT_FOLDER_CANDIDATES = ['INBOX.Sent', 'Sent', 'Sent Items', 'Sent Messages', '[Gmail]/Sent Mail']


def fetch_rams_emails():
    """Walk the Sent folder for RAMS Generator notifications.

    Identifies them by subject line starting with 'Your RAMS' (with optional em-dash variants).
    Memory-efficient: streams Message-IDs, skips already-processed ones cheaply.
    """
    mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    mail.login(IMAP_USER, IMAP_PASSWORD)

    # Find the Sent folder — name varies by server
    selected_folder = None
    for candidate in SENT_FOLDER_CANDIDATES:
        status, _ = mail.select(candidate, readonly=True)
        if status == "OK":
            selected_folder = candidate
            break
    if not selected_folder:
        mail.logout()
        raise RuntimeError("Could not find Sent folder — check IMAP folder list")

    since_dt = datetime.now(timezone.utc) - timedelta(days=SCAN_DAYS_BACK)
    since_str = since_dt.strftime("%d-%b-%Y")
    # Filter at the IMAP level: only emails with "RAMS" in the subject within our window
    status, data = mail.search(None, f'(SINCE "{since_str}" SUBJECT "RAMS")')
    if status != "OK":
        mail.logout()
        return [], 0

    all_ids = list(reversed(data[0].split()))  # newest first
    rams_emails = []
    seen_count = 0

    for msg_id in all_ids:
        if len(rams_emails) >= MAX_EMAILS_PER_SCAN:
            break

        # Cheap header-only fetch for dedup
        status, header_data = mail.fetch(msg_id, "(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID SUBJECT)])")
        if status != "OK":
            continue
        header_blob = header_data[0][1].decode("utf-8", errors="replace")
        msg_id_match = ""
        subject_match = ""
        for line in header_blob.splitlines():
            low = line.lower()
            if low.startswith("message-id:"):
                msg_id_match = line.split(":", 1)[1].strip()
            elif low.startswith("subject:"):
                subject_match = line.split(":", 1)[1].strip()
        if not msg_id_match:
            msg_id_match = f"no-id-{msg_id.decode()}"

        # Sanity check the subject — must start with "Your RAMS"
        if not subject_match.lower().startswith("your rams"):
            continue

        if rams_exists(msg_id_match):
            seen_count += 1
            continue

        # Full fetch — we need To: header, body, and attachment names
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

        subject = decode_value(msg.get("Subject", "")).strip()
        to_hdr  = decode_value(msg.get("To", "")).strip()
        # 'To' can be "Name <email@example.com>" or just an address
        contractor_name, contractor_email = email.utils.parseaddr(to_hdr)
        contractor_name = decode_value(contractor_name).strip()
        contractor_email = contractor_email.lower().strip()

        # Attachment filename + body text
        attachment_filename = ""
        body_text = ""
        if msg.is_multipart():
            for part in msg.walk():
                disp = str(part.get("Content-Disposition") or "")
                if "attachment" in disp.lower():
                    fname = part.get_filename()
                    if fname:
                        attachment_filename = decode_value(fname).strip()
                elif part.get_content_type() == "text/plain" and "attachment" not in disp.lower():
                    try:
                        body_text += part.get_payload(decode=True).decode(
                            part.get_content_charset() or "utf-8", errors="replace"
                        )
                    except Exception:
                        pass
                elif part.get_content_type() == "text/html" and not body_text and "attachment" not in disp.lower():
                    try:
                        html = part.get_payload(decode=True).decode(
                            part.get_content_charset() or "utf-8", errors="replace"
                        )
                        body_text = strip_html(html)
                    except Exception:
                        pass

        rams_emails.append({
            "message_id":          msg_id_match,
            "received":            received_dt.isoformat(),
            "subject":             subject,
            "contractor_email":    contractor_email,
            "contractor_name":     contractor_name,
            "attachment_filename": attachment_filename,
            "body_text":           body_text[:2000],
        })

        del msg, raw, msg_data

    mail.logout()
    return rams_emails, seen_count


def parse_rams_email(em):
    """Pull project name and hazards count out of a RAMS notification email."""
    subject = em["subject"]

    # Project name: subject is "Your RAMS — <project>" with various dash characters
    project = ""
    for sep in ["—", "–", "-"]:  # em-dash, en-dash, hyphen
        if sep in subject:
            after = subject.split(sep, 1)[1].strip()
            if after:
                project = after
                break
    if not project:
        # Fallback: strip "Your RAMS" prefix
        project = subject.replace("Your RAMS", "", 1).strip().lstrip("—–-:").strip()

    # Hazards count: e.g. "12 identified hazards" in the body
    hazards = None
    import re
    m = re.search(r"(\d+)\s+identified\s+hazards", em["body_text"] or "", re.IGNORECASE)
    if m:
        try:
            hazards = int(m.group(1))
        except ValueError:
            pass

    return {
        "project_name": project,
        "hazards_count": hazards,
    }


@app.route("/imap-folders", methods=["GET"])
def imap_folders():
    """Diagnostic: list all IMAP folders on the mailbox so we can see exact names."""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(IMAP_USER, IMAP_PASSWORD)
        status, folders = mail.list()
        mail.logout()
        if status != "OK":
            return jsonify({"error": "Could not list folders"}), 500
        folder_names = [f.decode("utf-8", errors="replace") for f in folders]
        return jsonify({"count": len(folder_names), "folders": folder_names})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/sent-subjects", methods=["GET"])
def sent_subjects():
    """Diagnostic: list recent Sent folder subject lines + recipient + Message-ID
       so we can see why RAMS emails aren't matching."""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(IMAP_USER, IMAP_PASSWORD)

        selected_folder = None
        for candidate in SENT_FOLDER_CANDIDATES:
            status, _ = mail.select(candidate, readonly=True)
            if status == "OK":
                selected_folder = candidate
                break
        if not selected_folder:
            mail.logout()
            return jsonify({"error": "Could not find Sent folder"}), 500

        # Get ALL emails in Sent folder (not filtered by subject)
        since_dt = datetime.now(timezone.utc) - timedelta(days=SCAN_DAYS_BACK)
        since_str = since_dt.strftime("%d-%b-%Y")
        status, data = mail.search(None, f'(SINCE "{since_str}")')
        if status != "OK":
            mail.logout()
            return jsonify({"error": "search failed"}), 500

        ids = list(reversed(data[0].split()))[:50]  # show last 50
        results = []
        for msg_id in ids:
            status, header_data = mail.fetch(msg_id, "(BODY.PEEK[HEADER.FIELDS (SUBJECT TO MESSAGE-ID DATE)])")
            if status != "OK":
                continue
            blob = header_data[0][1].decode("utf-8", errors="replace")
            subject = to = mid = date = ""
            for line in blob.splitlines():
                low = line.lower()
                if low.startswith("subject:"):
                    subject = decode_value(line.split(":", 1)[1].strip())
                elif low.startswith("to:"):
                    to = line.split(":", 1)[1].strip()
                elif low.startswith("message-id:"):
                    mid = line.split(":", 1)[1].strip()
                elif low.startswith("date:"):
                    date = line.split(":", 1)[1].strip()
            results.append({
                "subject": subject,
                "to": to,
                "message_id": mid,
                "date": date,
                "starts_with_your_rams": subject.lower().lstrip().startswith("your rams"),
            })

        mail.logout()
        return jsonify({"folder": selected_folder, "count": len(results), "emails": results})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/scan-rams", methods=["POST", "OPTIONS"])
def scan_rams():
    """Scan the Sent folder for RAMS Generator notifications and write to rams_submissions."""
    if request.method == "OPTIONS":
        return ("", 204)

    scan_started_at = datetime.now(timezone.utc).isoformat()

    try:
        emails, skipped = fetch_rams_emails()
    except Exception as e:
        app.logger.exception("RAMS scan failed")
        return jsonify({"error": f"RAMS scan failed: {e}"}), 500

    new_rams = []
    for em in emails:
        parsed = parse_rams_email(em)
        row = {
            "outlook_message_id":  em["message_id"],
            "date_sent":           em["received"],
            "contractor_email":    em["contractor_email"],
            "contractor_name":     em["contractor_name"] or "",
            "project_name":        parsed["project_name"],
            "attachment_filename": em["attachment_filename"],
            "hazards_count":       parsed["hazards_count"],
        }
        inserted = sb_insert("rams_submissions", row)
        if inserted:
            new_rams.append({
                "date":                em["received"][:16].replace("T", " "),
                "contractor_email":    row["contractor_email"],
                "contractor_name":     row["contractor_name"],
                "project_name":        row["project_name"],
                "attachment_filename": row["attachment_filename"],
                "hazards_count":       row["hazards_count"],
            })

    more = len(emails) >= MAX_EMAILS_PER_SCAN
    return jsonify({
        "rams":              new_rams,
        "scanCompletedAt":   scan_started_at,
        "scannedCount":      len(emails),
        "skippedDuplicates": skipped,
        "moreAvailable":     more,
    })


@app.route("/rams", methods=["GET", "OPTIONS"])
def list_rams():
    """Return all RAMS submissions from the database, newest first."""
    if request.method == "OPTIONS":
        return ("", 204)

    rows = sb_select("rams_submissions", {
        "select": "date_sent,contractor_email,contractor_name,project_name,attachment_filename,hazards_count,id",
        "order": "date_sent.desc",
        "limit": "500",
    })
    portal_rows = [{
        "id":                  r.get("id"),
        "date":                (r.get("date_sent") or "")[:16].replace("T", " "),
        "contractor_email":    r.get("contractor_email") or "",
        "contractor_name":     r.get("contractor_name") or "",
        "project_name":        r.get("project_name") or "",
        "attachment_filename": r.get("attachment_filename") or "",
        "hazards_count":       r.get("hazards_count"),
    } for r in rows]
    return jsonify({"count": len(portal_rows), "rams": portal_rows})


@app.route("/enquiries", methods=["GET", "OPTIONS"])
def list_enquiries():
    """Return all enquiries from the database, newest first, in the portal's expected shape."""
    if request.method == "OPTIONS":
        return ("", 204)

    rows = sb_select("enquiries", {
        "select": "received_date,from_name,from_email,company,subject,summary,type,role,stage,estimated_value,priority,id",
        "order": "received_date.desc",
        "limit": "500",
    })
    portal_rows = [{
        "id":       r.get("id"),
        "date":     (r.get("received_date") or "")[:16].replace("T", " "),
        "from":     r.get("from_name") or "",
        "email":    r.get("from_email") or "",
        "company":  r.get("company") or "",
        "subject":  r.get("subject") or "",
        "summary":  r.get("summary") or "",
        "type":     r.get("type") or "freelance",
        "role":     r.get("role"),
        "stage":    r.get("stage") or "lead",
        "value":    r.get("estimated_value") or 0,
        "priority": r.get("priority") or "warm",
    } for r in rows]
    return jsonify({"count": len(portal_rows), "enquiries": portal_rows})


@app.route("/contacts", methods=["GET", "OPTIONS"])
def list_contacts():
    """Return all contacts from the database, newest first."""
    if request.method == "OPTIONS":
        return ("", 204)

    rows = sb_select("contacts", {
        "select": "name,company,email,phone,id",
        "order": "created_at.desc",
        "limit": "500",
    })
    return jsonify({"count": len(rows), "contacts": rows})


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
        emails, skipped_dupes = fetch_recent_emails()
    except Exception as e:
        app.logger.exception("IMAP fetch failed")
        return jsonify({"error": f"Could not fetch mail: {e}"}), 500

    new_enquiries = []
    new_contacts  = []
    classified    = 0
    rejections    = 0
    parse_errors  = 0

    for em in emails:
        cls, raw_text = classify_email(em)
        classified += 1

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
            log_row.update({"is_enquiry": False, "reason": "Classification failed — see raw_response"})
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

    # Signal to the portal whether there's likely more to process
    more_available = classified >= MAX_EMAILS_PER_SCAN

    return jsonify({
        "enquiries":         new_enquiries,
        "contacts":          new_contacts,
        "scanCompletedAt":   scan_started_at,
        "scannedCount":      len(emails),
        "newlyClassified":   classified,
        "skippedDuplicates": skipped_dupes,
        "rejections":        rejections,
        "parseErrors":       parse_errors,
        "moreAvailable":     more_available,
        "batchSize":         MAX_EMAILS_PER_SCAN,
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
