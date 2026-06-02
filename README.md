# Schofield email agent

Tiny Python service that reads `info@schofieldconstruction.site` over IMAP, classifies inbound emails as project / freelance enquiries via Claude, and writes the results to Supabase. The CRM portal calls it via `POST /scan`.

## How it fits together

```
Portal (Hostinger static HTML)
   │  POST /scan { since: ISO_string }
   ▼
This service (Render free web service)
   │
   ├──► IMAP (imap.hostinger.com) — reads new emails
   ├──► Anthropic API — classifies each email
   └──► Supabase — stores enquiries + contacts, returns JSON
```

## Local development

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env                # then fill in your real values
export $(cat .env | xargs)          # load env vars (mac/linux)
python app.py
```

Service runs on `http://localhost:8000`. Test with:

```bash
curl -X POST http://localhost:8000/scan -H "Content-Type: application/json" -d '{}'
```

## Endpoints

- `GET /` — sanity check
- `GET /health` — returns `{ok: true, time: ...}`
- `POST /scan` — body `{ "since": "2026-05-14T08:00:00Z" }` (or empty); returns enquiries+contacts JSON

## Environment variables

See `.env.example`. All are required except `PORTAL_ORIGIN` (defaults to `*` during testing).

## Deploy to Render

1. Push this repo to GitHub.
2. On Render: New → Web Service → connect the repo.
3. It'll detect `render.yaml`. Confirm: Python, free plan, gunicorn start.
4. Set all the environment variables in the Render dashboard.
5. Deploy. Copy the resulting `https://schofield-email-agent.onrender.com` URL into the portal's Settings.

## Costs

- Render free tier: £0 (sleeps after 15 min inactive; ~30s cold start on first call)
- Supabase free tier: £0 (500 MB, pauses after 7 days inactive)
- Anthropic API: ~£0.001 per email classified (Claude Haiku)
