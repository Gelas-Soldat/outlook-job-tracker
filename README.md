# Outlook Job Search Tracker

[![License: MIT](https://img.shields.io/badge/License-MIT-3DA639?style=for-the-badge)](LICENSE)
[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-ryancreates-FFD000?style=for-the-badge&logo=buymeacoffee&logoColor=black)](https://buymeacoffee.com/ryancreates)
[![Ko-fi](https://img.shields.io/badge/Ko--fi-Gelassoldat-FF5E5B?style=for-the-badge&logo=kofi&logoColor=white)](https://ko-fi.com/gelassoldat)

A live job application tracker that connects to your personal Outlook inbox, automatically classifies job emails, and displays everything in a real-time dashboard. Built with Python, Flask, and vanilla HTML/JS. Includes Claude Desktop integration via MCP.

**[Live Demo](https://outlook-job-tracker.netlify.app/?demo=true)**

---

## The Problem

When you're applying to 100+ jobs across LinkedIn, Greenhouse, Ashby, Workday, Lever, and iCIMS simultaneously, your inbox becomes the only source of truth — and it's unstructured, noisy, and impossible to analyze manually.

No clear response rate. No way to distinguish a rejection from silence. No pattern visibility.

This project builds an end-to-end pipeline that turns your inbox into a structured, searchable dashboard.

---

## What It Does

**Ingestion** — A Python/Flask server connects to Microsoft Graph API and queries your inbox across 14 job-related keyword searches, paginating through results to maximize coverage.

**Classification** — A two-layer classifier processes each email:
- Layer 1: Rule-based classifier identifies sender domains (Greenhouse, Ashby, Workday, Lever, iCIMS, LinkedIn, Rippling, and more), extracts company name and role from subject patterns, and detects status signals — applied, interview, rejected, pending, or other.
- Layer 2: AI fallback using Claude Haiku classifies any emails where the rules couldn't extract a company name. Results are cached in localStorage so each email is only classified once.

**Data quality** — A manual override system lets you correct any misclassified records. Overrides persist in localStorage. A rejection sync scans for rejection language across 16 phrase variants and auto-updates matching records.

**Delivery** — A live HTML dashboard hosted on Netlify with real-time filtering, search, status breakdown chart, recent activity feed, and click-to-edit rows. Manual refresh. Demo mode available via `?demo=true` URL parameter.

**MCP integration** — The Flask server is wrapped as a local MCP server, connecting it to Claude Desktop so Claude can query your inbox and surface insights conversationally.

---

## Why This Exists

The built-in Microsoft 365 connector in Anthropic's MCP library only supports work and school accounts. Personal Outlook/Live accounts are not supported. This project started as a workaround for that gap.

---

## Architecture

```
Outlook Inbox
     |
     v
Microsoft Graph API
     |
     v
Flask Server (server.py)          <-- MCP server (mcp_server.py)
     |                                        |
     v                                        v
index.html (Netlify)           Claude Desktop (reads inbox directly)
     |
     v
Claude Haiku API (AI fallback classifier)
```

### Stack

- Python 3.12, Flask, Flask-CORS
- MSAL (Microsoft Authentication Library)
- Microsoft Graph API — Mail.Read, Calendars.Read, User.Read
- Render (Flask server hosting, free tier)
- Netlify (dashboard hosting, free tier)
- cron-job.org (keep-alive ping every 10 minutes)
- Claude Haiku API (AI email classification fallback)
- Chart.js, vanilla HTML/CSS/JS
- MCP (Model Context Protocol) for Claude Desktop

### Files

| File | Purpose |
|---|---|
| `server.py` | Flask API, Microsoft Graph queries, OAuth token management, Render env var persistence |
| `mcp_server.py` | MCP wrapper exposing server endpoints to Claude Desktop |
| `index.html` | Live dashboard — classifier, AI fallback, demo mode, all UI |
| `.env` | Local credentials (never committed) |
| `token_cache.json` | OAuth token cache (never committed) |
| `Procfile` | Render start command |
| `requirements.txt` | Python dependencies |

---

## Features

- Classifies emails from 30+ ATS platforms and job senders
- Extracts company name and role from subject line patterns
- Detects applied / interview / rejected / pending / other status
- AI fallback via Claude Haiku for unknown senders
- AI classification results cached in localStorage (only new emails cost API calls)
- Manual override system — click any row to edit
- Auto-fill from email body for company and role extraction
- Rejection sync — scans inbox for rejection language and updates records
- Search and filter by status
- Status breakdown donut chart
- Recent activity feed
- Mobile-responsive card layout
- Demo mode with synthetic finance/crypto data (`?demo=true`)
- Token cache persists to Render environment variables — survives restarts without re-login

---

## Setup

### Prerequisites

- Python 3.12+
- A personal Microsoft account (Outlook/Live)
- Render account (free tier)
- Netlify account (free tier)
- Anthropic API account (optional, for AI fallback classifier)
- Claude Desktop (optional, for MCP integration)

### 1. Register an Azure App

1. Go to https://portal.azure.com and sign in
2. Search for App registrations → New registration
3. Name it anything, set supported account types to Personal Microsoft accounts only
4. Click Register and copy the Application (client) ID
5. Go to Certificates & secrets → New client secret → copy the value immediately
6. Go to API permissions → Add → Microsoft Graph → Delegated: Mail.Read, Calendars.Read, User.Read
7. Go to Authentication → Add Web platform → add redirect URIs:
   - `http://localhost:8081/callback`
   - `https://your-render-service.onrender.com/callback`

### 2. Deploy to Render

1. Fork this repo and connect it to Render as a new Web Service
2. Build command: `pip install -r requirements.txt && pip install gunicorn`
3. Start command: `python -m gunicorn server:app --bind 0.0.0.0:$PORT`
4. Add environment variables:

| Key | Value |
|---|---|
| `CLIENT_ID` | Your Azure app client ID |
| `CLIENT_SECRET` | Your Azure client secret |
| `TENANT_ID` | `consumers` |
| `REDIRECT_URI` | `https://your-render-service.onrender.com/callback` |
| `API_KEY` | Generate with `python -c "import secrets; print(secrets.token_hex(32))"` |
| `RENDER_API_KEY` | From Render dashboard → Account → API Keys |
| `RENDER_SERVICE_ID` | Your service ID from the Render dashboard URL |

5. After deploy, visit `https://your-render-service.onrender.com/login` to authenticate

### 3. Set up keep-alive

Create a free account at cron-job.org and set up a job to ping your Render URL every 10 minutes. This prevents the free tier from spinning down.

### 4. Deploy dashboard to Netlify

1. Connect your GitHub repo to Netlify as a new static site
2. Set publish directory to `.`
3. Deploy — your dashboard will be live at `https://your-site.netlify.app`
4. On first visit, enter your API key when prompted. It saves to localStorage permanently.

### 5. Run locally (optional)

```bash
pip install flask flask-cors msal requests python-dotenv mcp gunicorn
```

Create `.env`:
```
CLIENT_ID=your-client-id
TENANT_ID=consumers
CLIENT_SECRET=your-client-secret
REDIRECT_URI=http://localhost:8081/callback
API_KEY=your-api-key
```

```bash
python server.py
```

Visit `http://localhost:8081/login` to authenticate, then open `index.html` in your browser.

### 6. Claude Desktop MCP integration (optional)

Find your Claude Desktop config file:
- Windows (Microsoft Store): `%LOCALAPPDATA%\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\claude_desktop_config.json`
- Windows (direct install): `%APPDATA%\Claude\claude_desktop_config.json`

Add to `mcpServers`:

```json
{
  "mcpServers": {
    "outlook-mail": {
      "command": "C:\\path\\to\\python.exe",
      "args": ["C:\\path\\to\\outlook-job-tracker\\mcp_server.py"],
      "env": {}
    }
  }
}
```

Restart Claude Desktop. Claude can now search your inbox, read emails, and pull job application data conversationally.

### 7. AI fallback classifier (optional)

1. Create an account at https://console.anthropic.com and add a small amount of credits ($5 covers months of use)
2. Create an API key
3. On your dashboard, click reset key and enter both your dashboard API key and your Anthropic API key
4. The AI classifier will automatically handle any emails the rule-based system can't parse, caching results in localStorage

---

## Demo

Visit the live demo at:

**https://outlook-job-tracker.netlify.app/?demo=true**

Shows 25 synthetic finance and crypto analyst applications with realistic statuses, companies, and roles. No API key required.

---

## Support

If this project helped you, consider buying me a coffee.

- Ko-fi: https://ko-fi.com/gelassoldat
- Buy Me a Coffee: https://buymeacoffee.com/ryancreates

## Author

Built by [Gelas-Soldat](https://github.com/Gelas-Soldat)
