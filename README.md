# Outlook Job Search Tracker

[![License: MIT](https://img.shields.io/badge/License-MIT-3DA639?style=for-the-badge)](LICENSE)
[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-ryancreates-FFD000?style=for-the-badge&logo=buymeacoffee&logoColor=black)](https://buymeacoffee.com/ryancreates)
[![Ko-fi](https://img.shields.io/badge/Ko--fi-Gelassoldat-FF5E5B?style=for-the-badge&logo=kofi&logoColor=white)](https://ko-fi.com/gelassoldat)

A live job application dashboard built on top of Microsoft Graph API. It connects to a personal Outlook inbox, automatically pulls and classifies job application emails, and displays them in a real-time web dashboard with search, filtering, and manual override capabilities.

Built with Python, Flask, and vanilla HTML/JS. Integrated with Claude Desktop as a local MCP server.

---

## The Problem

Job hunting at scale is a data problem most people don't treat like one.

When you're actively applying across LinkedIn, Greenhouse, Ashby, Workday, Lever, and iCIMS simultaneously, your inbox becomes the only source of truth — and it's unstructured, noisy, and nearly impossible to analyze manually. No clear response rate. No way to distinguish a rejection from silence. No pattern visibility.

I had over 100 applications in flight with no real picture of where any of them stood. So I built one.

## What It Does

Rather than a spreadsheet, this is an end-to-end data pipeline:

**Ingestion** — A Python/Flask server connects to Microsoft Graph API and queries the inbox across 14 job-related keyword searches, paginating through results to maximize coverage.

**Classification** — Each email runs through a multi-layered classifier that identifies the sender domain (Greenhouse, Ashby, Workday, Lever, iCIMS, etc.), extracts company name and role from subject line patterns, and detects status signals — applied, interview, rejected, pending — using keyword matching.

**Data quality** — A manual override system lets users correct misclassified records. Overrides persist in browser localStorage across refreshes. A dedicated rejection sync scans for rejection language across 14 phrase variants and auto-updates matching records.

**Delivery** — A live HTML dashboard with real-time filtering, search, status breakdown charts, and a recent activity feed. Auto-refreshes every 5 minutes. Hosted on Netlify, backend on Render.

**MCP integration** — The Flask server is wrapped as a local MCP server, connecting it to Claude Desktop so the AI assistant can query the inbox and surface insights conversationally.

## Why This Exists

The built-in Microsoft 365 connector in Anthropic's MCP library only supports work and school accounts. Personal Outlook/Live accounts aren't supported. This project started as a workaround for that gap and grew from there.

---

## Technical Overview

### Architecture

```
Outlook Inbox
     |
     v
Microsoft Graph API
     |
     v
Flask Server (server.py)     <-- local MCP server (mcp_server.py)
     |                                    |
     v                                    v
index.html (Netlify)        Claude Desktop (reads inbox directly)
```

### Stack

- Python 3.12
- Flask + Flask-CORS
- MSAL (Microsoft Authentication Library)
- Microsoft Graph API (Mail.Read, Calendars.Read, User.Read)
- Vanilla HTML/CSS/JS
- Chart.js
- MCP (Model Context Protocol) for Claude Desktop integration

### Files

| File | Purpose |
|---|---|
| `server.py` | Flask API server, Microsoft Graph queries, token management |
| `mcp_server.py` | MCP wrapper exposing server endpoints to Claude Desktop |
| `index.html` | Live web dashboard |
| `.env` | Your credentials (never committed) |
| `token_cache.json` | OAuth token cache (never committed) |
| `requirements.txt` | Python dependencies |

---

## Setup

### Prerequisites

- Python 3.12+
- A personal Microsoft account (Outlook/Live)
- Claude Desktop (optional, for MCP integration)

### 1. Register an Azure App

1. Go to https://portal.azure.com and sign in with your Microsoft account
2. Search for "App registrations" and click New registration
3. Name it anything, set supported account types to "Personal Microsoft accounts only"
4. Click Register
5. Copy the Application (client) ID
6. Go to Certificates & secrets, create a new client secret, copy the value immediately
7. Go to API permissions, add Microsoft Graph delegated permissions: Mail.Read, Calendars.Read, User.Read
8. Go to Authentication, add a Web platform with redirect URI: http://localhost:8080/callback

### 2. Configure credentials

Create a `.env` file in the project folder:

```
CLIENT_ID=your-client-id-here
TENANT_ID=consumers
CLIENT_SECRET=your-client-secret-here
REDIRECT_URI=http://localhost:8080/callback
API_KEY=your-generated-api-key-here
```

### 3. Install dependencies

```bash
pip install flask flask-cors msal requests python-dotenv mcp gunicorn
```

### 4. Run the server

```bash
python server.py
```

Visit http://localhost:8080/login in your browser and sign in with your Microsoft account. You only need to do this once as the token is saved locally.

### 5. Open the dashboard

Open `index.html` in your browser. On first load it will prompt for your API key. Enter it once and it saves to localStorage permanently.

### 6. Claude Desktop integration (optional)

Add this to your Claude Desktop config file at `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook-mail": {
      "command": "python",
      "args": ["C:\\path\\to\\outlook-job-tracker\\mcp_server.py"],
      "env": {}
    }
  }
}
```

Restart Claude Desktop. You should see the hammer icon indicating MCP tools are connected.

---

## Support

If this project helped you, consider buying me a coffee.

- Ko-fi: https://ko-fi.com/gelassoldat
- Buy Me a Coffee: https://buymeacoffee.com/ryancreates

## Author

Built by [Gelas-Soldat](https://github.com/Gelas-Soldat)
