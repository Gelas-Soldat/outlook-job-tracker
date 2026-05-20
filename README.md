# Outlook Job Search Tracker

A live job application dashboard built on top of Microsoft Graph API. It connects to a personal Outlook inbox, automatically pulls and classifies job application emails, and displays them in a real-time web dashboard with search, filtering, and manual override capabilities.

Built with Python, Flask, and vanilla HTML/JS. Integrated with Claude Desktop as a local MCP server.

---

## For Hiring Managers

If you are reviewing this project as part of evaluating me for a business analyst or data analyst role, this section is written with you in mind.

### The Problem

Job hunting at scale creates a data problem most people never think about. When you are actively applying to dozens of roles across multiple platforms — LinkedIn, Greenhouse, Ashby, Workday, Lever, iCIMS and more — your inbox becomes the only source of truth. But it is unstructured, noisy, and nearly impossible to analyze manually.

I was tracking over 100 applications with no clear picture of where each one stood. I did not know my actual response rate. I could not tell which companies had rejected me versus which were simply quiet. I had no way to spot patterns.

That is a data problem. And data problems have solutions.

### The Use Case

What I needed was a pipeline that could:

- Ingest raw, unstructured email data from multiple sender formats
- Classify each record into a meaningful status category
- Surface actionable insights in a simple, readable interface
- Update automatically as new data came in
- Allow manual corrections where automated classification fell short

This is exactly the kind of problem a business analyst or data analyst is asked to solve in a professional setting. The domain changes. The underlying challenge does not.

### The Solution

Rather than a spreadsheet or a manual log, I built an end-to-end data pipeline:

**Data ingestion:** A Python/Flask server connects to Microsoft Graph API and pulls emails matching job-related keywords across 14 search queries, paginating through results to maximize coverage.

**Classification logic:** Each email is run through a multi-layered classifier that identifies the sender domain (Greenhouse, Ashby, Workday, Lever, etc.), extracts company name and role from subject line patterns, detects status signals (applied, interview, rejected, pending) using keyword matching, and filters out noise from non-job senders.

**Data quality layer:** A manual override system lets users correct any misclassified record. Overrides persist in browser local storage and survive auto-refreshes. A dedicated rejection sync scans specifically for rejection language across 14 phrase variants and auto-updates matching records.

**Delivery:** A live HTML dashboard presents the data with real-time filtering, search, status breakdown charts, and recent activity feed. It auto-refreshes every 5 minutes.

**Integration:** The Flask server is also wrapped as a local MCP server, connecting it directly to Claude Desktop so the AI assistant can query the inbox, read emails, and surface insights conversationally.

### Why It Matters for a BA/DA Role

The skills demonstrated here map directly to what business and data analysts do day to day:

- Identifying a real problem and defining requirements before building anything
- Connecting to and querying an external data source via API
- Designing classification logic to turn unstructured data into structured records
- Building a feedback loop for data quality (the manual override and sync features)
- Presenting findings in a clear, accessible format for a non-technical audience
- Iterating based on what the data actually shows

I did not write the code myself. I worked with Claude to build it, asking questions at every step and understanding the decisions being made. I am actively upskilling in Python and SQL. This project is part of that process.

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
dashboard.html              Claude Desktop (reads inbox directly)
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
| `dashboard.html` | Live web dashboard |
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
```

### 3. Install dependencies

```bash
pip install flask flask-cors msal requests python-dotenv mcp
```

### 4. Run the server

```bash
python server.py
```

Visit http://localhost:8080/login in your browser and sign in with your Microsoft account. You only need to do this once as the token is saved locally.

### 5. Open the dashboard

Open `dashboard.html` in your browser. It connects to your local server at http://127.0.0.1:8080.

### 6. Claude Desktop integration (optional)

Add this to your Claude Desktop config file at `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook-mail": {
      "command": "python",
      "args": ["C:\\path\\to\\claude-mail-connector\\mcp_server.py"],
      "env": {}
    }
  }
}
```

Restart Claude Desktop. You should see the hammer icon indicating MCP tools are connected.

---

## Notes

The built-in Microsoft 365 connector in Anthropic's MCP library only supports work and school accounts. Personal Outlook/Live accounts are not supported through that integration. This project exists because of that gap — building a custom connector was the only way to make it work with a personal account.

---

## Author

Ryan Berkeley
ryanberkeley@live.com
