# WorkIQ v2 — Entra ID HR Platform
## Real SQLite persistence · Real Azure AD sync · 4 role profiles

---

## ⚡ Quick start (zero dependencies)
```bash
tar -xzf workiq-v2.tar.gz
cd workiq2
node server.js          # starts on http://localhost:3000
# or a specific port:
PORT=8080 node server.js
```
**Requires:** Node.js 22+ only. No npm install needed.

---

## 👤 Login credentials

| Role       | Email                    | Password      |
|------------|--------------------------|---------------|
| Employee   | alex@company.com         | employee123   |
| Manager    | maria@company.com        | manager123    |
| HR Admin   | sarah@company.com        | hr123         |
| Director   | david@company.com        | director123   |

---

## 🗄️ Database
- **Engine:** Node.js 22 built-in SQLite (`node:sqlite`) — no install needed
- **Location:** `data/workiq.db` (auto-created on first run)
- **All settings, calendars, employees, leave requests, balances and audit logs persist here**
- To reset: delete `data/workiq.db` and restart

---

## 🔵 Connecting Entra ID / Azure AD

### Step 1 — Create an App Registration in Azure Portal
1. Go to **Azure Portal → Microsoft Entra ID → App registrations → New registration**
2. Name: `WorkIQ`, Supported account types: **Single tenant**
3. Redirect URI: `https://your-domain.com/auth/callback`
4. Click **Register**

### Step 2 — Add API permissions
In your app registration:
1. Go to **API permissions → Add a permission → Microsoft Graph → Application permissions**
2. Add all of these:
   - `AuditLog.Read.All`
   - `Directory.Read.All`
   - `User.Read.All`
   - `SignInActivity.Read`
   - `Reports.Read.All`
3. Click **Grant admin consent**

### Step 3 — Create a client secret
1. Go to **Certificates & secrets → New client secret**
2. Copy the **Value** (not the ID) — you only see it once

### Step 4 — Enter credentials in WorkIQ
1. Log in as **HR Admin** (sarah@company.com)
2. Go to **Entra ID config** in the sidebar
3. Enter:
   - **Tenant ID** — from Azure Portal → Overview
   - **Client ID** — from App Registration → Overview
   - **Client Secret** — the value you copied
4. Click **Save credentials**, then **Test connection**
5. Click **Force sync now** to import users and sign-in logs immediately

The app will then sync every 5 minutes automatically (configurable).

---

## 🚀 Deploy to cloud (free tiers)

### Railway (easiest)
```bash
# Install Railway CLI: https://railway.app
railway login
railway init
railway up
# Set start command: node server.js
```

### Render
1. Push folder to GitHub
2. New Web Service → select repo
3. Build command: *(leave blank)*
4. Start command: `node server.js`
5. Set env var: `PORT=10000`

### Fly.io
```bash
fly launch --no-deploy
# Edit fly.toml: set [processes] app = "node server.js"
fly deploy
fly volumes create workiq_data --size 1  # for persistent DB
```
Set `DB_PATH=/data/workiq.db` as env var when using a mounted volume.

### VPS / own server
```bash
# With PM2 for auto-restart:
npm install -g pm2
pm2 start server.js --name workiq
pm2 save && pm2 startup
```

---

## 🔧 Environment variables

| Variable  | Default              | Description                      |
|-----------|----------------------|----------------------------------|
| `PORT`    | `3000`               | HTTP port                        |
| `DB_PATH` | `./data/workiq.db`   | SQLite database file path        |

---

## 📁 File structure
```
workiq2/
├── server.js        — HTTP server + all REST API routes
├── db.js            — SQLite schema, seed data, all query helpers
├── entra.js         — Microsoft Graph API sync (users + sign-in logs)
├── data/
│   └── workiq.db    — SQLite database (auto-created)
└── public/
    └── index.html   — Full single-page frontend
```
# HRadmin
