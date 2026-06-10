# Excel → Google Sheets Converter — Setup Guide

A fully client-side web app that converts `.xlsx`/`.xls` Excel files to Google Sheets.  
**No backend, no server, no data uploaded to any third party.**

---

## 🚀 Quick Start

1. Open `index.html` in your browser (or use VS Code Live Server)
2. Click ⚙️ in the top-right corner
3. Paste your **Google OAuth2 Client ID**
4. Click **Save Configuration**
5. Done! Start uploading Excel files.

---

## 🔑 Getting Your Google OAuth2 Client ID (5 minutes)

### Step 1 — Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click **Select a project** → **New Project**
3. Give it a name like `excel-sheets-converter`
4. Click **Create**

### Step 2 — Enable the Google Sheets API

1. In your project, go to **APIs & Services → Library**
2. Search for **Google Sheets API**
3. Click it and press **Enable**

### Step 3 — Configure OAuth Consent Screen

1. Go to **APIs & Services → OAuth consent screen**
2. Select **External** → Click **Create**
3. Fill in:
   - **App name**: Excel to Sheets Converter
   - **User support email**: your email
   - **Developer contact**: your email
4. Click **Save and Continue** through the remaining steps
5. On the **Test users** page, add your own Google email address
6. Click **Save and Continue**

### Step 4 — Create OAuth2 Credentials

1. Go to **APIs & Services → Credentials**
2. Click **+ Create Credentials → OAuth 2.0 Client ID**
3. Application type: **Web application**
4. Name: `Excel Converter`
5. Under **Authorized JavaScript origins**, add the URL where you'll open the app:
   - For local file: `http://localhost` (if using a local server)
   - For Live Server: `http://localhost:5500` (or whatever port VS Code uses)
   - For a hosted URL: your full domain, e.g., `https://yourdomain.com`
6. Click **Create**
7. Copy the **Client ID** (looks like `123456789-xxxx.apps.googleusercontent.com`)

### Step 5 — Paste into the App

1. Open `index.html` in the browser
2. Click ⚙️ (top-right)
3. Paste your Client ID
4. Click **Save Configuration**

---

## 🖥️ Running Locally

The easiest way is to use VS Code with the **Live Server** extension:

```bash
# Install Live Server in VS Code, then right-click index.html → Open with Live Server
# Default URL: http://localhost:5500
```

Alternatively, use Python's built-in server:

```bash
cd "path"
python -m http.server 8080
# Open: http://localhost:8080
```

Or Node.js:

```bash
npx serve .
# Opens on http://localhost:3000
```

> ⚠️ **Important**: You must open the app via `http://` (not `file://`) for Google OAuth2 to work. Make sure the URL you use matches the **Authorized JavaScript origins** you configured.

---

## ✨ Features

| Feature | Details |
|---|---|
| 📂 Drag & Drop | Drop `.xlsx`, `.xls`, `.ods` files |
| 👁️ Live Preview | See all sheets and data before converting |
| 🔐 Google Sign-In | Secure OAuth2 — no passwords stored |
| ✨ New Spreadsheet | Creates a Google Sheet with the Excel filename |
| 🎯 Write to Existing | Paste a Sheet URL to write data into it |
| 📋 Multi-sheet | All Excel tabs → Google Sheet tabs |
| 🎨 Header Formatting | First row is auto-bolded and frozen |
| 📊 Format Preservation | Numbers, dates, booleans handled correctly |
| 🔗 One-click Link | Direct link to open the converted Sheet |

---

## 🔒 Privacy & Security

- **Zero backend**: All processing happens in your browser
- **SheetJS** parses Excel locally — the file is never uploaded anywhere
- **Google OAuth2** tokens are stored in `sessionStorage` only (cleared when tab closes)
- **Client ID** is stored in `localStorage` — only in your own browser

---

## 🛠️ Project Structure

```
Ecxcel Convo/
├── index.html        — App UI
├── styles.css        — Dark glassmorphism design system
├── app.js            — Main orchestration logic
├── auth.js           — Google OAuth2 flow
├── excel-parser.js   — SheetJS wrapper (Excel → JSON)
├── sheets-api.js     — Google Sheets REST API client
└── README.md         — This file
```

---

## ❓ Troubleshooting

**"Popup was blocked"**  
→ Allow popups for `localhost` in your browser settings.

**"redirect_uri_mismatch" error from Google**  
→ The URL in your browser doesn't match what's in your OAuth2 credentials. Add the exact origin (e.g., `http://localhost:5500`) to **Authorized JavaScript origins** in Google Cloud Console.

**"Access blocked: This app has not been verified"**  
→ During development, click **Advanced → Go to Excel Converter (unsafe)**. To remove this warning, submit your app for Google verification.

**Token expired**  
→ Sign out and sign back in. Tokens last 1 hour.

---

## 📄 License

MIT — free to use, modify, and distribute.
