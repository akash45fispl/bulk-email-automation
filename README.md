# 🚀 AutoMailer PRO — Bulk Email Automation Platform

A high-performance, 100% free and open-source automated bulk email sending system designed to handle **1,000+ client emails** directly from Excel/CSV spreadsheets with personalized merge variables (`{{Name}}`, `{{Company}}`), SMTP rate throttling, real-time dispatch monitoring, and live audit logging.

---

## 🌟 Key Features

1. **Excel & CSV Direct Upload**:
   - Accepts `.xlsx`, `.xls`, or `.csv` files with up to 10,000+ rows.
   - Auto-detects email & name columns while converting all other columns into clickable merge tags.
   - Cleans and flags invalid/malformed email addresses automatically.

2. **Personalized Dynamic Templating**:
   - Supports personalized variables in both Subject line and Message Body (e.g., `Hello {{Name}}, welcome to {{Company}}!`).
   - Rich HTML and Plain text format toggles.
   - Live interactive email preview for each recipient in the spreadsheet.

3. **100% Free Email Delivery (SMTP Providers)**:
   - **Personal Gmail**: Up to 500 emails/day free with Google App Passwords.
   - **Google Workspace**: Up to 2,000 emails/day.
   - **Brevo (formerly Sendinblue)**: 300 free emails/day.
   - **Outlook / Microsoft 365 / Zoho / Custom Company SMTP**: Supported seamlessly.

4. **Spam Shield & Rate Throttling**:
   - Customizable delay pacing (e.g. 2–3 seconds per email) to protect your domain reputation, bypass spam filters, and prevent temporary SMTP blocks.

5. **Real-time Live Campaign Center**:
   - Circular progress gauge with live Sent, Failed, and Queued counters.
   - Real-time audit log terminal with instant error reporting.
   - Pause, Resume, and Cancel buttons during active runs.
   - One-click **Export Results CSV** to download delivery records.

---

## 🚀 Quick Start (1-Click Launch)

### Option 1: Double-Click
Simply double-click `start.bat` in this folder. It will start the server and open your browser at `http://localhost:3000`.

### Option 2: Command Line
```bash
npm install
node server.js
```
Then visit [http://localhost:3000](http://localhost:3000) in your browser.

---

## 🔑 How to Get a Free Gmail App Password (2 Minutes)

1. Open your [Google Account Security Settings](https://myaccount.google.com/security).
2. Ensure **2-Step Verification** is turned ON.
3. In the top search bar, type **"App passwords"** and click on it.
4. Name the app `AutoMailer` and click **Create**.
5. Copy the generated **16-character code** (e.g., `abcd efgh ijkl mnop`).
6. Paste it into the **Password / App Password** field in AutoMailer PRO.

---

## 📁 Project Structure

```
email-automation/
├── server.js              # Express backend, Nodemailer queue, Excel parser, SSE
├── package.json           # Dependencies (express, nodemailer, xlsx, multer, cors)
├── sample_clients.xlsx    # Sample Excel file for immediate testing
├── start.bat              # 1-Click Windows launcher
├── public/                # Web Dashboard
│   ├── index.html         # Modern responsive UI
│   ├── styles.css         # Glassmorphic dark design system
│   └── app.js             # Client interactivity, dynamic tags, live stream
└── uploads/               # Temporary file storage (auto-managed)
```
