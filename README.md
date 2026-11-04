# VICE VERSE 2026 - Problem Statement Selection Portal

A modern, real-time web portal for student teams to explore and claim official **Smart India Hackathon (SIH) 2025** Problem Statements for **VICE VERSE 2026 - VVCE, Mysuru**.

Connected directly with **Google Sheets** via **Google Apps Script**, featuring **atomic concurrency locking** so that **only one team can select each problem statement**, with live status synchronization across all devices.

---

## 🌟 Key Features

1. **Rich Domain & Category Explorer (Not a simple dropdown)**:
   - Filter by Domain: **CSE / AIML (15)**, **Electronics & Communication (10)**, and **Mechanical & Civil (5)**.
   - Filter by Category: **Software (💻)** vs **Hardware (⚙️)**.
   - Filter by Status: **Available Only (🟢)** vs **Claimed (🔒)**.
   - Instant search across PS Code, SIH ID, Title, Ministry/Org, and Themes.
   - **Custom Domain-Grouped Combobox / Navigation**: Quick navigation dropdown organized with domain headers, status badges, and search.
   - **Case Study Modal**: Full background text, problem understanding, and presentation guidelines (7-minute pitch + 3-minute Q&A) from the official booklet.

2. **Real-Time Google Sheets Synchronization**:
   - Background live polling every 4 seconds.
   - Live synchronization indicator (`🟢 Live Syncing`) with sync countdown timer.
   - If Team A claims a problem statement, it immediately flips to `🔒 Claimed by Team A` on everyone else's screen in real time.

3. **Concurrency-Safe Atomic Locking (`LockService`)**:
   - Powered by Google Apps Script `LockService.getScriptLock()`.
   - Prevents race conditions: If two teams click "Confirm" on the exact same problem statement within milliseconds of each other, the lock ensures only the first team succeeds and the second is immediately informed with a clear alert.
   - **1 Team = 1 Problem Statement Rule**: Prevents any team from claiming multiple problem statements.

4. **Team Selection Receipt**:
   - Generates a confirmation ticket with Team details, Leader Contact, PS Code, SIH Reference ID, and official timestamp.
   - Printable slip button for documentation.

5. **Audit Logging & Admin Controls**:
   - Maintains a dedicated `Selection_Audit_Log` sheet tracking all attempts, timestamps, and outcomes.
   - Admin functions to reset selections or release specific statements if needed.

---

## 📁 Repository Structure

```text
├── index.html                 # Main frontend application (Tailwind CSS + Lucide Icons)
├── app.js                     # Client logic, real-time sync engine, atomic claim flow
├── data.js                    # All 30 SIH 2025 Problem Statements from the booklet
├── problem_statements.json    # Structured JSON database of the 30 case studies
├── generate_data.py           # Python generator script
├── Code.gs                    # Google Apps Script Backend (Copy to Google Sheet)
└── README.md                  # Setup & Deployment Guide
```

---

## 🚀 Setup & Google Sheets Deployment Guide

### Step 1: Prepare the Google Sheet

1. Create a new Google Spreadsheet at [sheets.new](https://sheets.new) (e.g. name it `VVCE Ideathon 2025 - PS Selection`).
2. In the top menu, click **Extensions > Apps Script**.
3. Delete any default code in `Code.gs`.
4. Copy and paste the entire contents of [`Code.gs`](./Code.gs) into the editor.
5. Save the project (click the disk icon or press `Ctrl + S`).

### Step 2: Initialize the Sheets (1-Click Setup)

1. In the Apps Script toolbar, make sure the function dropdown has `setupSheet` selected.
2. Click **Run**.
3. Authorize the permissions when prompted by Google.
4. Once completed, your Google Sheet will instantly have two sheets created and styled:
   - `Problem_Statements`: All 30 problem statements pre-loaded with conditional formatting (Green = Available, Red = Claimed).
   - `Selection_Audit_Log`: Real-time audit log of all selections.

### Step 3: Deploy as a Web App

1. In the top right of Apps Script, click **Deploy > New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Fill in the deployment details:
   - **Description**: `Ideathon Selection API v1`
   - **Execute as**: `Me (your email)`
   - **Who has access**: **`Anyone`** *(Crucial: This allows team participants to connect without needing Google logins)*.
4. Click **Deploy**.
5. Copy the generated **Web App URL** (it looks like `https://script.google.com/macros/s/.../exec`).

### Step 4: Connect the Web App

1. Open `index.html` in your browser (or host it via GitHub Pages).
2. Click the **Google Sheets Setup** button in the top navigation bar.
3. Paste your copied Web App URL into the field.
4. Click **Test Connection** &rarr; you should see a green checkmark!
5. Click **Save & Connect**.
6. The sync badge will turn bright green: `🟢 Live Syncing`!

---

## 🌐 Deploy to GitHub Pages (Live for Participants)

1. In GitHub, go to your repository: `ivc-ps-select`.
2. Navigate to **Settings > Pages**.
3. Under **Branch**, select `main` (or `master`) and folder `/ (root)`.
4. Click **Save**.
5. Within 1-2 minutes, your website will be live at:
   `https://aniruddhamjois.github.io/ivc-ps-select/`

---

## 📊 Problem Statements Summary (30 Case Studies)

| Domain | Count | PS Codes | Category Split |
| :--- | :---: | :--- | :--- |
| **CSE / AIML** | 15 | `CS-01` to `CS-15` | 15 Software |
| **Electronics & Comm (ECE)** | 10 | `EC-01` to `EC-10` | 10 Hardware |
| **Mechanical & Civil (ME/CV)** | 7 | `ME-01` to `ME-03`, `CV-01` to `CV-04` | 2 Software, 5 Hardware |
| **Total** | **32** | Official SIH 2025 PS Numbers | **17 Software / 15 Hardware** |

---

## 🛡️ Concurrency Control Details

When multiple teams attempt to claim problem statements at the same time:
```javascript
const lock = LockService.getScriptLock();
const hasLock = lock.tryLock(10000); // Wait up to 10s for atomic access
try {
  // 1. Verify team hasn't already claimed any statement
  // 2. Verify target statement status is still 'Available'
  // 3. Atomically write team data to sheet
} finally {
  lock.releaseLock();
}
```
If a team loses the race by a millisecond, they receive an immediate popup informing them the statement was claimed, and their screen dynamically refreshes to show the remaining available choices.

---

&copy; 2025 Vidyavardhaka College of Engineering (VVCE), Mysuru. Developed for Ideathon 2025.
