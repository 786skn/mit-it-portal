# MIT ACSC – IT Section Portal

**MIT Arts, Commerce & Science College, Alandi, Pune**  
IT Section — Authorised Personnel Only

---

## 🌐 Live Portal

Once deployed: `https://YOUR-USERNAME.github.io/mit-it-portal/`

---

## 📁 Files

| File | Description |
|------|-------------|
| `index.html` | Main portal application |
| `Code.gs` | Google Apps Script backend (paste into GAS editor) |

---

## 🚀 Quick Setup

### 1. Google Apps Script
1. Open [script.google.com](https://script.google.com)
2. Create new project → paste `Code.gs` contents
3. Run `setupAllSheets()` → authorise → creates all sheet tabs
4. Run `setupAllTriggers()` → sets up daily email reports
5. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Copy the `/exec` URL

### 2. Portal
1. Log in as **admin** (default: `mitit@2025`)
2. Paste the GAS URL in the setup banner
3. Click **🔗 Test Connection** → should show ✅ green
4. Click **Save**

---

## 👤 Default Accounts

| Username | Role | First Login |
|----------|------|-------------|
| `admin` | Administrator | `mitit@2025` |
| `rutuj` | IT Tech | `rutuj@123` |
| `sandeep` | IT Tech | `sandeep@123` (must change) |
| `mangesh` | IT Tech | `mangesh@123` (must change) |
| `pankaj` | IT Tech | `pankaj@123` |
| `ziyaafshan` | IT Tech | `ziyaafshan@123` |
| `ashwni` | IT Tech | `ashwni@123` |
| `bhavik` | IT Tech | `bhavik@123` |
| `director` | View Only | `director@mit` (must change) |
| `registrar` | View Only | `reg@mit2025` (must change) |

**Change all passwords after first login.**  
Add new users via: More → 👤 User Management

---

## 📧 Email Reports

Configured to send to: `sknadaf@mitacsc.ac.in`

| Report | Schedule |
|--------|----------|
| Daily IT Summary | 7:00 AM daily |
| Pending Task Reminders | 8:00 AM daily (per-user) |
| Weekly Summary | 8:00 AM every Monday |
| Critical Alert | Instant (when Critical ticket logged) |

---

## 🛠 Modules

Task Log · Vendor WOs · Equipment Register · IT Asset Inventory  
Staff Register · Vendor Register · Labs Master · Budget · Delivery Challan  
Scrap Assets · Asset Handover · GST Invoice · Document Library  
Department Report · Global Search (admin) · User Management (admin)

---

*MIT ACSC IT Section | Alandi, Pune – 412105*
