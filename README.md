# 📋 Serverless ATS & Job Application Portal

A lightweight, zero-cost **Applicant Tracking System (ATS)** and **Job Application Portal** built using **React 18** (frontend) and **Google Apps Script** (backend), powered by Google Sheets, Google Drive, and Gmail.

[![Live Demo](https://img.shields.io/badge/Live-Demo-2e7d32?style=for-the-badge&logo=google-chrome&logoColor=white)](https://script.google.com/macros/s/AKfycbzIIJ1go0McvQLlswdTo9GmGa0stA9DzOpTsU0burRoM5PxzMC7EaIHpattSd1aonPiug/exec)

🔗 **Live Demo URL:** [Open Web App](https://script.google.com/macros/s/AKfycbzIIJ1go0McvQLlswdTo9GmGa0stA9DzOpTsU0burRoM5PxzMC7EaIHpattSd1aonPiug/exec)

---

## ✨ Key Features

### 👤 For Applicants
* **Dynamic Job Board:** Displays real-time job descriptions stored directly in Google Sheets.
* **Structured Application Form:** Captures personal details, academic history (SSC, HSC, Bachelor, Postgraduate), and work experience.
* **CV/Resume Upload:** Direct PDF upload (up to 5MB) converted to Base64 and saved into a dedicated Google Drive folder.
* **Automated Confirmation:** Instant on-screen confirmation modal and email acknowledgement via Gmail.

### 🛡️ For Recruiters / Admins
* **Password-Protected Portal:** Secure admin login with session management.
* **Interactive Analytics Dashboard:** Real-time KPI charts powered by **Chart.js** (shortlist/rejection rates, application trends, gender & experience distribution).
* **Candidate Pipeline Management:** Filter, search, and update candidate statuses (`Pending`, `Shortlisted`, `Rejected`).
* **CV Viewer:** Built-in PDF preview modal with direct Google Drive integration.
* **Interview Management:** Schedule and reschedule interviews with automatic calendar/email invite dispatch.
* **Email Automation:** Customizable email templates with placeholders and bulk-messaging capabilities.

---

## 🛠️ Tech Stack

* **Frontend:** React 18 (via CDN), Babel Standalone, Chart.js 4.4, Font Awesome 6
* **Backend:** Google Apps Script (V8 Engine)
* **Database:** Google Sheets
* **File Storage:** Google Drive
* **Mail Dispatch:** Gmail API / `MailApp`

---

## 📂 Project Structure

```text
├── Code.gs          # Apps Script backend API, Drive upload logic & Sheets database handlers
├── index.html       # Single-file React SPA (Application form + Admin dashboard)
└── README.md        # Project documentation
