# Clinical & Volunteer Operations SaaS Platform

A full-stack recruitment, onboarding, and tele-consultation management platform built to streamline volunteer intake, clinical staff credentialing, and real-time virtual interviews for youth-focused community clinic operations.

## Overview

Coordinating high-volume youth volunteer cohorts and clinical support personnel requires structured intake, rapid credential verification, and seamless communication. This platform replaces fragmented spreadsheet workflows with a unified **PHP/MySQL** application featuring **OAuth 2.0** authentication, **Google Gemini API**-assisted candidate screening, automated email dispatch, and embedded **WebRTC** video rooms for remote onboarding and triage consultations.

### Key Capabilities
* **Multi-Stage Intake & Credentialing Form Engine:** Captures academic qualifications (secondary through postgraduate), clinical/volunteer experience, demographic records, and document uploads (`CV_URL` / certification attachments).
* **AI-Assisted Screening (Gemini API):** Parses applicant backgrounds and uploaded CVs against role requirements (`Settings: JOB_DESCRIPTION`) to generate structured candidate summaries and role-fit scores.
* **Pipeline & Status Tracking:** Real-time state transitions (`Applied` -> `Shortlisted` -> `Interview Scheduled` -> `Onboarded` / `Rejected`) with full audit attribution (`Updated_By`, `Updated_At`).
* **Embedded WebRTC Video Consultations:** Peer-to-peer browser video sessions integrated directly into the interview scheduling module (`Interviews`), eliminating external meeting link dependencies.
* **Automated Mail-Merge & Audit Logging:** Dynamic template rendering (`confirmation`, `shortlist`, `interview`, `rejection`) with complete delivery tracking in `EmailLog`.

---

## System Architecture & Tech Stack

| Layer | Technology | Role |
| :--- | :--- | :--- |
| **Backend** | PHP 8.2+ (PDO) | RESTful routing, form validation, session management, and mail queue dispatch |
| **Database** | MySQL 8.0 | Relational persistence for intake records, status logs, templates, and interview schedules |
| **Authentication** | OAuth 2.0 (Google Identity) | Role-based access control (RBAC) for administrators, coordinators, and applicants |
| **AI Integration** | Google Gemini API | Automated CV parsing, qualification matching, and candidate briefing generation |
| **Real-Time Comms** | WebRTC + WebSockets | Low-latency video/audio rooms for remote volunteer interviews and clinical triage |
| **Storage Integration** | Google Drive API | Cloud document storage and retrieval for applicant CVs and certifications |

---

## Database Schema

The relational schema is normalized across six core operational tables:

1. **`applications`** — Master intake table storing personal details, academic history (`ssc_*`, `hsc_*`, `bachelor_*`, `postgraduate_*`), prior organizational experience, and cloud document pointers (`cv_url`).
2. **`candidate_status`** — Tracks pipeline state transitions (`status`, `updated_at`, `updated_by`) indexed by candidate email.
3. **`interviews`** — Manages scheduled virtual and on-site evaluations (`scheduled_at`, `candidate_email`, `interview_date`, `location` / WebRTC room ID, `interviewer`, `notes`, `status`).
4. **`email_templates`** — Stores reusable transactional templates (`confirmation`, `shortlist`, `interview`, `rejection`) with dynamic variable interpolation (`[Candidate Name]`, `[Position]`, `[Date]`, `[Time]`, `[Location/Meeting Link]`, `[Interviewer Name]`).
5. **`email_log`** — Immutable audit trail of all outbound notifications (`timestamp`, `recipient`, `subject`, `body`, `sent_by`).
6. **`settings` & `users`** — System-wide configuration parameters (active role descriptions, screening rubrics) and OAuth-linked administrator accounts.

---

## Local Development Setup

### Prerequisites
* PHP >= 8.2 with `pdo_mysql`, `curl`, and `openssl` extensions enabled
* MySQL >= 8.0 or MariaDB >= 10.6
* Composer
* Google Cloud Console project (configured for **OAuth 2.0**, **Drive API**, and **Gemini API**)

### Installation

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/](https://github.com/)<Gosa-n-geleto>/clinical-volunteer-saas.git
   cd clinical-volunteer-saas
