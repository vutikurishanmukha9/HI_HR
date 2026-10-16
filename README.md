# HiHR — Enterprise Cold Email Outreach & Recruitment Platform

A high-deliverability, enterprise-grade cold email outreach engine built specifically for recruitment, talent acquisition, and targeted executive outreach. HiHR empowers recruitment teams, founders, and sales leaders to execute hyper-personalized email campaigns at scale with human-like sending cadence, native multi-layer email verification, an advanced deliverability engine, and a rich library of battle-tested outreach templates.

Built with React 19, TypeScript, Tailwind CSS, Node.js, Express, Prisma, and SQLite/PostgreSQL.

---

## Table of Contents

- [Core Architectural Highlights](#core-architectural-highlights)
- [Key Features](#key-features)
  - [1. Multi-Layer Email Verification & Domain Safety](#1-multi-layer-email-verification--domain-safety)
  - [2. Advanced Deliverability & Spam Analysis Engine](#2-advanced-deliverability--spam-analysis-engine)
  - [3. Curated Outreach Templates Library](#3-curated-outreach-templates-library)
  - [4. Campaign Execution & Anti-Bot Smart Jitter](#4-campaign-execution--anti-bot-smart-jitter)
  - [5. Dynamic Contact Ingestion & Personalization](#5-dynamic-contact-ingestion--personalization)
  - [6. Real-Time Delivery Tracking & Analytics](#6-real-time-delivery-tracking--analytics)
  - [7. Enterprise Security & Cryptography](#7-enterprise-security--cryptography)
- [Design System & UI/UX Standards](#design-system--uiux-standards)
- [Technology Stack](#technology-stack)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Backend Setup](#backend-setup)
  - [Frontend Setup](#frontend-setup)
- [Usage Workflow](#usage-workflow)
- [API Reference](#api-reference)
  - [Authentication](#authentication-apiauth)
  - [Email Credentials](#email-credentials-apicredentials)
  - [Outreach Templates](#outreach-templates-apitemplates)
  - [Campaign Management](#campaign-management-apicampaigns)
  - [Tracking & Analytics](#tracking--analytics-apitrack)
- [Testing & Quality Assurance](#testing--quality-assurance)
- [License](#license)

---

## Core Architectural Highlights

- **Zero Third-Party Verification Fees** — Performs pre-flight RFC 5322 syntax validation, multi-provider disposable email identification, and native asynchronous DNS MX record lookups directly via Node.js DNS promises (`dns.promises.resolveMx`), eliminating the need for expensive third-party email verification APIs.
- **Smart Jitter Anti-Bot Dispatch** — Randomizes batch delivery (+/-20% variance) and per-recipient dispatch delays (300ms–800ms) to emulate authentic human sending habits and avoid email service provider (ESP) spam-trap rate limits.
- **Client-Side Spreadsheet Processing** — Employs native SheetJS (`xlsx`) bundled at build time to parse Excel (`.xlsx`, `.xls`) and CSV files client-side, dynamically discovering custom columns without transmitting raw files to cloud endpoints.
- **End-to-End Cryptography** — Sensitive SMTP passwords (such as 16-character Google App Passwords) are encrypted at rest using AES-256-CBC with SHA-256 key derivation.
- **Strict Rectangular Geometry System** — High-density, professional enterprise design language utilizing crisp `rounded-none` geometry across all modals, tables, badges, tabs, and action cards.
- **Zero-Emoji Policy** — Professional typographical layout utilizing semantic SVG vector icons and monochromatic Slate neutral tokens.

---

## Key Features

### 1. Multi-Layer Email Verification & Domain Safety
- **Dynamic Disposable & Temporary Mail Detection** — Deep registry covering 3,000+ throwaway and temporary mail services (e.g., `hi2.in`, `mailinator.com`, `10minutemail.com`, `tempmail.com`, `guerrilla.com`). Flags disposable mailboxes with dedicated high-visibility warning badges.
- **Native DNS MX Pre-Flight Checks** — Queries authoritative name servers to verify active Mail Exchange (MX) records before dispatching campaigns.
- **Mailbox Classification** — Automatically tags recipients into semantic categories:
  - *Corporate / Work Domain* (e.g., custom company domains with verified corporate MX records).
  - *Personal Webmail* (e.g., Gmail, Outlook, Yahoo, iCloud, Proton).
  - *Disposable / Temporary* (throwaway inboxes flagged for exclusion).
- **One-Click Dead & Disposable Exclusion** — Allows users to strip unverified or temporary email addresses from the active campaign queue with a single click, protecting domain sender reputation and eliminating hard bounces.

### 2. Advanced Deliverability & Spam Analysis Engine
- **Real-Time 0–100 Deliverability Scoring** — Evaluates subject lines and message bodies as you type, generating an instant deliverability rating (*Optimal*, *Moderate Risk*, *High Risk*).
- **250+ Spam Trigger Word Database** — Categorized detection across:
  - *High-Pressure Urgency* ("act now", "urgent response required", "expires today").
  - *Financial & Money Claims* ("guaranteed income", "free cash", "no risk").
  - *Exaggeration & Hype* ("100% free", "miracle", "once in a lifetime").
  - *Recruitment Spam Jargon* ("immediate wire transfer", "unclaimed funds").
- **Context-Aware Heuristics** — Flags excessive capitalization, multiple consecutive punctuation marks (`???`, `!!!`), lack of recipient personalization, excessive link counts, and oversized email bodies.
- **Actionable Remediation Tips** — Pinpoints exact trigger words by location (*Subject* vs *Body*) and provides instant, context-specific substitution guidance.

### 3. Curated Outreach Templates Library
- **30+ High-Conversion Starter Templates** — Pre-loaded across core talent acquisition and business development workflows:
  - *Candidate Sourcing & Cold Outreach* (Passive engineers, design leads, product managers).
  - *Executive & Leadership Recruiting* (VP/C-suite confidential outreach).
  - *Referral & Mutual Connections* (Warm introductions, alumni outreach).
  - *Follow-Up Sequences* (Polite follow-up, value-add touchpoint, breakup email).
  - *Sales & B2B Lead Generation* (Problem-focused pitches, case study sharing).
- **Custom Template Management** — Save, edit, categorize, and delete personal templates with instant insertion into the rich-text editor.
- **Dynamic Personalization Insertion** — Interpolates dynamic tags (`{name}`, `{firstName}`, `{company}`, `{role}`, and custom spreadsheet headers) directly at the editor cursor.

### 4. Campaign Execution & Anti-Bot Smart Jitter
- **Configurable Rate Limiting** — Set batch sizes (e.g., 5 to 50 emails per batch) and inter-batch intervals (e.g., 10 to 300 seconds) to match daily Gmail or SMTP provider quotas.
- **Smart Jitter Mechanics** — Introduces natural human variance to both the delay between batches and the pause between individual emails, bypassing algorithmic volume detection on recipient mail servers.
- **Scheduled Sending** — Schedule campaigns for automatic launch at specific calendar dates and times.
- **Real-Time WebSocket Streaming** — Live campaign status broadcast via Socket.IO with automatic fallback to status polling.

### 5. Dynamic Contact Ingestion & Personalization
- **Spreadsheet Flexibility** — Supports `.csv`, `.xlsx`, and `.xls` files up to 5MB.
- **Dynamic Header Mapping** — Every non-standard spreadsheet column (e.g., `github_handle`, `recent_project`, `mutual_connection`) is automatically registered as a dynamic tag.
- **Deduplication Engine** — Automatically detects and removes duplicate recipient addresses during file ingestion.
- **Live Dynamic Simulation** — Step through preview cards to view how each individual recipient will see their personalized email, including dynamic variable substitutions.

### 6. Real-Time Delivery Tracking & Analytics
- **Granular Status Tracking** — Real-time tracking of every recipient across four states: *Queued*, *Sending*, *Sent*, and *Failed*.
- **Detailed Error Diagnostics** — Captures specific SMTP error codes and failure reasons (e.g., authentication rejection, invalid recipient, connection timeout) for rapid troubleshooting.
- **Analytics Dashboard** — Visual aggregation of total dispatched emails, success rate percentages, and recent campaign performance.
- **Searchable Delivery Logs** — Filter dispatch logs by status, campaign, or recipient email address.

### 7. Enterprise Security & Cryptography
- **AES-256 Credential Encryption** — Google App Passwords and SMTP credentials are encrypted before storage using AES-256-CBC cipher with unique initialization vectors (IVs).
- **JWT Authentication** — 15-minute access tokens paired with secure HTTP-only refresh tokens (`/api/auth/refresh`) and session revocation.
- **Brute-Force & Rate Limiting** — Account lockout defense (5 failed attempts trigger a 30-minute lock) and Express rate-limit middleware on critical endpoints.
- **SMTP Header Injection Defense** — Strict carriage-return and line-feed (`\r\n`) sanitization on email subjects and recipient headers.
- **Attachment Filtering** — Rejection of executable extensions (`.exe`, `.bat`, `.cmd`, `.sh`, `.ps1`, `.msi`) and 10MB per-file size caps.

---

## Design System & UI/UX Standards

HiHR utilizes a monochromatic, warm slate/neutral design system engineered for enterprise productivity:

- **Strict Rectangular Geometry** — All cards, badges, chips, buttons, modals, and input fields utilize `rounded-none` edges, avoiding oval or pill shapes. Circular styling is reserved strictly for animated loading spinners.
- **No Text Overflow (Zero-Overflow Guarantee)** — All dynamic containers implement `min-w-0`, `break-words`, and capped truncation (`max-w-[200px] truncate` on table cells, `max-w-[240px]` on attachments, `max-w-[260px]` on error diagnostics) with native `title` hover attributes.
- **Clean Neutral Typography** — Styled using system fonts and Google Inter/Roboto typefaces with balanced text contrast hierarchies (`text-ink`, `text-ink-secondary`, `text-ink-muted`).
- **Zero-Emoji Policy** — Clutter-free design adhering to corporate communication standards with custom SVG vector icons.

---

## Technology Stack

### Frontend
| Component | Technology | Version | Purpose |
| :--- | :--- | :--- | :--- |
| Core Library | React | 19.x | Component-driven UI architecture |
| Language | TypeScript | 5.x | Strict end-to-end static typing |
| Build Tool | Vite | 6.x | Fast HMR dev server & production bundling |
| Styling | Tailwind CSS | 3.4.x | Utility-first styling with custom design tokens |
| Spreadsheet Parser | SheetJS (xlsx) | 0.18.5 | Client-side Excel & CSV parsing |
| HTML Sanitizer | DOMPurify | 3.x | XSS protection for email HTML previews |
| Real-Time Client | Socket.IO Client | 4.8.x | Real-time campaign dispatch event listeners |

### Backend
| Component | Technology | Version | Purpose |
| :--- | :--- | :--- | :--- |
| Runtime | Node.js | 18+ | High-performance asynchronous JavaScript runtime |
| Web Framework | Express | 4.x | RESTful routing and middleware engine |
| ORM | Prisma | 5.x | Type-safe database schema modeling & queries |
| Database | SQLite / PostgreSQL | - | Development (`dev.db`) / Production database |
| Mail Transport | Nodemailer | 6.x | SMTP connection pooling and email transmission |
| Cryptography | Node.js Crypto | Built-in | AES-256-CBC credential encryption |
| DNS Verification | Node.js DNS Promises | Built-in | Native asynchronous MX record lookups |
| Security Suite | Helmet, bcrypt, express-rate-limit | - | HTTP headers, password hashing, route throttling |

---

## Getting Started

### Prerequisites
- **Node.js**: v18.17.0 or higher
- **npm**: v9.0.0 or higher
- **Google Account**: A Gmail or Google Workspace account with 2-Step Verification enabled and a generated **16-character Google App Password**.

---

### Backend Setup

1. **Navigate to the backend directory**:
   ```bash
   cd backend
   ```

2. **Install dependencies**:
   ```bash
   npm install
   ```

3. **Configure Environment Variables**:
   Create a `.env` file in the `backend/` directory:
   ```env
   # Server Configuration
   PORT=5000
   NODE_ENV=development

   # Database (SQLite for local development)
   DATABASE_URL="file:./dev.db"

   # JWT Authentication Secrets (Must be at least 32 characters)
   JWT_SECRET=your-secure-jwt-access-secret-minimum-32-chars
   REFRESH_TOKEN_SECRET=your-secure-jwt-refresh-secret-minimum-32-chars
   JWT_EXPIRES_IN=15m
   REFRESH_TOKEN_EXPIRES_IN=7d

   # AES-256 Encryption Key (64-character hex string or 32-character ASCII string)
   ENCRYPTION_KEY=0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef

   # Allowed Frontend Origin
   FRONTEND_URL=http://localhost:3000
   ```

4. **Initialize Database Schema**:
   ```bash
   npx prisma db push
   ```

5. **Start Development Server**:
   ```bash
   npm run dev
   ```
   The backend API will start on `http://localhost:5000`.

---

### Frontend Setup

1. **Navigate to the frontend directory**:
   ```bash
   cd frontend
   ```

2. **Install dependencies**:
   ```bash
   npm install
   ```

3. **Configure Environment Variables**:
   Create a `.env` file in the `frontend/` directory:
   ```env
   VITE_API_URL=http://localhost:5000/api
   ```

4. **Start Development Server**:
   ```bash
   npm run dev
   ```
   The web application will be accessible at `http://localhost:3000`.

---

## Usage Workflow

```
[ Step 1: Connect Mailbox ] ──> [ Step 2: Upload Contacts ] ──> [ Step 3: Compose Email ] ──> [ Step 4: Review & Send ]
        │                               │                               │                               │
  Verify SMTP via                 Auto-Parse CSV/XLSX             Live 0-100 Deliverability       Smart Jitter Throttling
  Live Google Handshake           Run DNS MX Lookups              30+ Outreach Templates          Real-Time Dispatch Feed
  AES-256 Storage                 Exclude Dead/Disposable         Dynamic Variable Tags           Detailed Error Logs
```

1. **Sign In & Authenticate**:
   - Register or sign in using your credentials. Passwords are validated using an interactive strength meter.
2. **Connect Gmail / SMTP Mailbox**:
   - Enter your Gmail address and 16-character Google App Password.
   - HiHR performs a live pre-flight SMTP handshake before saving the encrypted credentials.
3. **Upload Contacts & Run Pre-Flight Verification**:
   - Upload your candidate or lead spreadsheet (`.csv`, `.xlsx`).
   - Click **Verify Domains (DNS MX)** to identify dead mail servers and disposable addresses.
   - Click **Exclude Dead Domains** or **Exclude Disposable** to clean the queue before sending.
4. **Compose Email & Optimize Deliverability**:
   - Insert dynamic tags (e.g. `{firstName}`, `{company}`) matching your spreadsheet columns.
   - Browse the **Templates Library** for high-performing cold outreach copy.
   - Review the **Deliverability & Spam Engine** to eliminate spam trigger words and optimize your spam score.
5. **Review Campaign & Launch**:
   - Toggle **Human-like Smart Jitter (Anti-Bot)** to randomize send intervals.
   - Configure batch sizes and delays.
   - Launch immediately or set a scheduled time.
   - Monitor real-time delivery logs on the **Analytics Dashboard**.

---

## API Reference

### Authentication (`/api/auth`)
| Method | Endpoint | Description | Auth Required |
| :--- | :--- | :--- | :--- |
| `POST` | `/api/auth/register` | Register a new user account | No |
| `POST` | `/api/auth/login` | Authenticate user and obtain access/refresh tokens | No |
| `POST` | `/api/auth/refresh` | Exchange valid refresh token for a new access token | No |
| `POST` | `/api/auth/logout` | Revoke current user session | No |
| `GET` | `/api/auth/me` | Fetch authenticated user profile | Yes |

### Email Credentials (`/api/credentials`)
| Method | Endpoint | Description | Auth Required |
| :--- | :--- | :--- | :--- |
| `GET` | `/api/credentials` | List saved sender accounts for the user | Yes |
| `POST` | `/api/credentials` | Verify with SMTP and save encrypted credentials | Yes |
| `DELETE` | `/api/credentials/:id` | Remove a saved sender account | Yes |

### Outreach Templates (`/api/templates`)
| Method | Endpoint | Description | Auth Required |
| :--- | :--- | :--- | :--- |
| `GET` | `/api/templates` | Retrieve starter templates and custom user templates | Yes |
| `POST` | `/api/templates` | Create a new custom template | Yes |
| `PUT` | `/api/templates/:id` | Update an existing template | Yes |
| `DELETE` | `/api/templates/:id` | Delete a custom template | Yes |

### Campaign Management (`/api/campaigns`)
| Method | Endpoint | Description | Auth Required |
| :--- | :--- | :--- | :--- |
| `POST` | `/api/campaigns/send` | Queue and launch cold email campaign with smart jitter | Yes |
| `POST` | `/api/campaigns/validate-domains` | Perform native DNS MX pre-flight domain checks | Yes |
| `GET` | `/api/campaigns/run/:runId/status` | Fetch status of an active campaign run | Yes |
| `GET` | `/api/campaigns` | List historical campaign dispatches | Yes |
| `GET` | `/api/campaigns/:id` | Get details and recipient logs of a campaign | Yes |
| `DELETE` | `/api/campaigns/:id` | Delete a campaign record | Yes |

### Tracking & Analytics (`/api/track`)
| Method | Endpoint | Description | Auth Required |
| :--- | :--- | :--- | :--- |
| `GET` | `/api/track/stats` | Aggregate delivery statistics (sent, failed, success rate) | Yes |
| `GET` | `/api/track/details` | Granular per-recipient dispatch logs with error diagnostics | Yes |
| `POST` | `/api/track/reset` | Clear tracking history and reset metric counters | Yes |

---

## Testing & Quality Assurance

### Frontend Type-Check & Production Build
```bash
cd frontend

# Verify static typing with TypeScript
npm run type-check

# Verify Vite production bundling
npm run build
```

### Backend Type-Check, Build & Feature Verification
```bash
cd backend

# Compile TypeScript backend
npm run build

# Run feature verification test suite
node scripts/test_new_features.js

# Execute Jest unit test suite
npm test
```

---

## License

Private and proprietary. Developed for high-deliverability talent acquisition, cold recruitment outreach, and executive sourcing.
