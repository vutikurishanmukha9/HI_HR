# HiHR — Smart Cold Email Outreach Platform

A high-deliverability, enterprise-grade cold email outreach platform for recruitment and talent acquisition. HiHR empowers recruiters and founders to execute personalized outreach campaigns with human-like sending cadence, real-time deliverability analytics, native DNS MX pre-flight verification, and a saved email template library. Built with React, TypeScript, Tailwind CSS, Express, and Prisma.

---

## Key Features

### HiHR Outreach Power & Safety Suite
- **Saved Email Template Library** — Save, categorize, and organize cold email templates with one-click editor insertion. Pre-loaded with 15 battle-tested starter templates across 5 categories (*Cold Outreach*, *Follow Up*, *Referral*, *Executive*, *Recruiter*).
- **Real-time Deliverability & Spam Score Analyzer** — Interactive 0–100 spam score meter analyzing copy in real time against 120+ high-risk trigger words, all-caps subject lines, excessive punctuation, and personalization tag coverage.
- **Domain DNS MX Pre-Flight Verification** — Native Node.js DNS engine (`dns.promises.resolveMx`) verifies mail server existence for recipient domains prior to dispatch. Excludes dead or typoed domains in one click to prevent hard bounces with zero third-party API costs.
- **Human-like Smart Jitter (Anti-Bot Sending)** — Dynamic randomization applied between email batches (+/-20% jitter) and individual recipient sends (300ms–800ms) to evade spam filter rate heuristics.
- **Pre-Flight SMTP Sender Verification** — Validates Gmail sender credentials directly against Google's live SMTP server before saving. Stores the authenticated account encrypted with AES-256 with 1-click continuation.

### Campaign Management & Composition
- **Dynamic Personalization Tags** — Automatically transforms any column header from uploaded CSV or Excel spreadsheets into reusable template tags (e.g., `{firstName}`, `{companyName}`, `{jobTitle}`).
- **Rich Text Editor** — Clean formatting toolbar supporting bold, italics, unordered lists, hyperlinks, and undo/redo shortcuts.
- **Attachment Support** — Drag-and-drop or browse attachments (PDF, DOCX, resumes) up to 10MB per file with automatic malicious executable filtering.
- **Future Scheduling** — Schedule campaigns for future dispatch with native date-time pickers.
- **Batching & Rate Throttling** — Configurable batch sizes and inter-batch delays to respect provider daily quotas.
- **Interactive Email Preview** — Preview the exact rendered email for any individual recipient from your contact list before dispatching.

### Contact List Management
- **Bulk Spreadsheet Ingestion** — Import contacts via CSV, XLSX, or XLS files up to 5MB.
- **Client-Side Parsing** — Native SheetJS integration bundled securely at build time without third-party script CDNs.
- **Deduplication** — Automatic identification and suppression of duplicate contact emails.
- **Selective Outreach Queue** — Select, deselect, or filter individual recipients before launch.

### Security & Architecture
- **AES-256 Encryption** — Google App Passwords encrypted at rest using AES-256-CBC with SHA-256 key entropy derivation.
- **Authenticated WebSocket Streaming** — Socket.IO real-time progress events secured via JWT handshake validation and strict multi-tenant room authorization (`run.userId === socket.userId`).
- **JWT Authentication** — 15-minute short-lived access tokens with secure token refresh (`/api/auth/refresh`) and logout revocation (`/api/auth/logout`).
- **Brute-Force & Rate Protection** — Account lockouts after 5 consecutive failed attempts (30-minute lockout window) and per-route rate limiters.
- **SMTP Header & File Defense** — Strict carriage-return and line-feed (CRLF) sanitization on subject lines; automatic rejection of executable file extensions (`.exe`, `.bat`, `.cmd`, `.sh`, `.ps1`, etc.).
- **Zero Emojis Principle** — Professional, clutter-free typography and clean semantic SVG vector icons throughout the interface and codebase.

---

## Technology Stack

### Frontend
| Component | Technology | Description |
| :--- | :--- | :--- |
| Framework | React 19 + TypeScript | Component-driven frontend with strict typing |
| Build Tool | Vite 6.x | Fast development server and optimized production bundler |
| Styling | Tailwind CSS 3.4 + tokens.css | Custom light-theme design system with Slate neutrals |
| Data Processing | SheetJS (xlsx) | Direct bundle-time Excel/CSV spreadsheet parsing |
| Sanitization | DOMPurify | XSS prevention for HTML email previews |
| Real-time | Socket.IO Client 4.x | Real-time campaign dispatch progress listeners |

### Backend
| Component | Technology | Description |
| :--- | :--- | :--- |
| Runtime | Node.js 18+ | JavaScript runtime environment |
| Framework | Express 4.x | RESTful API server |
| ORM | Prisma 5.x | Type-safe database queries and migrations |
| Database | SQLite (Dev) / PostgreSQL (Prod) | Structured relational storage (`prisma/dev.db`) |
| Mail Transport | Nodemailer 6.x | Gmail SMTP connection pooling and dispatch |
| Cryptography | Node.js Crypto | AES-256-CBC with SHA-256 derived cipher keys |
| Network DNS | Node.js DNS Promises | Native DNS MX record lookups for domain verification |
| Security | Helmet + Rate Limit | Security headers, CORS filtering, and rate limiting |

## Getting Started

### Prerequisites
- Node.js v18.x or higher
- npm v9.x or higher

### 1. Clone Repository
```bash
git clone https://github.com/vutikurishanmukha9/HR_Cold_Email.git
cd HR_Cold_Email
```

### 2. Configure Backend Environment
Navigate to the `backend/` directory and create `.env`:
```bash
cd backend
```

Create `backend/.env` with the following variables:
```env
# Server
NODE_ENV=development
PORT=5000

# Database (SQLite for development)
DATABASE_URL="file:./dev.db"

# JWT Authentication
JWT_SECRET=your-secure-jwt-access-secret-minimum-32-chars
REFRESH_TOKEN_SECRET=your-secure-jwt-refresh-secret-minimum-32-chars
JWT_EXPIRES_IN=15m
REFRESH_TOKEN_EXPIRES_IN=7d

# AES-256 Encryption (must be a 64-character hex string or 32-character key)
ENCRYPTION_KEY=0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef

# CORS
FRONTEND_URL=http://localhost:3000
```

Install backend dependencies and initialize the SQLite schema:
```bash
npm install
npx prisma db push
```

Start the backend development server:
```bash
npm run dev
```
The backend will run on `http://localhost:5000`.

### 3. Configure Frontend Environment
In a new terminal window, navigate to `frontend/`:
```bash
cd frontend
```

Create `frontend/.env`:
```env
VITE_API_URL=http://localhost:5000/api
```

Install frontend dependencies:
```bash
npm install
```

Start the frontend development server:
```bash
npm run dev
```
The frontend will run on `http://localhost:3000`.

---

## Usage Workflow

1. **Sign In or Register**:
   - Create an account using email and a secure password (evaluated via real-time password strength meter).
2. **Connect Gmail Mailbox**:
   - Choose from previously saved and verified accounts, or enter a new Gmail address with a dedicated 16-character Google App Password.
   - Credentials are verified via live SMTP connection before saving.
3. **Upload Contacts & Run Pre-Flight MX Check**:
   - Upload a CSV or Excel spreadsheet containing an `Email` column and optional custom columns (`Name`, `Company`, `Role`).
   - Click **Verify Domains (DNS MX)** to automatically check recipient domains for valid mail exchange servers.
   - Click **Exclude Dead Domains** to eliminate non-existent domains before queuing.
4. **Compose Email & Review Deliverability**:
   - Write your subject and body copy using dynamic tags (e.g. `Hi {firstName}`).
   - Click **Templates Library** to insert pre-built outreach templates or save custom drafts.
   - Inspect the real-time **Deliverability Score** (0–100) and address any flagged spam trigger words.
5. **Configure Throttling & Dispatch**:
   - Configure batch size and batch delays.
   - Keep **Human-like Smart Jitter (Anti-Bot)** enabled to randomize sending intervals.
   - Click **Launch Campaign** or schedule for future delivery.
   - Monitor real-time status updates (*Queued*, *Sending*, *Sent*, *Failed*) over WebSocket.

---

## API Reference

### Authentication (`/api/auth`)
| Method | Endpoint | Description | Auth Required |
| :--- | :--- | :--- | :--- |
| `POST` | `/api/auth/register` | Register a new user account | No |
| `POST` | `/api/auth/login` | Authenticate user and receive access/refresh tokens | No |
| `POST` | `/api/auth/refresh` | Exchange valid refresh token for a new access token | No |
| `POST` | `/api/auth/logout` | Revoke session and log out | No |
| `GET` | `/api/auth/me` | Fetch authenticated user profile | Yes |

### Email Credentials (`/api/credentials`)
| Method | Endpoint | Description | Auth Required |
| :--- | :--- | :--- | :--- |
| `GET` | `/api/credentials` | List all saved Gmail accounts for the user | Yes |
| `POST` | `/api/credentials` | Verify with Gmail SMTP and save encrypted credential | Yes |
| `DELETE` | `/api/credentials/:id` | Remove a saved Gmail sender account | Yes |

### Saved Templates (`/api/templates`)
| Method | Endpoint | Description | Auth Required |
| :--- | :--- | :--- | :--- |
| `GET` | `/api/templates` | List user custom templates + starter templates | Yes |
| `POST` | `/api/templates` | Create and save a new email template | Yes |
| `PUT` | `/api/templates/:id` | Update an existing saved template | Yes |
| `DELETE` | `/api/templates/:id` | Delete a custom saved template | Yes |

### Campaign Execution (`/api/campaigns`)
| Method | Endpoint | Description | Auth Required |
| :--- | :--- | :--- | :--- |
| `POST` | `/api/campaigns/send` | Queue and launch a cold email campaign with smart jitter | Yes |
| `POST` | `/api/campaigns/validate-domains` | Execute DNS MX lookups on recipient email domains | Yes |
| `GET` | `/api/campaigns/run/:runId/status` | Polling fallback for active campaign progress | Yes |
| `GET` | `/api/campaigns` | List historical campaigns | Yes |
| `GET` | `/api/campaigns/:id` | Retrieve campaign details | Yes |
| `DELETE` | `/api/campaigns/:id` | Delete campaign record | Yes |

### Delivery Analytics (`/api/track`)
| Method | Endpoint | Description | Auth Required |
| :--- | :--- | :--- | :--- |
| `GET` | `/api/track/stats` | Aggregate delivery counts (sent, failed, total) | Yes |
| `GET` | `/api/track/details` | Per-recipient dispatch logs and diagnostics | Yes |
| `POST` | `/api/track/reset` | Clear delivery history and reset counters | Yes |

---

## Testing & Quality Assurance

### Run Frontend Build Verification
```bash
cd frontend
npm run build
```

### Run Backend Tests & Compilation
```bash
cd backend
npm run build
node scripts/test_new_features.js
npm test
```

---

## License

Private and proprietary. Developed for high-deliverability talent acquisition and outreach workflows.
