# HiHR Backend API

The Express & TypeScript REST API server powering the HiHR Cold Email Outreach platform. Designed for high deliverability, security, and real-time dispatch monitoring.

---

## Core Capabilities

- **JWT Authentication & Session Control**: Access and refresh tokens with account lockout protection (5 failed attempts = 30-minute lockout).
- **AES-256 Encrypted Credential Vault**: Google App Passwords encrypted at rest using AES-256-CBC with SHA-256 key entropy derivation. Pre-authenticated via Gmail SMTP before saving.
- **Native DNS MX Pre-Flight Engine**: Validates recipient domain mail exchange records via Node's native `dns.promises.resolveMx` with in-memory caching and batching to prevent hard bounces.
- **Human-like Smart Jitter Dispatch**: Randomized batch delays (+/-20%) and individual email delays (300–800ms) to evade automated spam filter heuristics.
- **Saved Templates Engine**: Full CRUD support for custom email templates and automatic seeding of curated starter templates.
- **Authenticated WebSocket Progress**: Real-time Socket.IO dispatch updates secured via JWT authentication and strict campaign room multi-tenant isolation.
- **Input Sanitization**: CRLF injection prevention on email subjects, executable attachment blocking, and Zod schema validation.

---

## Prerequisites

- Node.js (v18 or higher)
- SQLite (default for development: `prisma/dev.db`) or PostgreSQL (production)
- npm or yarn

---

## Setup & Installation

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Configure environment variables:**
   Create `.env` in the `backend/` root directory:
   ```env
   NODE_ENV=development
   PORT=5000

   # Database
   DATABASE_URL="file:./dev.db"

   # JWT Configuration
   JWT_SECRET=your-jwt-access-secret-min-32-chars
   REFRESH_TOKEN_SECRET=your-jwt-refresh-secret-min-32-chars
   JWT_EXPIRES_IN=15m
   REFRESH_TOKEN_EXPIRES_IN=7d

   # AES-256 Key (64 hex characters or 32 characters)
   ENCRYPTION_KEY=0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef

   # Allowed CORS Origin
   FRONTEND_URL=http://localhost:3000
   ```

3. **Initialize the database:**
   ```bash
   npx prisma db push
   ```

4. **Start the development server:**
   ```bash
   npm run dev
   ```
   The backend API will start on `http://localhost:5000`.

---

## API Endpoints Overview

### Authentication (`/api/auth`)
- `POST /api/auth/register` - Create a new user account
- `POST /api/auth/login` - Authenticate with email/password
- `POST /api/auth/refresh` - Issue new access token using refresh token
- `POST /api/auth/logout` - Revoke session
- `GET /api/auth/me` - Get current authenticated user profile

### Credentials (`/api/credentials`)
- `GET /api/credentials` - List user's saved Gmail sender accounts
- `POST /api/credentials` - Verify SMTP connection and save encrypted credentials
- `DELETE /api/credentials/:id` - Delete saved credential

### Templates (`/api/templates`)
- `GET /api/templates` - List templates (custom + system starters)
- `POST /api/templates` - Create a new email template
- `PUT /api/templates/:id` - Update template name, subject, or body
- `DELETE /api/templates/:id` - Delete custom template

### Campaigns (`/api/campaigns`)
- `POST /api/campaigns/send` - Launch campaign with batching and smart jitter
- `POST /api/campaigns/validate-domains` - Verify MX DNS records for recipient domains
- `GET /api/campaigns/run/:runId/status` - Polling fallback for campaign progress
- `GET /api/campaigns` - List user campaigns (paginated)
- `GET /api/campaigns/:id` - Get specific campaign details
- `DELETE /api/campaigns/:id` - Delete campaign record

### Delivery Analytics (`/api/track`)
- `GET /api/track/stats` - Summary statistics (sent, failed, total)
- `GET /api/track/details` - Detailed per-recipient delivery logs
- `POST /api/track/reset` - Reset campaign statistics

---

## Verification & Testing

```bash
# Run TypeScript compilation
npm run build

# Run automated new features test suite
node scripts/test_new_features.js

# Run Jest unit test suite
npm test
```
