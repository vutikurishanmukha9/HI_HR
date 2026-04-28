# HiHR — Smart HR Email Outreach

A modern, enterprise-grade platform for managing cold email outreach campaigns with personalization, scheduling, batch sending, and email tracking. Built with React, TypeScript, and Node.js.

---

## Overview

HiHR streamlines HR recruitment and outreach workflows by enabling professionals to send personalized bulk emails efficiently. The platform features a premium **Warm Slate** dark-themed interface with glassmorphism design, per-step gradient identities, and luminous accent colors.

## Key Features

### Email Campaign Management
- **Personalized Emails** — Dynamic placeholders (`{fullName}`, `{companyName}`, `{jobTitle}`) for tailored messaging
- **Rich Text Editor** — Compose professional emails with formatting, lists, and hyperlinks
- **Attachments** — Support for files up to 10MB per attachment
- **Scheduling** — Schedule campaigns for future delivery
- **Batch Sending** — Configurable batch size and delays to optimize deliverability
- **Copy to Clipboard** — One-click copy of composed email content

### Data Management
- **Bulk Upload** — Import recipients from Excel files (.xlsx, .xls)
- **Server-side Excel Parsing** — Backend upload endpoint with Multer + SheetJS
- **Duplicate Detection** — Automatic identification and removal of duplicate emails
- **Recipient Selection** — Choose specific recipients before sending

### Monitoring & Analytics
- **Real-time Progress** — Live status updates during campaign execution
- **Campaign Results** — Detailed sent/failed statistics after completion
- **Estimated Time** — Remaining time calculations for active campaigns
- **Dashboard Stats** — Total sent, success rate, today's count, pending emails
- **Email Preview** — See exactly what recipients will receive before sending
- **Confetti Celebration** — Animation on successful campaign completion

### Email Tracking
- **Open Tracking** — Invisible 1×1 pixel embedded in emails to track opens
- **Click Tracking** — Links rewritten to track clicks and redirect to original URL
- **Open Rate Analytics** — View percentage of recipients who opened emails
- **Click Analytics** — Track which links were clicked and how many times
- **Tracking Dashboard** — Frontend component to visualize tracking stats
- **Per-Recipient Details** — See exactly who opened and when

### Security
- **Encrypted Credentials** — AES-256-CBC encryption for stored SMTP credentials
- **JWT Authentication** — Secure user sessions with access and refresh tokens
- **Account Lockout** — 5 failed login attempts = 30-minute lockout
- **Password Strength Meter** — Real-time visual feedback during registration
- **Rate Limiting** — Per-route rate limits (auth, email sending, file uploads)
- **CORS / CSP Headers** — Configured via Helmet middleware
- **Error Boundary** — Graceful error handling for React components
- **Audit Logging** — Track all email actions for compliance (JSONL format)
- **Sentry Integration** — Optional error tracking for production

### User Experience
- **Warm Slate Design System** — Premium dark palette (`#0c0e1a`) with luminous accents
- **Per-Step Gradient Identity** — Indigo→Violet (Connect), Emerald→Teal (Upload), Rose→Orange (Compose), Cyan→Blue (Send)
- **Glassmorphism Cards** — 24px blur with subtle inner-glow borders
- **Typography** — Outfit for branding, Inter for body text
- **Toast Notifications** — Color-coded success / error / warning / info alerts
- **Skeleton Loaders** — Warm shimmer effects during loading states
- **Empty State Illustrations** — SVG graphics for empty lists
- **Step Animations** — Smooth transitions between workflow steps
- **Mobile Responsive** — Optimized for all device sizes

---

## Technology Stack

### Frontend
| Technology | Version | Purpose |
|------------|---------|---------|
| React | 19.x | UI Framework |
| TypeScript | 5.8 | Type Safety |
| Tailwind CSS | 3.4 | Utility-first Styling |
| Vite | 6.x | Build Tool & Dev Server |
| SheetJS (xlsx) | 0.18 | Client-side Excel Parsing |
| Zod | 3.x | Schema Validation |

### Backend
| Technology | Version | Purpose |
|------------|---------|---------|
| Node.js | 18+ | Runtime |
| Express | 4.x | API Framework |
| TypeScript | 5.8 | Type Safety |
| Prisma | 5.x | ORM |
| SQLite | — | Database (Development) |
| PostgreSQL | — | Database (Production) |
| Nodemailer | 6.x | SMTP Email Sending |
| Multer | 2.x | File Upload Handling |
| Winston | 3.x | Structured Logging |
| Helmet | 8.x | Security Headers |
| Sentry | 10.x | Error Tracking (optional) |
| bcryptjs | 2.x | Password Hashing |
| jsonwebtoken | 9.x | JWT Token Management |

### DevOps
| Technology | Purpose |
|------------|---------|
| Docker Compose | Local PostgreSQL + Redis |
| Render | Cloud Deployment |
| Railway | Cloud Deployment |
| Jest | Unit Testing |

---

## Architecture Highlights

### Performance Optimizations
- **Email Connection Pooling** — SMTP connections are cached and reused (5–10× faster)
- **API Retry with Exponential Backoff** — Auto-retry on network failures (3 retries, 1s / 2s / 4s delays)
- **Request Timeout** — 30s default, 5min for campaign sends
- **Batch Email Processing** — Configurable batch size and inter-batch delays

### Code Organization
- **Pages Pattern** — `AuthPage` and `DashboardPage` for clear separation
- **Custom Hooks** — `useCampaign` extracts complex campaign state logic
- **Service Layer** — Business logic separated from controllers
- **Config Module** — Centralized environment config with validation (`config/env.ts`)
- **Request Logging** — Unique request IDs for tracing

### API Client Features
- Automatic retry on 5xx errors
- Request / response timeout handling
- Token refresh on 401 errors
- Structured error extraction

---

## Getting Started

### Prerequisites
- Node.js v18 or higher
- npm or yarn

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/vutikurishanmukha9/HR_Cold_Email.git
   cd HR_Cold_Email
   ```

2. **Install frontend dependencies**
   ```bash
   cd frontend
   npm install
   ```

3. **Install backend dependencies**
   ```bash
   cd backend
   npm install
   ```

4. **Configure environment variables**
   
   Create `backend/.env`:
   ```env
   # Server Configuration
   NODE_ENV=development
   PORT=5000

   # Database (SQLite for development)
   DATABASE_URL="file:./dev.db"

   # JWT Secrets (use strong random strings in production)
   JWT_SECRET=your-jwt-secret-key-min-32-characters
   REFRESH_TOKEN_SECRET=your-refresh-token-secret-min-32-chars

   # Encryption Key (MUST be exactly 32 characters)
   ENCRYPTION_KEY=your-32-character-encryption-key

   # CORS
   FRONTEND_URL=http://localhost:3000
   ```

   Create `frontend/.env`:
   ```env
   VITE_API_URL=http://localhost:5000
   ```

5. **Initialize the database**
   ```bash
   cd backend
   npx prisma db push
   ```

6. **Start the development servers**
   
   Terminal 1 (Backend):
   ```bash
   cd backend
   npm run dev
   ```
   
   Terminal 2 (Frontend):
   ```bash
   cd frontend
   npm run dev
   ```

7. **Open the application**
   
   Navigate to `http://localhost:3000`

---

## Usage Guide

### Step 1: Authentication
- Register a new account or login with existing credentials
- User sessions are managed via JWT tokens (15min access / 7-day refresh)

### Step 2: Connect Gmail Account
- Enter your Gmail address
- Provide your 16-character Google App Password (no spaces)
- Credentials are encrypted with AES-256-CBC and stored securely

> For information on creating a Google App Password, visit the [Google Help Center](https://support.google.com/accounts/answer/185833)

### Step 3: Upload Recipients
- Prepare an Excel file with columns:
  - `Full name` or `Name`
  - `Email address` or `Email`
  - `Company` or `Organization`
  - `Job title` (optional)
- Drag and drop or select your file (max 5MB)
- Review the parsed recipients and duplicates are auto-removed

### Step 4: Compose Email
- Write your subject line (max 200 characters)
- Use the rich text editor for email body
- Insert personalization tags from the sidebar (`{fullName}`, `{companyName}`, `{jobTitle}`)
- Add attachments if needed (max 10MB each)

### Step 5: Review and Send
- Review campaign summary
- Select / deselect individual recipients
- Configure batch size and delay settings
- Send immediately or schedule for later
- Monitor real-time progress with live status indicators

---

## Project Structure

```
HR_Cold_Email/
├── README.md
├── LICENSE
├── .gitignore
├── docker-compose.yml              # Local PostgreSQL + Redis
├── render.yaml                     # Render deployment config
├── railway.toml                    # Railway deployment config
├── RENDER_DEPLOYMENT.md            # Render deployment guide
├── RAILWAY_DEPLOYMENT.md           # Railway deployment guide
│
├── frontend/                       # React Frontend
│   ├── App.tsx                     # Root component (routing)
│   ├── index.html                  # Entry HTML
│   ├── index.css                   # Design system (Warm Slate theme)
│   ├── index.tsx                   # React entry point
│   ├── types.ts                    # Shared TypeScript types
│   ├── vite.config.ts              # Vite configuration
│   ├── tailwind.config.js          # Tailwind CSS configuration
│   ├── postcss.config.js           # PostCSS pipeline
│   ├── package.json
│   ├── pages/
│   │   ├── AuthPage.tsx            # Login / Register page
│   │   └── DashboardPage.tsx       # Main campaign workflow
│   ├── hooks/
│   │   ├── useCampaign.ts          # Campaign state management
│   │   └── useScript.tsx           # Script loading hook
│   ├── components/
│   │   ├── StepIndicator.tsx       # Per-step gradient progress bar
│   │   ├── CredentialsForm.tsx     # Gmail SMTP credential input
│   │   ├── RecipientUploader.tsx   # Excel drag-and-drop upload
│   │   ├── EmailComposer.tsx       # Rich text editor + placeholders
│   │   ├── ReviewAndSend.tsx       # Final review + send / schedule
│   │   ├── EmailPreview.tsx        # Email preview panel
│   │   ├── DashboardStats.tsx      # Campaign statistics cards
│   │   ├── EmailTrackingStats.tsx  # Open / click tracking dashboard
│   │   ├── Confetti.tsx            # Celebration animation
│   │   ├── EmptyState.tsx          # Empty list illustrations
│   │   ├── Toast.tsx               # Toast notification system
│   │   ├── Skeleton.tsx            # Shimmer loading skeletons
│   │   ├── ErrorBoundary.tsx       # React error boundary
│   │   └── PasswordStrengthMeter.tsx
│   ├── contexts/
│   │   └── AuthContext.tsx         # Auth state provider
│   ├── services/
│   │   └── api.ts                  # API client with retry logic
│   └── public/
│       └── favicon.svg
│
└── backend/                        # Node.js Backend
    ├── src/
    │   ├── server.ts               # Express app setup
    │   ├── config/
    │   │   ├── env.ts              # Environment variable validation
    │   │   └── database.ts         # Prisma client setup
    │   ├── routes/
    │   │   ├── auth.routes.ts      # Authentication routes
    │   │   ├── credential.routes.ts # Credential CRUD routes
    │   │   ├── campaign.routes.ts  # Campaign CRUD + send routes
    │   │   └── tracking.routes.ts  # Open / click tracking routes
    │   ├── controllers/
    │   │   ├── auth.controller.ts
    │   │   ├── credential.controller.ts
    │   │   └── campaign.controller.ts
    │   ├── middleware/
    │   │   ├── auth.ts             # JWT authentication
    │   │   ├── rateLimit.ts        # Per-route rate limiting
    │   │   ├── security.ts         # CORS, Helmet headers
    │   │   ├── validation.ts       # Zod schema validation
    │   │   ├── errorHandler.ts     # Global error handling
    │   │   └── requestLogger.ts    # Request ID & logging
    │   ├── services/
    │   │   ├── auth.service.ts     # Login, register, lockout
    │   │   ├── email.service.ts    # SMTP connection pooling
    │   │   ├── campaign.service.ts # Campaign logic
    │   │   ├── credential.service.ts # Credential encryption
    │   │   ├── tracking.service.ts # Open / click tracking
    │   │   └── audit.service.ts    # JSONL audit logging
    │   ├── utils/
    │   │   ├── logger.ts           # Winston structured logging
    │   │   ├── encryption.ts       # AES-256-CBC encryption
    │   │   ├── jwt.ts              # Token management
    │   │   ├── validation.ts       # Zod schemas
    │   │   ├── security.ts         # Security utilities
    │   │   ├── excel.ts            # Server-side Excel parsing
    │   │   └── sentry.ts          # Sentry error tracking setup
    │   └── types/
    │       └── index.ts            # Backend TypeScript types
    ├── prisma/
    │   └── schema.prisma
    └── package.json
```

---

## API Endpoints

### Authentication
| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/auth/register` | Register new user |
| POST | `/api/auth/login` | User login |
| POST | `/api/auth/refresh` | Refresh access token |
| POST | `/api/auth/logout` | User logout |
| GET | `/api/auth/me` | Get current user |

### Credentials
| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/credentials` | Save email credentials |
| GET | `/api/credentials` | List user credentials |
| DELETE | `/api/credentials/:id` | Delete credential |

### Campaigns
| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/campaigns` | Create campaign |
| GET | `/api/campaigns` | List campaigns (paginated) |
| GET | `/api/campaigns/:id` | Get campaign by ID |
| PATCH | `/api/campaigns/:id` | Update campaign |
| DELETE | `/api/campaigns/:id` | Delete campaign |
| POST | `/api/campaigns/upload-recipients` | Upload Excel file |
| POST | `/api/campaigns/send` | Send email campaign |

### Email Tracking
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/track/open/:token` | Tracking pixel (returns 1×1 GIF) |
| GET | `/api/track/click/:token` | Click redirect (302 to original URL) |
| GET | `/api/track/stats` | Aggregate tracking stats (auth required) |
| GET | `/api/track/details` | Per-recipient tracking details (auth required) |

---

## Design System — Warm Slate + Luminous Accents

### Color Palette
| Token | Hex | Usage |
|-------|-----|-------|
| `--bg-primary` | `#0c0e1a` | Page background |
| `--bg-secondary` | `#111327` | Section backgrounds |
| `--bg-tertiary` | `#141625` | Card backgrounds |
| `--accent-indigo` | `#6366f1` | Step 1 — Connect |
| `--accent-teal` | `#14b8a6` | Step 2 — Upload |
| `--accent-rose` | `#f43f5e` | Step 3 — Compose |
| `--accent-cyan` | `#06b6d4` | Step 4 — Send |

### Typography
| Font | Weight | Usage |
|------|--------|-------|
| **Outfit** | 600–900 | Brand elements, headings, logo |
| **Inter** | 400–800 | Body text, labels, UI elements |

### Text Hierarchy
| Color | Hex | Role |
|-------|-----|------|
| Primary | `#f1f5f9` | Headings, important text |
| Secondary | `#94a3b8` | Body, descriptions |
| Muted | `#64748b` | Labels, captions |

---

## Troubleshooting

### Email Sending Fails
- Verify your Google App Password is correct (16 characters, no spaces)
- Ensure 2-Factor Authentication is enabled on your Google account
- Check that the recipient email addresses are valid
- Check backend console for detailed SMTP error messages

### Decryption Error ("bad decrypt")
- The `ENCRYPTION_KEY` in `.env` may have changed since credentials were saved
- Solution: Reset the database and re-add your credentials
  ```bash
  cd backend
  npx prisma db push --force-reset
  ```

### File Upload Issues
- Ensure file is `.xlsx` or `.xls` format
- Check file size is under 5MB
- Verify column names match expected patterns (Name, Email, Company)

### Authentication Issues
- Clear browser cookies and local storage
- Check that the backend server is running on port 5000
- Verify environment variables are set correctly

### Tailwind CSS Not Loading
- Ensure `@tailwind base; @tailwind components; @tailwind utilities;` directives exist at the top of `index.css`
- Restart the Vite dev server after config changes
- Run `npm run build` to verify the PostCSS pipeline works

---

## Security Considerations

- Credentials are encrypted using AES-256-CBC before storage
- Passwords are hashed using bcrypt with salt rounds
- JWT tokens expire after 15 minutes (access) / 7 days (refresh)
- **Account Lockout**: 5 failed login attempts triggers 30-minute lockout
- **Password Strength Meter**: Visual feedback encourages strong passwords
- Rate limiting prevents brute force attacks
- Input validation using Zod schemas on both frontend and backend
- Helmet middleware for security headers
- CORS configured for specific frontend origin

---

## Production Deployment

Deployment guides are provided for two platforms:

- **Render** — See [`RENDER_DEPLOYMENT.md`](RENDER_DEPLOYMENT.md)
- **Railway** — See [`RAILWAY_DEPLOYMENT.md`](RAILWAY_DEPLOYMENT.md)
- **Docker** — `docker-compose.yml` provided for local PostgreSQL + Redis

Key considerations:
- Use PostgreSQL instead of SQLite
- Set `NODE_ENV=production`
- Use strong, random secrets for JWT and encryption keys
- Configure proper CORS origins
- Enable HTTPS

---

## License

This project is private and for demonstration purposes.

---

## Author

**V Shanmukha**

---

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Open a Pull Request
