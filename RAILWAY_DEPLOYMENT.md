# Railway Deployment Guide for HiHR

This guide covers deploying HiHR to [Railway](https://railway.app).

---

## Prerequisites

- A [Railway](https://railway.app) account (free tier available)
- Git repository pushed to GitHub

---

## Deployment Steps

### Step 1: Create a New Project

1. Go to [Railway Dashboard](https://railway.app/dashboard)
2. Click **"New Project"**
3. Select **"Deploy from GitHub repo"**
4. Connect your GitHub account and select the `HR_Cold_Email` repository

---

### Step 2: Add PostgreSQL Database

1. In your project, click **"+ New"** → **"Database"** → **"PostgreSQL"**
2. Railway automatically provisions the database
3. The `DATABASE_URL` environment variable is auto-injected

---

### Step 3: Deploy the Backend

1. Click **"+ New"** → **"GitHub Repo"** → Select your repo again
2. Railway will detect it as a Node.js app
3. Go to **Settings** → **Root Directory**: Set to `backend`
4. Railway will use the `railway.toml` configuration automatically

#### Required Environment Variables

Go to **Variables** tab and add:

```env
NODE_ENV=production
PORT=5000
JWT_SECRET=your-super-secure-jwt-secret-min-32-chars
REFRESH_TOKEN_SECRET=your-super-secure-refresh-secret-32-chars
ENCRYPTION_KEY=exactly-32-characters-for-aes!!
FRONTEND_URL=https://your-frontend.up.railway.app
```

> **Important**: Generate strong random strings for secrets. Use `openssl rand -base64 32`

---

### Step 4: Deploy the Frontend

1. Click **"+ New"** → **"GitHub Repo"** → Select your repo
2. Keep Root Directory as `/` (root)
3. Railway will use the frontend `railway.toml`

#### Required Environment Variables

```env
VITE_API_URL=https://your-backend.up.railway.app/api
```

---

### Step 5: Configure Networking

#### Backend Service:
1. Go to **Settings** → **Networking**
2. Click **"Generate Domain"** to get a public URL
3. Note this URL (e.g., `https://hihr-backend.up.railway.app`)

#### Frontend Service:
1. Go to **Settings** → **Networking**  
2. Click **"Generate Domain"**
3. Note this URL for your users

---

### Step 6: Update Environment Variables

1. Update Backend's `FRONTEND_URL` with the frontend's Railway URL
2. Update Frontend's `VITE_API_URL` with the backend's Railway URL + `/api`

---

## Troubleshooting

### Database Migration Fails
Check the deploy logs. Common issues:
- `DATABASE_URL` not linked from PostgreSQL service
- Solution: In backend service → **Variables** → Click **"Add Reference"** → Select PostgreSQL → `DATABASE_URL`

### CORS Errors
- Ensure `FRONTEND_URL` in backend exactly matches the frontend URL (including `https://`)

### Build Fails
Check logs for:
- Missing dependencies
- TypeScript errors
- Prisma generation issues

### Email Sending Issues
- Ensure Google App Password is correct
- Gmail requires 2FA enabled for App Passwords

---

## Environment Variables Reference

### Backend
| Variable | Description | Example |
|----------|-------------|---------|
| `DATABASE_URL` | PostgreSQL URL (auto-injected) | `postgresql://...` |
| `NODE_ENV` | Environment mode | `production` |
| `PORT` | Server port (auto-set by Railway) | `5000` |
| `JWT_SECRET` | JWT signing key (32+ chars) | Random string |
| `REFRESH_TOKEN_SECRET` | Refresh token key (32+ chars) | Random string |
| `ENCRYPTION_KEY` | AES key (exactly 32 chars) | Random string |
| `FRONTEND_URL` | Frontend origin for CORS | `https://...` |

### Frontend
| Variable | Description | Example |
|----------|-------------|---------|
| `VITE_API_URL` | Backend API URL | `https://backend.up.railway.app/api` |

---

## Local Development After Deployment

To switch back to SQLite for local development:

1. Change `backend/prisma/schema.prisma`:
   ```prisma
   provider = "sqlite"
   ```

2. Update `backend/.env`:
   ```env
   DATABASE_URL="file:./dev.db"
   ```

3. Run:
   ```bash
   cd backend
   npx prisma db push
   ```

---

## Cost Estimate

Railway's pricing (as of 2024):
- **Hobby Plan**: $5/month credit (usually enough for small apps)
- **PostgreSQL**: ~$0.000231/hour (~$0.17/month idle)
- **Backend**: ~$0.000231/hour when active
- **Frontend**: ~$0.000231/hour when active

For a small app with moderate usage, expect **~$2-5/month**.

---

## Security Checklist

- [ ] Environment variables are set (not in code)
- [ ] Strong random secrets (32+ characters)
- [ ] HTTPS enabled (automatic on Railway)
- [ ] FRONTEND_URL matches exactly
- [ ] Rate limiting enabled (already configured)
- [ ] CORS configured correctly

