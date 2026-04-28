# Render Deployment Guide for HiHR

This guide covers deploying HiHR to [Render](https://render.com).

---

## Option 1: Blueprint Deployment (Recommended)

The easiest way - deploy everything with one click using the `render.yaml` file.

### Steps:
1. Go to [Render Dashboard](https://dashboard.render.com)
2. Click **"New"** → **"Blueprint"**
3. Connect your GitHub repo
4. Render will detect `render.yaml` and show services to create
5. Click **"Apply"** to deploy all services

### After Deployment:
1. Go to your **backend service** → **Environment**
2. Set `ENCRYPTION_KEY` (exactly 32 characters)
3. Set `FRONTEND_URL` to your frontend's Render URL
4. Go to your **frontend service** → **Environment**
5. Set `VITE_API_URL` to `https://your-backend.onrender.com/api`

---

## Option 2: Manual Deployment

### Step 1: Create PostgreSQL Database

1. Dashboard → **"New"** → **"PostgreSQL"**
2. Name: `hihr-db`
3. Region: Choose closest to your users
4. Plan: Free (90 days) or Starter ($7/mo)
5. Click **"Create Database"**
6. Copy the **Internal Database URL** for later

---

### Step 2: Deploy Backend

1. Dashboard → **"New"** → **"Web Service"**
2. Connect your GitHub repo
3. Configure:

| Setting | Value |
|---------|-------|
| Name | `hihr-backend` |
| Region | Same as database |
| Root Directory | `backend` |
| Runtime | Node |
| Build Command | `npm install && npx prisma generate && npm run build` |
| Start Command | `npx prisma migrate deploy && npm run start` |

4. Add Environment Variables:

| Variable | Value |
|----------|-------|
| `NODE_ENV` | `production` |
| `PORT` | `5000` |
| `DATABASE_URL` | *Paste Internal Database URL* |
| `JWT_SECRET` | *Generate 32+ char random string* |
| `REFRESH_TOKEN_SECRET` | *Generate 32+ char random string* |
| `ENCRYPTION_KEY` | *Exactly 32 characters* |
| `FRONTEND_URL` | `https://your-frontend.onrender.com` |

5. Click **"Create Web Service"**

---

### Step 3: Deploy Frontend

1. Dashboard → **"New"** → **"Static Site"**
2. Connect your GitHub repo
3. Configure:

| Setting | Value |
|---------|-------|
| Name | `hihr-frontend` |
| Root Directory | `.` (leave empty or root) |
| Build Command | `npm install && npm run build` |
| Publish Directory | `dist` |

4. Add Environment Variables:

| Variable | Value |
|----------|-------|
| `VITE_API_URL` | `https://your-backend.onrender.com/api` |

5. Add Rewrite Rule (for SPA routing):
   - Source: `/*`
   - Destination: `/index.html`
   - Action: Rewrite

6. Click **"Create Static Site"**

---

## Environment Variables Reference

### Backend
| Variable | Required | Description |
|----------|----------|-------------|
| `NODE_ENV` | Yes | `production` |
| `PORT` | Yes | `5000` |
| `DATABASE_URL` | Yes | PostgreSQL connection string |
| `JWT_SECRET` | Yes | Random 32+ characters |
| `REFRESH_TOKEN_SECRET` | Yes | Random 32+ characters |
| `ENCRYPTION_KEY` | Yes | **Exactly 32 characters** |
| `FRONTEND_URL` | Yes | Frontend URL for CORS |

### Frontend
| Variable | Required | Description |
|----------|----------|-------------|
| `VITE_API_URL` | Yes | Backend URL + `/api` |

---

## Troubleshooting

### Build Fails
- Check Render build logs
- Ensure `rootDir` is correct (`backend` for backend)
- Verify Node.js version compatibility

### Database Connection Fails
- Use **Internal Database URL** (not External)
- Ensure database is in same region as backend

### CORS Errors
- Verify `FRONTEND_URL` exactly matches frontend URL (including `https://`)

### Frontend Routing Issues
- Add rewrite rule: `/* → /index.html`

### Cold Starts (Free Tier)
- Free tier services sleep after 15 minutes of inactivity
- First request after sleep takes ~30 seconds
- Upgrade to Starter ($7/mo) for always-on

---

## Cost Estimate (Render)

| Service | Free Tier | Paid |
|---------|-----------|------|
| Backend | 750 hrs/month, sleeps | $7/mo (Starter) |
| Frontend | Unlimited | $0 (Static is free) |
| PostgreSQL | 90 days free | $7/mo (Starter) |

**Total Paid**: ~$14/month for always-on

---

## Generate Secrets

Use [RandomKeygen.com](https://randomkeygen.com):
- **JWT_SECRET**: Fort Knox Passwords section
- **REFRESH_TOKEN_SECRET**: Fort Knox Passwords section
- **ENCRYPTION_KEY**: CodeIgniter Keys section (exactly 32 chars)
