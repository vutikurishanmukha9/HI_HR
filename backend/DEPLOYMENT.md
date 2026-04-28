# Backend Deployment Configuration for Render/Railway

# Build and start commands for deployment platforms

# For Render.com:
# - Build Command: npm install && npx prisma generate && npm run build
# - Start Command: npm start

# For Railway:
# - Build Command: npm install && npx prisma generate && npm run build
# - Start Command: npm start

# Environment Variables Required:
# NODE_ENV=production
# PORT=5000
# DATABASE_URL=postgresql://user:password@host:5432/database
# JWT_SECRET=your-secure-jwt-secret-min-32-chars
# REFRESH_TOKEN_SECRET=your-secure-refresh-secret-min-32-chars
# ENCRYPTION_KEY=your-32-character-encryption-key!
# FRONTEND_URL=https://your-frontend-domain.vercel.app
# BACKEND_URL=https://your-backend-domain.onrender.com  # For email tracking pixels
# SENTRY_DSN=optional-sentry-dsn-for-error-tracking
