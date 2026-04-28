@echo off
echo Creating backend .env file...
echo.

(
echo # Server Configuration
echo NODE_ENV=development
echo PORT=5000
echo API_URL=http://localhost:5000
echo.
echo # Database - SQLite ^(no installation needed!^)
echo DATABASE_URL="file:./dev.db"
echo.
echo # JWT Secrets
echo JWT_SECRET=streammail-jwt-secret-dev-2024-change-in-production
echo JWT_EXPIRES_IN=15m
echo REFRESH_TOKEN_SECRET=streammail-refresh-token-dev-2024-change-in-production
echo REFRESH_TOKEN_EXPIRES_IN=7d
echo.
echo # Encryption ^(32 characters for AES-256^)
echo ENCRYPTION_KEY=streammail-encryption-key-32ch
echo.
echo # CORS
echo FRONTEND_URL=http://localhost:3000
echo.
echo # Email Configuration ^(for system notifications - optional^)
echo SMTP_HOST=smtp.gmail.com
echo SMTP_PORT=587
echo SMTP_SECURE=false
echo SMTP_USER=your-email@gmail.com
echo SMTP_PASS=your-app-password
echo.
echo # File Upload
echo MAX_FILE_SIZE=10485760
echo UPLOAD_DIR=./uploads
echo.
echo # Rate Limiting
echo RATE_LIMIT_WINDOW_MS=900000
echo RATE_LIMIT_MAX_REQUESTS=100
) > .env

echo .env file created successfully!
echo.
echo Now running database migrations...
call npm run prisma:generate
call npm run prisma:migrate

echo.
echo Setup complete! You can now start the backend with: npm run dev
pause
