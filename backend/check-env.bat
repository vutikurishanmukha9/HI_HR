@echo off
echo Checking backend .env file...
echo.

if not exist .env (
    echo .env file not found! Creating it now...
    copy .env.example .env
    echo.
    echo IMPORTANT: Please edit backend\.env and set:
    echo - DATABASE_URL=file:./dev.db
    echo - ENCRYPTION_KEY must be exactly 32 characters
    echo.
    pause
    exit /b 1
)

echo .env file exists!
echo.
echo Checking if database exists...
if exist dev.db (
    echo Database exists!
) else (
    echo Database not found. Running migrations...
    call npm run prisma:migrate
)

echo.
echo Restarting backend server...
echo Press Ctrl+C to stop the current server, then run: npm run dev
pause
