@echo off
setlocal
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js was not found.
  echo Please install Node.js and run this file again.
  pause
  exit /b 1
)

if not exist node_modules (
  echo Installing dependencies...
  npm install
  if errorlevel 1 (
    echo npm install failed.
    pause
    exit /b 1
  )
)

if not exist .next (
  echo Building Next.js app...
  npm run build
  if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
  )
)

echo Starting Excel tool...
echo Keep this window open while using smart fill.
echo.

start "Open Excel tool" cmd /c "timeout /t 3 /nobreak >nul && rundll32 url.dll,FileProtocolHandler http://127.0.0.1:4173/"

set PORT=4173
npm start
pause
