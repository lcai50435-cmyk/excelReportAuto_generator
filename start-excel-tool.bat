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

echo Starting Excel tool...
echo Keep this window open while using smart fill.
echo.

start "Open Excel tool" cmd /c "timeout /t 2 /nobreak >nul && rundll32 url.dll,FileProtocolHandler http://127.0.0.1:4173/"

node server.js
pause
