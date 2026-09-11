@echo off
REM Starts the local dev server for the Weekly Status Agent.
REM Real tenant values come from config.local.js (gitignored).
cd /d "%~dp0"
echo Starting local server on http://localhost:5500/
echo Press Ctrl+C to stop.
echo.
npx --yes http-server -p 5500 -c-1 --cors
