@echo off
title AutoMailer PRO - Bulk Email Automation
echo ========================================================
echo       AutoMailer PRO - Bulk Email Automation Engine
echo ========================================================
echo.
echo [1/2] Checking dependencies...
if not exist node_modules (
    echo Installing node dependencies...
    call npm install
)
echo.
echo [2/2] Starting local server...
start http://localhost:3000
node server.js
pause
