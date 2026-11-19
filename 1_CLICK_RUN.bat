@echo off
setlocal
title FactoryMind AI - Instant Platform Launcher

echo ===============================================================================
echo        FACTORYMIND AI - PREDICTIVE MANUFACTURING PLATFORM
echo          SY MCA Major Project - TKIET Warananagar (2026-2027)
echo ===============================================================================
echo.
echo [*] Starting FactoryMind AI Local Web Console on http://127.0.0.1:8000 ...
echo [*] No external installations required (Uses Windows Built-In Web Server).
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0backend\server.ps1" -Port 8000

if %errorlevel% neq 0 (
    echo.
    echo [!] Server port busy or restricted. Opening directly in web browser...
    start "" "%~dp0backend\static\login.html"
)

pause
