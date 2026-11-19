@echo off
setlocal enabledelayedexpansion
title FactoryMind AI - Instant Free Public HTTPS Domain

echo ===============================================================================
echo       FACTORYMIND AI - INSTANT FREE PUBLIC HTTPS TUNNEL
echo ===============================================================================
echo.
echo Choose tunnel provider:
echo  [1] Cloudflare Quick Tunnel (Zero setup, high-speed worldwide HTTPS, recommended)
echo  [2] Localhost.run (Built-in Windows SSH, zero downloads)
echo  [3] Pinggy (Fast Webhook & HTTPS tunnel)
echo.
set /p opt="Select provider [1-3] (Default: 1): "
if "%opt%"=="" set opt=1

if "%opt%"=="1" goto CLOUDFLARE
if "%opt%"=="2" goto LOCALHOST_RUN
if "%opt%"=="3" goto PINGGY

:CLOUDFLARE
echo.
echo [1/2] Checking cloudflared binary...
where cloudflared >nul 2>nul
if %errorlevel% neq 0 (
    echo Cloudflared CLI not found in PATH.
    echo Downloading portable official Cloudflare Tunnel client...
    powershell -Command "Invoke-WebRequest -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile '%TEMP%\cloudflared.exe'"
    set CF_BIN="%TEMP%\cloudflared.exe"
) else (
    set CF_BIN=cloudflared
)

echo [2/2] Exposing local port 8000 to public HTTPS domain...
echo.
echo Look for your public URL below: https://...trycloudflare.com
echo.
!CF_BIN! tunnel --url http://127.0.0.1:8000
pause
exit /b 0

:LOCALHOST_RUN
echo.
echo Connecting to localhost.run via built-in Windows OpenSSH...
ssh -R 80:127.0.0.1:8000 nokey@localhost.run
pause
exit /b 0

:PINGGY
echo.
echo Connecting to Pinggy tunnel...
ssh -p 443 -R0:localhost:8000 a.pinggy.io
pause
exit /b 0
