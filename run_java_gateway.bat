@echo off
setlocal enabledelayedexpansion
title FactoryMind AI - Java Industrial Edge Gateway

echo ========================================================
echo       FACTORYMIND AI - JAVA INDUSTRIAL IIoT GATEWAY
echo            High-Throughput PLC Polling Daemon
echo ========================================================
echo.

set JAVA_BIN=
set JAVAC_BIN=

where javac >nul 2>nul
if %errorlevel% equ 0 (
    set JAVAC_BIN=javac
    set JAVA_BIN=java
) else (
    for /d %%D in ("C:\Program Files\Java\jdk*" "C:\Program Files\Eclipse Adoptium\jdk*" "C:\Program Files\Microsoft\jdk*") do (
        if exist "%%~fD\bin\javac.exe" (
            set JAVAC_BIN="%%~fD\bin\javac.exe"
            set JAVA_BIN="%%~fD\bin\java.exe"
        )
    )
)

if "!JAVAC_BIN!"=="" (
    echo [NOTICE] Java Development Kit (JDK 17 or 21) was not detected in PATH.
    echo.
    echo In enterprise industry:
    echo   - Java is used for high-frequency PLC / SCADA Edge Gateways.
    echo   - Python is used for AI / ML Predictive Analytics.
    echo.
    echo Java source code is ready at:
    echo   java-edge-gateway\src\com\factorymind\gateway\IndustrialEdgeGateway.java
    echo.
    echo Options:
    echo  [1] Install OpenJDK via winget: winget install Microsoft.OpenJDK.21
    echo  [2] Or press any key to switch to the Python IoT Edge Streamer now!
    echo.
    set /p choice="Launch Python Edge Streamer instead? (Y/N): "
    if /i "!choice!"=="Y" (
        cd /d "%~dp0backend"
        set PY_BIN=py
        where py >nul 2>nul
        if !errorlevel! neq 0 (
            set PY_BIN="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
        )
        !PY_BIN% simulator\iot_simulator.py --interval 2.0
    )
    exit /b 0
)

cd /d "%~dp0java-edge-gateway"
if not exist "bin" mkdir bin

echo [1/2] Compiling Java Industrial Edge Gateway (pure Java, zero external deps)...
!JAVAC_BIN! -d bin src\com\factorymind\gateway\IndustrialEdgeGateway.java

if %errorlevel% neq 0 (
    echo [ERROR] Java compilation failed.
    pause
    exit /b 1
)

echo [2/2] Launching Java IIoT Edge Gateway Service...
echo Ingesting real-time machine telemetry into FastAPI at http://127.0.0.1:8000
echo.
!JAVA_BIN! -cp bin com.factorymind.gateway.IndustrialEdgeGateway

pause
