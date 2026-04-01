@echo off
chcp 65001 >nul
:: ============================================================
::  DOU Trend Update Workstation - Startup Script
::  Run this on the host machine to start the web server.
::  Team members then open: http://<this-machine-ip>:5000
:: ============================================================

setlocal
set PORT=5000
set APP_DIR=%~dp0

echo.
echo  ============================================================
echo   DOU Trend Update Workstation
echo  ============================================================
echo.

:: -- Check Python --------------------------------------------
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found. Please install Python 3.10+
    pause & exit /b 1
)

:: -- Install / upgrade dependencies --------------------------
echo  [1/2] Installing dependencies...
pip install -q -r "%APP_DIR%requirements.txt"
if errorlevel 1 (
    echo  [ERROR] pip install failed.
    pause & exit /b 1
)

:: -- Print access URL ----------------------------------------
echo.
echo  [2/2] Starting server on port %PORT%...
echo.

for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    set IP=%%a
    goto :found_ip
)
:found_ip
set IP=%IP: =%

echo  +--------------------------------------------------+
echo  ^|  Local  :  http://localhost:%PORT%                    ^|
echo  ^|  Network:  http://%IP%:%PORT%               ^|
echo  ^|                                                  ^|
echo  ^|  Share the Network URL with your team members.   ^|
echo  +--------------------------------------------------+
echo.
echo  Press Ctrl+C to stop the server.
echo.

:: -- Launch with waitress (production) or flask (fallback) ---
python -c "import waitress" >nul 2>&1
if errorlevel 1 (
    echo  [WARN] waitress not found, using Flask dev server...
    python "%APP_DIR%app.py"
) else (
    python "%APP_DIR%app.py"
)

pause