@echo off
setlocal
title EDOS Weekly Review Cloud - shared link
cd /d "%~dp0"

echo ============================================================
echo   EDOS Weekly Review Cloud
echo   Starting with a PUBLIC LINK.
echo ============================================================
echo.

REM ---- Node.js, same requirement as START.bat --------------------------------
where node >nul 2>nul
if errorlevel 1 (
  echo   [X] Node.js is not installed.
  echo.
  echo       Install the LTS version from  https://nodejs.org
  echo       then double-click SHARE.bat again.
  echo.
  pause
  exit /b 1
)

REM ---- a password must be set, or run.js will refuse anyway -------------------
findstr /C:"\"password\": \"\"" sharing.json >nul 2>nul
if not errorlevel 1 (
  echo   [X] No password is set.
  echo.
  echo       Open  sharing.json  in Notepad, put a password between the
  echo       quotes on the "password" line, save it, and run SHARE.bat again.
  echo.
  echo       Example:   "password": "weekly-review-2026",
  echo.
  pause
  exit /b 1
)

REM ---- cloudflared: one file, fetched once ------------------------------------
REM This is Cloudflare's own signed release. It dials out from this laptop; it does
REM not open a port and does not need administrator rights.
if not exist "cloudflared.exe" (
  echo   cloudflared.exe is not here yet - downloading it once ^(about 60 MB^).
  echo   This needs internet access and takes a minute.
  echo.
  powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop';" ^
    "try { [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;" ^
    "  Invoke-WebRequest -UseBasicParsing -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile 'cloudflared.exe';" ^
    "  exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
  if errorlevel 1 (
    echo.
    echo   [X] The download failed - a proxy or firewall usually blocks it.
    echo.
    echo       Download this file by hand, rename it to cloudflared.exe,
    echo       and put it in this folder:
    echo.
    echo       https://github.com/cloudflare/cloudflared/releases/latest
    echo       ^(the file called cloudflared-windows-amd64.exe^)
    echo.
    pause
    exit /b 1
  )
  echo   Downloaded.
  echo.
)

REM ---- everything the package needs travels with it ---------------------------
if not exist "backend\dist\server.js" goto broken
if not exist "backend\node_modules" goto broken
if not exist "web\index.html" goto broken
if not exist "edos_config.json" goto broken

node run.js --share

echo.
echo   Stopped. The public link is now dead.
pause
exit /b 0

:broken
echo   [X] This folder is incomplete.
echo.
echo       Extract the ZIP again and keep all of its folders together,
echo       then double-click SHARE.bat from inside the extracted folder.
echo.
pause
exit /b 1
