@echo off
title EDOS Weekly Review Cloud
cd /d "%~dp0"

echo ============================================================
echo   EDOS Weekly Review Cloud
echo   Starting up - this takes a few seconds.
echo ============================================================
echo.

REM Node.js is the only thing that has to be installed.
where node >nul 2>nul
if errorlevel 1 (
  echo   [X] Node.js is not installed.
  echo.
  echo       Install the LTS version from  https://nodejs.org
  echo       then double-click START.bat again.
  echo.
  pause
  exit /b 1
)

for /f "delims=" %%v in ('node --version') do set NODEVER=%%v
echo   Node.js %NODEVER% found.
echo.

REM Everything the package needs travels with it. No install step, no internet.
if not exist "backend\dist\server.js" goto broken
if not exist "backend\node_modules" goto broken
if not exist "web\index.html" goto broken
if not exist "edos_config.json" goto broken

node run.js

echo.
echo   The application has stopped.
pause
exit /b 0

:broken
echo   [X] This folder is incomplete.
echo.
echo       Extract the ZIP again and keep all of its folders together,
echo       then double-click START.bat from inside the extracted folder.
echo.
pause
exit /b 1
