@echo off
title DevShop Report Studio - Build

:: ── Always run from the folder this script lives in ──────────────────────────
cd /d "%~dp0"

:: ── Disable code signing — prevents the symlink/permissions error ─────────────
set CSC_IDENTITY_AUTO_DISCOVERY=false
set WIN_CSC_LINK=
set WIN_CSC_KEY_PASSWORD=

echo.
echo ============================================================
echo   DevShop Report Studio v2.0 - Build Script
echo ============================================================
echo.
echo Working directory: %CD%
echo.

:: Check Node.js
node --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
  echo [ERROR] Node.js not found. Install from https://nodejs.org
  pause
  exit /b 1
)

echo [1/4] Installing dependencies...
call npm install
if %ERRORLEVEL% neq 0 (
  echo [ERROR] npm install failed.
  pause
  exit /b 1
)

echo.
echo [2/4] Installing Electron and electron-builder...
call npm install --save-dev electron electron-builder
if %ERRORLEVEL% neq 0 (
  echo [ERROR] Failed to install dev dependencies.
  pause
  exit /b 1
)

echo.
echo [3/4] Creating assets folder if needed...
if not exist "assets" mkdir assets

:: Clear any corrupted winCodeSign cache from previous failed attempts
echo     Clearing cached build artifacts...
set CACHE_DIR=%LOCALAPPDATA%\electron-builder\Cache\winCodeSign
if exist "%CACHE_DIR%" (
  rmdir /s /q "%CACHE_DIR%"
  echo     Cache cleared.
)

echo.
echo [4/4] Building portable .exe (no install required, no admin needed)...
call npm run build
if %ERRORLEVEL% neq 0 (
  echo.
  echo [ERROR] Build failed. Check output above.
  echo.
  echo Common fixes:
  echo   - Delete the dist\ folder and try again
  echo   - Delete node_modules\ folder and run BUILD.bat again from scratch
  echo.
  pause
  exit /b 1
)

echo.
echo ============================================================
echo   BUILD COMPLETE
echo.
echo   Output: dist\DevShop_Report_Studio.exe
echo.
echo   Share that single .exe with your team.
echo   They just double-click it — no install, no admin required.
echo ============================================================
echo.
pause
