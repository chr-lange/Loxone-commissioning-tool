@echo off
title Build LoxoneCommissioning.exe
echo ============================================================
echo  Building Loxone Commissioning Tool -- Windows EXE
echo ============================================================
echo.

:: ── 1. Make sure PyInstaller is installed ────────────────────
echo [1/3] Checking / installing PyInstaller ...
py -m pip install --quiet pyinstaller
if errorlevel 1 (
    echo ERROR: Could not install PyInstaller.
    echo Make sure Python is installed and try again.
    pause
    exit /b 1
)

:: ── 2. Make sure runtime dependencies are installed ──────────
echo [2/3] Checking runtime dependencies ...
py -m pip install --quiet requests reportlab pillow
if errorlevel 1 (
    echo ERROR: Could not install dependencies.
    pause
    exit /b 1
)

:: ── 3. Run PyInstaller ────────────────────────────────────────
echo [3/3] Running PyInstaller (this may take 1-2 minutes) ...
cd /d "%~dp0"
py -m PyInstaller LoxoneCommissioning.spec --clean --noconfirm

if errorlevel 1 (
    echo.
    echo BUILD FAILED. See errors above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  SUCCESS!
echo  EXE is in:  %~dp0dist\LoxoneCommissioning.exe
echo ============================================================
echo.
pause
