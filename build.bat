@echo off
REM ============================================================
REM  Unimex FTZ Processor - Build Script
REM  Run this once on your Windows machine to produce the .exe.
REM  Requires Python 3.10+ already installed.
REM ============================================================

REM --- Pick a Python interpreter -----------------------------
REM  Prefer the Windows launcher "py -3": it resolves the real
REM  installed Python from the registry and ignores whatever
REM  happens to be first on PATH (e.g. an unrelated venv). Fall
REM  back to plain "python" only if the launcher isn't present.
where py >nul 2>nul && (set "PY=py -3") || (set "PY=python")

echo.
echo === Using interpreter ===
%PY% --version
if errorlevel 1 (
    echo.
    echo [ERROR] No usable Python found. Install Python 3.10+ from
    echo         python.org and make sure "py" or "python" works in a
    echo         new terminal, then run this again.
    pause
    exit /b 1
)

echo.
echo === Installing required packages ===
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERROR] pip install failed. Make sure Python is installed and on your PATH.
    pause
    exit /b 1
)

echo.
echo === Syncing ftz_processor.py to docs\ for web version ===
copy /Y ftz_processor.py docs\ftz_processor.py >nul
if errorlevel 1 (
    echo [WARN] could not sync docs\ftz_processor.py -- web version may be stale
)

echo.
echo === Building OCEAN executable with PyInstaller ===
REM --onefile    : produce a single .exe instead of a folder
REM --console    : keep the console window (we want users to see output)
REM --name       : sets the output filename (UnimexFTZ.exe)
REM --clean      : wipe PyInstaller caches first to avoid stale builds
%PY% -m PyInstaller --onefile --console --clean --name UnimexFTZ ftz_processor.py
if errorlevel 1 (
    echo.
    echo [ERROR] PyInstaller failed building UnimexFTZ. See messages above.
    pause
    exit /b 1
)

echo.
echo === Building AIR executable with PyInstaller ===
REM Separate program for AIR shipments (###-######## MWB or a carrier
REM booking reference like ZIMUSHH..., 4-letter bag prefixes).
%PY% -m PyInstaller --onefile --console --clean --name UnimexAir air_processor.py
if errorlevel 1 (
    echo.
    echo [ERROR] PyInstaller failed building UnimexAir. See messages above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Build complete!
echo.
echo  The .exe files are at:
echo    dist\UnimexFTZ.exe   (ocean)
echo    dist\UnimexAir.exe   (air)
echo.
echo  To distribute: copy the relevant .exe + README.txt to
echo  wherever the customs team will run it.
echo ============================================================
echo.
pause
