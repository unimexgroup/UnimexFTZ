@echo off
REM ============================================================
REM  Unimex FTZ Processor - Build Script
REM  Run this once on your Windows machine to produce the .exe.
REM  Requires Python 3.10+ already installed.
REM ============================================================

echo.
echo === Installing required packages ===
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
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
python -m PyInstaller --onefile --console --clean --name UnimexFTZ ftz_processor.py
if errorlevel 1 (
    echo.
    echo [ERROR] PyInstaller failed building UnimexFTZ. See messages above.
    pause
    exit /b 1
)

echo.
echo === Building AIR executable with PyInstaller ===
REM Separate program for AIR shipments (###-######## MWB, 4-letter bag prefixes).
python -m PyInstaller --onefile --console --clean --name UnimexAir air_processor.py
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
