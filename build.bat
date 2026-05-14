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
echo === Building executable with PyInstaller ===
REM --onefile    : produce a single .exe instead of a folder
REM --console    : keep the console window (we want users to see output)
REM --name       : sets the output filename (UnimexFTZ.exe)
REM --clean      : wipe PyInstaller caches first to avoid stale builds
python -m PyInstaller --onefile --console --clean --name UnimexFTZ ftz_processor.py
if errorlevel 1 (
    echo.
    echo [ERROR] PyInstaller failed. See messages above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Build complete!
echo.
echo  The .exe is at:  dist\UnimexFTZ.exe
echo.
echo  To distribute: copy UnimexFTZ.exe + README.txt to wherever
echo  the customs team will run it.
echo ============================================================
echo.
pause
