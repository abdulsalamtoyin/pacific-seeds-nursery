@echo off
REM ============================================================
REM  Pacific Seeds — Windows build script
REM  Run this on a Windows machine with Python 3.11+ installed.
REM ============================================================

setlocal
cd /d "%~dp0\.."

echo.
echo === Pacific Seeds Nursery Fieldbook — Windows build ===
echo.

REM 1. Create venv if not present
if not exist ".venv-win\Scripts\python.exe" (
    echo Creating virtual environment...
    py -3 -m venv .venv-win
)

REM 2. Activate and install build deps
call .venv-win\Scripts\activate.bat

echo Installing runtime + build dependencies...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller pystray pillow

REM 3. Build with PyInstaller
echo.
echo Running PyInstaller...
pyinstaller dist_windows\PacificSeeds.spec --clean --noconfirm

REM 4. Report
if exist dist\PacificSeeds.exe (
    echo.
    echo  Build succeeded:
    echo    dist\PacificSeeds.exe
    echo.
    echo  Double-click that .exe to test. It opens your browser at the landing page.
) else (
    echo.
    echo  Build FAILED — see messages above.
    exit /b 1
)

endlocal
