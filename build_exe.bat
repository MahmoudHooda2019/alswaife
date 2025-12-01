@echo off
REM Build standalone Windows executable for Al Sawife Factory app

REM Activate venv if it exists
IF EXIST ".venv\Scripts\activate.bat" (
    call ".venv\Scripts\activate.bat"
)

REM Install dependencies (safe to re-run)
IF EXIST "requirements.txt" (
    pip install -r requirements.txt
)

REM Build using PyInstaller (include res folder)
pyinstaller ^
  --noconfirm ^
  --onedir ^
  --windowed ^
  --name "AlSawifeFactory" ^
  --add-data "res;res" ^
  main.py

echo.
echo Build finished. Output folder:
echo   dist\AlSawifeFactory
echo.
pause


