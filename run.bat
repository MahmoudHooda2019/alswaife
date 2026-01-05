@echo off
echo ========================================
echo Starting AlSawifeFactory Application
echo ========================================
echo.

echo Checking Python installation...
python --version
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH!
    pause
    exit /b 1
)

echo.
echo Checking and installing required libraries...
echo.

python -m pip install -r requirements.txt --quiet

if %errorlevel% neq 0 (
    echo ERROR: Failed to install required libraries!
    pause
    exit /b 1
)

echo Libraries are ready!
echo.

if not exist src\main.py (
    echo ERROR: src\main.py not found!
    pause
    exit /b 1
)

echo Starting application...
echo.

python src\main.py

if %errorlevel% neq 0 (
    echo.
    echo Application exited with error!
    pause
    exit /b 1
)

echo.
echo Application finished.
pause
