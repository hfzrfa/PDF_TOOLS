@echo off
echo ==========================================
echo   PDF Compressor - Setup & Launch
echo ==========================================
echo.

echo [1/2] Installing dependencies...
pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies.
    echo Make sure Python and pip are installed.
    pause
    exit /b 1
)

echo [2/2] Launching PDF Compressor...
echo.
python pdf_compressor.py
pause
