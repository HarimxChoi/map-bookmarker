@echo off
chcp 949 > nul 2>&1

echo.
echo  ========================================
echo   Map Favorite Registrar
echo  ========================================
echo.

:: Python check
python --version > nul 2>&1
if errorlevel 1 (
    echo  [!] Python not found. Run install.bat first.
    pause
    exit /b 1
)

:: Mode select
echo  Select mode:
echo.
echo    1. GUI (Recommended)
echo    2. CLI - All (Kakao + Naver)
echo    3. CLI - Kakao only
echo    4. CLI - Naver only
echo    5. CLI - Dry run (preview only)
echo.
set /p choice=Enter number (default 1):

if "%choice%"=="2" (
    python src/main.py
) else if "%choice%"=="3" (
    python src/main.py --kakao-only
) else if "%choice%"=="4" (
    python src/main.py --naver-only
) else if "%choice%"=="5" (
    python src/main.py --dry-run
) else (
    python run_gui.py
)

echo.
pause
