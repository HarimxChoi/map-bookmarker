@echo off

echo.
echo  ========================================
echo   Map Favorite Registrar - INSTALL
echo  ========================================
echo.

:: 1. Python check
echo [1/3] Checking Python...
python --version > nul 2>&1
if errorlevel 1 goto NO_PYTHON
python --version
echo  OK: Python found
goto STEP2

:NO_PYTHON
echo.
echo  [!] Python is NOT installed.
echo.
echo  Download: https://www.python.org/downloads/
echo  IMPORTANT: Check "Add Python to PATH" when installing!
echo.
set /p open_url=Open download page? (Y/N):
if /i "%open_url%"=="Y" start https://www.python.org/downloads/
echo.
echo  After installing Python, run this file again.
pause
exit /b 1

:STEP2
:: 2. pip packages
echo.
echo [2/3] Installing packages...
python -m pip install --upgrade pip -q
python -m pip install -r requirements.txt -q
if errorlevel 1 goto PKG_FAIL
echo  OK: Packages installed
goto STEP3

:PKG_FAIL
echo  [!] Package install failed. Check internet connection.
pause
exit /b 1

:STEP3
:: 3. Playwright Chromium
echo.
echo [3/3] Installing Chromium browser... (1-2 min)
python -m playwright install chromium
if errorlevel 1 goto CHROME_FAIL
echo  OK: Chromium installed
goto DONE

:CHROME_FAIL
echo  [!] Chromium install failed. Check internet connection.
pause
exit /b 1

:DONE
echo.
echo  ========================================
echo   INSTALL COMPLETE!
echo.
echo   Next: Double-click run.bat
echo  ========================================
echo.
pause
