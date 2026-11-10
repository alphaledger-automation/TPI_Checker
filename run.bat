@echo off
title Taxpayer Profile Inquiry
cd /d "%~dp0"

call :find_python
if defined PY goto :launch

:: -- Python missing: try to install it automatically (Windows 10/11) --
echo Python is not installed. Trying to install it automatically...
winget --version >nul 2>&1
if errorlevel 1 goto :manual
winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
:: PATH isn't refreshed in this window, so look in the default install folder
if exist "%LocalAppData%\Programs\Python\Python312\python.exe" (
    set "PY="%LocalAppData%\Programs\Python\Python312\python.exe""
    goto :launch
)
call :find_python
if defined PY goto :launch

:manual
echo.
echo Could not install Python automatically.
echo Please install Python 3.10+ from https://www.python.org/downloads/
echo and tick "Add Python to PATH" during setup, then run this file again.
start "" "https://www.python.org/downloads/"
pause
exit /b 1

:launch
:: main.py installs any missing packages, the Chromium browser and Tesseract-OCR itself.
%PY% main.py
if errorlevel 1 (
    echo.
    echo The application exited with an error.
    echo Setup log: %LocalAppData%\TPI_Checker\setup.log
    echo.
    pause
)
exit /b 0

:find_python
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if defined PY exit /b 0
python --version >nul 2>&1 && set "PY=python"
exit /b 0
