@echo off
REM Quick installer for anthropic library in pyRevit
REM Run this as Administrator

echo.
echo ============================================
echo   Anthropic Library Installer for pyRevit
echo   ValtriaPyTools Extension
echo ============================================
echo.

set SITE_PACKAGES=C:\Program Files\pyRevit-Master\site-packages
set PYTHON_EXE=C:\Program Files\pyRevit-Master\bin\cengines\CPY3123\python.exe

echo Checking for system Python...
where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [OK] System Python found!
    echo.
    echo Installing anthropic library...
    python -m pip install anthropic --target="%SITE_PACKAGES%"

    if %ERRORLEVEL% EQU 0 (
        echo.
        echo [SUCCESS] anthropic library installed!
        echo.
        echo Testing installation...
        "%PYTHON_EXE%" -c "import sys; sys.path.insert(0, r'%SITE_PACKAGES%'); import anthropic; print('Version:', anthropic.__version__ if hasattr(anthropic, '__version__') else 'OK')"
        echo.
        echo [DONE] You can now use the Claude button in ValtriaPyTools!
    ) else (
        echo.
        echo [ERROR] Installation failed.
        echo Try running this script as Administrator.
    )
) else (
    echo [NOT FOUND] No system Python detected.
    echo.
    echo Please install Python from:
    echo   https://www.python.org/downloads/
    echo.
    echo Or use WinPython (portable):
    echo   https://winpython.github.io/
    echo.
    echo Or see the manual installation instructions in the Claude button folder.
)

echo.
pause
