@echo off
echo Launching ICP/MP Expert Automation GUI...
echo.

REM Check if IronPython is available
where ipy.exe >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: IronPython not found in PATH
    echo Please run setup_dev_environment.bat first
    echo.
    pause
    exit /b 1
)

REM Launch the GUI application
echo Starting GUI application...
ipy.exe gui_automation_app.py

echo.
echo Application closed.
pause
