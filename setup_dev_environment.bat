@echo off
echo Setting up ICP/MP Expert Automation Development Environment...
echo.

REM Check if IronPython is available
where ipy.exe >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: IronPython not found in PATH
    echo Please install IronPython 2.7 from: https://github.com/IronLanguages/ironpython2/releases
    echo.
    pause
    exit /b 1
)

echo ✓ IronPython found
ipy.exe --version

echo.
echo Testing syntax validation...
ipy.exe tests\test_gui_syntax.py

echo.
echo Testing IronPython compatibility...
ipy.exe tests\test_ironpython.py

echo.
echo Development environment setup complete!
echo.
echo To run the main application:
echo   ipy.exe gui_automation_app.py
echo.
echo To run tests:
echo   ipy.exe tests\comprehensive_test.py
echo.
pause
