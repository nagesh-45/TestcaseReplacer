@echo off
REM Batch file to run Excel Replacer using PowerShell

REM Check if PowerShell is available
where powershell >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ERROR: PowerShell is not installed or not in PATH
    echo Please install PowerShell or add it to your PATH
    pause
    exit /b 1
)

REM Get the directory where the batch file is located
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Run the PowerShell script
echo Running Excel Replacer...
powershell -ExecutionPolicy Bypass -File "%~dp0run_excel_replacer.ps1"

REM Check if the script ran successfully
if %ERRORLEVEL% neq 0 (
    echo.
    echo ERROR: The script failed to run
    echo Please check the error messages above
    pause
    exit /b 1
)

echo.
echo Script completed successfully!
pause 