@echo off
setlocal enabledelayedexpansion

:: Set current directory as base path
set "CURRENT_DIR=%~dp0"
cd /d "%CURRENT_DIR%"

:: Set paths relative to current directory
set "RESOURCES_DIR=src\main\resources"
set "OUTPUT_DIR=output"

:: Create output directory if it doesn't exist
if not exist "%OUTPUT_DIR%" (
    echo Creating output directory...
    mkdir "%OUTPUT_DIR%"
)

:menu
cls
echo Excel Replacer - Menu
echo ====================
echo Current Directory: %CURRENT_DIR%
echo.
echo 1. Run Excel Replacer
echo 2. Open Output Folder
echo 3. Check Files
echo 4. Exit
echo.
set /p choice="Enter your choice (1-4): "

if "%choice%"=="1" goto run
if "%choice%"=="2" goto open_output
if "%choice%"=="3" goto check_files
if "%choice%"=="4" goto end
echo Invalid choice. Please try again.
timeout /t 2 >nul
goto menu

:check_files
cls
echo Checking required files...
echo.
echo Current Directory: %CURRENT_DIR%
echo.

if exist "%RESOURCES_DIR%\config.properties" (
    echo [OK] config.properties found
) else (
    echo [ERROR] config.properties not found in %RESOURCES_DIR%
)

if exist "%RESOURCES_DIR%\IDPE Onus Tcs.xlsm" (
    echo [OK] Excel file found
) else (
    echo [ERROR] Excel file not found in %RESOURCES_DIR%
)

echo.
echo Press any key to return to menu...
pause >nul
goto menu

:run
cls
echo Running Excel Replacer...
echo Current Directory: %CURRENT_DIR%
echo.

:: Check if required files exist
if not exist "%RESOURCES_DIR%\config.properties" (
    echo Error: config.properties not found
    echo Please ensure the file exists in: %RESOURCES_DIR%
    echo.
    echo Current files in %RESOURCES_DIR%:
    dir "%RESOURCES_DIR%"
    pause
    goto menu
)

if not exist "%RESOURCES_DIR%\IDPE Onus Tcs.xlsm" (
    echo Error: Excel file not found
    echo Please ensure the file exists in: %RESOURCES_DIR%
    echo.
    echo Current files in %RESOURCES_DIR%:
    dir "%RESOURCES_DIR%"
    pause
    goto menu
)

:: Create timestamp for output file
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YYYY=%dt:~0,4%"
set "MM=%dt:~4,2%"
set "DD=%dt:~6,2%"
set "HH=%dt:~8,2%"
set "Min=%dt:~10,2%"
set "Sec=%dt:~12,2%"
set "timestamp=%YYYY%%MM%%DD%_%HH%%Min%%Sec%"

:: Copy and rename the Excel file
echo Copying Excel file...
copy "%RESOURCES_DIR%\IDPE Onus Tcs.xlsm" "%OUTPUT_DIR%\IDPE Onus Tcs_%timestamp%.xlsm"
if errorlevel 1 (
    echo Error: Failed to copy file
    echo Please check file permissions
    pause
    goto menu
)

echo.
echo Process completed successfully!
echo Output file created: IDPE Onus Tcs_%timestamp%.xlsm
echo in folder: %OUTPUT_DIR%
echo.
pause
goto menu

:open_output
if exist "%OUTPUT_DIR%" (
    explorer "%OUTPUT_DIR%"
) else (
    echo Output directory not found
    pause
)
goto menu

:end
echo.
echo Thank you for using Excel Replacer!
timeout /t 2 >nul
exit /b 0 