@echo off
setlocal enabledelayedexpansion

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "RESOURCES_DIR=%PROJECT_ROOT%src\main\resources"
set "OUTPUT_DIR=%PROJECT_ROOT%output"

:: Create output directory if it doesn't exist
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

:menu
cls
echo Excel Replacer - Menu
echo ====================
echo 1. Run Excel Replacer
echo 2. Open Output Folder
echo 3. Exit
echo.
set /p choice="Enter your choice (1-3): "

if "%choice%"=="1" goto run
if "%choice%"=="2" goto open_output
if "%choice%"=="3" goto end
echo Invalid choice. Please try again.
timeout /t 2 >nul
goto menu

:run
cls
echo Running Excel Replacer...
echo.

:: Check if required files exist
if not exist "%RESOURCES_DIR%\config.properties" (
    echo Error: config.properties not found
    echo Please ensure the file exists in: %RESOURCES_DIR%
    pause
    goto menu
)

if not exist "%RESOURCES_DIR%\IDPE Onus Tcs.xlsm" (
    echo Error: Excel file not found
    echo Please ensure the file exists in: %RESOURCES_DIR%
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
copy "%RESOURCES_DIR%\IDPE Onus Tcs.xlsm" "%OUTPUT_DIR%\IDPE Onus Tcs_%timestamp%.xlsm" >nul

echo.
echo Process completed successfully!
echo Output file created: IDPE Onus Tcs_%timestamp%.xlsm
echo.
pause
goto menu

:open_output
explorer "%OUTPUT_DIR%"
goto menu

:end
echo.
echo Thank you for using Excel Replacer!
timeout /t 2 >nul
exit /b 0 