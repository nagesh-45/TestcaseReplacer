@echo off
setlocal enabledelayedexpansion

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "FILES_DIR=%PROJECT_ROOT%Files"
set "OUTPUT_DIR=%PROJECT_ROOT%Output"
set "JAR_FILE=%PROJECT_ROOT%target\testcasereplacer-1.0-SNAPSHOT.jar"

:: Debug information
echo Current directory: %CD%
echo Project root: %PROJECT_ROOT%
echo Files directory: %FILES_DIR%
echo.

:: Create output directory if it doesn't exist
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

:: Check for required files
if not exist "%FILES_DIR%" (
    echo ERROR: Files folder not found!
    echo Please create a 'Files' folder in: %PROJECT_ROOT%
    dir "%PROJECT_ROOT%"
    pause
    exit /b 1
)

:: List contents of Files folder with detailed information
echo Contents of Files folder:
dir "%FILES_DIR%" /X
echo.

:menu
cls
echo Excel Replacer
echo ==============
echo.
echo 1. Use default files (IDPE Onus Tcs.xlsm and config.properties)
echo 2. Select different Excel file
echo 3. Select different config file
echo 4. Exit
echo.
set /p choice="Enter your choice (1-4): "

if "%choice%"=="1" goto default_files
if "%choice%"=="2" goto select_excel
if "%choice%"=="3" goto select_config
if "%choice%"=="4" exit /b 0

echo Invalid choice! Please try again.
timeout /t 2 >nul
goto menu

:select_excel
cls
echo Available Excel files in %FILES_DIR%:
echo.
echo Regular names and 8.3 names:
echo ---------------------------
dir "%FILES_DIR%\*.xls*" /X
echo.
set /p excel_file="Enter Excel filename (with extension): "
set "EXCEL_FOUND="

:: Try to find file with exact name first
if exist "%FILES_DIR%\%excel_file%" (
    set "EXCEL_FOUND=%FILES_DIR%\%excel_file%"
) else (
    :: If not found, try to match 8.3 name
    for %%f in ("%FILES_DIR%\*.*") do (
        if /i "%%~snxf"=="%excel_file%" set "EXCEL_FOUND=%%f"
    )
)

if not defined EXCEL_FOUND (
    echo File not found!
    echo Please check the filename and try again.
    timeout /t 2 >nul
    goto select_excel
)
goto check_config

:select_config
cls
echo Available config files in %FILES_DIR%:
echo.
echo Regular names and 8.3 names:
echo ---------------------------
dir "%FILES_DIR%\*.properties" /X
echo.
set /p config_file="Enter config filename (with extension): "
set "CONFIG_FOUND="

:: Try to find file with exact name first
if exist "%FILES_DIR%\%config_file%" (
    set "CONFIG_FOUND=%FILES_DIR%\%config_file%"
) else (
    :: If not found, try to match 8.3 name
    for %%f in ("%FILES_DIR%\*.*") do (
        if /i "%%~snxf"=="%config_file%" set "CONFIG_FOUND=%%f"
    )
)

if not defined CONFIG_FOUND (
    echo File not found!
    echo Please check the filename and try again.
    timeout /t 2 >nul
    goto select_config
)
goto check_excel

:default_files
:: Find files using exact 8.3 names
set "CONFIG_FOUND="
set "EXCEL_FOUND="

for %%f in ("%FILES_DIR%\*.*") do (
    if /i "%%~snxf"=="CONFIG~1.PRO" set "CONFIG_FOUND=%%f"
    if /i "%%~snxf"=="IDPEON~1.XLS" set "EXCEL_FOUND=%%f"
)

:check_config
if not defined CONFIG_FOUND (
    echo ERROR: config.properties not found!
    echo Looking for file with 8.3 name: CONFIG~1.PRO
    echo Current files in Files folder:
    dir "%FILES_DIR%" /X
    echo.
    echo Please select a config file from the menu.
    pause
    goto menu
)

:check_excel
if not defined EXCEL_FOUND (
    echo ERROR: Excel file not found!
    echo Looking for file with 8.3 name: IDPEON~1.XLS
    echo Current files in Files folder:
    dir "%FILES_DIR%" /X
    echo.
    echo Please select an Excel file from the menu.
    pause
    goto menu
)

:: Check for Java
java -version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Java is not installed!
    echo Please install Java 8 or later from: https://www.java.com/download/
    pause
    exit /b 1
)

:: Check for JAR file
if not exist "%JAR_FILE%" (
    echo ERROR: Java program not found!
    echo Please make sure the JAR file exists at: %JAR_FILE%
    dir "%PROJECT_ROOT%target"
    pause
    exit /b 1
)

:: Create timestamp for output file
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "timestamp=%dt:~0,8%_%dt:~8,6%"

:: Get base name of Excel file for output
for %%i in ("%EXCEL_FOUND%") do set "base_name=%%~ni"

:: Run Java program to replace values
echo Running Excel Replacer...
echo Input Excel: %EXCEL_FOUND%
echo Config File: %CONFIG_FOUND%
echo Output File: %OUTPUT_DIR%\%base_name%_%timestamp%.xlsm
echo.

java -jar "%JAR_FILE%" "%EXCEL_FOUND%" "%CONFIG_FOUND%" "%OUTPUT_DIR%\%base_name%_%timestamp%.xlsm"

if errorlevel 1 (
    echo ERROR: Failed to process Excel file!
    echo Please make sure:
    echo 1. The Excel file is not open in Excel
    echo 2. You have write permissions to this folder
    echo 3. The config.properties file has the correct format
    pause
    exit /b 1
)

echo.
echo SUCCESS: Excel file processed successfully!
echo Output file: %base_name%_%timestamp%.xlsm
echo Location: %OUTPUT_DIR%
echo.
pause 