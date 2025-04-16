@echo off
setlocal enabledelayedexpansion

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "FILES_DIR=%PROJECT_ROOT%Files"
set "OUTPUT_DIR=%PROJECT_ROOT%Output"
set "JAR_FILE=%PROJECT_ROOT%target\testcasereplacer-1.0-SNAPSHOT.jar"

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

:: Show available files
echo Available files in %FILES_DIR%:
dir "%FILES_DIR%" /X
echo.

:: Find Excel file (.xlsm first, then other Excel formats)
set "EXCEL_FOUND="
for %%f in ("%FILES_DIR%\*.xlsm") do (
    echo Found Excel file: %%~nxf (8.3 name: %%~snxf)
    set "EXCEL_FOUND=%%f"
    goto :found_excel
)

:: If no .xlsm found, try other Excel formats
if not defined EXCEL_FOUND (
    for %%f in ("%FILES_DIR%\*.xls*") do (
        echo Found Excel file: %%~nxf (8.3 name: %%~snxf)
        set "EXCEL_FOUND=%%f"
        goto :found_excel
    )
)

:found_excel
if not defined EXCEL_FOUND (
    echo ERROR: No Excel file found!
    pause
    exit /b 1
)

:: Find config file
set "CONFIG_FOUND="
for %%f in ("%FILES_DIR%\*.properties") do (
    echo Found config file: %%~nxf (8.3 name: %%~snxf)
    set "CONFIG_FOUND=%%f"
)

if not defined CONFIG_FOUND (
    echo ERROR: No config file found!
    pause
    exit /b 1
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