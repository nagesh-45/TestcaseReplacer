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

:: Find files using 8.3 names
set "CONFIG_FOUND="
set "EXCEL_FOUND="

for %%f in ("%FILES_DIR%\*.*") do (
    if /i "%%~snxf"=="CONFIG~.PRO" set "CONFIG_FOUND=%%f"
    if /i "%%~snxf"=="IDPEON~1.XLS" set "EXCEL_FOUND=%%f"
)

if not defined CONFIG_FOUND (
    echo ERROR: config.properties not found!
    echo Looking for file with 8.3 name: CONFIG~.PRO
    echo Current files in Files folder:
    dir "%FILES_DIR%" /X
    echo.
    pause
    exit /b 1
)

if not defined EXCEL_FOUND (
    echo ERROR: Excel file not found!
    echo Looking for file with 8.3 name: IDPEON~1.XLS
    echo Current files in Files folder:
    dir "%FILES_DIR%" /X
    echo.
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

:: Run Java program to replace values
echo Running Excel Replacer...
echo Input Excel: %EXCEL_FOUND%
echo Config File: %CONFIG_FOUND%
echo Output File: %OUTPUT_DIR%\IDPE Onus Tcs_%timestamp%.xlsm
echo.

java -jar "%JAR_FILE%" "%EXCEL_FOUND%" "%CONFIG_FOUND%" "%OUTPUT_DIR%\IDPE Onus Tcs_%timestamp%.xlsm"

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
echo Output file: IDPE Onus Tcs_%timestamp%.xlsm
echo Location: %OUTPUT_DIR%
echo.
pause 