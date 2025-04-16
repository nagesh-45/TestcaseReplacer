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
    pause
    exit /b 1
)

if not exist "%FILES_DIR%\IDPE Onus Tcs.xlsm" (
    echo ERROR: Excel file not found!
    echo Please copy 'IDPE Onus Tcs.xlsm' to: %FILES_DIR%
    pause
    exit /b 1
)

if not exist "%FILES_DIR%\config.properties" (
    echo ERROR: config.properties not found!
    echo Please copy 'config.properties' to: %FILES_DIR%
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
    pause
    exit /b 1
)

:: Create timestamp for output file
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "timestamp=%dt:~0,8%_%dt:~8,6%"

:: Run Java program to replace values
echo Running Excel Replacer...
echo Input Excel: %FILES_DIR%\IDPE Onus Tcs.xlsm
echo Config File: %FILES_DIR%\config.properties
echo Output File: %OUTPUT_DIR%\IDPE Onus Tcs_%timestamp%.xlsm
echo.

java -jar "%JAR_FILE%" "%FILES_DIR%\IDPE Onus Tcs.xlsm" "%FILES_DIR%\config.properties" "%OUTPUT_DIR%\IDPE Onus Tcs_%timestamp%.xlsm"

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