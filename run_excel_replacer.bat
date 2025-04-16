@echo off
setlocal

:: Set paths
set "FILES_DIR=Files"
set "OUTPUT_DIR=Output"
set "JAR_FILE=target\testcasereplacer-1.0-SNAPSHOT-jar-with-dependencies.jar"
set "RESOURCES_DIR=src\main\resources"

:: Create required directories
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"
if not exist "%RESOURCES_DIR%" mkdir "%RESOURCES_DIR%"

:: Check for required files
echo Checking files...
dir /X "%FILES_DIR%"

:: Set Excel file
set "EXCEL_FILE=%FILES_DIR%\IDPE Onus Tcs.xlsm"
set "CONFIG_FILE=%FILES_DIR%\config.properties"

:: Check if files exist
if not exist "%EXCEL_FILE%" (
    echo ERROR: Excel file not found: %EXCEL_FILE%
    exit /b 1
)

if not exist "%CONFIG_FILE%" (
    echo ERROR: Config file not found: %CONFIG_FILE%
    exit /b 1
)

:: Copy files to resources directory
echo Copying files to resources directory...
copy "%EXCEL_FILE%" "%RESOURCES_DIR%\IDPE Onus Tcs.xlsm" >nul
copy "%CONFIG_FILE%" "%RESOURCES_DIR%\config.properties" >nul

:: Create timestamp
set "TIMESTAMP=%date:~-4%%date:~-7,2%%date:~-10,2%_%time:~0,2%%time:~3,2%%time:~6,2%"
set "TIMESTAMP=%TIMESTAMP: =0%"
set "OUTPUT_FILE=%OUTPUT_DIR%\IDPE Onus Tcs_%TIMESTAMP%.xlsm"

:: Create temp directory
set "TEMP_DIR=%TEMP%\excel_replacer_%RANDOM%"
mkdir "%TEMP_DIR%"

:: Copy files to temp directory
copy "%RESOURCES_DIR%\IDPE Onus Tcs.xlsm" "%TEMP_DIR%\input.xlsm" >nul
copy "%RESOURCES_DIR%\config.properties" "%TEMP_DIR%\config.properties" >nul

:: Run Java
echo Running Excel Replacer...
java -jar "%JAR_FILE%" "%TEMP_DIR%\input.xlsm" "%TEMP_DIR%\config.properties" "%OUTPUT_FILE%"

:: Clean up
rmdir /s /q "%TEMP_DIR%"

echo Done!
pause 