@echo off
setlocal

:: Set paths
set "FILES_DIR=Files"
set "OUTPUT_DIR=Output"
set "JAR_FILE=target\testcasereplacer-1.0-SNAPSHOT-jar-with-dependencies.jar"

:: Create required directories
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

:: Check for required files
echo Checking files in %FILES_DIR%...
dir /X "%FILES_DIR%"

:: Set Excel file
set "EXCEL_FILE=%FILES_DIR%\IDPE Onus Tcs.xlsm"
set "CONFIG_FILE=%FILES_DIR%\config.properties"

:: Check if files exist
if not exist "%EXCEL_FILE%" (
    echo ERROR: Excel file not found: %EXCEL_FILE%
    pause
    exit /b 1
)

if not exist "%CONFIG_FILE%" (
    echo ERROR: Config file not found: %CONFIG_FILE%
    pause
    exit /b 1
)

:: Create timestamp
set "TIMESTAMP=%date:~-4%%date:~-7,2%%date:~-10,2%_%time:~0,2%%time:~3,2%%time:~6,2%"
set "TIMESTAMP=%TIMESTAMP: =0%"
set "OUTPUT_FILE=%OUTPUT_DIR%\IDPE Onus Tcs_%TIMESTAMP%.xlsm"

:: Create temp directory
set "TEMP_DIR=%TEMP%\excel_replacer_%RANDOM%"
mkdir "%TEMP_DIR%"

:: Copy files to temp directory
copy "%EXCEL_FILE%" "%TEMP_DIR%\input.xlsm" >nul
copy "%CONFIG_FILE%" "%TEMP_DIR%\config.properties" >nul

:: Check Java installation
java -version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ERROR: Java is not installed or not in PATH
    echo Please install Java and add it to your PATH
    pause
    exit /b 1
)

:: Display Java version
echo Java version:
java -version

:: Run Java
echo Running Excel Replacer...
java -jar "%JAR_FILE%" "%TEMP_DIR%\input.xlsm" "%TEMP_DIR%\config.properties" "%OUTPUT_FILE%"
if %ERRORLEVEL% neq 0 (
    echo ERROR: Java program failed
    echo Common issues:
    echo 1. Check if the Excel file is not corrupted
    echo 2. Verify the config.properties format
    echo 3. Check Java version compatibility
    echo 4. Ensure all required dependencies are in the JAR
    echo 5. Check if Excel is not running (close Excel if open)
    echo 6. Verify Java is in your PATH
    echo 7. Check file permissions
    echo 8. Make sure the JAR file exists at: %JAR_FILE%
    pause
    exit /b 1
)

:: Clean up
rmdir /s /q "%TEMP_DIR%"

echo Excel Replacer completed successfully!
echo Output file created at: %OUTPUT_FILE%
pause 