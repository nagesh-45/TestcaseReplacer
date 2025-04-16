@echo off
setlocal enabledelayedexpansion

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "FILES_DIR=%PROJECT_ROOT%Files"
set "OUTPUT_DIR=%PROJECT_ROOT%Output"
set "JAR_FILE=%PROJECT_ROOT%target\testcasereplacer-1.0-SNAPSHOT-jar-with-dependencies.jar"

:: Debug information
echo Current directory: %CD%
echo Project root: %PROJECT_ROOT%
echo Files directory: %FILES_DIR%
echo Output directory: %OUTPUT_DIR%
echo JAR file: %JAR_FILE%

:: Create output directory if it doesn't exist
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

:: Check for required files
echo Checking for required files in %FILES_DIR%...
if not exist "%FILES_DIR%" (
    echo ERROR: Files directory not found at %FILES_DIR%
    exit /b 1
)

:: List all files in the Files directory with 8.3 names
echo Available files in %FILES_DIR% (with 8.3 names):
dir /X "%FILES_DIR%"

:: Check for Excel file (case-insensitive and 8.3 name aware)
set "EXCEL_FILE="
set "EXCEL_SHORT="
for /r "%FILES_DIR%" %%f in (*.xlsm) do (
    if not defined EXCEL_FILE (
        set "EXCEL_FILE=%%f"
        set "EXCEL_SHORT=%%~sf"
    )
)
if not defined EXCEL_FILE (
    echo ERROR: No Excel file (.xlsm) found in %FILES_DIR%
    echo Looking for files with pattern: *.xlsm
    dir /X "%FILES_DIR%\*.xlsm" 2>nul
    exit /b 1
)
echo Found Excel file: %~nxEXCEL_FILE%
echo Full path: %EXCEL_FILE%
echo 8.3 name: %EXCEL_SHORT%

:: Check for config file (case-insensitive and 8.3 name aware)
set "CONFIG_FILE="
set "CONFIG_SHORT="
for /r "%FILES_DIR%" %%f in (*.properties) do (
    if not defined CONFIG_FILE (
        set "CONFIG_FILE=%%f"
        set "CONFIG_SHORT=%%~sf"
    )
)
if not defined CONFIG_FILE (
    echo ERROR: No config file (.properties) found in %FILES_DIR%
    echo Looking for files with pattern: *.properties
    dir /X "%FILES_DIR%\*.properties" 2>nul
    exit /b 1
)
echo Found config file: %~nxCONFIG_FILE%
echo Full path: %CONFIG_FILE%
echo 8.3 name: %CONFIG_SHORT%

:: Check for Java
where java >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ERROR: Java is not installed or not found in PATH
    exit /b 1
)

:: Display Java version
echo Java version:
java -version

:: Display Java home
echo Java home:
echo %JAVA_HOME%

:: Check for JAR file
if not exist "%JAR_FILE%" (
    echo ERROR: JAR file not found at %JAR_FILE%
    echo Please run 'mvn clean package' to build the JAR file
    exit /b 1
)
echo JAR file exists at: %JAR_FILE%
for %%A in ("%JAR_FILE%") do echo JAR file size: %%~zA bytes

:: Create timestamp for output file
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set datetime=%%I
set "TIMESTAMP=%datetime:~0,8%_%datetime:~8,6%"
set "OUTPUT_FILE=%OUTPUT_DIR%\%~nEXCEL_FILE%_%TIMESTAMP%.xlsm"

echo Running Excel Replacer...
echo Input Excel: %EXCEL_FILE%
echo Config File: %CONFIG_FILE%
echo Output File: %OUTPUT_FILE%

:: Create temporary directory for processing
set "TEMP_DIR=%TEMP%\excel_replacer_%RANDOM%"
mkdir "%TEMP_DIR%"
echo Created temporary directory: %TEMP_DIR%

:: Copy files to temporary directory using 8.3 names
echo Copying files to temporary directory...
copy "%EXCEL_FILE%" "%TEMP_DIR%\input.xlsm" >nul
if %ERRORLEVEL% neq 0 (
    echo ERROR: Failed to copy Excel file
    echo Trying with 8.3 name...
    copy "%EXCEL_SHORT%" "%TEMP_DIR%\input.xlsm" >nul
    if %ERRORLEVEL% neq 0 (
        echo ERROR: Failed to copy Excel file even with 8.3 name
        rmdir /s /q "%TEMP_DIR%"
        exit /b 1
    )
)

copy "%CONFIG_FILE%" "%TEMP_DIR%\config.properties" >nul
if %ERRORLEVEL% neq 0 (
    echo ERROR: Failed to copy config file
    echo Trying with 8.3 name...
    copy "%CONFIG_SHORT%" "%TEMP_DIR%\config.properties" >nul
    if %ERRORLEVEL% neq 0 (
        echo ERROR: Failed to copy config file even with 8.3 name
        rmdir /s /q "%TEMP_DIR%"
        exit /b 1
    )
)

:: Run Java program
echo Running Java program...
echo Java command: java -jar "%JAR_FILE%" "%TEMP_DIR%\input.xlsm" "%TEMP_DIR%\config.properties" "%OUTPUT_FILE%"
java -jar "%JAR_FILE%" "%TEMP_DIR%\input.xlsm" "%TEMP_DIR%\config.properties" "%OUTPUT_FILE%"
if %ERRORLEVEL% neq 0 (
    echo ERROR: Java program failed with exit code %ERRORLEVEL%
    echo Please check the Java output above for details.
    echo Common issues:
    echo 1. Check if the Excel file is not corrupted
    echo 2. Verify the config.properties format
    echo 3. Check Java version compatibility
    echo 4. Ensure all required dependencies are in the JAR
    rmdir /s /q "%TEMP_DIR%"
    exit /b 1
)

:: Clean up temporary directory
rmdir /s /q "%TEMP_DIR%"

echo Excel Replacer completed successfully!
echo Output file created at: %OUTPUT_FILE%

endlocal 