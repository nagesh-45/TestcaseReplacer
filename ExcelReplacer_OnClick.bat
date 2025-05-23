@echo off
setlocal enabledelayedexpansion

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "FILES_DIR=%PROJECT_ROOT%Files"
set "OUTPUT_DIR=%PROJECT_ROOT%Output"
set "JAR_FILE=%PROJECT_ROOT%target\testcasereplacer-1.0-SNAPSHOT.jar"

:: Create output directory if it doesn't exist
if not exist "%OUTPUT_DIR%" (
    echo Creating output directory: %OUTPUT_DIR%
    mkdir "%OUTPUT_DIR%"
)

:: Check for required files
if not exist "%FILES_DIR%" (
    echo ERROR: Files folder not found!
    echo Please create a 'Files' folder in: %PROJECT_ROOT%
    dir "%PROJECT_ROOT%"
    echo Press any key to exit...
    pause >nul
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
    echo Press any key to exit...
    pause >nul
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
    echo Press any key to exit...
    pause >nul
    exit /b 1
)

:: Check for Java
java -version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Java is not installed or not found in PATH
    echo Please install Java
    pause
    exit /b 1
)

:: Check for JAR file
if not exist "%JAR_FILE%" (
    echo ERROR: JAR file not found!
    echo Please build the project first: mvn clean package
    pause
    exit /b 1
)

:: Create timestamp for output file
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "timestamp=%dt:~0,8%_%dt:~8,6%"

:: Get base name of Excel file for output
for %%i in ("%EXCEL_FOUND%") do set "base_name=%%~ni"
set "OUTPUT_FILE=%OUTPUT_DIR%\%base_name%_%timestamp%.xlsm"

:: Run Java program to replace values
echo Running Excel Replacer...
echo Input Excel: %EXCEL_FOUND%
echo Config File: %CONFIG_FOUND%
echo Output File: %OUTPUT_FILE%
echo.

:: Ensure files are not in use
echo Checking if files are in use...
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>NUL | find /I /N "EXCEL.EXE">NUL
if "%ERRORLEVEL%"=="0" (
    echo ERROR: Excel is running. Please close Excel and try again.
    echo Press any key to exit...
    pause >nul
    exit /b 1
)

:: Copy files to temporary location to avoid path issues
set "TEMP_DIR=%TEMP%\ExcelReplacer_%RANDOM%"
mkdir "%TEMP_DIR%"
copy "%EXCEL_FOUND%" "%TEMP_DIR%\input.xlsm" >nul
copy "%CONFIG_FOUND%" "%TEMP_DIR%\config.properties" >nul

:: Run Java program with temporary files and capture output
echo Running Java program...
java -jar "%JAR_FILE%" "%TEMP_DIR%\input.xlsm" "%TEMP_DIR%\config.properties" "%OUTPUT_FILE%" 2>&1 | tee "%TEMP_DIR%\java_output.txt"

:: Check Java program exit code
set "JAVA_EXIT_CODE=%ERRORLEVEL%"
echo.
echo Java program output:
echo -------------------
type "%TEMP_DIR%\java_output.txt"
echo -------------------
echo.

:: Clean up temporary files
del "%TEMP_DIR%\input.xlsm"
del "%TEMP_DIR%\config.properties"
del "%TEMP_DIR%\java_output.txt"
rmdir "%TEMP_DIR%"

if %JAVA_EXIT_CODE% neq 0 (
    echo ERROR: Java program failed with exit code %JAVA_EXIT_CODE%
    echo Please check the Java output above for details.
    echo Press any key to exit...
    pause >nul
    exit /b 1
)

:: Verify output file was created
if not exist "%OUTPUT_FILE%" (
    echo ERROR: Output file was not created!
    echo Expected location: %OUTPUT_FILE%
    echo Current contents of Output folder:
    dir "%OUTPUT_DIR%"
    echo Press any key to exit...
    pause >nul
    exit /b 1
)

echo.
echo SUCCESS: Excel file processed successfully!
echo Output file: %base_name%_%timestamp%.xlsm
echo Exact location: %OUTPUT_FILE%
echo.
echo Verifying output file...
dir "%OUTPUT_FILE%"
echo.
echo Press any key to exit...
pause >nul 