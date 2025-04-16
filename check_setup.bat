@echo off
setlocal enabledelayedexpansion

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "FILES_DIR=%PROJECT_ROOT%Files"
set "OUTPUT_DIR=%PROJECT_ROOT%Output"
set "JAR_FILE=%PROJECT_ROOT%target\testcasereplacer-1.0-SNAPSHOT.jar"

echo Checking setup...
echo =================
echo.

:: Check current directory
echo Current Directory:
echo %PROJECT_ROOT%
echo.

:: Check Files folder
echo Checking Files folder...
if exist "%FILES_DIR%" (
    echo Files folder exists at: %FILES_DIR%
    echo Contents of Files folder:
    dir "%FILES_DIR%" /X
) else (
    echo ERROR: Files folder not found at: %FILES_DIR%
)
echo.

:: Check Excel file
echo Checking Excel file...
set "EXCEL_FOUND="
for %%f in ("%FILES_DIR%\*.xlsm") do (
    echo Found Excel file: %%~nxf (8.3 name: %%~snxf)
    set "EXCEL_FOUND=%%f"
)

if not defined EXCEL_FOUND (
    echo ERROR: No Excel file (.xlsm) found in Files folder
)
echo.

:: Check config file
echo Checking config file...
set "CONFIG_FOUND="
for %%f in ("%FILES_DIR%\*.properties") do (
    echo Found config file: %%~nxf (8.3 name: %%~snxf)
    set "CONFIG_FOUND=%%f"
)

if not defined CONFIG_FOUND (
    echo ERROR: No config file (.properties) found in Files folder
)
echo.

:: Check JAR file
echo Checking JAR file...
if exist "%JAR_FILE%" (
    echo JAR file exists at: %JAR_FILE%
) else (
    echo ERROR: JAR file not found at: %JAR_FILE%
)
echo.

:: Check Java installation
echo Checking Java installation...
java -version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Java is not installed or not found in PATH
) else (
    echo Java is installed and available
    java -version
)
echo.

:: Check Output folder
echo Checking Output folder...
if exist "%OUTPUT_DIR%" (
    echo Output folder exists at: %OUTPUT_DIR%
    echo Current contents:
    dir "%OUTPUT_DIR%"
) else (
    echo Output folder will be created when running the program
)
echo.

echo =================
echo Setup Check Complete
echo.
echo If you see any ERROR messages above, please fix those issues first.
echo.
pause 