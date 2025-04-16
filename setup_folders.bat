@echo off
setlocal enabledelayedexpansion

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "FILES_DIR=%PROJECT_ROOT%Files"
set "OUTPUT_DIR=%PROJECT_ROOT%Output"
set "TARGET_DIR=%PROJECT_ROOT%target"

:: Create directory structure
echo Creating required folder structure...
echo.

:: Create Files directory
if not exist "%FILES_DIR%" (
    echo Creating Files directory...
    mkdir "%FILES_DIR%"
    echo ✓ Files directory created at: %FILES_DIR%
) else (
    echo ✓ Files directory already exists at: %FILES_DIR%
)

:: Create Output directory
if not exist "%OUTPUT_DIR%" (
    echo Creating Output directory...
    mkdir "%OUTPUT_DIR%"
    echo ✓ Output directory created at: %OUTPUT_DIR%
) else (
    echo ✓ Output directory already exists at: %OUTPUT_DIR%
)

:: Create target directory
if not exist "%TARGET_DIR%" (
    echo Creating target directory...
    mkdir "%TARGET_DIR%"
    echo ✓ Target directory created at: %TARGET_DIR%
) else (
    echo ✓ Target directory already exists at: %TARGET_DIR%
)

:: Verify folder structure
echo.
echo Verifying folder structure...
echo.
echo Current folder structure:
echo %PROJECT_ROOT%
dir /X "%PROJECT_ROOT%"

:: Check for required files
echo.
echo Checking for required files...
echo.

:: Check for Excel file
set "EXCEL_FOUND="
for %%f in ("%FILES_DIR%\*.xlsm") do set "EXCEL_FOUND=%%f"
if defined EXCEL_FOUND (
    echo ✓ Excel file found: %~nxEXCEL_FOUND%
    echo   8.3 name: %~snxEXCEL_FOUND%
) else (
    echo ✗ No Excel file (.xlsm) found in Files directory
    echo   Please place your Excel file in: %FILES_DIR%
)

:: Check for config file
set "CONFIG_FOUND="
for %%f in ("%FILES_DIR%\*.properties") do set "CONFIG_FOUND=%%f"
if defined CONFIG_FOUND (
    echo ✓ Config file found: %~nxCONFIG_FOUND%
    echo   8.3 name: %~snxCONFIG_FOUND%
) else (
    echo ✗ No config file (.properties) found in Files directory
    echo   Please place your config file in: %FILES_DIR%
)

:: Check for JAR file
set "JAR_FILE=%TARGET_DIR%\testcasereplacer-1.0-SNAPSHOT-jar-with-dependencies.jar"
if exist "%JAR_FILE%" (
    echo ✓ JAR file found: %~nxJAR_FILE%
    echo   8.3 name: %~snxJAR_FILE%
) else (
    echo ✗ JAR file not found
    echo   Please build the project using: mvn clean package
)

:: Display final instructions
echo.
echo Setup complete!
echo.
echo Required folder structure:
echo %PROJECT_ROOT%
echo ├── Files\
echo │   ├── IDPE Onus Tcs.xlsm
echo │   └── config.properties
echo ├── Output\
echo └── target\
echo     └── testcasereplacer-1.0-SNAPSHOT-jar-with-dependencies.jar
echo.
echo Next steps:
echo 1. Place your Excel file (.xlsm) in the Files directory
echo 2. Place your config file (.properties) in the Files directory
echo 3. Build the project using: mvn clean package
echo 4. Run the Excel Replacer using: run_excel_replacer.bat
echo.
echo Press any key to exit...
pause >nul 