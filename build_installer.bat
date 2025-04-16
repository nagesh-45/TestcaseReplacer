@echo off
setlocal

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "INSTALLER_DIR=%PROJECT_ROOT%dist\installer"
set "RESOURCES_DIR=%PROJECT_ROOT%src\main\resources"

echo Building Excel Replacer Installer...
echo Project Root: %PROJECT_ROOT%
echo.

:: Check if Java is installed
echo Checking Java installation...
java -version >nul 2>&1
if errorlevel 1 (
    echo Error: Java is not installed or not found in PATH
    echo Please install Java 14 or later
    pause
    exit /b 1
)

:: Check if jpackage is available
echo Checking jpackage...
jpackage --version >nul 2>&1
if errorlevel 1 (
    echo Error: jpackage not found
    echo Please install Java 14 or later with jpackage support
    echo You can download it from: https://adoptium.net/
    pause
    exit /b 1
)

:: Create installer directory if it doesn't exist
if not exist "%INSTALLER_DIR%" (
    echo Creating installer directory: %INSTALLER_DIR%
    mkdir "%INSTALLER_DIR%"
)

:: Build the project
echo Building project...
call mvn clean install
if errorlevel 1 (
    echo Error: Build failed
    pause
    exit /b 1
)

:: Verify JAR file exists
if not exist "target\testcasereplacer-1.0-SNAPSHOT.jar" (
    echo Error: JAR file not found in target directory
    pause
    exit /b 1
)

:: Create the installer
echo Creating installer...
echo Running jpackage command...
jpackage ^
  --name "Excel Replacer" ^
  --app-version "1.0.0" ^
  --description "Excel Replacer Tool" ^
  --vendor "Your Company" ^
  --type msi ^
  --input target ^
  --main-jar testcasereplacer-1.0-SNAPSHOT.jar ^
  --main-class com.example.ExcelReplacer ^
  --win-console ^
  --win-shortcut ^
  --win-menu ^
  --win-menu-group "Excel Tools" ^
  --dest "%INSTALLER_DIR%"

if errorlevel 1 (
    echo Error: Failed to create installer
    echo Please check if you have the required permissions
    pause
    exit /b 1
)

:: Verify installer was created
if not exist "%INSTALLER_DIR%\Excel Replacer.msi" (
    echo Error: Installer was not created
    echo Please check the jpackage output above for errors
    pause
    exit /b 1
)

:: Copy resources to installer directory
echo Copying resources...
if not exist "%INSTALLER_DIR%\resources" mkdir "%INSTALLER_DIR%\resources"
xcopy "%RESOURCES_DIR%\*.*" "%INSTALLER_DIR%\resources\" /E /I /Y

echo.
echo ===========================================
echo Installer created successfully!
echo Location: %INSTALLER_DIR%
echo.
echo Contents:
dir "%INSTALLER_DIR%"
echo ===========================================
echo.
pause 