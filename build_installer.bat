@echo off
setlocal

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "INSTALLER_DIR=%PROJECT_ROOT%dist\installer"
set "RESOURCES_DIR=%PROJECT_ROOT%src\main\resources"

echo Building Excel Replacer Installer...
echo Project Root: %PROJECT_ROOT%
echo.

:: Create installer directory if it doesn't exist
if not exist "%INSTALLER_DIR%" mkdir "%INSTALLER_DIR%"

:: Build the project
echo Building project...
call mvn clean install
if errorlevel 1 (
    echo Error: Build failed
    pause
    exit /b 1
)

:: Create the installer
echo Creating installer...
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

:: Copy resources to installer directory
echo Copying resources...
xcopy "%RESOURCES_DIR%\*.*" "%INSTALLER_DIR%\resources\" /E /I /Y

echo.
echo ===========================================
echo Installer created in: %INSTALLER_DIR%
echo.
echo Contents:
echo - Excel Replacer.msi (Windows Installer)
echo - resources\ (Configuration files)
echo ===========================================
echo.
pause 