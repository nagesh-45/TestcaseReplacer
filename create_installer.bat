@echo off
echo Building Excel Replacer Installer...
echo.

:: Check if Java is installed
java -version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Java is not installed!
    echo Please install Java 14 or later from: https://www.java.com/download/
    pause
    exit /b 1
)

:: Build the project
echo Building project...
mvn clean install
if errorlevel 1 (
    echo ERROR: Failed to build project!
    pause
    exit /b 1
)

:: Create installer
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
    --dest "installer"

if errorlevel 1 (
    echo ERROR: Failed to create installer!
    echo Please make sure you have Java 14 or later installed.
    pause
    exit /b 1
)

echo.
echo SUCCESS: Installer created successfully!
echo Location: installer folder
echo.
pause 