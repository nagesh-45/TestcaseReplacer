@echo off
setlocal enabledelayedexpansion

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "JAR_FILE=%PROJECT_ROOT%target\testcasereplacer-1.0-SNAPSHOT.jar"
set "OUTPUT_DIR=%PROJECT_ROOT%dist"
set "APP_NAME=ExcelReplacer"
set "APP_VERSION=1.0.0"
set "VENDOR=YourCompany"

:: Check for Java 14+
java -version 2>&1 | findstr /i "version" | findstr /i "14" >nul
if errorlevel 1 (
    echo ERROR: Java 14 or later is required for jpackage
    echo Please install Java 14 or later
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

:: Create output directory
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

:: Create installer using jpackage
echo Creating Windows installer...
jpackage ^
  --name "%APP_NAME%" ^
  --app-version "%APP_VERSION%" ^
  --vendor "%VENDOR%" ^
  --input "%PROJECT_ROOT%target" ^
  --main-jar "testcasereplacer-1.0-SNAPSHOT.jar" ^
  --main-class "com.example.ExcelReplacer" ^
  --win-menu ^
  --win-shortcut ^
  --win-dir-chooser ^
  --dest "%OUTPUT_DIR%" ^
  --type msi ^
  --java-options "-Xmx512m" ^
  --icon "%PROJECT_ROOT%icon.ico" ^
  --description "Excel Replacer Tool" ^
  --copyright "Copyright (c) 2024 YourCompany"

if errorlevel 1 (
    echo ERROR: Failed to create installer
    pause
    exit /b 1
)

echo.
echo SUCCESS: Installer created successfully!
echo Location: %OUTPUT_DIR%\%APP_NAME%-%APP_VERSION%.msi
echo.
pause 