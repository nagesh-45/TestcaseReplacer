@echo off
setlocal enabledelayedexpansion

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "JAR_FILE=%PROJECT_ROOT%target\testcasereplacer-1.0-SNAPSHOT.jar"
set "OUTPUT_DIR=%PROJECT_ROOT%dist"
set "APP_NAME=ExcelReplacer"
set "LAUNCH4J_DIR=%PROJECT_ROOT%launch4j"

:: Check for JAR file
if not exist "%JAR_FILE%" (
    echo ERROR: JAR file not found!
    echo Please build the project first: mvn clean package
    pause
    exit /b 1
)

:: Create output directory
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

:: Create Launch4j configuration file
echo Creating Launch4j configuration...
(
echo ^<?xml version="1.0" encoding="UTF-8"?^>
echo ^<launch4jConfig^>
echo   ^<dontWrapJar^>false^</dontWrapJar^>
echo   ^<headerType^>gui^</headerType^>
echo   ^<jar^>%JAR_FILE%^</jar^>
echo   ^<outfile^>%OUTPUT_DIR%\%APP_NAME%.exe^</outfile^>
echo   ^<errTitle^>Error^</errTitle^>
echo   ^<cmdLine^>^</cmdLine^>
echo   ^<chdir^>^</chdir^>
echo   ^<priority^>normal^</priority^>
echo   ^<downloadUrl^>http://java.com/download^</downloadUrl^>
echo   ^<supportUrl^>^</supportUrl^>
echo   ^<stayAlive^>false^</stayAlive^>
echo   ^<restartOnCrash^>false^</restartOnCrash^>
echo   ^<manifest^>^</manifest^>
echo   ^<icon^>%PROJECT_ROOT%icon.ico^</icon^>
echo   ^<jre^>
echo     ^<path^>^</path^>
echo     ^<bundledJre64Bit^>true^</bundledJre64Bit^>
echo     ^<minVersion^>1.8.0^</minVersion^>
echo     ^<maxVersion^>^</maxVersion^>
echo     ^<jdkPreference^>preferJre^</jdkPreference^>
echo     ^<runtimeBits^>64/32^</runtimeBits^>
echo   ^</jre^>
echo   ^<versionInfo^>
echo     ^<fileVersion^>1.0.0.0^</fileVersion^>
echo     ^<txtFileVersion^>1.0.0^</txtFileVersion^>
echo     ^<fileDescription^>Excel Replacer Tool^</fileDescription^>
echo     ^<copyright^>Copyright (c) 2024 YourCompany^</copyright^>
echo     ^<productVersion^>1.0.0.0^</productVersion^>
echo     ^<txtProductVersion^>1.0.0^</txtProductVersion^>
echo     ^<companyName^>YourCompany^</companyName^>
echo     ^<internalName^>ExcelReplacer^</internalName^>
echo     ^<originalFilename^>ExcelReplacer.exe^</originalFilename^>
echo   ^</versionInfo^>
echo ^</launch4jConfig^>
) > "%PROJECT_ROOT%launch4j.xml"

:: Check if Launch4j is installed
if not exist "%LAUNCH4J_DIR%\launch4j.exe" (
    echo Launch4j not found. Downloading...
    powershell -Command "Invoke-WebRequest -Uri 'https://sourceforge.net/projects/launch4j/files/launch4j-3/3.14/launch4j-3.14-win32.zip/download' -OutFile 'launch4j.zip'"
    powershell -Command "Expand-Archive -Path 'launch4j.zip' -DestinationPath 'launch4j' -Force"
    del launch4j.zip
)

:: Create executable
echo Creating executable...
"%LAUNCH4J_DIR%\launch4j.exe" "%PROJECT_ROOT%launch4j.xml"

if errorlevel 1 (
    echo ERROR: Failed to create executable
    pause
    exit /b 1
)

echo.
echo SUCCESS: Executable created successfully!
echo Location: %OUTPUT_DIR%\%APP_NAME%.exe
echo.
pause 