@echo off
setlocal enabledelayedexpansion

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "JAR_FILE=%PROJECT_ROOT%target\testcasereplacer-1.0-SNAPSHOT.jar"
set "TASK_NAME=ExcelReplacer"
set "TASK_DESCRIPTION=Excel Replacer Tool - Scheduled Task"

:: Check for JAR file
if not exist "%JAR_FILE%" (
    echo ERROR: JAR file not found!
    echo Please build the project first: mvn clean package
    pause
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

:: Create the task
echo Creating scheduled task...
schtasks /create /tn "%TASK_NAME%" ^
    /tr "java -jar \"%JAR_FILE%\"" ^
    /sc ONLOGON ^
    /rl HIGHEST ^
    /ru "%USERNAME%" ^
    /f

if errorlevel 1 (
    echo ERROR: Failed to create scheduled task
    echo Please run this script as Administrator
    pause
    exit /b 1
)

echo.
echo SUCCESS: Task created successfully!
echo Task Name: %TASK_NAME%
echo Description: %TASK_DESCRIPTION%
echo Trigger: On User Logon
echo.
echo To modify the task, use Task Scheduler or run:
echo schtasks /change /tn "%TASK_NAME%"
echo.
pause 