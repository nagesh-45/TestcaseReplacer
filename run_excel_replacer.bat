@echo off
setlocal enabledelayedexpansion

echo Excel Replacer - Build and Run Tool
echo ==================================
echo.

:: Check if Java is installed
echo Checking Java installation...
java -version >nul 2>&1
if errorlevel 1 (
    echo Error: Java is not installed or not found in PATH
    echo Please install Java and try again
    pause
    exit /b 1
)

:: Check if Maven is installed
echo Checking Maven installation...
mvn -version >nul 2>&1
if errorlevel 1 (
    echo Error: Maven is not installed or not found in PATH
    echo Please install Maven and try again
    pause
    exit /b 1
)

:menu
cls
echo Excel Replacer - Menu
echo ====================
echo 1. Run Excel Replacer
echo 2. Rebuild Project
echo 3. Clean and Rebuild Project
echo 4. Exit
echo.
set /p choice="Enter your choice (1-4): "

if "%choice%"=="1" goto run
if "%choice%"=="2" goto build
if "%choice%"=="3" goto clean_build
if "%choice%"=="4" goto end
echo Invalid choice. Please try again.
timeout /t 2 >nul
goto menu

:clean_build
echo.
echo Cleaning and rebuilding project...
call mvn clean install
if errorlevel 1 (
    echo Error: Build failed
    pause
    goto menu
)
echo Build successful!
echo.
choice /c YN /m "Do you want to run the program now?"
if errorlevel 2 goto menu
if errorlevel 1 goto run

:build
echo.
echo Building project...
call mvn install
if errorlevel 1 (
    echo Error: Build failed
    pause
    goto menu
)
echo Build successful!
echo.
choice /c YN /m "Do you want to run the program now?"
if errorlevel 2 goto menu
if errorlevel 1 goto run

:run
echo.
echo Running Excel Replacer...
echo.

if not exist "target\testcasereplacer-1.0-SNAPSHOT.jar" (
    echo Error: JAR file not found. Building project first...
    goto build
)

:: Check if resources exist
if not exist "src\main\resources\config.properties" (
    echo Error: config.properties not found in src\main\resources
    echo Please ensure all required files are in place
    pause
    goto menu
)

:: Run the JAR file
java -jar target\testcasereplacer-1.0-SNAPSHOT.jar
echo.
echo Execution completed.
pause
goto menu

:end
echo.
echo Thank you for using Excel Replacer!
timeout /t 2 >nul
exit /b 0 