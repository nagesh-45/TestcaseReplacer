@echo off
echo Building Excel Replacer JAR...
echo.

:: Check if Maven is installed
mvn -version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Maven is not installed!
    echo Please install Maven from: https://maven.apache.org/download.cgi
    pause
    exit /b 1
)

:: Build the project
echo Building project...
mvn clean install

if errorlevel 1 (
    echo ERROR: Failed to build project!
    echo Please check the error messages above.
    pause
    exit /b 1
)

echo.
echo SUCCESS: JAR file created successfully!
echo Location: target\testcasereplacer-1.0-SNAPSHOT.jar
echo.
pause 