@echo off
echo Building Excel Replacer Installer...
echo.

:: Build the project
call mvn clean install

:: Create the installer
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

echo.
echo Installer created in the 'installer' folder
pause 