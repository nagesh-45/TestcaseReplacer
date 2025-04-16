@echo off
setlocal enabledelayedexpansion

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "DIST_DIR=%PROJECT_ROOT%dist"
set "FILES_DIR=%DIST_DIR%\Files"
set "TARGET_DIR=%DIST_DIR%\target"
set "OUTPUT_DIR=%DIST_DIR%\Output"

:: Create distribution directory structure
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
mkdir "%DIST_DIR%"
mkdir "%FILES_DIR%"
mkdir "%TARGET_DIR%"
mkdir "%OUTPUT_DIR%"

:: Copy required files
echo Copying files...
copy "%PROJECT_ROOT%run_excel_replacer.bat" "%DIST_DIR%"
copy "%PROJECT_ROOT%target\testcasereplacer-1.0-SNAPSHOT.jar" "%TARGET_DIR%"
copy "%PROJECT_ROOT%Files\IDPE Onus Tcs.xlsm" "%FILES_DIR%"
copy "%PROJECT_ROOT%Files\config.properties" "%FILES_DIR%"

:: Create README
echo Creating README...
(
echo Excel Replacer Tool
echo ==================
echo.
echo Folder Structure:
echo ExcelReplacer/
echo ├── run_excel_replacer.bat    (Run this file to start the program)
echo ├── README.txt               (This file)
echo ├── Files/
echo │   ├── IDPE Onus Tcs.xlsm   (Excel template)
echo │   └── config.properties    (Configuration file)
echo ├── target/
echo │   └── testcasereplacer-1.0-SNAPSHOT.jar  (Java program)
echo └── Output/                  (Output files will appear here)
echo     └── IDPE Onus Tcs_[timestamp].xlsm
echo.
echo Requirements:
echo - Windows operating system
echo - Java 8 or later (Download from: https://www.java.com/download/)
echo.
echo Installation:
echo 1. Install Java if not already installed
echo    - Download from: https://www.java.com/download/
echo    - Run the installer
echo    - Restart your computer if prompted
echo.
echo 2. Extract all files to a folder
echo    - Keep the folder structure as shown above
echo    - Do not move or rename any files
echo.
echo Usage:
echo 1. Double-click run_excel_replacer.bat
echo 2. The program will process the Excel file
echo 3. Check the Output folder for results
echo.
echo Troubleshooting:
echo If you get an error about Java not being found:
echo 1. Make sure Java is installed
echo 2. Try restarting your computer
echo 3. Contact support if the problem persists
) > "%DIST_DIR%\README.txt"

:: Create zip file
echo Creating distribution package...
powershell Compress-Archive -Path "%DIST_DIR%\*" -DestinationPath "%PROJECT_ROOT%ExcelReplacer.zip" -Force

echo.
echo SUCCESS: Distribution package created!
echo Location: ExcelReplacer.zip
echo.
echo To use on another machine:
echo 1. Extract ExcelReplacer.zip
echo 2. Install Java 8 or later
echo 3. Run run_excel_replacer.bat
echo.
pause 