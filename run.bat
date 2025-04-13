@echo off
setlocal

:: Get the directory where the script is located
set "SCRIPT_DIR=%~dp0"

:: Check if Java is installed
java -version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Java is not installed. Please install Java first.
    pause
    exit /b 1
)

:: Check if Maven is installed
mvn -version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Maven is not installed. Please install Maven first.
    pause
    exit /b 1
)

:: Create desktop shortcut if it doesn't exist
if not exist "%USERPROFILE%\Desktop\Run Excel Replacer.bat" (
    echo @echo off > "%USERPROFILE%\Desktop\Run Excel Replacer.bat"
    echo cd /d "%~dp0" >> "%USERPROFILE%\Desktop\Run Excel Replacer.bat"
    echo call run.bat >> "%USERPROFILE%\Desktop\Run Excel Replacer.bat"
    echo pause >> "%USERPROFILE%\Desktop\Run Excel Replacer.bat"
    echo Created desktop shortcut
)

:: Run the program
echo Starting Excel Replacer...
cd /d "%SCRIPT_DIR%"
call mvn exec:java -Dexec.mainClass="com.example.ExcelReplacer"

:: Check if the program ran successfully
if %ERRORLEVEL% EQU 0 (
    echo Program completed successfully!
    :: Open the output folder
    start "" "%SCRIPT_DIR%\output"
) else (
    echo Program failed to run. Please check the error messages above.
)

pause 