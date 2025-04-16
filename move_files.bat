@echo off
setlocal enabledelayedexpansion

:: Set paths
set "PROJECT_ROOT=%~dp0"
set "RESOURCES_DIR=%PROJECT_ROOT%src\main\resources"
set "FILES_DIR=%PROJECT_ROOT%Files"

:: Create Files directory if it doesn't exist
if not exist "%FILES_DIR%" mkdir "%FILES_DIR%"

:: Copy files from resources to Files folder
echo Moving files...
copy "%RESOURCES_DIR%\IDPE Onus Tcs.xlsm" "%FILES_DIR%"
copy "%RESOURCES_DIR%\config.properties" "%FILES_DIR%"

echo.
echo SUCCESS: Files moved successfully!
echo Files are now in: %FILES_DIR%
echo.
pause 