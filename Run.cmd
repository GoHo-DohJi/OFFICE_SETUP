:: SETTINGS :: 
@echo off


:: ADMIN CHECK ::
net session >nul 2>&1 || ( PowerShell -Command "Start-Process \"%~f0\" -Verb RunAs" & exit /B )


:: VARIABLES ::
set "REPO=OFFICE_SETUP"
set "URL=https://github.com/GoHo-DohJi/%REPO%/archive/refs/heads/main.zip"
set "TEMP_DIR=%TEMP%"
set "ZIP=%TEMP_DIR%\%REPO%.zip"
set "SCRIPT=%TEMP_DIR%\%REPO%-main\SRC\MAIN.ps1"


:: DOWNLOAD ::
cls
echo [*] DOWNLOADING...
curl.exe -sSL "%URL%" -o "%ZIP%"

:: EXTRACT ::
echo [*] EXTRACTING...
tar -xf "%ZIP%" -C "%TEMP_DIR%"
del /F /Q "%ZIP%"
