:: SETTINGS :: 
@echo off


:: ADMIN ::
net session >nul 2>&1 || ( PowerShell -Command "Start-Process \"%~f0\" -Verb RunAs" & exit /B )


:: MAIN ::
PowerShell -ExecutionPolicy "Bypass" -NoProfile -WindowStyle "Maximized" -File "%~dp0SRC\MAIN.ps1"