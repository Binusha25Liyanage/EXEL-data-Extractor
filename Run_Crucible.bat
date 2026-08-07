@echo off
REM Double-click this file to launch Crucible.
REM Keep this .bat file in the SAME FOLDER as report_builder.py.

cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python was not found on your PATH.
    echo Install it from https://python.org and tick "Add python.exe to PATH".
    pause
    exit /b 1
)

python report_builder.py
if %errorlevel% neq 0 (
    echo.
    echo Crucible closed with an error - see the message above.
    pause
)
