@echo off
setlocal

set "PROJECT=%~dp0"
set "VENV=%PROJECT%venv"
set "PYTHON=%VENV%\Scripts\python.exe"
set "REQS=%PROJECT%requirements.txt"
set "APP=%PROJECT%src\app.py"

:: ── Check Python is installed ────────────────────────────────────────────────
where python >nul 2>&1
if errorlevel 1 (
    echo Python not found.
    echo Please install Python 3.11 or later from https://www.python.org/downloads/
    echo Make sure to tick "Add Python to PATH" during installation.
    pause
    exit /b 1
)

:: ── Create venv if missing ───────────────────────────────────────────────────
if not exist "%PYTHON%" (
    echo Setting up Pdf-Mithra for the first time...
    python -m venv "%VENV%"
    if errorlevel 1 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo Installing dependencies ^(this takes about a minute^)...
    "%VENV%\Scripts\pip" install --quiet -r "%REQS%"
    if errorlevel 1 (
        echo Failed to install dependencies.
        pause
        exit /b 1
    )
    echo Setup complete!
)

:: ── Launch the app ───────────────────────────────────────────────────────────
start "" "%PYTHON%" "%APP%"
