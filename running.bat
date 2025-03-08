@echo off
setlocal

:: Define environment and script variables
set VENV_DIR=venv
set REQUIREMENTS_FILE=requirements.txt
set PYTHON_SCRIPT=Traditional_Methods_Bakri.py

:: Check if Python is installed
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python is not installed. Please install Python and try again.
    pause
    exit /b
)

:: Create virtual environment if it doesn't exist
if not exist %VENV_DIR% (
    echo Creating virtual environment...
    python -m venv %VENV_DIR%
)

:: Activate virtual environment
call %VENV_DIR%\Scripts\activate

:: Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip

:: Install dependencies
if exist %REQUIREMENTS_FILE% (
    echo Installing dependencies from %REQUIREMENTS_FILE%...
    pip install -r %REQUIREMENTS_FILE%
) else (
    echo Requirements file not found. Skipping dependencies installation.
)

:: Run the Python script
if exist %PYTHON_SCRIPT% (
    echo Running %PYTHON_SCRIPT%...
    python %PYTHON_SCRIPT%
) else (
    echo Python script not found. Please make sure %PYTHON_SCRIPT% exists.
)

:: Deactivate virtual environment
deactivate

echo Done.
pause