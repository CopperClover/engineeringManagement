@echo off

REM Get the directory of the script
set "SCRIPT_DIR=%~dp0"                 REM Directory of the batch file
set "REPO_DIR=%SCRIPT_DIR%"            REM Repository directory (same as script directory)
set "APP_SCRIPT=main.py"               REM Replace with the name of your Python application
set "VENV_DIR=%REPO_DIR%venv"          REM Virtual environment folder inside the repo

REM Step 1: Navigate to the repository
echo Navigating to repository...
cd /d "%REPO_DIR%"
if errorlevel 1 (
    echo Repository not found! Exiting.
    exit /b 1
)

REM Step 2: Pull the latest changes from the repo
echo Pulling the latest updates...
git reset --hard
git pull origin main
if errorlevel 1 (
    echo Failed to pull the latest updates! Exiting.
    exit /b 1
)

REM Step 3: Remove the old virtual environment (if it exists)
if exist "%VENV_DIR%" (
    echo Removing old virtual environment...
    rmdir /s /q "%VENV_DIR%"
)

REM Step 4: Create a new virtual environment
echo Creating a new virtual environment...
python -m venv "%VENV_DIR%"
if errorlevel 1 (
    echo Failed to create virtual environment! Exiting.
    exit /b 1
)

REM Step 5: Activate the virtual environment and install dependencies
echo Activating virtual environment and installing dependencies...
call "%VENV_DIR%\Scripts\activate"
pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install dependencies! Exiting.
    call deactivate
    exit /b 1
)

REM Step 6: Run the application
echo Running the application...
python "%APP_SCRIPT%"
call deactivate
