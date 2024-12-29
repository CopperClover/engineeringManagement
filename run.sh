#!/bin/bash

# Directory of the script
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"  # Gets the full path to the directory of this script
REPO_DIR="$SCRIPT_DIR"                      # Assumes the repo is in the same directory as the script
APP_SCRIPT="app.py"                         # Replace with the name of your Python application
VENV_DIR="$REPO_DIR/venv"                   # Virtual environment folder inside the repo

# Step 1: Navigate to the repository
echo "Navigating to repository..."
cd "$REPO_DIR" || { echo "Repository not found! Exiting."; exit 1; }

# Step 2: Pull the latest changes from the repo
echo "Pulling the latest updates..."
git reset --hard  # Resets any local changes
git pull origin main || { echo "Failed to pull the latest updates! Exiting."; exit 1; }

# Step 3: Remove the old virtual environment (if it exists)
if [ -d "$VENV_DIR" ]; then
    echo "Removing old virtual environment..."
    rm -rf "$VENV_DIR"
fi

# Step 4: Create a new virtual environment
echo "Creating a new virtual environment..."
python3 -m venv "$VENV_DIR" || { echo "Failed to create virtual environment! Exiting."; exit 1; }

# Step 5: Activate the virtual environment and install dependencies
echo "Activating virtual environment and installing dependencies..."
source "$VENV_DIR/bin/activate"
pip install --upgrade pip
pip install -r requirements.txt || { echo "Failed to install dependencies! Exiting."; deactivate; exit 1; }

# Step 6: Run the application
echo "Running the application..."
python "$APP_SCRIPT"

# Deactivate the virtual environment after execution
deactivate
