@echo off
REM ================================
REM Setup Virtual Environment
REM ================================

REM Check if .venv directory exists
if not exist .venv (
    echo Creating virtual environment...
    python -m venv .venv
)

REM Activate the virtual environment
call .venv\Scripts\activate

REM Upgrade pip and install dependencies
echo Upgrading pip...
pip install --upgrade pip

echo Installing dependencies...
pip install -e .[dev]

echo Setup complete. You can now run Vector_Issue using the command 'Vector_Issue' or 'Vector_Issue run'.