
@REM ================================
@REM Prepare workspace
@REM ================================

@REM This file supports you to setup the SmkTool environment
@REM The only prerequisite is that a Python installation exists

@REM __author__ = 'Dominik Schubert'
@REM __copyright__ = 'Copyright 2023, Marquardt'
@REM __credits__ = []
@REM __license__ = 'Marquardt'
@REM __version__ = '0.0.0'
@REM __maintainer__ = 'Dominik Schubert'
@REM __email__ = 'dominik.schubert@marquardt.com'
@REM __status__ = 'Development'

@ECHO OFF

@REM Install the virtualenv Package
pip install virtualenv

@REM Create the virtual environment
@REM The env shall be in the Project root so vscode can find the environment
virtualenv .venv

@REM Activate the virtual environment
call .\.venv\Scripts\activate.bat

@REM Install the SmkTool working directory into the virtual environment
@REM The Tool does include all dependencies
pip install -e .[dev]

@REM Activate the virtual environment
call .\.venv\Scripts\deactivate.bat