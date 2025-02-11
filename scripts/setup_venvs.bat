@echo off
setlocal

rem -------------------- BEFORE YOU START --------------------

rem Are you using miniconda3 or anaconda3?
set CONDA=miniconda3

rem Name of your virtual environment (for main modules)
set VENV_NAME_MAIN=TRNSYSAuto

rem Name of your virtual environment (for TRNSYS/python modules)
set VENV_NAME_TRNSYS=TRNSYS

rem ----------------------------------------------------------

echo Activating %CONDA% prompt
call "%USERPROFILE%\%CONDA%\Scripts\activate.bat"

echo Creating virtual environments "%VENV_NAME_MAIN%" and "%VENV_NAME_TRNSYS%"
call conda create --name %VENV_NAME_MAIN% python -y
call conda create --name %VENV_NAME_TRNSYS% python -y

rem Change to parent directory, where the requirements Files are located
cd ..

echo Installing packages from requirements%VENV_NAME%.txt
call conda activate %VENV_NAME_MAIN%
pip install -r requirements%VENV_NAME_MAIN%.txt
call conda deactivate
call conda activate %VENV_NAME_TRNSYS%
pip install -r requirements%VENV_NAME_TRNSYS%.txt

echo Virtual environments %VENV_NAME_MAIN% and %VENV_NAME_TRNSYS% set up successfully!
endlocal
pause