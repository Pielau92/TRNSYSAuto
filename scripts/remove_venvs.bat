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

echo Removing virtual environments "%VENV_NAME_MAIN%" and "%VENV_NAME_TRNSYS%"
call conda remove --name %VENV_NAME_MAIN% --all %VENV_NAME_MAIN% -y
call conda remove --name %VENV_NAME_TRNSYS% --all %VENV_NAME_TRNSYS% -y

echo Virtual environments %VENV_NAME_MAIN% and %VENV_NAME_TRNSYS% removed successfully!
endlocal
pause