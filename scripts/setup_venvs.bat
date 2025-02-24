@echo off
setlocal

rem Load variables
call set_variables.bat

echo Activating %CONDA% prompt
call "%USERPROFILE%\%CONDA%\Scripts\activate.bat"

echo Creating virtual environments "%VENV_NAME_MAIN%" and "%VENV_NAME_TRNSYS%"
call conda create --name %VENV_NAME_MAIN% python -y
call conda create --name %VENV_NAME_TRNSYS% python -y

rem Change to parent directory, where the requirements Files are located
cd ..

echo Installing packages from requirements%VENV_NAME%.txt
call conda activate %VENV_NAME_MAIN%
%PYTHON_PATH% -m pip install -r requirements%VENV_NAME_MAIN%.txt
call conda deactivate
call conda activate %VENV_NAME_TRNSYS%
%PYTHON_PATH% -m pip install -r requirements%VENV_NAME_TRNSYS%.txt

echo Virtual environments %VENV_NAME_MAIN% and %VENV_NAME_TRNSYS% set up successfully!
endlocal
pause