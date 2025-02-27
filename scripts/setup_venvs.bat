@echo off
setlocal

rem Load variables
call set_variables.bat

echo Activating %CONDA% prompt
call %CONDA_PATH%\Scripts\activate.bat

echo Creating virtual environments "%VENV_NAME_MAIN%" and "%VENV_NAME_TRNSYS%"
call conda create --name %VENV_NAME_MAIN% python -y
call conda create --name %VENV_NAME_TRNSYS% python -y

rem Change to parent directory, where the requirements Files are located
cd ..

echo Installing packages from requirements%VENV_NAME_MAIN%.txt
call conda activate %VENV_NAME_MAIN%
%PYTHON_PATH_MAIN% -m pip install -r requirements%VENV_NAME_MAIN%.txt || (echo Error while installing packages for %VENV_NAME_MAIN% & exit /b %errorlevel%)

call conda deactivate

echo Installing packages from requirements%VENV_NAME_TRNSYS%.txt
call conda activate %VENV_NAME_TRNSYS%
%PYTHON_PATH_TRNSYS% -m pip install -r requirements%VENV_NAME_TRNSYS%.txt || (echo Error while installing packages for %VENV_NAME_TRNSYS%  & exit /b %errorlevel%)

echo Virtual environments %VENV_NAME_MAIN% and %VENV_NAME_TRNSYS% set up successfully!
endlocal
pause