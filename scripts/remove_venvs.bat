@echo off
setlocal

rem Load variables
call set_variables.bat

echo Activating %CONDA% prompt
call "%USERPROFILE%\%CONDA%\Scripts\activate.bat"

echo Removing virtual environments "%VENV_NAME_MAIN%" and "%VENV_NAME_TRNSYS%"
call conda remove --name %VENV_NAME_MAIN% --all %VENV_NAME_MAIN% -y
call conda remove --name %VENV_NAME_TRNSYS% --all %VENV_NAME_TRNSYS% -y

echo Virtual environments %VENV_NAME_MAIN% and %VENV_NAME_TRNSYS% removed successfully!
endlocal
pause