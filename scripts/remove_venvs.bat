@echo off
setlocal

rem Load variables
call set_variables.bat

echo Activating %CONDA% prompt
call "%USERPROFILE%\%CONDA%\Scripts\activate.bat"

echo Removing virtual environments "%VENV_NAME_MAIN%" and "%VENV_NAME_TRNSYS%"

call conda remove --name %VENV_NAME_MAIN% --all %VENV_NAME_MAIN% -y || (echo Error while removing %VENV_NAME_MAIN% & exit /b %errorlevel%)


call conda remove --name %VENV_NAME_TRNSYS% --all %VENV_NAME_TRNSYS% -y || (echo Error while removing %VENV_NAME_TRNSYS% & exit /b %errorlevel%)
if %errorlevel% neq 0 (
    echo An error occured while creating installation .setup-file
    exit /b %errorlevel%
)

echo Virtual environments %VENV_NAME_MAIN% and %VENV_NAME_TRNSYS% removed successfully!
endlocal
pause