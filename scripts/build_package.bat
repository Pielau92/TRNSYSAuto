@echo off
setlocal

rem -------------------- BEFORE YOU START --------------------

rem Are you using miniconda3 or anaconda3?
set CONDA=miniconda3

rem Name of your virtual environment
set VENV_NAME=TRNSYSAuto

rem Do not forget to "pip install wheel setuptools twine" beforehand!!

rem ----------------------------------------------------------

rem Path variables
for %%i in ("%CD%\..") do set PROJECT_PATH=%%~fi
set PYTHON_PATH="%USERPROFILE%\%CONDA%\envs\%VENV_NAME%\python.exe"

rem Build package
echo Building package...
cd %PROJECT_PATH%
%PYTHON_PATH% setup.py bdist_wheel sdist

if %errorlevel% neq 0 (
    echo An error occured while building package
    exit /b %errorlevel%
)

echo Package built successfully!
endlocal
pause