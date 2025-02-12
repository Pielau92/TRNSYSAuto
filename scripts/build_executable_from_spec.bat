@echo off
setlocal

rem -------------------- BEFORE YOU START --------------------

rem Are you using miniconda3 or anaconda3?
set CONDA=miniconda3

rem Name of your virtual environment
set VENV_NAME=TRNSYSAuto

rem Name of spec file
set SPEC_NAME=main.spec

rem ----------------------------------------------------------

rem Path variables
set VENV_PATH=%USERPROFILE%\%CONDA%\envs\%VENV_NAME%
set PYINSTALLER_PATH=%VENV_PATH%\Scripts\pyinstaller.exe

rem Change to parent directory
cd ..

echo Creating .exe file with PyInstaller...
%PYINSTALLER_PATH% scripts\%SPEC_NAME%

echo Copying executable into %CD%
xcopy "dist\main.exe" . /y

if %errorlevel% neq 0 (
    echo An error occured while creating executable .exe-file
    exit /b %errorlevel%
)

echo Executable created successfully!
endlocal
pause
