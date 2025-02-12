@echo off
setlocal

rem -------------------- BEFORE YOU START --------------------

rem Are you using miniconda3 or anaconda3?
set CONDA=miniconda3

rem Name of your virtual environment
set VENV_NAME=TRNSYSAuto

rem ----------------------------------------------------------

rem Path variables
set VENV_PATH=%USERPROFILE%\%CONDA%\envs\%VENV_NAME%
set PYINSTALLER_PATH=%VENV_PATH%\Scripts\pyinstaller.exe
for %%i in ("%CD%\..") do set PROJECT_PATH=%%~fi
for %%i in ("%CD%\..") do set PROJECT_NAME=%%~nxi
set MAIN_PATH=%PROJECT_PATH%\%PROJECT_NAME%\main.py

rem call PyInstaller to create executable
echo Creating .exe file with PyInstaller...
cd %PROJECT_PATH%
%PYINSTALLER_PATH% %MAIN_PATH% --clean --onefile --add-binary "C:/Users/pierre/miniconda3/DLLs/pyexpat.pyd;dlls"
rem --hidden-import=openpyxl.cell._writer --hidden-import=win32timezone --hidden-import=tkinter --hidden-import=pyexpat  --collect-all xml.parsers.expat

if %errorlevel% neq 0 (
    echo An error occured while creating executable .exe-file
    exit /b %errorlevel%
)

echo Executable created successfully!
endlocal
pause
