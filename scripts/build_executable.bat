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

cd %PROJECT_PATH%

echo Creating .exe file with PyInstaller...
%PYINSTALLER_PATH% %MAIN_PATH% --clean --onefile --add-binary "C:/Users/pierre/miniconda3/envs/TRNSYSAuto/DLLs/pyexpat.pyd;dlls" --add-binary "C:/Users/pierre/miniconda3/envs/TRNSYSAuto/Library/bin/libexpat.dll;." --add-binary "C:/Users/pierre/miniconda3/envs/TRNSYSAuto/Library/bin/libcrypto-3-x64.dll;." --add-binary "C:/Users/pierre/miniconda3/envs/TRNSYSAuto/Library/bin/libssl-3-x64.dll;." --add-binary "C:/Users/pierre/miniconda3/envs/TRNSYSAuto/Library/bin/liblzma.dll;." --add-binary "C:/Users/pierre/miniconda3/envs/TRNSYSAuto/Library/bin/LIBBZ2.dll;." --add-binary "C:/Users/pierre/miniconda3/envs/TRNSYSAuto/Library/bin/sqlite3.dll;." --add-binary "C:/Users/pierre/miniconda3/envs/TRNSYSAuto/Library/bin/tk86t.dll;." --add-binary "C:/Users/pierre/miniconda3/envs/TRNSYSAuto/Library/bin/tcl86t.dll;." --add-binary "C:/Users/pierre/miniconda3/envs/TRNSYSAuto/Library/bin/ffi.dll;." --add-binary "C:/Users/pierre/miniconda3/envs/TRNSYSAuto/Library/bin/libmpdec-4.dll;."

rem --collect-all xml.parsers.expat --hidden-import=pyexpat --collect-all pyexpat --hidden-import=openpyxl.cell._writer --hidden-import=win32timezone --hidden-import=tkinter 

echo Copying executable into %PROJECT_PATH%
xcopy "%PROJECT_PATH%\dist\main.exe" . /y

if %errorlevel% neq 0 (
    echo An error occured while creating executable .exe-file
    exit /b %errorlevel%
)

echo Executable created successfully!
endlocal
if "%1" neq "nopause" pause
