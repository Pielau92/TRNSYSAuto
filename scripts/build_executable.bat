@echo off
setlocal

rem Load variables
call set_variables.bat

rem Change directory
cd %PROJECT_PATH%

echo Creating .exe file with PyInstaller...
%PYINSTALLER_PATH% %MAIN_PATH% --clean --onefile --add-binary "%VENV_PATH_MAIN%/DLLs/pyexpat.pyd;dlls" --add-binary "%VENV_PATH_MAIN%/Library/bin/libexpat.dll;." --add-binary "%VENV_PATH_MAIN%/Library/bin/libcrypto-3-x64.dll;." --add-binary "%VENV_PATH_MAIN%/Library/bin/libssl-3-x64.dll;." --add-binary "%VENV_PATH_MAIN%/Library/bin/liblzma.dll;." --add-binary "%VENV_PATH_MAIN%/Library/bin/LIBBZ2.dll;." --add-binary "%VENV_PATH_MAIN%/Library/bin/sqlite3.dll;." --add-binary "%VENV_PATH_MAIN%/Library/bin/tk86t.dll;." --add-binary "%VENV_PATH_MAIN%/Library/bin/tcl86t.dll;." --add-binary "%VENV_PATH_MAIN%/Library/bin/ffi.dll;." --add-binary "%VENV_PATH_MAIN%/Library/bin/libmpdec-4.dll;."

rem Check for error
if %errorlevel% neq 0 (
    echo An error occured while creating executable .exe-file
    exit /b %errorlevel%
)

rem --collect-all xml.parsers.expat --hidden-import=pyexpat --collect-all pyexpat --hidden-import=openpyxl.cell._writer --hidden-import=win32timezone --hidden-import=tkinter 

rem echo Copying executable into %SRC_PATH%\%PROJECT_NAME%
rem xcopy %PROJECT_PATH%\dist\main.exe %SRC_PATH%\%PROJECT_NAME% /y || (exit /b %errorlevel%)

echo Executable created successfully!
endlocal

rem Do not pause if batch file was called by an other batch file
if "%1" neq "nopause" pause
