@echo off
setlocal

rem -------------------- BEFORE YOU START --------------------

rem Name of Inno Setup Script
set INNO_SETUP_SCRIPT_NAME=TRNSYSAuto.iss

rem ----------------------------------------------------------

rem Path variables
for %%i in ("%CD%\..") do set PROJECT_PATH=%%~fi
set INNO_SETUP_EXE_PATH="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
set INNO_SETUP_SCRIPT_PATH="%CD%\%INNO_SETUP_SCRIPT_NAME%"

rem Build executable .exe-file
rem build_executable.bat

if %errorlevel% neq 0 (
    echo An error occured while creating executable .exe-file, setup aborted!
    exit /b %errorlevel%
)

rem Create installation .setup-file
echo Creating installation .setup-file with Inno Setup...

cd %PROJECT_PATH%
%INNO_SETUP_EXE_PATH% %INNO_SETUP_SCRIPT_PATH%

if %errorlevel% neq 0 (
    echo An error occured while creating installation .setup-file
    exit /b %errorlevel%
)

echo installation .setup file created successfully!
endlocal
pause
