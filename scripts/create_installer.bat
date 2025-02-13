@echo off
setlocal

rem -------------------- BEFORE YOU START --------------------

rem Name of Inno Setup Script
set INNO_SETUP_SCRIPT_NAME=TRNSYSAuto.iss

rem Name of Project
set PROJECT_NAME=TRNSYSAuto

rem ----------------------------------------------------------

rem Path variables
for %%i in ("%CD%\..") do set PROJECT_PATH=%%~fi
set INNO_SETUP_EXE_PATH="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
set INNO_SETUP_SCRIPT_PATH="%CD%\%INNO_SETUP_SCRIPT_NAME%"
set PROJECT_PATH=%USERPROFILE%\PycharmProjects\%PROJECT_NAME%

rem Build executable .exe-file
call build_executable.bat nopause

rem update MPC modules in dist\assets\mpc_code directory
xcopy "%PROJECT_PATH%\assets\mpc_code\main_mpc.py" "%PROJECT_PATH%\dist\assets\mpc_code\" /y
xcopy "%PROJECT_PATH%\assets\mpc_code\MPCModule.py" "%PROJECT_PATH%\dist\assets\mpc_code\" /y
xcopy "%PROJECT_PATH%\assets\mpc_code\settingsMPC.ini" "%PROJECT_PATH%\dist\assets\mpc_code\" /y
xcopy "%PROJECT_PATH%\settings.ini" "%PROJECT_PATH%\dist\settings.ini"/y

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
