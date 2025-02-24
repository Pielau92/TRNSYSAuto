@echo off
setlocal

rem Load variables
call set_variables.bat

rem Build executable .exe-file
call build_executable.bat nopause

rem update MPC modules in dist\assets\mpc_code directory
xcopy "%PROJECT_PATH%\assets\mpc_code\main_mpc.py" "%PROJECT_PATH%\dist\assets\mpc_code\" /y
xcopy "%PROJECT_PATH%\assets\mpc_code\MPCModule.py" "%PROJECT_PATH%\dist\assets\mpc_code\" /y
xcopy "%PROJECT_PATH%\assets\mpc_code\settingsMPC.ini" "%PROJECT_PATH%\dist\assets\mpc_code\" /y
xcopy "%PROJECT_PATH%\settings.ini" "%PROJECT_PATH%\dist\settings.ini"/y
xcopy %PROJECT_PATH%\scripts %PROJECT_PATH%\dist\scripts /i /y

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
