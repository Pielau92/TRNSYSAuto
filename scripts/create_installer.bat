@echo off
setlocal

rem Load variables
call set_variables.bat

rem Build executable .exe-file
call build_executable.bat nopause

rem Copy files
xcopy %PROJECT_PATH%\assets\settingsMPC.ini %PROJECT_PATH%\dist\assets\ /y
xcopy %PROJECT_PATH%\configs.ini %PROJECT_PATH%\dist\ /y
xcopy %PROJECT_PATH%\requirements%VENV_NAME_MAIN%.txt %PROJECT_PATH%\dist\ /y
xcopy %PROJECT_PATH%\requirements%VENV_NAME_TRNSYS%.txt %PROJECT_PATH%\dist\ /y
xcopy %PROJECT_PATH%\misc\Installationsanleitung.pdf %PROJECT_PATH%\dist\ /y
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
