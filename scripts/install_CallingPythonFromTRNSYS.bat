@echo off
setlocal

rem Load variables
call set_variables.bat


echo Copying TRNSYS files into %TRNSYS_PATH%
xcopy %PROJECT_PATH%\assets\CallingPython-Cffi %TRNSYS_PATH%\ /e /i || (exit /b %errorlevel%)
echo TRNSYS files copied successfully!

echo Rebuilding python interface
cd C:\Trnsys18\TRNLib\CallingPython-Cffi\PythonInterface
%PYTHON_PATH_TRNSYS% TrnsysPythonInterfaceBuilder.py

echo Testing python interface
PythonInterfaceTesterWithCondaEnvironment.bat

endlocal
pause