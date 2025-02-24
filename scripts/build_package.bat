@echo off
setlocal

rem Load variables
call set_variables.bat

echo Building package...
cd %PROJECT_PATH%
%PYTHON_PATH_MAIN% setup.py bdist_wheel sdist

if %errorlevel% neq 0 (
    echo An error occured while building package
    exit /b %errorlevel%
)

echo Package built successfully!
endlocal
pause