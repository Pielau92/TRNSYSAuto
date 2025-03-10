@echo off
setlocal

echo Downloading miniconda
curl https://repo.anaconda.com/miniconda/Miniconda3-latest-Windows-x86_64.exe --output %USERPROFILE%\Downloads\Miniconda3-latest-Windows-x86_64.exe

echo Starting miniconda installer
%USERPROFILE%\Downloads\Miniconda3-latest-Windows-x86_64.exe

endlocal
pause