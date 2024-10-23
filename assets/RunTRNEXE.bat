:: Batch file to run the TRNSYS Studio and use the Python (CFFI) interface with a miniconda environment
:: ----------------------------------------------------------------------------------------------------
:: Edit the %condaenvname% variable (set to TRNSYS in the example file)
:: If TRNSYS is not installed in C:\TRNSYS18, edit the path to the Studio accordingly
::
:: Set the name of the conda environment to be used (should have cffi and numpy installed at the minimum, edit if required)
set condaenvname=TRNSYS
::
:: Set required environment variables for the conda environment to be found and used by the TRNSYS Python Interface 
::
:: Add directory with python to the path (to the front of the path!)
set path=C:\Users\%username%\miniconda3\condabin;%path%
set path=C:\Users\%username%\miniconda3\envs\%condaenvname%;%path%
set path=C:\Users\%username%\miniconda3\envs\%condaenvname%\bin;%path%
set path=C:\Users\%username%\miniconda3\envs\%condaenvname%\Library\mingw-w64\bin;%path%
set path=C:\Users\%username%\miniconda3\envs\%condaenvname%\Library\bin;%path%
set path=C:\Users\%username%\miniconda3\envs\%condaenvname%\Library\usr\bin;%path%
set path=C:\Users\%username%\miniconda3\envs\%condaenvname%\Scripts;%path%
:: Set PYTHONHOME to the same directory 
set PYTHONHOME=C:\Users\%username%\miniconda3\envs\%condaenvname%
:: Set PYTHONPATH to the site-packages directory (which is within your environment\Lib)
set PYTHONPATH=C:\Users\%username%\miniconda3\envs\%condaenvname%\Lib\site-packages
::
:: Launch TRNEXE
::
C:\TRNSYS18\Exe\TrnEXE64.exe
