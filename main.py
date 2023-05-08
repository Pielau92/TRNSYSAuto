import sys
# sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED - needed as tkinter has compatibility issues with pywinauto
import os
import shutil
import functions
import tkinter as tk

from datetime import datetime
from tkinter import filedialog

current_time = datetime.now().strftime("%d.%m.%Y_%H.%M")    # current time

# region paths and directories

# ask directory with base folder "Basisordner" in it
root = tk.Tk()
root.withdraw()
current_dir = filedialog.askdirectory()
# current_dir = os.getcwd()  # current directory

path_sim_series = os.path.join(current_dir, current_time)   # path of simulation series folder
path_base = os.path.join(current_dir, 'Basisordner')        # path of base folder (templates, load profiles, weather...)
path_exe = 'C:\TRNSYS18\Exe\TrnEXE64.exe'                   # path of the simulation .exe file

os.makedirs(path_sim_series)  # create new directory for simulation series

# endregion

# import Excel file
path_excel = os.path.join(path_base, 'Simulationsvarianten.xlsx')
sim_list, weather_series, df_dck, df_b18 = functions.import_input_excel(path_excel)

for sim in sim_list:

    # region create simulation folders

    path_sim = os.path.join(path_sim_series, sim)   # path of simulation folder
    os.makedirs(path_sim)                           # create new directory for simulation
    # shutil.copytree(path_base, path_sim)          # copy all files from base to simulation folder

    # list of specified file names to be copied
    file_names = ["Building.b18", "Building.dck", "Lastprofil.txt", weather_series[sim] + '.txt']

    # copy specified files into simulation folder
    for file_name in file_names:
        src_file = os.path.join(path_base, file_name)  # source path
        dst_file = os.path.join(path_sim, file_name)  # destination path
        shutil.copy(src_file, dst_file)  # copy file

    functions.find_and_replace_param(path_sim + '/Building.b18', df_b18.loc[sim])  # replace parameters in .b18 File
    functions.find_and_replace_param(path_sim + '/Building.dck', df_dck.loc[sim])  # replace parameters in .dck File

    # endregion

    # perform simulation
    # functions.start_sim(path_exe, os.path.join(path_sim, 'Building.dck'))







# Output-Files (.txt) in eine Excel-Datei (vorbereitetes Auswertungs-Excel) laden
# ?
