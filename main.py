import sys
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED - needed as tkinter has compatibility issues with pywinauto
import os
import shutil
import functions
import tkinter as tk
import re

from datetime import datetime
from tkinter import filedialog

current_time = datetime.now().strftime('%d.%m.%Y_%H.%M')    # current time

# region paths and directories

# ask directory with base folder 'Basisordner' in it
root = tk.Tk()
root.withdraw()
current_dir = filedialog.askdirectory()
# current_dir = os.getcwd()  # current directory

path_sim_series = os.path.join(current_dir, current_time)   # path of simulation series folder
path_base = os.path.join(current_dir, 'Basisordner')        # path of base folder (templates, load profiles, weather...)
path_exe = 'C:\TRNSYS18\Exe\TrnEXE64.exe'                   # path of the simulation .exe file
path_excel = os.path.join(path_base, 'Simulationsvarianten.xlsx')

os.makedirs(path_sim_series)  # create new directory for simulation series

# endregion

# import Excel file
sim_list, weather_series, df_dck, b18_series = functions.import_input_excel(path_excel)

for sim in sim_list:

    # region create simulation folders

    path_sim = os.path.join(path_sim_series, sim)   # path of simulation folder
    os.makedirs(path_sim)                           # create new directory for simulation
    # shutil.copytree(path_base, path_sim)          # copy all files from base to simulation folder

    # region source/destination file paths for copying process
    src_file = [
        os.path.join(path_base, 'templateDck.dck'),
        os.path.join(path_base, 'Lastprofil.txt'),
        os.path.join(path_base, 'b18', b18_series[sim]),
        os.path.join(path_base, 'Wetterdaten', weather_series[sim])]

    dst_file = [
        os.path.join(path_sim, 'templateDck.dck'),
        os.path.join(path_sim, 'Lastprofil.txt'),
        os.path.join(path_sim, b18_series[sim]),
        os.path.join(path_sim, weather_series[sim])]

    # endregion

    # copy specified files into simulation folder
    for index in range(len(src_file)):
        shutil.copy(src_file[index], dst_file[index])  # copy file

    # replace parameters in .dck File
    functions.find_and_replace_param(dst_file[0], r'(@\w+)\s*=\s*([\d.]+)', df_dck.loc[sim])

    # region find and replace .b18/.b17 file name

    with open(dst_file[0], 'r') as file:
        text = file.read()  # read file
        new_text = re.sub(r'(\*ASSIGN "b17")', r'ASSIGN "' + b18_series[sim] + '"', text)
    with open(dst_file[0], 'w') as file:
        file.write(new_text)  # overwrite file

    # endregion

    # region find and replace weather data file name

    with open(dst_file[0], 'r') as file:
        text = file.read()  # read file
        new_text = re.sub(r'(\*ASSIGN "tm2")', r'ASSIGN "' + weather_series[sim] + '"', text)
    with open(dst_file[0], 'w') as file:
        file.write(new_text)  # overwrite file

    # endregion

    functions.find_and_replace_param(dst_file[0], r'(\*ASSIGN "b17")', df_dck.loc[sim])

    # new_text = re.sub(pattern, r'* ASSIGN "replacement"', text)

    # endregion

    # perform simulation
    functions.start_sim(path_exe, os.path.join(path_sim, dst_file[0]))

# text = re.sub(r'(*\w+)\s*=\s*([\d.]+)',





# Output-Files (.txt) in eine Excel-Datei (vorbereitetes Auswertungs-Excel) laden
# ?
