import os
import shutil
import functions

from datetime import datetime

current_time = datetime.now().strftime("%d.%m.%Y_%H.%M")    # current time

# paths and directories
current_dir = os.getcwd()                               # current directory
path_sim_series = current_dir + "./" + current_time     # path of simulation series folder
path_base = current_dir + "./Basisordner"               # path of base folder (templates, load profiles, weather...)

os.makedirs(path_sim_series)                            # create new directory for simulation series

# import Excel file
path_excel = path_base + '/Simulationsvarianten.xlsx'
sim_list, weather_series, parameters = functions.import_input_excel(path_excel)

for sim in sim_list:
    path_sim = path_sim_series + "./" + sim  # path of simulation folder
    shutil.copytree(path_base, path_sim)  # copy files from base to simulation folder

    functions.find_and_replace_param(path_sim + '/Building.b18', parameters.loc[sim])  # replace parameters in .b18 File
    functions.find_and_replace_param(path_sim + '/Building.dck', parameters.loc[sim])  # replace parameters in .dck File
