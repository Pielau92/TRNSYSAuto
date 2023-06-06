import multiprocessing
import sys
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED - needed as tkinter has compatibility issues with pywinauto    #todo: prüfen ob noch nötig
import classes
import os
import tkinter as tk
from tkinter import filedialog  #todo direkt über tk aufrufen

if __name__ == '__main__':

    """ For some unknown reason main is also affected by multiprocessing (directory is asked multiple times), therefore
    the function freeze_support is necessary
    """
    multiprocessing.freeze_support()

    # ask directory with base folder 'Basisordner' in it
    root = tk.Tk()
    root.withdraw()
    current_dir = filedialog.askdirectory()

    # create simulation series object
    sim_series = classes.SimulationSeries(current_dir,
                                          path_exe='C:\TRNSYS18\Exe\TrnEXE64.exe',
                                          name_excel_file='Simulationsvarianten.xlsx',
                                          name_excelsheet='Simulationsvarianten',
                                          name_base_folder='Basisordner',
                                          filename_dck_template='templateDck.dck',
                                          timeout=30*60,
                                          cpu_threshold=60,
                                          start_time_buffer=15,
                                          multiprocessing_max=3)

    # create new folder for simulation series
    os.makedirs(sim_series.path_sim_series)

    # import routine for input Excel file
    sim_series.import_input_excel()

    # sim_series.start_sim_series()       # start simulation - linear computing
    # sim_series.start_sim_series_par()   # start simulation - parallel computing
    sim_series.start_sim_series_par_fixed_amount()   # start simulation - parallel computing fixed amount of parallel simulations
