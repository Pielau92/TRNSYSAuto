import multiprocessing
import sys
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED - needed as tkinter has compatibility issues with pywinauto    #todo: prüfen ob noch nötig
import classes
import functions
import os
import tkinter as tk
from tkinter import filedialog  #todo direkt über tk aufrufen

if __name__ == '__main__':

    """ For some unknown reason main is also affected by multiprocessing (directory is asked multiple times), therefore
    the function freeze_support is necessary
    """
    multiprocessing.freeze_support()

    # create simulation series object
    sim_series = classes.SimulationSeries()

    # ask simulation variants Excel file path
    root = tk.Tk()
    root.withdraw()
    sim_series.path_sim_variants_excel = filedialog.askopenfilename().replace("/", "\\")
    sim_series.set_paths()

    # import and apply settings Excel file
    sim_series.import_settings_excel()
    sim_series.set_settings()

    # create new folder for simulation series
    os.makedirs(sim_series.dir_sim_series)

    # import routine for input Excel file
    sim_series.import_input_excel()

    # sim_series.start_sim_series()       # start simulation - linear computing
    # sim_series.start_sim_series_par()   # start simulation - parallel computing
    sim_series.start_sim_series_par_fixed_amount()   # start simulation - parallel computing fixed amount of parallel simulations
