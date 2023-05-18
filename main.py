import sys
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED - needed as tkinter has compatibility issues with pywinauto

import classes
import os

import tkinter as tk
from tkinter import filedialog

if __name__ == '__main__':
    # ask directory with base folder 'Basisordner' in it
    root = tk.Tk()
    root.withdraw()
    current_dir = filedialog.askdirectory()

    sim_series = classes.SimulationSeries(current_dir, 'C:\TRNSYS18\Exe\TrnEXE64.exe', 'Simulationsvarianten.xlsx')

    os.makedirs(sim_series.path_sim_series)  # create new directory for simulation series

    sim_series.import_input_excel()
    # sim_series.start_sim_series()
    sim_series.start_sim_series_par()



