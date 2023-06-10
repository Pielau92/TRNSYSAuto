import multiprocessing
import classes
import os
import tkinter as tk
from tkinter import filedialog

# in case the script asks directory/path multiple times, use this
# import sys
# sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED

if __name__ == '__main__':
    """ For some unknown reason, when producing an exe-File main is also affected by multiprocessing (directory/file is
    asked multiple times), therefore the function freeze_support is necessary
    """
    multiprocessing.freeze_support()

    # ask simulation variants Excel file path
    root = tk.Tk()
    root.withdraw()
    path_sim_variants_excel = filedialog.askopenfilenames()

    sim_queue = list()

    for path in path_sim_variants_excel:
        path = path.replace("/", "\\")

        # create simulation series object
        sim_queue.append(classes.SimulationSeries(path))

    for sim_series in sim_queue:

        # ask simulation variants Excel file path
        sim_series.set_paths()

        # import and apply settings Excel file
        sim_series.import_settings_excel()
        sim_series.set_settings()

        # create new folder for simulation series
        os.makedirs(sim_series.dir_sim_series)

        # import routine for input Excel file
        sim_series.import_input_excel()

        # start simulation
        sim_series.start_sim_series()
