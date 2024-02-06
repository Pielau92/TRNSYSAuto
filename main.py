import multiprocessing
import classes
import os
import tkinter as tk

from tkinter import filedialog  # explicit import required, as calling from tk.filedialog does not work


# region FIX askdirectory/askfilename window does not open
# in case the askdirectory/askfile window does not open, try this (fixes compatibility issues tkinter <=> pywinauto)
# import sys
# sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED
# endregion


def main():
    """Main method."""

    # region FIX directory/file asked multiple times
    # for some unknown reason, when producing an exe-File main is also affected by multiprocessing (therefore
    # directory/file is asked multiple times), therefore the function "freeze_support" is necessary
    multiprocessing.freeze_support()
    # endregion

    sim_queue = create_sim_queue()  # create queue of simulation series (list of SimulationSeries object(s))
    start_sim_queue(sim_queue)  # start calculation


def create_sim_queue():
    """Create queue of simulation series.

    Opens the explorer and asks for one or multiple simulation variants Excel files. For each selected Excel file, an
    additional SimulationSeries object is created and added to the list sim_queue.

    Returns
    -------
    sim_queue : list of SimulationSeries
        list with SimulationSeries objects.
    """

    # ask simulation variants Excel file path(s)
    root = tk.Tk()
    root.withdraw()
    path_sim_variants_excel = filedialog.askopenfilenames()

    # create SimulationSeries object for each simulation variants Excel file selected
    sim_queue = list()
    for path in path_sim_variants_excel:
        path = path.replace("/", "\\")

        # create simulation series object
        sim_queue.append(classes.SimulationSeries(path))

    return sim_queue


def start_sim_queue(sim_queue):
    """Start simulation series queue.

    Starts each SimulationSeries object stored in sim_queue successively. Additionally, after each simulation series an
    evaluation routine starts, provided the parameter "autostart_evaluation" is set to True.

    Parameters
    ----------
    sim_queue : list of SimulationSeries
        list with SimulationSeries objects.
    """

    for sim_series in sim_queue:

        # store essential paths in SimulationSeries object
        # sim_series.set_paths()

        # import and apply settings Excel file
        sim_series.import_settings_excel(filename_settings_excel='Einstellungen.xlsx',
                                         name_excelsheet_settings='Einstellungen')

        # create new folder for simulation series
        os.makedirs(sim_series.dir_sim_series)

        # initialize logging file
        sim_series.initialize_logging()

        # import simulation variants Excel file
        sim_series.import_input_excel()

        # WORKAROUND - perform Mapping
        sim_series.mapping_routine()

        # start simulation series
        sim_series.start_sim_series()

        # start evaluation routine, if enabled
        if sim_series.autostart_evaluation:
            sim_series.evaluation()


if __name__ == '__main__':
    main()
