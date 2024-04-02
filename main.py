import multiprocessing
import classes
import os
import tkinter as tk

from tkinter import filedialog  # explicit import required, as calling from tk.filedialog does not work properly

# region FIX askdirectory/askfilename window does not open
"""In case the askdirectory/askfile window does not open, try this (fixes compatibility issues between tkinter and
 pywinauto)."""


# import sys
# sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED
# endregion


def main():
    """Main method."""

    # region FIX directory/file asked multiple times
    """for some unknown reason (when producing an exe-File), main is also affected by multiprocessing (therefore
    directory/file is asked multiple times), therefore the function "freeze_support" is necessary."""
    multiprocessing.freeze_support()
    # endregion

    # start_gui()
    sim_queue = create_sim_queue()  # create queue of simulation series (list of SimulationSeries object(s))
    start_sim_queue(sim_queue)  # start calculation, evaluation included


def start_gui():
    """Start GUI"""

    def simulate_and_evaluate():
        window.destroy()

    def simulate():
        window.destroy()

    def evaluate():
        window.destroy()

    window = tk.Tk()
    label = tk.Label(text="Aktion auswählen")
    label.pack()

    btn_sim_and_eval = tk.Button(
        window,
        text="Simulate and evaluate",
        width=25,
        height=5,
        command=simulate_and_evaluate,
    )
    btn_sim_and_eval.pack()

    btn_sim = tk.Button(
        window,
        text="Simulate only",
        width=25,
        height=5,
        command=simulate,
    )
    btn_sim.pack()

    btn_eval = tk.Button(
        window,
        text="Evaluate only",
        width=25,
        height=5,
        command=evaluate,
    )
    btn_eval.pack()

    window.mainloop()


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

        # create simulation series object and append to list
        sim_queue.append(classes.SimulationSeries(path))

    return sim_queue


def start_sim_queue(sim_queue):
    """Start simulation series queue.

    Starts each SimulationSeries object stored in sim_queue successively. Additionally, after each simulation series an
    evaluation routine starts, provided the parameter "autostart_evaluation" is set to True in the settings Excel file.

    Parameters
    ----------
    sim_queue : list of SimulationSeries
        list with SimulationSeries objects.
    """

    for sim_series in sim_queue:

        # import and apply settings Excel file
        sim_series.import_settings_excel(filename_settings_excel='Einstellungen.xlsx',
                                         name_excelsheet_settings='Einstellungen')

        # create new directory for the simulation series
        os.makedirs(sim_series.dir_sim_series)

        # initialize logging file
        sim_series.initialize_logging()

        # import simulation variants Excel file
        sim_series.import_sim_variants_excel()

        # start simulation series
        sim_series.start_sim_series()

        # start evaluation routine, if enabled
        if sim_series.autostart_evaluation:
            sim_series.evaluation()


if __name__ == '__main__':
    main()
