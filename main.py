import multiprocessing
import classes
import os
import tkinter as tk

from functions import load
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

    start_gui()


def start_gui():
    """Start GUI

    https://realpython.com/python-gui-tkinter/
    """

    def simulate_and_evaluate():
        window.destroy()    # close GUI window

        sim_queue = create_sim_queue()  # create queue of simulation series (list of SimulationSeries object(s))

        for sim_series in sim_queue:
            sim_series.start()  # start calculation
            sim_series.evaluation()     # start evaluation

    def simulate():
        window.destroy()    # close GUI window

        sim_queue = create_sim_queue()  # create queue of simulation series (list of SimulationSeries object(s))

        for sim_series in sim_queue:
            sim_series.start()  # start calculation

    def evaluate():
        window.destroy()    # close GUI window

    def continue_simulation():
        window.destroy()  # close GUI window

        # ask simulation series savefile
        sim_series_path = functions.ask_filename()

        sim_series = load(sim_series_path)

        # todo: sim_series herannehmen und Simulation anstoßen, nachdem check_success ausgeführt wurde.

    # region GUI

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

    btn_continue_sim = tk.Button(
        window,
        text="Continue interrupted Simulation",
        width=25,
        height=5,
        command=continue_simulation,
    )
    btn_continue_sim.pack()

    window.mainloop()

    # endregion


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
    path_sim_variants_excel = functions.ask_filenames()

    # create SimulationSeries object for each simulation variants Excel file selected
    sim_queue = list()
    for path in path_sim_variants_excel:
        path = path.replace("/", "\\")

        # create simulation series object and append to list
        sim_queue.append(classes.SimulationSeries(path))

    return sim_queue

if __name__ == '__main__':
    main()
