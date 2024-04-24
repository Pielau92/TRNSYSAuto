# region FIX askdirectory window does not open
"""There are compatibility issues between filedialog.askdirectory() and pywinauto, which cause the askdirectory window
 not to open. To fix this, use the following lines. This must happen before importing pywinauto and tkinter."""
# import sys
# import warnings

# deactivate warnings as workaround for higher stability, but it is not optimal as other warnings are also suppressed
# warnings.simplefilter("ignore", UserWarning)

# sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED
# endregion

import multiprocessing
import classes
import functions
import os
import tkinter as tk


def main():
    """Main method."""

    # region FIX directory/file asked multiple times
    """For some unknown reason (when producing an exe-File), main is also affected by multiprocessing (therefore
    directory/file is asked multiple times), therefore the function "freeze_support" is necessary."""
    multiprocessing.freeze_support()
    # endregion

    start_gui()


def start_gui():
    """Start GUI

    https://realpython.com/python-gui-tkinter/
    """

    def simulate_and_evaluate():
        window.destroy()  # close GUI window

        # for each simulation series...
        for sim_series in get_sim_queue():
            sim_series.start()  # start simulation
            sim_series.evaluate()  # start evaluation

        window.quit()

    def simulate():
        window.destroy()  # close GUI window

        # for each simulation series...
        for sim_series in get_sim_queue:
            sim_series.start()  # start simulation

        window.quit()

    def evaluate():
        window.destroy()  # close GUI window

        path_savefile = functions.ask_filename()  # ask for pickle savefile
        sim_series = functions.load(path_savefile)  # load SimulationSeries object

        # adapt simulation series directory path, in case the simulation was done in another directory/another machine
        sim_series.dir_sim_series = os.path.dirname(path_savefile)

        # initialize logging file
        sim_series.initialize_logging()

        # start evaluation
        sim_series.evaluate()  # start evaluation

        window.quit()

    def continue_simulation():
        window.destroy()  # close GUI window

        path_savefile = functions.ask_filename()  # ask for pickle savefile
        sim_series = functions.load(path_savefile)  # load SimulationSeries object

        # initialize logging file
        sim_series.initialize_logging()

        sim_series.check_sim_success()
        sim_series.start_sim_series()

        window.quit()

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
        text="Continue incomplete simulation",
        width=25,
        height=5,
        command=continue_simulation,
    )
    btn_continue_sim.pack()

    window.mainloop()

    # endregion


def get_sim_queue():
    """Ask for simulation variants Excel files and create list of SimulationSeries objects accordingly.

    Opens the explorer and asks for one or multiple simulation variants Excel files. For each selected Excel file, an
    additional SimulationSeries object is created and added a list."""

    return [classes.SimulationSeries(path) for path in functions.ask_filenames()]


if __name__ == '__main__':
    main()
