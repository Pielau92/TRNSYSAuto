# region FIX askdirectory/askfilename window does not open
"""In case the askdirectory/askfile window does not open, try this (fixes compatibility issues between tkinter and
 pywinauto). This must happen before importing pywinauto and tkinter."""
import sys
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED
# endregion

import multiprocessing
import classes
import functions
import os
import tkinter as tk


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

        window.quit()

    def simulate():
        window.destroy()    # close GUI window

        sim_queue = create_sim_queue()  # create queue of simulation series (list of SimulationSeries object(s))

        for sim_series in sim_queue:
            sim_series.start()  # start calculation

        window.quit()

    def evaluate():
        window.destroy()    # close GUI window

        # ask simulation series directory
        sim_series_dir = functions.ask_dir()

        # todo: load SimulationSeries object here

        # find path to simulation variants excel file
        path_sim_variants_excel = os.path.join(
            sim_series_dir,[filename for filename in os.listdir(sim_series_dir) if ".xlsx" in filename][0])

        # create SimulationSeries object
        sim_series = classes.SimulationSeries(path_sim_variants_excel)

        # replace object attributes  with those of the selected simulation series to be evaluated
        sim_series.dir_sim_series = sim_series_dir
        sim_series.dir_save_path_evaluation = os.path.join(sim_series.dir_sim_series, 'evaluation')
        sim_series.file_save_path_cumulative_evaluation = os.path.join(sim_series.dir_save_path_evaluation, 'gesamt.xlsx')
        sim_series.filename_sim_variants_excel = os.path.basename(sim_series.path_sim_variants_excel).split('.')[0]
        sim_series.dir_logfile = os.path.join(sim_series.dir_sim_series, sim_series.logger_filename)

        # initialize logging file
        sim_series.initialize_logging()

        # start evaluation
        sim_series.evaluation()  # start evaluation

        window.quit()

    def continue_simulation():
        window.destroy()  # close GUI window

        filename = 'SimulationSeries.pickle'

        # ask simulation series directory
        sim_series_path = functions.ask_dir()
        savefile_path = os.path.join(sim_series_path, filename)
        sim_series = functions.load(savefile_path)

        window.quit() # todo: sim_series herannehmen und Simulation anstoßen, nachdem check_success ausgeführt wurde.

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
