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
import sys
import time
import classes
import functions
import os
import tkinter as tk

"""For some unknown reason (when producing an exe-File), main is also affected by multiprocessing (therefore
directory/file is asked multiple times), freeze_support() prevents this."""
multiprocessing.freeze_support()


def main():
    """Main method.

    GUI guide: https://realpython.com/python-gui-tkinter/"""

    def simulate_and_evaluate():
        window.destroy()  # close GUI window

        # for each simulation series...
        for sim_series in create_sim_queue():
            sim_series.setup_simulation()  # set simulation up
            sim_series.start_sim_series()   # start simulation series

            sim_series.setup_evaluation()   # set evaluation up
            sim_series.start_evaluation()  # start evaluation

        window.quit()

    def simulate():
        window.destroy()  # close GUI window

        # for each simulation series in the queue...
        for sim_series in create_sim_queue():
            sim_series.setup_simulation()  # start simulation
            sim_series.start_sim_series()  # start simulation series

        window.quit()

    def evaluate():
        window.destroy()  # close GUI window

        path_savefile = functions.ask_filename()  # ask for pickle savefile
        sim_series = functions.load(path_savefile)  # load SimulationSeries object

        # initialize logging file
        sim_series.initialize_logging()

        # start evaluation
        sim_series.setup_evaluation()   # set evaluation up
        sim_series.start_evaluation()  # start evaluation

        window.quit()

    def continue_simulation():
        window.destroy()  # close GUI window

        path_savefile = functions.ask_filename()  # ask for pickle savefile
        sim_series = functions.load(path_savefile)  # load SimulationSeries object

        # initialize logging file
        sim_series.initialize_logging()

        sim_series.check_sim_success(reset=True)
        sim_series.start_sim_series()

        window.quit()

    def continue_evaluation():
        window.destroy()  # close GUI window

        path_savefile = functions.ask_filename()  # ask for pickle savefile
        sim_series = functions.load(path_savefile)  # load SimulationSeries object

        # initialize logging file
        sim_series.initialize_logging()

        sim_series.start_evaluation()

        window.quit()

    def add_button(button_text, button_command):

        button = tk.Button(
            window,
            text=button_text,
            width=25,
            height=5,
            command=button_command,
        )
        button.pack()

    # Start GUI ()
    window = tk.Tk()
    label = tk.Label(text="Aktion auswählen")
    label.pack()

    check_cwd()  # check if base directory is located in current working directory, otherwise exit

    button_dict = {"Simulate and evaluate": simulate_and_evaluate,
                   "Simulate only": simulate,
                   "Evaluate only": evaluate,
                   "Continue incomplete simulation": continue_simulation,
                   "Continue incomplete evaluation": continue_evaluation,
                   }
    
    for text, command in button_dict.items():
        add_button(text, command)

    window.mainloop()


def check_cwd():
    """Check if base directory is located in current working directory.

    Checks if current working directory (directory where the used main.exe file is located) is located in
    the same directory as the base directory ("Basisordner"), as it is a requirement to perform simulations. If not,
    issues a message and exits the program."""

    if not os.path.exists(os.path.join(os.getcwd(), 'Basisordner')):
        input('main.exe file must be located in the same directory as the base directory ("Basisordnder"). Press ENTER '
              'to exit.')
        sys.exit()


def create_sim_queue():
    """Ask for simulation variants Excel files and create list of SimulationSeries objects accordingly.

    Opens the explorer and asks for one or multiple simulation variants Excel files. For each selected Excel file, an
    additional SimulationSeries object is created and added a list."""

    return [classes.SimulationSeries(path) for path in functions.ask_filenames()]


if __name__ == '__main__':
    main()
