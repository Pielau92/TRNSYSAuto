import os
import utils
import pickle

import tkinter as tk

from typing import Callable
from TRNSYSAuto.simulation import SimulationSeries


def gui(root_dir: str) -> None:
    """Graphical user interface (see GUI guide: https://realpython.com/python-gui-tkinter/)."""

    def load_sim_series() -> SimulationSeries:
        """Load simulation series from pickle savefile."""

        window.destroy()  # close GUI window

        # ask for pickle savefile and load SimulationSeries instance
        initial_dir = os.path.join(root_dir, '../data', 'results')
        path_savefile = utils.ask_filename(initialdir=initial_dir)

        # load savefile
        with open(path_savefile, 'rb') as file:
            sim_series = pickle.load(file)

        sim_series.init_logger()  # initialize logging file

        return sim_series

    def simulate_and_evaluate() -> None:
        window.destroy()  # close GUI window

        # for each simulation series...
        for sim_series in create_sim_queue(root_dir):
            sim_series.setup_simulation()  # set simulation up
            sim_series.start_sim_series()  # start simulation series

            sim_series.setup_evaluation()  # set evaluation up
            sim_series.start_evaluation()  # start evaluation

        window.quit()

    def simulate() -> None:
        window.destroy()  # close GUI window

        # for each simulation series in the queue...
        for sim_series in create_sim_queue(root_dir):
            sim_series.setup_simulation()  # start simulation
            sim_series.start_sim_series()  # start simulation series

        window.quit()

    def evaluate() -> None:
        sim_series = load_sim_series()

        # start evaluation
        sim_series.setup_evaluation()  # set evaluation up
        sim_series.start_evaluation()  # start evaluation

        window.quit()

    def continue_simulation() -> None:
        sim_series = load_sim_series()

        sim_series.check_sim_success(reset=True)
        sim_series.start_sim_series()

        window.quit()

    def continue_evaluation() -> None:
        sim_series = load_sim_series()

        sim_series.start_evaluation()

        window.quit()

    def add_button(button_text: str, button_command: Callable) -> None:

        button = tk.Button(
            window,
            text=button_text,
            width=25,
            height=5,
            command=button_command,
        )
        button.pack()

    # define window
    window = tk.Tk()
    label = tk.Label(text="Aktion auswählen")
    label.pack()

    # define buttons
    button_dict = {"Simulate and evaluate": simulate_and_evaluate,
                   "Simulate only": simulate,
                   "Evaluate only": evaluate,
                   "Continue incomplete simulation": continue_simulation,
                   "Continue incomplete evaluation": continue_evaluation,
                   }

    # add all buttons to gui
    for text, command in button_dict.items():
        add_button(text, command)

    # run gui
    window.mainloop()


def create_sim_queue(root_dir: str) -> list[SimulationSeries]:
    """Ask for simulation variants Excel files and create list of SimulationSeries objects accordingly.

    Opens the explorer and asks for one or multiple simulation variants Excel files. For each selected Excel file, an
    additional instance of SimulationSeries is created and added to a list.

    :param str root_dir: root directory
    :return:
    """

    initial_dir = os.path.join(root_dir, '../data', 'input')
    path_config = os.path.join(root_dir, 'configs.ini')

    return [SimulationSeries(path_config, root_dir, path) for path in utils.ask_filenames(initialdir=initial_dir)]


if __name__ == '__main__':
    gui(utils.get_root_dir())
