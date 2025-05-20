import sys, os
import multiprocessing
import time
import csv
import logging
import shutil
import pickle

import numpy as np
import pandas as pd
import TRNSYSAuto.utils as utils

from datetime import datetime
from typing import Optional
from dataclasses import dataclass
from pywinauto.application import Application
from TRNSYSAuto.configs import Configs, Paths, load_from_ini
from TRNSYSAuto.evaluation import Evaluation


class SimulationSeries:
    """Simulation series class.

    A Simulation series is a series of TRNSYS simulations, which can be computed using multiprocessing."""

    def __init__(self, path_config: str, path_root: str, original_sim_variants_excel: str):
        """ Initialize simulation series object.

            A simulation series is saved in a directory in the same directory as the base folder (which contains templates,
            b17/18 files, dck file, the simulation variants Excel file, weather data, load profiles, ...). The simulation
            series directory is named after its corresponding simulation variants Excel file, followed by a timestamp at the
            time of the execution of the main method. TRNSYS has to be installed in order to perform the simulations
            successfully.

            Parameters
            ----------
            path.original_sim_variants_excel : str
                path to original simulation variants Excel file corresponding to the simulation series.
            """
        self.simulations: dict[Simulation] = None  # simulations within simulation series
        self.evaluation: Evaluation = None
        self.configs: Configs = load_from_ini(path=path_config)
        self.path = Paths(_configs=self.configs,
                          root=path_root,
                          config=path_config,
                          original_sim_variants_excel=original_sim_variants_excel)
        self.logger: Optional[logging.Logger] = None

        # set runtime configurations
        kwargs = {
            'execution_time': datetime.now().strftime('%d.%m.%Y_%H.%M'),
            'filename_sim_variants_excel': os.path.basename(self.path.original_sim_variants_excel).split('.')[0],
        }
        self.configs.runtime = Configs.Runtime(**kwargs)

        # todo: Zeilen ab hier verbesserungswürdig
        self.setup_sim_series_dir()
        utils.set_env_and_paths(self.conda_venv_name)

    def init_logger(self):
        """Initialize logging file."""

        # create logger
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)

        # create file handler
        f_handler = logging.FileHandler(self.path.logfile, mode='w')

        # create stream handler (outputs messages in the console additionally)
        s_handler = logging.StreamHandler(sys.stdout)  # sys.stdout prevents messages to be formatted like errors (red)
        s_handler.setLevel(logging.INFO)

        # create formatter and add it to the handlers
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        f_handler.setFormatter(formatter)
        s_handler.setFormatter(formatter)

        # add the handlers to the logger
        self.logger.addHandler(f_handler)
        self.logger.addHandler(s_handler)

        self.logger.info(('Log file created successfully in {}.'.format(self.path.logfile)))

    def save(self):
        """Pickle save SimulationSeries instance."""

        self.logger.info(f'Saving progress in {self.path.savefile}.')

        with open(self.path.savefile, 'wb') as file:
            pickle.dump(self, file)

    def setup_sim_series_dir(self):  # todo rename setup_dir
        """Set up new directory for the simulation series."""

        # if directory already exists, delete
        if os.path.exists(self.path.sim_series_dir):
            shutil.rmtree(self.path.sim_series_dir)

        # create directory
        os.makedirs(self.path.sim_series_dir)

        # create logfile
        self.init_logger()

        # save copy of simulation variants Excel file
        shutil.copy(os.path.join(self.path.original_sim_variants_excel), self.path.sim_variants_excel)

    def check_sim_success(self, reset: bool = False) -> None:
        """Check simulation success.

        Checks if the simulation was calculated successfully, for each simulation inside the simulation series. If so,
        its sim_success flag is switched from False to True. If reset is set to True, the sim_success flags are reset
        first.

        :param bool reset: Determines if the sim_success flags should be reset before checking.
        """

        if reset:
            self.logger.info('Resetting simulation success flags.')

        self.logger.info('Checking for failed simulations.')

        self.sim_success = [sim.check_success() for sim in self.simulations if not (sim.success and not reset)]

        # log simulation success status
        if all(self.sim_success):
            self.logger.info(f'"Simulation of {self.filename_sim_variants_excel}" completed successfully.')
        else:
            self.logger.info(
                f'{sum(self.sim_success)} out of {len(self.sim_success)} simulations completed successfully.')

    def start_sim_series(self):
        """Start simulation series.

        Starts the calculation of the simulation series. Multiple simulations may run simultaneously, depending on the
        attribute "multiprocessing_max". After all simulations are done, the method checks if all simulations were
        calculated successfully. If needed, unsuccessful simulations are calculated and checked again (this process is
        repeated until all simulations were calculated successfully, unless some simulations are on the "ignore" list
        anyway).
        """

        # initialize lock, if multiprocessing is enabled
        if self.multiprocessing_max > 1:
            lock = multiprocessing.Lock()

        while not all(np.logical_or(self.sim_success, self.sim_ignore)):  # check for remaining simulations

            # initialize progress bar
            progress = 0
            total = len(self.sim_list) - sum(np.logical_or(self.sim_success, self.sim_ignore))
            utils.progress_bar(progress, total)

            self.logger.info(f'Starting simulation series "{self.filename_sim_variants_excel}".')

            for index in range(len(self.sim_list)):

                if not self.sim_success[index] and not self.sim_ignore[index]:
                    sim = self.sim_list[index]  # name of simulation
                    path_dck = os.path.join(self.path.sim_series_dir, sim,
                                            self.filename_dck_template)  # path of dck-file

                    try:

                        if self.multiprocessing_max > 1:

                            # create a new process instance
                            process = multiprocessing.Process(target=self.start_sim,
                                                              args=(path_dck, lock))
                            with lock:
                                start_time = time.time()
                                while len(multiprocessing.active_children()) >= self.multiprocessing_max:
                                    time.sleep(5)  # pause until number of active simulations drops below maximum
                                    if time.time() - start_time > self.timeout:
                                        sys.exit('Timeout of ' + str(self.timeout) + ' sec reached, program ended.')
                                time.sleep(5)
                                process.start()  # start process
                            lock.acquire()

                        elif self.multiprocessing_max == 1:
                            self.start_sim(path_dck)

                    except Exception:

                        self.logger.error(f'Error occurred during simulation of {sim}.')

                    progress += 1
                    utils.progress_bar(progress, total)

            # after all simulations were triggered, wait until all are done before proceeding
            while len(multiprocessing.active_children()) > 0:
                time.sleep(5)
                if time.time() - start_time > self.timeout:
                    sys.exit('Timeout of ' + str(self.timeout) + ' sec reached, program ended.')

            # check for each simulation if it was successful
            self.check_sim_success()

            # save progress
            self.save()

    def setup_simulation(self):
        """Set up simulation series.

        Setting up the simulation series is only necessary once, continuing the simulation at a later time does not need
        an additional setup. Doing so anyway results in a reset of the simulation progress.
        """

        self.logger.info('Setting up simulation.')

        # import simulation variants Excel file
        self.import_sim_variants_excel()

        # initialize simulation success/ignore flags
        self.sim_success = [False] * len(self.sim_list)
        self.sim_ignore = [False] * len(self.sim_list)

        self.setup_sim_subdirectories()

        # save SimulationSeries object
        self.save()

    def import_sim_variants_excel(self):
        """Import simulation variants Excel file.

        Imports the simulation variants Excel file and applies the data to the SimulationSeries object.
        """

        self.logger.info(f'Importing simulation variants Excel file from {self.path.original_sim_variants_excel}.')

        # todo: hier wird zwei mal der Inhalt des Simulationsvariantenfiles importiert, auf 1 mal reduzieren
        # read simulation variant parameters
        self.variant_parameter_df = pd.read_excel(self.path.original_sim_variants_excel,
                                                  sheet_name='Simulationsvarianten')
        self.variant_parameter_df.columns = [str(parameter) for parameter in self.variant_parameter_df.columns]

        # read simulation variants Excel file
        excel_data = pd.ExcelFile(self.path.original_sim_variants_excel)

        # convert Excel data into pandas DataFrame
        df = excel_data.parse(self.sheet_name_sim_variants, index_col=0)

        # transpose data
        df_weather = df[df.index == 'Wetterdaten'].transpose()
        b18_series = df[df.index == 'b18'].transpose()
        df_dck = df[df.index == 'dck'].transpose()
        df_mpc = df[df.index == 'mpc'].transpose()

        # convert into series
        self.weather_series = df_weather[1:].squeeze()
        self.b18_series = b18_series[1:].squeeze()

        # use first row as header
        df_dck.columns = df_dck.iloc[0]
        df_mpc.columns = df_mpc.iloc[0]
        self.df_dck = df_dck[1:]
        self.df_mpc = df_mpc[1:]

        # list of simulation variants
        self.sim_list = df.columns[1:].astype(str).tolist()

        # in case there is only 1 simulation variant, make sure those attributes are still pandas.Series
        if len(self.sim_list) == 1:
            self.weather_series = pd.Series(self.weather_series)
            self.weather_series.index = [self.sim_list[0]]
            self.b18_series = pd.Series(self.b18_series)
            self.b18_series.index = [self.sim_list[0]]

        # convert index into string (for stability reasons)
        self.weather_series.index = self.weather_series.index.map(str)
        self.b18_series.index = self.b18_series.index.map(str)
        self.df_dck.index = self.df_dck.index.map(str)
        self.df_mpc.index = self.df_mpc.index.map(str)

    def setup_sim_subdirectories(self):
        """Set up simulation subdirectories.

        1)  Create a subdirectory (inside the simulation series directory) for each TRNSYS simulation
        2)  Save a copy of each file and template necessary for the simulation
        3)  Overwrite parameter inside the copy of the template .dck file, according to the variant's specifications
            from the simulation variants Excel file

        After the simulation is complete, the simulation results from TRNSYS are also saved inside its respective
        subdirectory.
        """

        self.logger.info('Setting up simulation subdirectories.')

        def create_sim_subdir():
            """Create simulation subdirectory within simulation series directory.

            Creates a subdirectory within the simulation series directory, with a copy of all simulation assets and
            template files. Each subdirectory corresponds to one TRNSYS simulation.
            """

            os.makedirs(path_sim)  # create new empty simulation subdirectory

            # filename list of files to be copied into simulation subdirectories
            file_list = [self.filename_dck_template, 'Lastprofil.txt', 'SzenarioAneu.txt', 'Qelww_CHR55025.txt',
                         'Windetc20190804.txt', 'StrahlungBruck.txt',
                         'EXAA_Day Ahead Preise & CO2-Intensität 2015-2022_1_2019.txt']

            # source paths
            src_file_list = file_list + [os.path.join('b18', self.b18_series[sim]),
                                         os.path.join('Wetterdaten', self.weather_series[sim]),
                                         os.path.join('mpc_code', 'main_mpc.py'),
                                         os.path.join('mpc_code', 'MPCModule.py'),
                                         os.path.join('mpc_code', 'settingsMPC.ini')]
            # destination paths
            dst_file_list = file_list + [self.b18_series[sim], self.weather_series[sim], 'main_mpc.py', 'MPCModule.py',
                                         'settingsMPC.ini']

            # Turn into absolute paths
            src_file_list = [os.path.join(self.path.assets_dir, f) for f in src_file_list]
            dst_file_list = [os.path.join(path_sim, f) for f in dst_file_list]

            # copy specified files into simulation subdirectory
            errors = utils.copy_files(src_file_list, dst_file_list)

            if errors:
                self.logger.error(f'FileNotFoundError: {", ".join(errors)}')
                self.sim_ignore[index] = True  # simulation variant will be ignored
                raise FileNotFoundError  # program will end if error is raised

        def overwrite_dck_file_parameters():
            """Overwrite parameters inside .dck File.

            Overwrites the parameters inside the .dck File, according to the corresponding simulation description in
            the simulation variants Excel file.
            """

            # replace weather data file name in .dck file
            utils.find_and_replace(
                path_dck, pattern=r'(ASSIGN "tm2")', replacement=r'ASSIGN "' + self.weather_series[sim] + '"')

            # replace .b17/.b18 file name in .dck file
            utils.find_and_replace(
                path_dck, pattern=r'(ASSIGN "b17")', replacement=r'ASSIGN "' + self.b18_series[sim] + '"')

            # replace parameter values
            utils.replace_parameter_values(path_dck, self.df_dck.loc[sim])

        for index, sim in enumerate(self.sim_list):
            path_sim = os.path.join(self.path.sim_series_dir, sim)  # path of simulation subdirectory
            path_dck = os.path.join(path_sim, self.filename_dck_template)  # path of .dck file
            path_mpc = os.path.join(path_sim, self.filename_mpc_settings)  # path of settingsMPC.ini file

            create_sim_subdir()
            overwrite_dck_file_parameters()

            # replace parameters inside settingsMPC.ini file
            utils.replace_parameter_values(path_mpc, self.df_mpc.loc[sim])


class Simulation:

    def __init__(self):
        name: str  # name of simulation
        success: bool = False  # True, if simulated successfully
        ignore: bool = False  # if True, do not simulate
        param: SimParameters

    def start_sim(self, path_dck_file, lock=None):
        """Start simulation.

        Starts a TRNSYS simulation using a specified dck-file. If multiprocessing is used, a lock is passed which
        ensures no other simulation starts until a specific point is reached. In this case, the lock is released as soon
        as the TRNSYS simulation window opens. Optionally, start_time_buffer acts as a time buffer before releasing the
        lock.

        Parameters
        ----------
        path_dck_file : str
            Path of the dck-File necessary for the TRNSYS simulation.
        lock : multiprocessing.Lock
            Lock object from the multiprocessing module.
        """

        def delete_redundant_files():
            """Delete redundant files generated by TRNSYS, to save disk space."""

            path_sim = os.path.dirname(path_dck_file)

            for redundant_file in self.filenames_redundant:
                try:
                    os.remove(os.path.join(path_sim, redundant_file))
                except FileNotFoundError:
                    pass

        # start application
        app = Application(backend='uia')
        app.start(self.path_exe)

        try:
            app.connect(title="Öffnen", timeout=2)
            app.Öffnen.wait('visible')
            app.Öffnen.set_focus()

            # insert .dck file path
            app.Öffnen.FileNameEdit.set_edit_text(path_dck_file)

            # press start button
            Button = app.Öffnen.child_window(title="Öffnen", auto_id="1", control_type="Button").wrapper_object()
            Button.click_input()

            # wait for the simulation window to open
            app.Öffnen.wait_not('visible', timeout=10)

        except Exception:  # TimeoutError:  todo: Add specific exceptions
            self.logger.error(f'Unknown error occured during simulation of {path_dck_file}.')

            # if an exception/error occurs, the window closes and the lock is released so the next simulation can start
            app.kill()
            if lock is not None:
                lock.release()
            return

        # add a time buffer before releasing the lock, which delays the next simulation
        time.sleep(self.start_time_buffer)
        if lock is not None:
            lock.release()

        window_title = 'TRNSYS: ' + path_dck_file
        window_title = window_title.replace('documents', 'Documents')  # workaround, as search is case sensitive

        success_message = app.window(title=window_title)  # .window(control_type="Text")
        try:
            success_message.wait('visible', timeout=self.timeout)
        except TimeoutError:
            pass  # goes ahead and closes window after time out

        app.kill()  # close window
        time.sleep(5)

        delete_redundant_files()

    def check_success(self):  # todo anpassen
        # path of output file
        path_output = os.path.join(self.path.sim_series_dir, self.sim_list[index], self.filename_trnsys_output)

        try:
            with open(path_output) as f:
                data = list(csv.reader(f, delimiter="\t"))

            # simulation was successful, if hourly data is complete (8760 entries)
            self.success = not len(data) < 8762
        except FileNotFoundError:  # no file found
            self.success = False

        return self.success


@dataclass
class SimParameters:
    dck: dict
    mpc: dict
    b18: str
    weather: str

    def from_df(self):
        pass
