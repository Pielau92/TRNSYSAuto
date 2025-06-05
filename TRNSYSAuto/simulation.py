import sys, os
import multiprocessing
import time
import csv
import logging
import shutil
import pickle
import mpccontroller

import TRNSYSAuto.utils as utils

from datetime import datetime
from typing import Optional
from tqdm import tqdm
from pywinauto.application import Application
from TRNSYSAuto.configs import Configs, Paths, load_from_ini
from TRNSYSAuto.datalayer import ExcelData, SimParameters

from importlib import resources


class SimulationSeries:
    """A Simulation series is a series of TRNSYS simulations, which can be computed using multiprocessing."""

    def __init__(self, path_config: str, path_root: str, path_original_sim_variants_excel: str):
        self.simulations: dict[Simulation] = {}  # simulations within simulation series
        self.excel_data = ExcelData()
        self.configs: Configs = load_from_ini(path=path_config)
        self.path = Paths(_configs=self.configs,
                          root=path_root,
                          config=path_config,
                          original_sim_variants_excel=path_original_sim_variants_excel)
        self.logger: Optional[logging.Logger] = None
        self.sim_success: list[bool] = []

        # set runtime configurations
        kwargs = {
            'execution_time': datetime.now().strftime('%d.%m.%Y_%H.%M'),
            'filename_sim_variants_excel': os.path.basename(self.path.original_sim_variants_excel).split('.')[0],
        }
        self.configs.runtime = Configs.Runtime(**kwargs)

        # self.evaluation = Evaluation()

    def setup(self):
        """Set up simulation series, as preparation for the simulation process ."""

        logs = []  # save log messages until logfile is created

        # create directory
        if os.path.exists(self.path.sim_series_dir):
            logs.append(f'Directory {self.path.sim_series_dir} already exists - deleting directory.')
            shutil.rmtree(self.path.sim_series_dir)

        logs.append(f'Creating new simulation series directory at {self.path.sim_series_dir}.')
        os.makedirs(self.path.sim_series_dir)

        self.init_logger(logs)

        self.logger.info(f'Saving copy of simulation variants Excel file at {self.path.sim_variants_excel}.')
        shutil.copy(os.path.join(self.path.original_sim_variants_excel), self.path.sim_variants_excel)

        self.init_simulations()

        self.create_sim_subdirectories()

        utils.set_env_and_paths(self.configs.general.conda_venv_name)

        self.sim_success = [sim.success for _, sim in self.simulations.items()]

        # save progress
        self.save()

    def init_logger(self, logs: list[str] = None):
        """Initialize logging file.

        Optionally, logs that would have been due, before the logger is initialized, can be passed. Those are logged in
        immediately after the logger is initialized.

        :param list[str] logs: list of messages to be logged in after initialization
        """

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

        if logs:
            [self.logger.info(log) for log in logs]

        self.logger.info(('Log file created successfully in {}.'.format(self.path.logfile)))

    def init_simulations(self):
        """Initialize simulations based on the simulation variants Excel file."""

        self.logger.info(f'Importing simulation variants Excel file from {self.path.original_sim_variants_excel}.')
        self.excel_data.import_excel(
            path=self.path.original_sim_variants_excel,
            sheet_name=self.configs.sheetnames.sim_variants
        )

        # get simulation parameters from imported Excel data
        self.excel_data.get_sim_params()

        variants = self.excel_data.parameters.keys()
        for variant in variants:
            self.simulations[variant] = Simulation(
                name=variant,
                params=self.excel_data.parameters[variant],
                configs=self.configs,
                paths=self.path,
                logger=self.logger
            )

    def create_sim_subdirectories(self):
        """Create a simulation subdirectory.

        Within the simulation series directory, creates a subdirectory for each simulation. Also fill each subdirectory
        with template files and overwrites certain values inside them, based on the simulation variant Excel file data."""

        self.logger.info(f'Creating simulation subdirectories inside {self.path.sim_series_dir}.')

        # path to MPCModule.py from mpccontroller package (necessary for CallingPythonFromTrnsys)
        with resources.path(mpccontroller, "MPCModule.py") as path:
            path_MPCModule = path.__str__()

        for key in self.simulations.keys():
            sim = self.simulations[key]
            path_sim = os.path.join(self.path.sim_series_dir, sim.name)

            os.makedirs(path_sim)  # create new empty simulation subdirectory

            # todo Lösung finden, wo kein 'assets' und 'mpc' Ordner nötig sind
            os.makedirs(os.path.join(path_sim, 'assets'))  # create assets directory within it

            # (relative) source paths for copying process
            src_file_list = self.configs.filenames.templates[:]
            src_file_list += [
                self.configs.filenames.dck_template,
                os.path.join('b18', sim.params.b18),
                os.path.join('Wetterdaten', sim.params.weather),
                path_MPCModule
            ]

            # (relative) destination paths for copying process
            dst_file_list = [os.path.basename(filename) for filename in src_file_list]

            # add files to be copied inside assets directory
            src_file_list += self.configs.filenames.templates_assets
            dst_file_list += \
                [os.path.join('assets', filename) for filename in self.configs.filenames.templates_assets]

            # turn into absolute paths
            src_file_list = [os.path.join(self.path.assets_dir, f) for f in src_file_list]
            dst_file_list = [os.path.join(path_sim, f) for f in dst_file_list]

            # copy specified files into simulation subdirectory
            errors = utils.copy_files(src_file_list, dst_file_list)
            if errors:
                self.logger.error(f'FileNotFoundError: {", ".join(errors)}')
                self.simulations[key].ignore = True  # simulation variant will be ignored
                raise FileNotFoundError  # program will end if error is raised

            sim.overwrite_dck_file_parameters()

    def save(self):
        """Pickle save SimulationSeries instance."""

        self.logger.info(f'Saving progress in {self.path.savefile}.')

        with open(self.path.savefile, 'wb') as file:
            pickle.dump(self, file)

    def simulate(self):
        """Start simulation series.

        Calculates all simulations that were neither already simulated not marked to be ignored. Then, checks if all
        simulations were calculated successfully. If necessary, unsuccessful simulations are calculated and checked
        again (this process is repeated until all simulations were calculated successfully).
        Also, multiple simulations may run simultaneously, depending on the multiprocessing configurations in
        self.configs.
        """

        def check_sim_flags() -> list[bool]:
            """Check for each simulation if enabled for calculation (neither simulated successfully nor marked to be
            ignored)."""

            return utils.logical_or([
                [_sim.success for _, _sim in self.simulations.items()],  # simulation success flags
                [_sim.ignore for _, _sim in self.simulations.items()]  # ignore flags
            ])

        # if multiprocessing is enabled, initialize lock
        if self.configs.general.multiprocessing_max > 1:
            lock = multiprocessing.Lock()
        else:
            lock = None

        # check which simulations are enabled for simulation
        sim_flags = check_sim_flags()

        # initialize progress bar
        pbar = tqdm(total=len(self.sim_success) - sum(sim_flags))

        while not all(sim_flags):  # check for remaining simulations

            self.logger.info(f'Starting simulation series "{self.configs.runtime.filename_sim_variants_excel}".')

            # for index in range(len(self.sim_list)):
            for _, sim in self.simulations.items():

                # if already successful or to be ignored, skip
                if sim.success or sim.ignore:
                    continue

                try:

                    if lock:

                        # create a new process instance
                        process = multiprocessing.Process(target=sim.start,
                                                          args=(lock,))
                        with lock:
                            start_time = time.time()
                            while len(multiprocessing.active_children()) >= self.configs.general.multiprocessing_max:
                                time.sleep(5)  # pause until number of active simulations drops below maximum
                                if time.time() - start_time > self.configs.general.timeout:
                                    sys.exit(
                                        f'Timeout of {str(self.configs.general.timeout)} sec reached, program ended.')
                            time.sleep(5)
                            process.start()
                        lock.acquire()

                    else:  # no lock
                        sim.start()

                except Exception:

                    self.logger.error(f'Error occurred during simulation of {sim.name}.')

                pbar.update(1)

            # after all simulations were triggered, wait until all are done before proceeding
            while len(multiprocessing.active_children()) > 0:
                time.sleep(5)
                if time.time() - start_time > self.configs.general.timeout:
                    sys.exit(f'Timeout of {str(self.configs.general.timeout)} sec reached, program ended.')

            # check for each simulation if it was successful
            self.check_sim_success()
            sim_flags = check_sim_flags()

            # save progress
            self.save()

        pbar.close()

    def check_sim_success(self, reset: bool = False) -> None:
        """Check simulation success.

        Checks for each simulation inside the simulation series, if the simulation was calculated successfully, If so,
        its sim_success flag is switched from False to True. If reset is set to True, the sim_success flags are reset
        first.

        :param bool reset: determines if the sim_success flags should be reset before checking.
        """

        if reset:
            self.logger.info('Resetting simulation success flags.')

        self.logger.info('Checking for failed simulations.')

        self.sim_success = [sim.check_success() for _, sim in self.simulations.items() if
                            not (sim.success and not reset)]

        # log simulation success status
        if all(self.sim_success):
            self.logger.info(
                f'"Simulation of {self.configs.runtime.filename_sim_variants_excel}" completed successfully.')
        else:
            self.logger.info(
                f'{sum(self.sim_success)} out of {len(self.sim_success)} simulations completed successfully.')


class Simulation:

    def __init__(self, name: str, params: SimParameters, configs: Configs, paths: Paths, logger: logging.Logger):
        self.name = name  # name of simulation
        self.params = params
        self.path = paths
        self.configs = configs
        self.logger = logger

        self.success: bool = False  # True, if simulated successfully
        self.ignore: bool = False  # if True, do not simulate

    @property
    def path_dck(self) -> str:
        """Path to dck file."""
        return os.path.join(self.path.sim_series_dir, self.name, self.configs.filenames.dck_template)

    def start(self, lock: multiprocessing.Lock = None):
        """Start simulation.

        Starts a TRNSYS simulation and uses a specified dck-file as input. If multiprocessing is used, a lock is passed
        which ensures no other simulation starts until a specific point is reached. In this case, the lock is released
        as soon as the TRNSYS simulation window opens.
        Optionally, start_time_buffer acts as a time buffer before releasing the lock (see self.configs).

        :param multiprocessing.Lock lock:  lock object from the multiprocessing module.
        """

        # start application
        app = Application(backend='uia')
        app.start(self.configs.general.path_exe)

        try:
            app.connect(title="Öffnen", timeout=2)
            app.Öffnen.wait('visible')
            app.Öffnen.set_focus()

            # insert .dck file path
            app.Öffnen.FileNameEdit.set_edit_text(self.path_dck)

            # press start button
            Button = app.Öffnen.child_window(title="Öffnen", auto_id="1", control_type="Button").wrapper_object()
            Button.click_input()

            # wait for the simulation window to open
            app.Öffnen.wait_not('visible', timeout=10)

        except Exception as e:  # TimeoutError:  todo: Add specific exceptions
            self.logger.error(f'{e} error occured during simulation of {self.path_dck}.')

            # if an exception/error occurs, the window closes and the lock is released so the next simulation can start
            app.kill()
            if lock is not None:
                lock.release()
            return

        # add a time buffer before releasing the lock, which delays the next simulation
        time.sleep(self.configs.general.start_time_buffer)
        if lock is not None:
            lock.release()

        window_title = 'TRNSYS: ' + self.path_dck
        window_title = window_title.replace('documents', 'Documents')  # workaround, as search is case sensitive

        success_message = app.window(title=window_title)  # .window(control_type="Text")
        try:
            success_message.wait('visible', timeout=self.configs.general.timeout)
        except TimeoutError:
            pass  # goes ahead and closes window after time out

        app.kill()  # close window
        time.sleep(5)

        # delete redundant files
        path_sim = os.path.dirname(self.path_dck)
        redundant_file_paths = [os.path.join(path_sim, file) for file in self.configs.filenames.redundant]
        utils.delete_files(redundant_file_paths)

    def check_success(self) -> bool:
        """Check if simulation was calculated successfully, based on the TRNSYS output file(s).

        :return: success flag as boolean
        """
        # path of output file
        path_output = os.path.join(self.path.sim_series_dir, self.name, self.configs.filenames.trnsys_output)

        try:
            with open(path_output) as f:
                data = list(csv.reader(f, delimiter="\t"))

            # simulation was successful, if hourly data is complete (8760 entries)
            self.success = not len(data) < 8762
        except FileNotFoundError:  # no file found
            self.success = False

        return self.success

    def overwrite_dck_file_parameters(self):
        """Overwrite parameters inside .dck File.

        Overwrites the parameters inside the .dck File, according to the corresponding simulation description in
        the simulation variants Excel file.
        """

        # replace weather data file name inside .dck file
        utils.find_and_replace(
            self.path_dck, pattern=r'(ASSIGN "tm2")', replacement=r'ASSIGN "' + self.params.weather + '"')

        # replace .b17/.b18 file name inside .dck file
        utils.find_and_replace(
            self.path_dck, pattern=r'(ASSIGN "b17")', replacement=r'ASSIGN "' + self.params.b18 + '"')

        # replace parameter values
        utils.replace_parameter_values(self.path_dck, self.params.dck)
