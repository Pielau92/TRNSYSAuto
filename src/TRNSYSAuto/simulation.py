import sys, os
import multiprocessing
import time
import logging
import shutil
import pickle
import mpccontroller

from datetime import datetime
from typing import Optional
from tqdm import tqdm
from importlib import resources

import TRNSYSAuto.utils as utils

from TRNSYSAuto.paths import Paths
from TRNSYSAuto.datalayer import ExcelData
from TRNSYSAuto.configs import Configs, Runtime
from trnsys_simulation.simulation import Simulation


class SimulationSeries:
    """A Simulation series is a series of TRNSYS simulations, which can be computed using multiprocessing."""

    def __init__(self, path_config: str, path_root: str, path_original_sim_variants_excel: str):
        self.simulations: dict[str, Simulation] = {}  # simulations within simulation series
        self.excel_data: ExcelData | None = None
        self.configs = Configs(path_config)
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
        kwargs.update({'dirname_sim_series': f'{kwargs["filename_sim_variants_excel"]}_{kwargs["execution_time"]}'})
        self.configs.runtime = Runtime(**kwargs)

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
        self.excel_data = ExcelData(
            path_excel=self.path.original_sim_variants_excel,
            sheet_name=self.configs.general.sheet_name_sim_variants
        )

        filenames = Simulation.Configs.Filenames(
            dck_template=self.configs.filenames.dck_template,
            trnsys_output=self.configs.filenames.trnsys_output,
            mpc_configs=self.configs.filenames.mpc_configs,
            redundant=self.configs.filenames.redundant,
        )

        configs = Simulation.Configs(
            path_exe=self.configs.general.path_exe,
            timeout_sim=self.configs.time.timeout_sim,
            timeout_open_dck_window=self.configs.time.timeout_open_dck_window,
            timeout_open_sim_window=self.configs.time.timeout_open_sim_window,
            buffer_sim_start=self.configs.time.buffer_sim_start,
            filenames=filenames,
        )

        variants = self.excel_data.parameters.keys()
        for variant in variants:
            path_sim = os.path.join(self.path.sim_series_dir, variant)
            self.simulations[variant] = Simulation(
                path_dir=path_sim,
                path_exe=self.configs.general.path_exe,
                params=self.excel_data.parameters[variant],
                configs=configs,
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
        if not os.path.isfile(path_MPCModule):
            # todo: Workaround mit hardcoded path_MPCModule entfernen.
            """ ACHTUNG: bis dahin nicht vergessen MPCModule.py in
            "C:/Users/pierre/PycharmProjects/TRNSYSAuto/dist/assets" aktuell zu halten, wenn im Projekt "mpccontroller"
            Änderungen vorgenommen werden! """
            alt_path = os.path.join(self.path.assets_dir, 'MPCModule.py')
            self.logger.warning(f'No file found inside {path_MPCModule}, will look at {alt_path} instead.')
            path_MPCModule = alt_path

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
                sim.params.mpc,
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

            sim.setup()

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
                        process = multiprocessing.Process(target=sim.simulate,
                                                          args=(lock,))
                        with lock:
                            start_time = time.time()
                            while len(multiprocessing.active_children()) >= self.configs.general.multiprocessing_max:
                                time.sleep(5)  # pause until number of active simulations drops below maximum
                                if time.time() - start_time > self.configs.time.timeout_sim:
                                    sys.exit(
                                        f'Timeout of {str(self.configs.time.timeout_sim)} sec reached, program ended.')
                            time.sleep(5)
                            process.start()
                        lock.acquire()

                    else:  # no lock
                        sim.simulate()

                except Exception:

                    self.logger.error(f'Error occurred during simulation of {sim.name}.')

                pbar.update(1)

            # after all simulations were triggered, wait until all are done before proceeding
            while len(multiprocessing.active_children()) > 0:
                time.sleep(5)
                if time.time() - start_time > self.configs.time.timeout_sim:
                    sys.exit(f'Timeout of {str(self.configs.time.timeout_sim)} sec reached, program ended.')

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


