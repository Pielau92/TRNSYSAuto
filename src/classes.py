import sys
import os
import shutil
import multiprocessing
import time
import csv
import logging
import math
import pickle
import openpyxl

import numpy as np
import pandas as pd
import src.functions as functions
import src.settings as settings

from pywinauto.application import Application
from datetime import datetime


class SimulationSeries:
    """Simulation series class.

    A Simulation series is a series of TRNSYS simulations, which can be computed using multiprocessing.
    """

    def __init__(self, path_original_sim_variants_excel, root_dir):
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

        # simulation series data
        self.sim_list = None  # list of simulation variant names
        self.sim_success = None  # boolean list, documenting the successful simulation of each simulation variant
        self.sim_ignore = None  # boolean list, documenting the simulation variants that are to be ignored
        self.df_dck = None  # pandas DataFrame with the simulation parameters to be replaced in the .dck Files
        self.b18_series = None  # pandas series with the .b18 data file names
        self.weather_series = None  # pandas series with the weather data file names

        # evaluation
        self.date_df = functions.create_date_column(2024)
        self.variant_parameter_df = None
        self.eval_success = None
        self.variant_result_columns = None
        self.zone_1_with_df = None
        self.zone_1_without_df = None
        self.zone_3_with_df = None
        self.zone_3_without_df = None

        # log file
        self.logger = None

        # region initialize attributes belonging to settings (including datatype)

        # general
        self.path_exe = str()  # path to TRNSYS executable file
        self.timeout = int()  # if timeout (sec) is reached without starting another simulation, stop program
        self.start_time_buffer = int()  # time buffer (sec) between two simulations, for increased stability (optional)
        self.multiprocessing_max = int()  # maximum number of simulations that can be calculated simultaneously
        self.multiprocessing_autodetect = bool()  # if true, override multiprocessing_max with number of cpu cores
        self.eval_save_interval = int()  # the evaluation progress is saved after each save interval

        # filenames
        self.filename_dck_template = str()  # name of the dck-File template
        self.filename_logger = str()
        self.filename_trnsys_output = str()
        self.filename_savefile = str()
        self.filenames_redundant = list()  # list of redundant TRNSYS files that are to be deleted after the simulation

        # Excel sheet names
        self.sheet_name_sim_variants = str()  # name of the Excel sheet containing the simulation variants table
        self.sheet_name_variant_input = str()
        self.sheet_name_calculation = str()
        self.sheet_name_cumulative_input = str()
        self.sheet_name_zone_1_input = str()
        self.sheet_name_zone_3_input = str()
        self.sheet_name_zone_1_with_operating_time = str()
        self.sheet_name_zone_1_without_operating_time = str()
        self.sheet_name_zone_3_with_operating_time = str()
        self.sheet_name_zone_3_without_operating_time = str()

        # column headers
        self.var_list_zone1 = list()
        self.var_list_zone2 = list()
        self.var_list_zone3 = list()
        self.col_headers_result_column = list()
        self.col_headers_trnsys_output = list()
        self.col_headers_sim_variant = list()

        # endregion

        # paths and settings
        self.path = settings.PathSettings(self, path_original_sim_variants_excel, root_dir)
        self.settings = settings.Settings(self)
        self.settings.load_settings()
        self.settings.apply_settings()

        # current time when the main.exe file was executed
        self.execution_time = datetime.now().strftime('%d.%m.%Y_%H.%M')

        self.filename_sim_variants_excel = os.path.basename(self.path.original_sim_variants_excel).split('.')[0]
        self.dirname_sim_series = self.filename_sim_variants_excel + '_' + self.execution_time

    # @property
    # def cwd(self):
    #     """Current working directory."""
    #     return os.getcwd()

    def setup_simulation(self):
        """Set simulation series up.

        Setting up the simulation series is only necessary once, continuing the simulation at a later time does not need
        an additional setup. Doing so anyway results in a reset of the simulation progress."""

        # import simulation variants Excel file
        self.import_sim_variants_excel()

        # initialize simulation success/ignore flags
        self.sim_success = [False] * len(self.sim_list)
        self.sim_ignore = [False] * len(self.sim_list)

        # create simulation series directory
        self.create_dir_sim_series()

        # save SimulationSeries object
        self.save()

    def setup_evaluation(self):
        """Set evaluation of simulation series up.

        Setting up the evaluation of a simulation series is only necessary once, continuing the evaluation at a later
        time does not need an additional setup. Doing so anyway results in a reset of the evaluation progress."""

        # create evaluation directory
        os.makedirs(self.path.evaluation_save_dir, exist_ok=True)

        # create copy of cumulative evaluation file template
        shutil.copy(self.path.cumulative_evaluation_template, self.path.cumulative_evaluation_save_file)

        # initialize evaluation success list
        self.eval_success = [False] * len(self.sim_list)

        # initialize tables for cumulative evaluation
        self.variant_result_columns = pd.DataFrame(columns=self.sim_list)
        self.zone_1_with_df = pd.DataFrame(columns=self.sim_list)
        self.zone_1_without_df = pd.DataFrame(columns=self.sim_list)
        self.zone_3_with_df = pd.DataFrame(columns=self.sim_list)
        self.zone_3_without_df = pd.DataFrame(columns=self.sim_list)

    def create_dir_sim_series(self):
        """Create simulation series directory.

        Creates a directory to save the results of the simulation series. Also fills the directory with the following:
        -   copy of the simulation variants Excel file
        -   separate subdirectories for each simulation within the simulation series, containing a copy of each file and
        template necessary for the simulation.

        Afterwards, the simulation parameters from the simulation variants Excel file are applied by modifying the
        copied template files.
        """

        def create_sim_subdirectory():
            """Create simulation subdirectory within simulation series directory.

            Creates a subdirectory within the simulation series directory, with a copy of all template files from the
            base directory "Basisordner". Each subdirectory corresponds to one simulation.
            """

            os.makedirs(path_sim)  # create new empty simulation subdirectory

            # source paths
            src_file_list = file_list + [os.path.join('b18', self.b18_series[sim]),
                                         os.path.join('Wetterdaten', self.weather_series[sim]),
                                         os.path.join('mpc_code', 'main_mpc.py'),
                                         os.path.join('mpc_code', 'MPCModule.py')]
            # destination paths
            dst_file_list = file_list + [self.b18_series[sim], self.weather_series[sim], 'main_mpc.py', 'MPCModule.py']

            # copy specified files into simulation subdirectory
            for file_index in range(len(src_file_list)):
                try:
                    shutil.copy(
                        os.path.join(self.path.assets_dir, src_file_list[file_index]),
                        os.path.join(path_sim, dst_file_list[file_index]))
                except FileNotFoundError:
                    message = 'File ' + os.path.join(self.path.assets_dir, src_file_list[file_index] +
                                                     ' could not be found, simulation variant added to ignore list.')
                    self.logger.error(message)
                    print(message)

                    self.sim_ignore[index] = True  # simulation variant will be ignored

        def overwrite_dck_file_parameters():
            """Overwrite parameters inside .dck File.

            Overwrites the parameters inside the .dck File, according to the corresponding simulation description in
            the simulation variants Excel file.
            """

            # replace weather data file name in .dck file
            functions.find_and_replace(
                path_dck, pattern=r'(ASSIGN "tm2")', replacement=r'ASSIGN "' + self.weather_series[sim] + '"')

            # replace .b17/.b18 file name in .dck file
            functions.find_and_replace(
                path_dck, pattern=r'(ASSIGN "b17")', replacement=r'ASSIGN "' + self.b18_series[sim] + '"')

            # replace parameter values
            functions.replace_parameter_values(path_dck, self.df_dck.loc[sim])

        # create new directory for the simulation series
        try:
            os.makedirs(self.path.sim_series_dir)
        except FileExistsError:  # if directory already exists, delete and create new one
            shutil.rmtree(self.path.sim_series_dir)
            os.makedirs(self.path.sim_series_dir)

        # initialize logging file
        self.initialize_logging()

        # copy simulation variants Excel file into simulation series directory
        shutil.copy(os.path.join(self.path.original_sim_variants_excel), self.path.sim_variants_excel)

        # todo: braucht Überarbeitung, da verwirrend. Vieleicht gleich alles hardgecodete ins settings.ini
        # file name list of files to be copied into simulation subdirectories
        file_list = [self.filename_dck_template, 'Lastprofil.txt', 'SzenarioAneu.txt', 'Qelww_CHR55025.txt',
                     'Windetc20190804.txt', 'StrahlungBruck.txt']

        for index, sim in enumerate(self.sim_list):
            path_sim = os.path.join(self.path.sim_series_dir, sim)  # path of simulation subdirectory
            path_dck = os.path.join(path_sim, self.filename_dck_template)  # path of .dck file

            create_sim_subdirectory()
            overwrite_dck_file_parameters()

    def initialize_logging(self):
        """Initialize logging file."""

        self.logger = logging.getLogger(__name__)
        logging.basicConfig(level=logging.DEBUG, filemode='w', filename=self.filename_logger)
        handler = logging.FileHandler(self.path.logfile)
        self.logger.addHandler(handler)

        self.logger.info(('Log file created successfully in {}.'.format(self.path.logfile)))

    def import_sim_variants_excel(self):
        """Import simulation variants Excel file.

        Imports the simulation variants Excel file and applies the data to the SimulationSeries object.
        """

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

        # convert into series
        self.weather_series = df_weather[1:].squeeze()
        self.b18_series = b18_series[1:].squeeze()

        # use first row as header
        df_dck.columns = df_dck.iloc[0]
        self.df_dck = df_dck[1:]

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

    def save(self):
        """Save SimulationSeries object in simulation series directory."""

        with open(self.path.savefile, 'wb') as file:
            pickle.dump(self, file)

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
            """Delete redundant files after the simulation.

            Deletes redundant files generated by TRNSYS, to save disk space.
            """

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
            app.connect(title="Öffnen", timeout=2)  # self.timeout)
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

            message = 'Unknown error occured during simulation of {}.'.format(path_dck_file)
            self.logger.error(message)
            print(message)

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

        success_message = app.window(title=window_title)  # .window(control_type="Text")
        try:
            success_message.wait('visible', timeout=self.timeout)
        except TimeoutError:
            pass  # goes ahead and closes window after time out

        app.kill()  # close window
        time.sleep(5)

        delete_redundant_files()

    def start_sim_series(self):
        """Start simulation series.

        Starts the calculation of the simulation series. Multiple simulations may run simultaneously, depending on the
        attribute "multiprocessing_max". After all simulations are done, the method checks if all simulations were
        calculated successfully. If needed, unsuccessful simulations are calculated and checked again (this process is
        repeated until all simulations were calculated successfully, unless some simulations are on the "ignore" list
        anyway).
        """

        functions.set_env_and_paths()

        # initialize lock, if multiprocessing is enabled
        if self.multiprocessing_max > 1:
            lock = multiprocessing.Lock()

        while not all(np.logical_or(self.sim_success, self.sim_ignore)):  # check for remaining simulations

            # initialize progress bar
            progress = 0
            total = len(self.sim_list) - sum(np.logical_or(self.sim_success, self.sim_ignore))
            functions.progress_bar(progress, total)

            message = 'Starting simulation series from "{}"'.format(self.filename_sim_variants_excel)
            self.logger.info(message)
            print(message)

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

                        message = 'Error occurred during simulation of {}.'.format(sim)
                        self.logger.error(message)
                        print(message)

                    progress += 1
                    functions.progress_bar(progress, total)

            # after all simulations were triggered, wait until all are done before proceeding
            while len(multiprocessing.active_children()) > 0:
                time.sleep(5)
                if time.time() - start_time > self.timeout:
                    sys.exit('Timeout of ' + str(self.timeout) + ' sec reached, program ended.')

            # check for each simulation if it was successful
            self.check_sim_success()

            # save progress
            self.save()

    def check_sim_success(self, reset=False):
        """Check simulation success.

        Checks for each simulation inside the simulation series, if the simulation was calculated successfully. If so,
        its sim_success flag is switched from False to True. If reset is set to True, the sim_success flags are reset
        first.

        Parameters
        ----------
        reset : bool
            Determines if the sim_success flags should be reset before checking.
        """

        if reset:
            message = 'Resetting simulation success flags'
            self.logger.info(message)
            print(message)

            self.sim_success = [False] * len(self.sim_list)

        message = 'Checking for failed simulations'
        self.logger.info(message)
        print(message)

        for index in range(len(self.sim_list)):
            # path of output file
            path_output = os.path.join(self.path.sim_series_dir, self.sim_list[index], self.filename_trnsys_output)

            if not self.sim_success[index]:
                try:
                    with open(path_output) as f:
                        data = list(csv.reader(f, delimiter="\t"))

                    # simulation was successful, if hourly data is complete (8760 entries)
                    self.sim_success[index] = not len(data) < 8762
                except FileNotFoundError:  # no file found
                    self.sim_success[index] = False

        # log simulation success status
        if all(self.sim_success):
            message = '"{}" completed successfully'.format(self.filename_sim_variants_excel)
            self.logger.info(message)
            print(message)
        else:
            message = \
                '{} out of {} simulations completed successfully'.format(sum(self.sim_success), len(self.sim_success))
            self.logger.info(message)
            print(message)

    def start_evaluation(self):
        """Start evaluation.

        Starts the evaluation process, which includes the evaluation of each individual simulation variant, followed by
        a cumulative evaluation using the combined results of those individual evaluations.
        """

        if not all(self.eval_success):
            # initialize progress bar
            progress = 0
            total = len(self.eval_success) - sum(self.eval_success)
            functions.progress_bar(progress, total)

            # logger entry "start"
            message = 'Starting evaluation for {}'.format(self.filename_sim_variants_excel)
            self.logger.info(message)
            print(message)

            # evaluate variants
            for variant_index, variant_name in enumerate(self.sim_list):

                if not self.eval_success[variant_index]:
                    self.evaluate_variant(variant_name, variant_index)
                    progress += 1
                    functions.progress_bar(progress, total)
                    if progress % 5 == 0:  # save evaluation progress
                        self.save()

            # save progress
            self.save()

        # cumulative evaluation
        self.cumulative_evaluation()

        # logger entry "finish"
        message = 'Evaluation done.'
        self.logger.info(message)
        print(message)

    def evaluate_variant(self, variant_name, variant_index):
        """Evaluate simulation variant.

        Performs a Schweiker model evaluation and exports the result to a simulation variant evaluation template Excel
        file.

        Parameters
        ----------
        variant_name : str
            Name of the simulation variant.
        variant_index : int
            Index of the simulation variant inside the sim_list attribute.
        """

        def create_schweiker_model(var_list_zone, zone):
            """Create SchweikerModel object.

            Creates a SchweikerModel object for a specified zone (the current model has 3).

            Parameters
            ----------
            var_list_zone : str
                Variable list of the zone.
            zone : int
                Index of the zone.
            """
            sm = SchweikerDataFrame()

            sm._df = trnsys_df[var_list_zone].reindex(var_list_zone, axis=1)

            zone = str(zone)

            # adapt column headers
            sm.df.columns = ['Period', 'ta', 'tzone', 'TMSURF_ZONE', 'relh', 'vel', 'pmv', 'ppd', 'clo', 'met',
                             'work']

            # insert date columns
            sm._df = pd.concat([self.date_df[0:len(sm.df)], sm.df], axis=1)

            # schweiker main
            sm.calculate()

            # remove redundant columns
            sm.df.drop(['Tag', 'Monat', 'Jahr', 'Stunde', 'Minute', 'index', 'Period'], axis=1, inplace=True)

            # numerate column names for each zone
            sm.df.columns = ['schweiker_' + string + zone for string in sm.df.columns]

            return sm

        path_variant_directory = os.path.join(self.path.sim_series_dir, variant_name)
        path_variant_file = os.path.join(path_variant_directory, self.filename_trnsys_output)
        save_path_variant_output = os.path.join(self.path.evaluation_save_dir, variant_name + '.xlsx')

        # region CHECK IF...

        # ...the trnsys output file is actually there
        if not os.path.exists(path_variant_file):
            message = f'File {path_variant_file} does not exist!'
            self.logger.error(message)
            print(message)
            return

        # ...the variant has a corresponding directory
        if variant_name not in self.sim_list:
            message = f'Did not find {variant_name} in {self.path.sim_variants_excel}'
            self.logger.error(message)
            print(message)
            return

        # endregion

        # read trnsys output file
        trnsys_df = pd.read_csv(path_variant_file, sep='\s+', skiprows=1, skipfooter=0, engine='python')

        # create schweiker models
        sm1 = create_schweiker_model(self.var_list_zone1, 1)
        sm2 = create_schweiker_model(self.var_list_zone2, 2)
        sm3 = create_schweiker_model(self.var_list_zone3, 3)

        # concatenate output
        result = pd.concat([trnsys_df[self.col_headers_trnsys_output], sm1.df, sm2.df, sm3.df], axis=1)

        # sort columns
        result = result[self.col_headers_sim_variant]

        # save copy of variant evaluation template
        shutil.copy(self.path.variant_evaluation_template, save_path_variant_output)

        # save data
        functions.excel_export_variant_evaluation(
            self.sheet_name_variant_input, result, variant_name, save_path_variant_output, self.variant_parameter_df)

        # update excel to receive cross-referenced values and updates calculations
        functions.update_excel_file(save_path_variant_output)

        # create single column with all hourly values, for the cumulative evaluation Excel file
        result_column = functions.to_single_column(result[self.col_headers_result_column])

        # save single column
        self.variant_result_columns[variant_name] = result_column

        self.eval_success[variant_index] = True

        # message = 'Finished evaluation for variant {}'.format(variant_name)
        # self.logger.info(message)
        # print(message)

    def cumulative_evaluation(self):
        """Perform cumulative evaluation.

        Performs a cumulative evaluation by accessing the evaluation results of the individual variants, combining them
        and exporting the result into the cumulative evaluation template Excel file.
        """

        def read(sheet_name, usecols):
            return pd.read_excel(save_path_variant_output, sheet_name=sheet_name, usecols=usecols, header=None,
                                 nrows=None, skiprows=None)

        # initialize progress bar
        progress = 0
        total = len(self.sim_list)
        functions.progress_bar(progress, total)

        # logger entry "start"
        message = 'Reading variant evaluation files for the cumulative evaluation.'
        self.logger.info(message)
        print(message)

        for variant_index, variant_name in enumerate(self.sim_list):
            save_path_variant_output = os.path.join(self.path.evaluation_save_dir, variant_name + '.xlsx')

            # read data from variant evaluation excel file, for the cumulative evaluation excel file
            self.zone_1_with_df[variant_name] = read(sheet_name=self.sheet_name_zone_1_input, usecols=[3])
            self.zone_1_without_df[variant_name] = read(sheet_name=self.sheet_name_zone_1_input, usecols=[2])
            self.zone_3_with_df[variant_name] = read(sheet_name=self.sheet_name_zone_3_input, usecols=[3])
            self.zone_3_without_df[variant_name] = read(sheet_name=self.sheet_name_zone_3_input, usecols=[2])

            # update progress bar
            progress += 1
            functions.progress_bar(progress, total)

        # logger entry "export"
        message = 'Exporting cumulative evaluation results.'
        self.logger.info(message)
        print(message)

        # copy into cumulative evaluation excel file
        self.excel_export_cumulative_evaluation()

        # update cumulative excel
        functions.update_excel_file(self.path.cumulative_evaluation_save_file)

    def excel_export_cumulative_evaluation(self):
        """Write data into cumulative evaluation file."""

        def export(df, sheetname, startrow, startcol, header=False):
            df.to_excel(writer, sheet_name=sheetname, startrow=startrow, startcol=startcol, index=False, header=header)

        with pd.ExcelWriter(
                self.path.cumulative_evaluation_save_file, mode="a", engine="openpyxl", if_sheet_exists='overlay') \
                as writer:
            export(self.variant_parameter_df, self.sheet_name_cumulative_input, 1, 0, header=True)
            export(self.zone_1_with_df, self.sheet_name_zone_1_with_operating_time, 1, 7)
            export(self.zone_1_without_df, self.sheet_name_zone_1_without_operating_time, 1, 7)
            export(self.zone_3_with_df, self.sheet_name_zone_3_with_operating_time, 1, 7)
            export(self.zone_3_without_df, self.sheet_name_zone_3_without_operating_time, 1, 7)
            export(self.variant_result_columns, self.sheet_name_cumulative_input, 60, 2)


class SchweikerDataFrame:
    """Modified pandas Dataframe for the Schweiker-Model."""

    def __init__(self, *args, **kwargs):
        self._df = pd.DataFrame(*args, **kwargs)

    def __getattr__(self, attr):
        return getattr(self._df, attr)

    @property
    def df(self):
        return self._df

    def calculate(self):

        self.calcFloatingAverageTemperature()

        # adapt metabolic rate
        self.df['metAdaptedColumn'] = self.df['met'] - (0.234 * self.Aussentemp_floating_average) / 58.2

        # determine clothing factor
        self.df['clo'] = 10 ** (-0.172 - 0.000485 * self.df['Aussentemp_floating_average']
                                + 0.0818 * self.df['metAdaptedColumn']
                                - 0.00527 * self.df['Aussentemp_floating_average'] * self.df['metAdaptedColumn'])

        # calculate comfort
        [pmv, ppd] = self.calcComfort()
        self.df['pmv'] = pmv
        self.df['ppd'] = ppd

    def calcFloatingAverageTemperature(self):
        """Calculate floating average temperature.

        Foreign code without documentation"""

        floating_alpha = 0.8
        values_name = 'ta'
        dates_name = 'index'

        if self.df[dates_name].isnull().values.any() or self.df[values_name].isnull().values.any():
            raise ValueError('Values are not allowed to be NaN, interpolate if necessary!')

        average_name = f'{values_name}_mean'
        floating_average_name = f'{values_name}_floating_average'

        df = pd.DataFrame()
        df[dates_name] = self.df[dates_name].copy()
        df[values_name] = self.df[values_name].copy()
        df = df.sort_values(dates_name)
        df['ymd'] = pd.to_datetime(df[dates_name]).dt.date
        mean_df = df.groupby('ymd').mean(numeric_only=False)
        mean_df = mean_df.rename(columns={values_name: average_name})

        df = df.merge(mean_df, how='left', on='ymd')
        day_counter = 1
        next_datapoint_time_delta = pd.Timedelta(0)
        # temporary list which consists the index of the first row of each new day
        new_day_datapoints = [0]

        for i in range(len(df)):
            # print(f'row: {i}')

            if i != 0:
                next_datapoint_time_delta = df.loc[i, 'ymd'] - df.loc[new_day_datapoints[-1], 'ymd']
                # print(next_datapoint_time_delta)

            if next_datapoint_time_delta >= pd.Timedelta('2D'):
                # reset time counter when a time gap happens
                new_day_datapoints = [i]
                day_counter = 1
            elif next_datapoint_time_delta >= pd.Timedelta('1D'):
                new_day_datapoints.append(i)
                day_counter = day_counter + 1

            if day_counter > 8:
                df.loc[i, floating_average_name] = (1 - floating_alpha) * df.loc[
                    new_day_datapoints[-2], average_name] + floating_alpha * df.loc[
                                                       new_day_datapoints[-2], floating_average_name]
            elif day_counter > 7:
                # DIN 1525251 Formula
                df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
                    new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name] + 0.5 *
                                                    df.loc[new_day_datapoints[-5], average_name] + 0.4 * df.loc[
                                                        new_day_datapoints[-6], average_name] + 0.3 * df.loc[
                                                        new_day_datapoints[-7], average_name] + 0.2 * df.loc[
                                                        new_day_datapoints[-8], average_name]) / 3.8
            elif day_counter > 6:
                df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
                    new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name] + 0.5 *
                                                    df.loc[new_day_datapoints[-5], average_name] + 0.4 * df.loc[
                                                        new_day_datapoints[-6], average_name] + 0.3 * df.loc[
                                                        new_day_datapoints[-7], average_name]) / 3.6
            elif day_counter > 5:
                df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
                    new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name] + 0.5 *
                                                    df.loc[new_day_datapoints[-5], average_name] + 0.4 * df.loc[
                                                        new_day_datapoints[-6], average_name]) / 3.3
            elif day_counter > 4:
                df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
                    new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name] + 0.5 *
                                                    df.loc[new_day_datapoints[-5], average_name]) / 2.9
            elif day_counter > 3:
                df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
                    new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name]) / 2.4
            elif day_counter > 2:
                df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
                    new_day_datapoints[-3], average_name]) / 1.8
            elif day_counter > 1:
                df.loc[i, floating_average_name] = df.loc[new_day_datapoints[-2], average_name]
            elif day_counter <= 1:
                df.loc[i, floating_average_name] = df.loc[i, average_name]
            else:
                raise ValueError('day_counter case is not considered! Fix it!')

        self.df['Aussentemp_mean'] = df[average_name]
        self.df['Aussentemp_floating_average'] = df[floating_average_name]

    def calcComfort(self):
        """Calculate comfort.

         Foreign code without documentation"""

        # PMV und PPD Berechnung
        # nach DIN EN ISO 7730 (mit Berichtigung)

        floatingAvgOutdoorTempColumn = self.df['Aussentemp_floating_average']

        row_count = self.df.shape[0]

        # region INITIALIZATION

        P1 = 0
        P2 = 0
        P3 = 0
        P4 = 0
        P5 = 0
        xn = 0
        xf = 0
        hcn = 0
        hc = 0
        tcl = 0
        F = 0
        n = 0
        eps = 0

        HL1 = 0
        HL2 = 0
        HL3 = 0
        HL4 = 0
        HL5 = 0
        HL6 = 0
        TS = 0
        pmv = 0
        ppd = 0
        dr = 0

        a = 7.5
        b = 237.3

        pmvColumn = np.zeros(row_count)
        ppdColumn = np.zeros(row_count)

        # endregion

        for i in range(row_count):

            # in einzelne elemente speichern
            tempAir = self.df.loc[i, 'ta']
            tempRad = self.df.loc[i, 'TMSURF_ZONE']
            hum = self.df.loc[i, 'relh'] / 100
            vAir = self.df.loc[i, 'vel']
            clo = self.df.loc[i, 'clo']
            met = self.df.loc[i, 'metAdaptedColumn']
            wme = self.df.loc[i, 'work']

            # Wasserpartikeldampfdurck
            # Magnus Formel
            # pa=hum*10*exp(16.6536-4030.183/(tempAir+235))
            pa = hum * 6.1078 * 10 ** (a * tempAir / (b + tempAir)) * 100

            # thermal insulation of the clothing
            icl = 0.155 * clo

            # metabolic rate in W/m²
            m = met * 58.15

            # external work in W/m²
            w = wme * 58.15

            # internal heat production in the human body
            mw = m - w

            # clothing area factor
            if icl <= 0.078:
                fcl = 1 + 1.29 * icl
            else:
                fcl = 1.05 + 0.645 * icl

            # heat transfer coefficient by forced convection
            hcf = 12.1 * math.sqrt(vAir)

            # air/ mean radiant temperature in Kelvin
            taa = tempAir + 273
            tra = tempRad + 273

            # first guess for surface temperature clothing
            tcla = taa + (35.5 - tempAir) / (3.5 * (6.45 * (icl + 0.1)))

            P1 = icl * fcl
            P2 = P1 * 3.96
            P3 = P1 * 100
            P4 = P1 * taa
            P5 = 308.7 - 0.028 * mw + P2 * (tra / 100) ** 4

            xn = tcla / 100
            xf = xn
            eps = 0.0015

            n = 0
            while True:
                xf = (xf + xn) / 2
                hcn = 2.38 * abs(100 * xf - taa) ** 0.25

                if hcf > hcn:
                    hc = hcf
                else:
                    hc = hcn

                xn = (P5 + P4 * hc - P2 * xf ** 4) / (100 + P3 * hc)

                if n > 150:
                    break
                n = n + 1

                if abs(xn - xf) > eps:
                    continue
                else:
                    break

            if n > 150:
                pmvColumn[i] = np.nan
                ppdColumn[i] = np.nan
                continue
            tcl = 100 * xn - 273

            # heat loss components
            HL1 = 3.05 * 0.001 * (5733 - 6.99 * mw - pa)  # heat loss diff.through skin
            if mw > 58.15:
                HL2 = 0.42 * (mw - 58.15)
            else:
                HL2 = 0

            HL3 = 1.7 * 0.00001 * m * (5867 - pa)
            HL4 = 0.0014 * m * (34 - tempAir)
            HL5 = 3.96 * fcl * (xn ** 4 - (tra / 100) ** 4)
            HL6 = fcl * hc * (tcl - tempAir)

            # calculate PMV and PPD
            TS = 0.303 * math.exp(-0.036 * m) + 0.028
            thermal_load = (mw - HL1 - HL2 - HL3 - HL4 - HL5 - HL6)

            if 'floatingAvgOutdoorTempColumn' not in locals():
                pmv = TS * thermal_load
            else:
                pmv = \
                    1.484 + 0.0276 * thermal_load \
                    - 0.960 * met \
                    - 0.0342 * self.df['Aussentemp_floating_average'][i] \
                    + 0.000226 * thermal_load * self.df['Aussentemp_floating_average'][i] \
                    + 0.0187 * met * self.df['Aussentemp_floating_average'][i] \
                    - 0.000291 * thermal_load * met * self.df['Aussentemp_floating_average'][i]

            ppd = 100 - 95 * math.exp(-0.03353 * pmv ** 4 - 0.2179 * pmv ** 2)

            pmvColumn[i] = pmv
            ppdColumn[i] = ppd

        return pmvColumn, ppdColumn
