import sys
import os
import shutil
import multiprocessing
import time
import csv
import logging
import math
import functions
import pickle
import openpyxl

import numpy as np
import pandas as pd

from pywinauto.application import Application
from datetime import datetime

# todo: docstrings aktualisieren
class SimulationSeries:
    """Simulation series class.

    A Simulation series is a series of TRNSYS simulations, which can be computed using multiprocessing.
    """

    def __init__(self, path_original_sim_variants_excel):
        """ Initialize simulation series object.

        A simulation series is saved in a directory in the same directory as the base folder (which contains templates,
        b17/18 files, dck file, the simulation variants Excel file, weather data, load profiles, ...). The simulation
        series directory is named after its corresponding simulation variants Excel file, followed by a timestamp at the
        time of the execution of the main method. TRNSYS has to be installed in order to perform the simulations
        successfully.

        Parameters
        ----------
        path_original_sim_variants_excel : str
            path to original simulation variants Excel file corresponding to the simulation series.
        """

        # region SET NAMES

        # filenames
        self.filename_logger = 'log.log'
        self.filename_trnsys_output = 'out5.txt'
        self.filename_settings_excel = 'Einstellungen.xlsx'
        self.filename_savefile = 'SimulationSeries.pickle'

        # Excel sheet names
        self.sheet_name_settings = 'Einstellungen'
        self.sheet_name_variant_input = 'Rohdaten'
        self.sheet_name_calculation = 'Berechn1'
        self.sheet_name_cumulative_input = 'Rohinputs'
        self.sheet_name_zone_1_input = 'Zusamm1'
        self.sheet_name_zone_3_input = 'Zusamm3'
        self.sheet_name_zone_1_with_operating_time = 'Zone1_Betrieb'
        self.sheet_name_zone_1_without_operating_time = 'Zone1ges'
        self.sheet_name_zone_3_with_operating_time = 'Zone3_Betrieb'
        self.sheet_name_zone_3_without_operating_time = 'Zone3ges'

        # column names
        self.var_list_zone1 \
            = ['Period', 'ta', 'tzone1', 'TMSURF_ZONE1', 'relh1', 'vel1', 'pmv1', 'ppd1', 'clo1', 'met1', 'work1']
        self.var_list_zone2 \
            = ['Period', 'ta', 'tzone1.1', 'TMSURF_ZONE1.1', 'relh2', 'vel2', 'pmv2', 'ppd2', 'clo2', 'met2', 'work2']
        self.var_list_zone3 \
            = ['Period', 'ta', 'tzone1.2', 'TMSURF_ZONE1.2', 'relh3', 'vel3', 'pmv3', 'ppd3', 'clo3', 'met3', 'work3']
        self.col_headers_result_column = \
            ['top1', 'top2', 'top3', 'Qventfges', 'qvolgesh', 'qc1', 'qc2', 'qc3', 'pmv1', 'pmv2', 'pmv3']
        self.col_headers_trnsys_output \
            = ['ta', 'top1', 'top2', 'top3', 'tzone1', 'tzone2', 'tzone3', 'Qventfges', 'qvolgesh', 'qc1', 'qc2', 'qc3',
               'qh1', 'qh2', 'qh3', 'pmv1', 'pmv2', 'pmv3', 'ppd1', 'ppd2', 'ppd3', 'clo1', 'clo2', 'clo3', 'met1',
               'met2', 'met3']
        self.col_headers_sim_variant \
            = ['ta', 'top1', 'top2', 'top3', 'tzone1', 'tzone2', 'tzone3', 'Qventfges', 'qvolgesh', 'qc1', 'qc2', 'qc3',
               'qh1', 'qh2', 'qh3', 'pmv1', 'pmv2', 'pmv3', 'ppd1', 'ppd2', 'ppd3', 'clo1', 'clo2', 'clo3', 'met1',
               'met2', 'met3', 'schweiker_pmv1', 'schweiker_pmv2', 'schweiker_pmv3', 'schweiker_ppd1', 'schweiker_ppd2',
               'schweiker_ppd3', 'schweiker_clo1', 'schweiker_clo2', 'schweiker_clo3', 'schweiker_met1',
               'schweiker_met2', 'schweiker_met3']

        # endregion

        # current time when the main.exe file was executed
        self.execution_time = datetime.now().strftime('%d.%m.%Y_%H.%M')

        # path to original simulation variants Excel file within the base directory "Basisordner"
        self.path_original_sim_variants_excel = path_original_sim_variants_excel

        self.filename_sim_variants_excel = os.path.basename(self.path_original_sim_variants_excel).split('.')[0]
        self.dirname_sim_series = self.filename_sim_variants_excel + '_' + self.execution_time
        self.path_sim_series_dir = os.path.abspath(self.dirname_sim_series)  # path to simulation series directory

        # simulation series data
        self.sim_list = None  # list of simulation variant names
        self.sim_success = None  # boolean list, documenting the successful simulation of each simulation variant
        self.sim_ignore = None  # boolean list, documenting the simulation variants that are to be ignored
        self.df_dck = None  # pandas DataFrame with the simulation parameters to be replaced in the .dck Files
        self.b18_series = None  # pandas series with the .b18 data file names
        self.weather_series = None  # pandas series with the weather data file names

        # settings
        self.settings = None  # Pandas series with miscellaneous simulation settings
        self.path_exe = None  # path to TRNSYS executable file
        self.name_excelsheet_sim_variants = None  # name of the Excel sheet containing the simulation variants table
        self.filename_dck_template = None  # name of the dck-File template
        self.timeout = None  # if timeout (sec) is reached without starting another simulation, stop program
        self.start_time_buffer = None  # time buffer (sec) between two simulations, for increased stability (optional)
        self.multiprocessing_max = None  # maximum number of simulations that can be calculated simultaneously
        self.autostart_evaluation = False  # start the evaluation routine for the simulation results afterwards if True
        self.filenames_redundant = None  # list of redundant TRNSYS files that are to be deleted after the simulation
        self.eval_save_interval = None  # the evaluation progress is saved after each save interval

        # region EVALUATION

        self.date_df = functions.create_date_column(2024)
        self.variant_parameter_df = None
        self.eval_success = None

        self.variant_result_columns = pd.DataFrame()
        self.zone_1_with_df = pd.DataFrame()
        self.zone_1_without_df = pd.DataFrame()
        self.zone_3_with_df = pd.DataFrame()
        self.zone_3_without_df = pd.DataFrame()

        # endregion

        self.logger = None

    # region PROPERTIES

    @property
    def cwd(self):
        """Current working directory."""
        return os.getcwd()

    @property
    def path_base_dir(self, base_name='Basisordner'):
        """Path to base directory "Basisordner"."""
        return os.path.abspath(base_name)

    @property
    def path_sim_variants_excel(self):
        """Path to simulation series Excel file, copied from the base directory "Basisordner"."""
        return os.path.join(self.path_sim_series_dir, self.filename_sim_variants_excel)

    @property
    def path_logfile(self):
        """Path to logfile."""
        return os.path.join(self.path_sim_series_dir, self.filename_logger)

    @property
    def path_evaluation_save_dir(self, dir_name='evaluation'):
        """Path to directory, where evaluation results are saved."""
        return os.path.join(self.path_sim_series_dir, dir_name)

    @property
    def path_cumulative_evaluation_save_file(self, filename='gesamt.xlsx'):
        """Path to cumulative evaluation file."""
        return os.path.join(self.path_evaluation_save_dir, filename)

    @property
    def path_cumulative_evaluation_template(self, filename='Auswertung_Gesamt.xlsx'):
        """Path to cumulative evaluation template file."""
        return os.path.join(self.path_base_dir, filename)

    @property
    def path_variant_evaluation_template(self, filename='Auswertung_Variante.xlsx'):
        """Path to variant evaluation template file."""
        return os.path.join(self.path_base_dir, filename)

    @property
    def path_savefile(self):
        return os.path.join(self.path_sim_series_dir, self.filename_savefile)

    # endregion

    def setup_simulation(self):
        """Set simulation series up.

        Setting up the simulation series is only necessary once, continuing the simulation at a later time does not need
        an additional setup."""

        # import and apply settings Excel file
        self.import_settings_excel()

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
        time does not need an additional setup."""

        # create evaluation directory
        os.makedirs(self.path_evaluation_save_dir, exist_ok=True)

        # create copy of cumulative evaluation file template
        shutil.copy(self.path_cumulative_evaluation_template, self.path_cumulative_evaluation_save_file)

        # initialize evaluation success list
        self.eval_success = [False] * len(self.sim_list)

    def create_dir_sim_series(self):
        """Create simulation series directory.

        Creates a directory to save the results of the simulation series. Also fills the directory with the following:
        -   copy of the simulation variants Excel file
        -   separate subdirectories for each simulation within the simulation series, containing a copy of each file and
        template necessary for the simulation.

        Afterwards, the copied template files are modified in order to apply specific simulation parameters from the
        simulation variants Excel file.
        """

        def create_sim_subdirectory():
            """Create simulation subdirectory within simulation series directory.

            Creates a subdirectory within the simulation series directory, with a copy of all template files from the
            base directory "Basisordner". Each subdirectory corresponds to one simulation.
            """

            os.makedirs(path_sim)  # create new empty simulation subdirectory

            # source paths
            src_file_list = file_list + [os.path.join('b18', self.b18_series[sim]),
                                         os.path.join('Wetterdaten', self.weather_series[sim])]
            # destination paths
            dst_file_list = file_list + [self.b18_series[sim], self.weather_series[sim]]

            # copy specified files into simulation subdirectory
            for file_index in range(len(src_file_list)):
                try:
                    shutil.copy(
                        os.path.join(self.path_base_dir, src_file_list[file_index]),
                        os.path.join(path_sim, dst_file_list[file_index]))
                except FileNotFoundError:
                    message = 'File ' + os.path.join(self.path_base_dir, src_file_list[file_index] +
                                                     ' could not be found, simulation variant added to ignore' 'list.')
                    self.logger.error(message)
                    print(message)

                    self.sim_ignore[index] = True  # simulation variant will be ignored

        def overwrite_dck_file_parameters():
            """Overwrite parameters inside .dck File.

            Overwrites the parameters inside the .dck File, according to the corresponding simulation description in
            the simulation variants Excel file.
            """

            # find and replace weather data file name in .dck file
            functions.find_and_replace(
                path_dck, pattern=r'(\*ASSIGN "tm2")', replacement=r'ASSIGN "' + self.weather_series[sim] + '"')

            # find and replace .b17/.b18 file name in .dck file
            functions.find_and_replace(
                path_dck, pattern=r'(\*ASSIGN "b17")', replacement=r'ASSIGN "' + self.b18_series[sim] + '"')

            # find and replace
            functions.find_and_replace_parameter_values(path_dck, r'(@\w+)\s*=\s*([\d.]+)', self.df_dck.loc[sim])
            functions.find_and_replace_parameter_values(path_dck, r'(\*ASSIGN "b17")', self.df_dck.loc[sim])

        # create new directory for the simulation series
        os.makedirs(self.path_sim_series_dir)

        # initialize logging file
        self.initialize_logging()

        # copy simulation variants Excel file into simulation series directory
        shutil.copy(os.path.join(self.path_original_sim_variants_excel), self.path_sim_variants_excel)

        # file name list of files to be copied into simulation subdirectories
        file_list = [self.filename_dck_template, 'Lastprofil.txt', 'SzenarioAneu.txt', 'Qelww_CHR55025.txt',
                     'Windetc20190804.txt', 'StrahlungBruck.txt']

        for index, sim in enumerate(self.sim_list):
            path_sim = os.path.join(self.path_sim_series_dir, sim)  # path of simulation subdirectory
            path_dck = os.path.join(path_sim, self.filename_dck_template)  # path of .dck file

            create_sim_subdirectory()
            overwrite_dck_file_parameters()

    def initialize_logging(self):
        """Initialize logging file."""

        self.logger = logging.getLogger(__name__)
        logging.basicConfig(level=logging.DEBUG, filemode='w', filename=self.filename_logger)
        handler = logging.FileHandler(self.path_logfile)
        self.logger.addHandler(handler)

        self.logger.info(('Log file created successfully in {}.'.format(self.path_logfile)))

    def import_settings_excel(self):
        """Import simulation series settings from settings Excel file.

        Imports simulation series settings from the settings Excel file and applies them to the corresponding attributes
        of the SimulationSeries object. Parameter names in the settings Excel file (column "Parameter") must correspond
        to an attribute name of the SimulationSeries class.
        """

        # read Excel data
        excel_data = pd.ExcelFile(os.path.join(self.path_base_dir, self.filename_settings_excel))

        # convert Excel data into pandas DataFrame
        df = excel_data.parse(self.sheet_name_settings, index_col=0)

        # extract values from the column "Wert"
        self.settings = df.Wert

        # apply imported settings
        self.apply_settings()

    def apply_settings(self):
        """Apply imported settings to SimulationSeries object.

        Applies the imported settings from the settings Excel file to the corresponding attributes of the
        SimulationSeries object with the same name.
        """

        for index, value in enumerate(self.settings):
            attribute_name = self.settings.index[index]
            if not hasattr(self, attribute_name):
                raise AttributeError(f'Unknown setting "{attribute_name}" in settings Excel file found.')
            setattr(self, attribute_name, value)

        if self.multiprocessing_max == 'auto':
            self.multiprocessing_max = multiprocessing.cpu_count()

        self.filenames_redundant = self.filenames_redundant.split(', ')

    def import_sim_variants_excel(self):
        """Import simulation variants Excel file.

        Imports the simulation variants Excel file and applies the data to the SimulationSeries object.
        """

        # todo: hier wird zwei mal der Inhalt des Simulationsvariantenfiles importiert, auf 1 mal reduzieren
        # read simulation variant parameters
        self.variant_parameter_df = pd.read_excel(self.path_original_sim_variants_excel,
                                                  sheet_name='Simulationsvarianten')
        self.variant_parameter_df.columns = [str(parameter) for parameter in self.variant_parameter_df.columns]

        # read simulation variants Excel file
        excel_data = pd.ExcelFile(self.path_original_sim_variants_excel)

        # convert Excel data into pandas DataFrame
        df = excel_data.parse(self.name_excelsheet_sim_variants, index_col=0)

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

        # convert index into string (for stability reasons)
        self.weather_series.index = self.weather_series.index.map(str)
        self.b18_series.index = self.b18_series.index.map(str)
        self.df_dck.index = self.df_dck.index.map(str)

    def save(self):
        """Save SimulationSeries object in simulation series directory."""

        with open(self.path_savefile, 'wb') as file:
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

            Deletes redundant files to save disk space.
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
            success_message.wait('visible', timeout=60 * 10)
        except TimeoutError:
            pass  # goes ahead and closes window after time out

        app.kill()  # close window
        time.sleep(5)

        delete_redundant_files()

    def start_sim_series(self):
        """Start simulation series.

        Starts the calculation of the simulation series and creates a simulation series directory to store the
        simulation results. Multiple simulation may run simultaneously, depending on the class attribute
        "multiprocessing_max". After all simulations are done, the method checks if all simulations were calculated
        successfully. If needed, unsuccessful simulations are calculated and checked again (this process is repeated
        until all simulations were calculated successfully, unless some simulations are on the "ignore" list anyway).
        """

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
                    path_dck = os.path.join(self.path_sim_series_dir, sim,
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

    def check_sim_success(self):
        """Check simulation success.

        Checks for each simulation inside the simulation series, if the simulation was calculated successfully. If so,
        its sim_success flag is switched from False to True.
        """

        message = 'Checking for failed simulations'
        self.logger.info(message)
        print(message)

        for index in range(len(self.sim_list)):
            # path of output file
            path_output = os.path.join(self.path_sim_series_dir, self.sim_list[index], self.filename_trnsys_output)

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

        # cumulative evaluation
        self.cumulative_evaluation()

        # logger entry "finish"
        message = 'Evaluation done.'
        self.logger.info(message)
        print(message)

    def evaluate_variant(self, variant_name, variant_index):    # todo: Auswertungsergebnisse nicht mehr durch concat speichern, sondern explizit über Variantennamen

        def create_schweiker_model(var_list_zone, zone):
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

        path_variant_directory = os.path.join(self.path_sim_series_dir, variant_name)
        path_variant_file = os.path.join(path_variant_directory, self.filename_trnsys_output)
        save_path_variant_output = os.path.join(self.path_evaluation_save_dir, variant_name + '.xlsx')

        # region CHECK IF...

        # ...the trnsys output file is actually there
        if not os.path.exists(path_variant_file):
            message = f'File {path_variant_file} does not exist!'
            self.logger.error(message)
            print(message)
            return

        # ...the variant has a corresponding directory
        if variant_name not in self.sim_list:
            message = f'Did not find {variant_name} in {self.path_sim_variants_excel}'
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
        shutil.copy(self.path_variant_evaluation_template, save_path_variant_output)

        # save data
        functions.excel_export_variant_evaluation(
            self.sheet_name_variant_input, result, variant_name, save_path_variant_output, self.variant_parameter_df)

        # update excel to receive cross-referenced values and updates calculations
        functions.update_excel_file(save_path_variant_output)

        # create single column with all hourly values, for the cumulative evaluation Excel file
        result_column = functions.to_single_column(result[self.col_headers_result_column])

        # save single column
        self.variant_result_columns = pd.concat([self.variant_result_columns, result_column], axis=1)

        self.eval_success[variant_index] = True     # todo: eval_success durch DataFrame ersetzen, um über den Variantennamen anstatt index eintragen zu können. Dann index als inputparameter entfernen

        message = 'Finished evaluation for variant {}'.format(variant_name)
        self.logger.info(message)
        print(message)

    def cumulative_evaluation(self):    # todo: Auswertungsergebnisse nicht mehr durch concat speichern, sondern explizit über Variantennamen

        for variant_index, variant_name in enumerate(self.sim_list):

            save_path_variant_output = os.path.join(self.path_evaluation_save_dir, variant_name + '.xlsx')
            # region CUMULATIVE EVALUATION

            # read data from variant evaluation excel file, for the cumulative evaluation excel file
            self.zone_1_with_df = \
                pd.concat([self.zone_1_with_df,
                           pd.read_excel(save_path_variant_output, sheet_name=self.sheet_name_zone_1_input,
                                         usecols=[3], header=None, nrows=None, skiprows=None)], axis=1)
            self.zone_1_without_df = \
                pd.concat([self.zone_1_without_df,
                           pd.read_excel(save_path_variant_output, sheet_name=self.sheet_name_zone_1_input,
                                         usecols=[2], header=None, nrows=None, skiprows=None)], axis=1)
            self.zone_3_with_df = \
                pd.concat([self.zone_3_with_df,
                           pd.read_excel(save_path_variant_output, sheet_name=self.sheet_name_zone_3_input,
                                         usecols=[3], header=None, nrows=None, skiprows=None)], axis=1)
            self.zone_3_without_df = \
                pd.concat([self.zone_3_without_df,
                           pd.read_excel(save_path_variant_output, sheet_name=self.sheet_name_zone_3_input,
                                         usecols=[2], header=None, nrows=None, skiprows=None)], axis=1)

            # copy into cumulative evaluation excel file
            self.excel_export_cumulative_evaluation()

            # update cumulative excel
            functions.update_excel_file(self.path_cumulative_evaluation_save_file)

    def excel_export_cumulative_evaluation(self):
        """Write data into cumulative evaluation file."""

        def export(df, sheetname, startrow, startcol, header=False):
            df.to_excel(writer, sheet_name=sheetname, startrow=startrow, startcol=startcol, index=False, header=header)

        with pd.ExcelWriter(
                self.path_cumulative_evaluation_save_file, mode="a", engine="openpyxl", if_sheet_exists='overlay') \
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
