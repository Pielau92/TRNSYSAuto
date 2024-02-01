import sys
import os
import shutil
import functions
import re
import multiprocessing
import time
import csv
import logging

import pandas as pd

from pywinauto.application import Application
from datetime import datetime


class SimulationSeries: #todo: Durch Vererbung erweitern, damit auch andere Programme als TRNSYS leichter automatisiert werden können
    """Simulation series class.

    A Simulation series is a series of TRNSYS simulations, which are computed using multiprocessing. todo allgemeiner ausdrücken
    """

    def __init__(self, path_sim_variants_excel):
        """ Initialize simulation series object.

        A simulation series is saved in a folder in the same directory as the base folder (which contains templates,
        b17/18 files, dck file, the simulation variants Excel file, weather data, load profiles, ...). The simulation
        series folder is named after its corresponding simulation variants Excel file, followed by a timestamp at the
        time of the execution of the main method. TRNSYS has to be installed in order to perform the simulations
        successfully.

        Parameters
        ----------
        path_sim_variants_excel : str
            path to simulation variants Excel file corresponding to the simulation series.
        """

        # current time when the main.exe file was executed
        self.current_time = datetime.now().strftime('%d.%m.%Y_%H.%M')

        self.path_sim_variants_excel = path_sim_variants_excel  # path to simulation variants Excel file

        self.logger = None
        self.logger_filename = 'log.log'

        # filenames and directories
        self.dir_sim_variants_excel = os.path.dirname(self.path_sim_variants_excel)
        self.dir_base_folder = os.path.dirname(self.dir_sim_variants_excel)
        self.filename_sim_variants_excel = os.path.basename(self.path_sim_variants_excel).split('.')[0]
        # simulation series folder in same directory as base folder
        self.dir_sim_series = \
            os.path.join(self.dir_base_folder, self.filename_sim_variants_excel + '_' + self.current_time)
        self.dir_logfile = os.path.join(self.dir_sim_series, self.logger_filename)

        # simulation series data
        self.sim_list = None        # list of simulation variant names
        self.sim_success = None     # boolean list, documenting the successful simulation of each simulation variant
        self.df_dck = None          # pandas DataFrame with the simulation parameters to be replaced in the .dck Files
        self.b18_series = None      # pandas series with the .b18 data file names
        self.weather_series = None  # pandas series with the weather data file names

        # Settings
        self.settings = None                        # Pandas series with miscellaneous simulation settings
        self.path_exe = None                        # path to TRNSYS executable file
        self.name_excelsheet_sim_variants = None    # name of the Excel sheet containing the simulation variants table
        self.filename_dck_template = None           # name of the dck-File template
        self.timeout = None             # if timeout (sec) is reached without starting another simulation, stop program
        self.start_time_buffer = None   # time buffer (sec) between two simulations, for increased stability (optional)
        self.multiprocessing_max = None     # maximum number of simulations that can be calculated simultaneously
        self.autostart_evaluation = False   # start the evaluation routine for the simulation results afterwards if True

    def initialize_logging(self):
        """Initialize logging file."""
        # Set up logging file
        self.logger = logging.getLogger(__name__)
        logging.basicConfig(level=logging.DEBUG, filemode='w', filename=self.logger_filename)
        handler = logging.FileHandler(self.dir_logfile)
        self.logger.addHandler(handler)

    def import_settings_excel(self, filename_settings_excel, name_excelsheet_settings):
        """Import simulation series settings.

        Imports simulation series settings from the settings Excel file and applies them to the corresponding attributes
        of the SimulationSeries object. Parameter names in the settings Excel file (column "Parameter" must correspond
        to an attribute name of the SimulationSeries class.

        Parameters
        ----------
        filename_settings_excel : str
            Filename of the settings Excel file.
        name_excelsheet_settings : str
            Name of the Excel sheet within the settings Excel file, where the settings are stored.
        """
        # read Excel data
        excel_data = pd.ExcelFile(os.path.join(self.dir_sim_variants_excel, filename_settings_excel))

        # convert Excel data into pandas DataFrame
        df = excel_data.parse(name_excelsheet_settings, index_col=0)

        self.settings = df.Wert

        # apply imported settings
        self.apply_settings()

    def apply_settings(self): #todo automatischer Namenabgleich einführen, am besten mit Prüfung und Meldung bei Unstimmigkeiten
        """Apply imported settings to SimulationSeries object.

        Applies the imported settings from the settings Excel file to attributes of the SimulationSeries object with the
        same name.
        """

        self.path_exe = self.settings.loc['path_exe']
        self.name_excelsheet_sim_variants = self.settings.loc['name_excelsheet_sim_variants']
        self.filename_dck_template = self.settings.loc['filename_dck_template']
        self.timeout = self.settings.loc['timeout']
        self.start_time_buffer = self.settings.loc['start_time_buffer']
        self.autostart_evaluation = bool(self.settings.loc['autostart_evaluation'])

        if self.settings.loc['multiprocessing_max'] == 'auto':
            self.multiprocessing_max = multiprocessing.cpu_count()
        else:
            self.multiprocessing_max = self.settings.loc['multiprocessing_max']

    def import_input_excel(self):
        """Import simulation variants Excel file."""

        # read input Excel file
        excel_data = pd.ExcelFile(self.path_sim_variants_excel)

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

        # initialize simulation success flags
        self.sim_success = [False] * len(self.sim_list)

        # convert index into string (for stability reasons)
        self.weather_series.index = self.weather_series.index.map(str)
        self.b18_series.index = self.b18_series.index.map(str)
        self.df_dck.index = self.df_dck.index.map(str)

    def create_sim_series_folder(self):
        """Create simulation series folder.

        Creates a folder for storing data of the simulation series. Also fills the folder with the following:
        -   copy of the simulation variants Excel file
        -   separate subfolder for each simulation within the simulation series
        Each simulation subfolder contains a copy of each file and template necessary for the simulation. Some files are
        also modified afterwards, in order to apply specific simulation parameters from the simulation variants Excel.
        """

        for sim_index in range(0, len(self.sim_list)):
            sim = self.sim_list[sim_index]

            path_sim = os.path.join(self.dir_sim_series, sim)  # path of simulation folder
            os.makedirs(path_sim)  # create new simulation subfolder

            # region SOURCE/DESTINATION FILE PATHS FOR COPYING PROCESS

            file_list = [self.filename_dck_template, 'Lastprofil.txt', 'SzenarioAneu.txt','Qelww_CHR55025.txt',
                         'Windetc20190804.txt', 'StrahlungBruck.txt']
            src_file_list = file_list + \
                            [os.path.join('b18', self.b18_series[sim]),
                             os.path.join('Wetterdaten', self.weather_series[sim])]
            dst_file_list = file_list + [self.b18_series[sim], self.weather_series[sim]]

            # endregion

            # copy specified files into simulation folder
            for file_index in range(len(src_file_list)):
                try:
                    shutil.copy(
                        os.path.join(self.dir_sim_variants_excel, src_file_list[file_index]),
                        os.path.join(path_sim, dst_file_list[file_index]))
                except FileNotFoundError:
                    self.logger.error('File ' + os.path.join(self.dir_sim_variants_excel, src_file_list[file_index]
                                                             + ' could not be found.'))
                    self.sim_success[sim_index] = True      #todo: als Übergangslösung wird eine erfolgreiche Simulation vorgetäuscht, um Simulationen wo kritische Daten fehlen überspringen zu können. Langfrisitg soll sim_success aber nur tatsächlich erfolgreiche Simulationen dokumentieren!

            path_dck = os.path.join(path_sim, dst_file_list[0])
            # find and replace weather data file name in .dck file
            functions.find_and_replace(
                path_dck, pattern=r'(\*ASSIGN "tm2")', repl=r'ASSIGN "' + self.weather_series[sim] + '"')

            # find and replace .b17/.b18 file name in .dck file
            functions.find_and_replace(
                path_dck, pattern=r'(\*ASSIGN "b17")', repl=r'ASSIGN "' + self.b18_series[sim] + '"')

            # region FIND AND REPLACE PARAMETERS IN .dck FILE

            functions.find_and_replace_text(path_dck, r'(@\w+)\s*=\s*([\d.]+)', self.df_dck.loc[sim])
            functions.find_and_replace_text(path_dck, r'(\*ASSIGN "b17")', self.df_dck.loc[sim])

            # endregion

        # copy simulation variants Excel file into simulation series folder
        shutil.copy(self.path_sim_variants_excel, self.dir_sim_series)

    def start_sim(self, path_dck_file, lock):
        """Start simulation.

        Starts a TRNSYS simulation using a specified dck-file. Also, a lock is passed which ensures no other simulation
        starts until a specific point is reached. In this case, the lock is released as soon as the program presses the
        button "Öffnen" from the explorer window. Optionally, start_time_buffer acts as a time buffer before releasing
        the lock.

        Parameters
        ----------
        path_dck_file : str
            Path of the dck-File necessary for the TRNSYS simulation.
        lock : multiprocessing.Lock
            Lock object from the multiprocessing module.
        """

        # start application
        app = Application(backend='uia')
        app.start(self.path_exe)

        try:
            app.connect(title="Öffnen", timeout=2) #self.timeout)

            # insert .dck file path
            app.Öffnen.FileNameEdit.set_edit_text(path_dck_file)

            # press start button
            Button = app.Öffnen.child_window(title="Öffnen", auto_id="1", control_type="Button").wrapper_object()
            Button.click_input()

            # wait for the simulation window to open
            while len(app.windows()) < 1:
                time.sleep(1)

        except Exception:   #TimeoutError:
            lock.release()
            return

        # add a time buffer before releasing the lock, which delays the next simulation
        time.sleep(self.start_time_buffer)
        lock.release()

        # check if simulation has ended and asks if the online plotter should be closed
        interval = 5  # checking interval in seconds
        start_time = time.time()
        while time.time() - start_time < self.timeout and len(app.windows()) < 2 and app.is_process_running():
            time.sleep(interval)

        app.kill()  # close window
        time.sleep(5)

        # region DELETE REDUNDANT FILES

        path_sim = os.path.dirname(path_dck_file)
        # os.remove(path_dck_file[:-3] + 'lst')
        redundant_file_list =['out11.txt', 'out8.txt', 'out6.txt', 'out7.txt', 'out10.txt', 'Speicher1_step.out']
        for redundant_file in redundant_file_list:
            try:
                os.remove(os.path.join(path_sim, redundant_file))
            except FileNotFoundError:
                pass

        # endregion

    def start_sim_series(self):
        """Start simulation series.

        Creates a simulation series folder and starts the calculation of the simulation series, multiple simulation may
        run simultaneously, depending on the class attribute "multiprocessing_max". After all simulations are done, the
        method checks if all simulations were calculated successfully. If needed, unsuccessful simulations are
        calculated and checked again (this process is repeated until all simulations were calculated successfilly).
        """

        # create simulation series folder
        self.create_sim_series_folder()

        lock = multiprocessing.Lock()

        while not all(self.sim_success):    # check if any simulation has not been simulated successfully yet

            self.logger.info('Starting simulation series from "{}"'.format(self.filename_sim_variants_excel))

            for index in range(len(self.sim_list)):

                if not self.sim_success[index]:
                    sim = self.sim_list[index]  # name of simulation
                    path_dck = os.path.join(self.dir_sim_series, sim, self.filename_dck_template)  # path of dck-file

                    # self.start_sim(path_dck, lock)  # for debugging only

                    # create a new process instance
                    process = multiprocessing.Process(target=self.start_sim,
                                                      args=(path_dck, lock))
                    with lock:
                        start_time = time.time()
                        while len(multiprocessing.active_children()) >= self.multiprocessing_max:
                            time.sleep(5)   # pause until number of active simulations drops below maximum
                            if time.time() - start_time > self.timeout:
                                sys.exit('Timeout of ' + str(self.timeout) + ' sec reached, program ended.')
                        time.sleep(5)
                        process.start()     # start process
                    lock.acquire()

            # after all simulations have started, wait until all are done
            while len(multiprocessing.active_children()) > 0:
                time.sleep(5)
                if time.time() - start_time > self.timeout:
                    sys.exit('Timeout of ' + str(self.timeout) + ' sec reached, program ended.')

            # check for each simulation if it was successful
            self.check_sim_success()

    def check_sim_success(self):
        """Check simulation success.

        Checks for each simulation inside the simulation series, if the simulation was calculated successfully. If so,
        its sim_success flag is switched from False to True."""

        self.logger.info('Checking for failed simulations')

        for index in range(len(self.sim_list)):
            sim = self.sim_list[index]

            path_output = os.path.join(self.dir_sim_series, sim, 'out5.txt')    # path of output file
            if not self.sim_success[index]:
                try:
                    with open(path_output) as f:
                        reader = csv.reader(f, delimiter="\t")
                        d = list(reader)

                    # simulation was successful, if hourly data is complete (8760 entries)
                    self.sim_success[index] = not len(d) < 8762
                except FileNotFoundError:  # no file found
                    self.sim_success[index] = False

        # log simulation success status
        if all(self.sim_success):
            self.logger.info(('"{}" completed successfully'.format(self.filename_sim_variants_excel)))
        else:
            self.logger.info('{} out of {} simulations completed successfully'.format(
                sum(self.sim_success), len(self.sim_success)))