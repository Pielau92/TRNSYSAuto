import sys
import os
import shutil
import functions
import re
import multiprocessing
import time
import csv

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

        # filenames and directories
        self.dir_sim_variants_excel = os.path.dirname(self.path_sim_variants_excel)
        self.dir_base_folder = os.path.dirname(self.dir_sim_variants_excel)
        self.filename_sim_variants_excel = os.path.basename(self.path_sim_variants_excel).split('.')[0]
        # simulation series folder in same directory as base folder
        self.dir_sim_series = \
            os.path.join(self.dir_base_folder, self.filename_sim_variants_excel + '_' + self.current_time)

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

    def create_sim_series_folder(self):
        """Create simulation series folder.

        Creates a folder for storing data of the simulation series. Also fills the folder with the following:
        -   copy of the simulation variants Excel file
        -   separate subfolder for each simulation within the simulation series
        Each simulation subfolder contains a copy of each file and template necessary for the simulation. Some files are
        also modified afterwards, in order to apply specific simulation parameters from the simulation variants Excel.
        """

        for sim in self.sim_list:

            path_sim = os.path.join(self.dir_sim_series, sim)  # path of simulation folder
            os.makedirs(path_sim)  # create new simulation subfolder
            # shutil.copytree(dir_base_folder, path_sim)          # copy all files from base to simulation folder

            # region SOURCE/DESTINATION FILE PATHS FOR COPYING PROCESS

            src_file = [  # todo: dynamisch neu machen, u.U. Funktion draus machen
                os.path.join(self.dir_sim_variants_excel, 'templateDck.dck'),
                os.path.join(self.dir_sim_variants_excel, 'Lastprofil.txt'),
                os.path.join(self.dir_sim_variants_excel, 'b18', self.b18_series[sim]),
                os.path.join(self.dir_sim_variants_excel, 'Wetterdaten', self.weather_series[sim]),
                os.path.join(self.dir_sim_variants_excel, 'SzenarioAneu.txt'),
                os.path.join(self.dir_sim_variants_excel, 'Qelww_CHR55025.txt'),
                os.path.join(self.dir_sim_variants_excel, 'Windetc20190804.txt'),
                os.path.join(self.dir_sim_variants_excel, 'StrahlungBruck.txt')]

            dst_file = [
                os.path.join(path_sim, 'templateDck.dck'),
                os.path.join(path_sim, 'Lastprofil.txt'),
                os.path.join(path_sim, self.b18_series[sim]),
                os.path.join(path_sim, self.weather_series[sim]),
                os.path.join(path_sim, 'SzenarioAneu.txt'),
                os.path.join(path_sim, 'Qelww_CHR55025.txt'),
                os.path.join(path_sim, 'Windetc20190804.txt'),
                os.path.join(path_sim, 'StrahlungBruck.txt')]

            # endregion

            # copy specified files into simulation folder
            for index in range(len(src_file)):
                shutil.copy(src_file[index], dst_file[index])

            # region FIND AND REPLACE WEATHER DATA FILE NAME IN .dck FILE todo: eigene Funktion draus machen

            with open(dst_file[0], 'r') as file:
                text = file.read()  # read file
                new_text = re.sub(r'(\*ASSIGN "tm2")', r'ASSIGN "' + self.weather_series[sim] + '"', text)
            with open(dst_file[0], 'w') as file:
                file.write(new_text)  # overwrite file

            # endregion

            # region FIND AND REPLACE .b18/b.17 FILE NAME IN .dck FILE

            with open(dst_file[0], 'r') as file:
                text = file.read()  # read file
                new_text = re.sub(r'(\*ASSIGN "b17")', r'ASSIGN "' + self.b18_series[sim] + '"', text)
            with open(dst_file[0], 'w') as file:
                file.write(new_text)  # overwrite file

            # endregion

            # region FIND AND REPLACE PARAMETERS IN .dck FILE
            functions.find_and_replace_text(dst_file[0], r'(@\w+)\s*=\s*([\d.]+)', self.df_dck.loc[sim])
            functions.find_and_replace_text(dst_file[0], r'(\*ASSIGN "b17")', self.df_dck.loc[sim])

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
        app.connect(title="Öffnen", timeout=self.timeout)

        # insert .dck file path
        app.Öffnen.FileNameEdit.set_edit_text(path_dck_file)

        # press start button
        Button = app.Öffnen.child_window(title="Öffnen", auto_id="1", control_type="Button").wrapper_object()
        Button.click_input()

        # wait for the simulation window to open
        while len(app.windows()) < 1:
            time.sleep(1)

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
        os.remove(os.path.join(path_sim, 'out11.txt'))
        os.remove(os.path.join(path_sim, 'out8.txt'))
        os.remove(os.path.join(path_sim, 'out6.txt'))
        os.remove(os.path.join(path_sim, 'out7.txt'))
        os.remove(os.path.join(path_sim, 'out10.txt'))
        os.remove(os.path.join(path_sim, 'Speicher1_step.out'))

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

        processes = []  #todo: benutzen oder löschen
        lock = multiprocessing.Lock()

        while not all(self.sim_success):    # check if any simulation has not been simulated successfully yet

            for index in range(len(self.sim_list)):
                sim = self.sim_list[index]  # name of simulation
                path_dck = os.path.join(self.dir_sim_series, sim, self.filename_dck_template)   # path of dck-file

                # self.start_sim(path_dck, lock)  #debug

                if not self.sim_success[index]:
                    # create a new process instance
                    process = multiprocessing.Process(target=self.start_sim,
                                                      args=(path_dck, lock))  # todo: Schleife auf Basis von processes?
                    processes.append(process)   #todo: benutzen oder löschen
                    with lock:
                        start_time = time.time()
                        while len(multiprocessing.active_children()) >= self.multiprocessing_max:
                            time.sleep(5)   # pause until number of active simulations drops below maximum
                            if time.time() - start_time > self.timeout:
                                sys.exit('Timeout of ' + str(self.timeout) + ' sec reached, program ended.')
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

        for index in range(len(self.sim_list)):
            sim = self.sim_list[index]

            path_output = os.path.join(self.dir_sim_series, sim, 'out5.txt')    # path of output file
            try:
                with open(path_output) as f:
                    reader = csv.reader(f, delimiter="\t")
                    d = list(reader)

                # simulation was successful, if hourly data is complete (8760 entries)
                self.sim_success[index] = not len(d) < 8762
            except:  # no file found todo: Exception hinzufügen
                self.sim_success[index] = False
