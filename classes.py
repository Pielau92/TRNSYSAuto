import sys
import os
import shutil
import functions
import re
import multiprocessing
import time
import csv
# import psutil

import pandas as pd

from pywinauto.application import Application
from datetime import datetime


class SimulationSeries:
    """ Simulation series class

    A Simulation series is a series of TRNSYS simulations, which are computed in a linear fashion or using
    multiprocessing.
    """

    def __init__(self, path_sim_variants_excel):
        """ Initialize simulation series object

        A simulation series is saved in a folder at the same level as the base folder (which contains templates, b17/18
        files, dck file, the Excel input file, weather data, load profiles, ...). The simulation series folder is named
        using a timestamp corresponding to the time of the execution of main.exe. TRSYS has to be installed in order to
        start the simulations.

        Parameters
        ----------
        dir_base_folder : string
            Current directory of the executed main.exe file
        path_exe : string
            Path to the TRNSYS simulation executable file
        name_excel_file : string
            File name of the input Excel file inside the base folder
        filename_dck_template : string
            File name of the .dck template file
        timeout : int
            Time limitation in sec for a single simulation, before it is canceled and closed #todo(0 means no timeout)
        cpu_threshold : int
            Threshold of the cpu usage that must be undercut to initiate the next simulation
        """

        self.dir_base_folder = None
        self.dir_sim_variants_excel = None

        self.path_sim_variants_excel = path_sim_variants_excel
        self.path_exe = None
        self.dir_sim_series = None
        self.dir_base_folder = None

        self.filename_sim_variants_excel = None
        self.filename_dck_template = None

        self.name_excelsheet = None
        self.name_base_folder = None

        self.current_time = datetime.now().strftime('%d.%m.%Y_%H.%M')   # current time when the main.exe file was executed
        self.timeout = None
        self.cpu_threshold = None
        self.start_time_buffer = None
        self.settings = None  # Pandas series with simulation settings
        self.sim_list = None  # list of simulation variant names
        self.sim_success = None
        self.df_dck = None  # pandas DataFrame with the simulation parameters to be replaced in the .dck Files
        self.b18_series = None  # Series with the .b18 data file names
        self.weather_series = None  # Series with the weather data file names
        self.autostart_evaluation = False

    def set_paths(self):

        self.dir_sim_variants_excel = os.path.dirname(self.path_sim_variants_excel)
        self.dir_base_folder = os.path.dirname(self.dir_sim_variants_excel)
        self.filename_sim_variants_excel = os.path.basename(self.path_sim_variants_excel).split('.')[0]

        # simulation series folder in same directory as base folder
        self.dir_sim_series = \
            os.path.join(self.dir_base_folder, self.filename_sim_variants_excel + '_' + self.current_time)

    def import_settings_excel(self):

        excel_data = pd.ExcelFile(
            os.path.join(self.dir_sim_variants_excel, 'Einstellungen.xlsx'))  # read input Excel file
        df = excel_data.parse('Einstellungen', index_col=0)  # Excel data as DataFrame
        self.settings = df.Wert

    def set_settings(self):

        self.path_exe = self.settings.loc['path_exe']
        self.name_excelsheet = self.settings.loc['name_excelsheet_sim_variants']
        self.name_base_folder = self.settings.loc['name_base_folder']
        self.filename_dck_template = self.settings.loc['filename_dck_template']
        self.timeout = self.settings.loc['timeout']
        self.start_time_buffer = self.settings.loc['start_time_buffer']
        self.cpu_threshold = self.settings.loc['cpu_threshold']

        self.autostart_evaluation = bool(self.settings.loc['autostart_evaluation'])

        if self.settings.loc['multiprocessing_max'] == 'auto':  # todo Prüfung ob sinnvoller Wert eingegeben wurde
            multiprocessing.cpu_count()
        else:
            self.multiprocessing_max = self.settings.loc['multiprocessing_max']

    def import_input_excel(self):
        """ Input Excel file import routine."""

        excel_data = pd.ExcelFile(self.path_sim_variants_excel)  # read input Excel file
        df = excel_data.parse(self.name_excelsheet, index_col=0)  # Excel data as DataFrame

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

        # initialize simulation success flag
        self.sim_success = [False] * len(self.sim_list)

    def create_sim_folders(self):
        """Create simulation folder.

        Parameters
        ----------
        sim : string
            name of the simulation variant
        """
        for sim in self.sim_list:

            path_sim = os.path.join(self.dir_sim_series, sim)  # path of simulation folder
            os.makedirs(path_sim)  # create new directory for simulation
            # shutil.copytree(dir_base_folder, path_sim)          # copy all files from base to simulation folder

            # region SOURCE/DESTINATION FILE PATHS FOR COPYING PROCESS

            src_file = [  # todo: dynamisch neu machen
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

            # region FIND AND REPLACE WEATHER DATA FILE NAME IN .dck FILE

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
            functions.find_and_replace_param(dst_file[0], r'(@\w+)\s*=\s*([\d.]+)', self.df_dck.loc[sim])
            functions.find_and_replace_param(dst_file[0], r'(\*ASSIGN "b17")', self.df_dck.loc[sim])

            # endregion

        # copy Input Excel file into simulation series folder
        shutil.copy(self.path_sim_variants_excel, self.dir_sim_series)

    def start_sim(self, path_dck_file):

        app = Application(backend='uia')
        app.start(self.path_exe)
        app.connect(title="Öffnen", timeout=60 * 30)

        app.Öffnen.FileNameEdit.set_edit_text(path_dck_file)  # insert .dck file path
        Button = app.Öffnen.child_window(title="Öffnen", auto_id="1", control_type="Button").wrapper_object()
        Button.click_input()

        time.sleep(1)  # give the simulation some time to start

        # region wait for simulation completion
        # Solution 1: wait until the cpu kernel is no longer needed (simulation has ended)
        # app.wait_cpu_usage_lower(threshold=0, timeout=60 * 15)

        # Solution 2: wait until second window pops up
        # (simulation has ended and asks if the online plotter should be closed)
        interval = 5  # checking interval in seconds
        timeout = 60 * 60  # timeout in seconds
        start_time = time.time()
        while not len(app.windows()) > 1 and time.time() - start_time < timeout:
            time.sleep(interval)
        # endregion

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

        processes = []

        self.create_sim_folders()

        while not all(self.sim_success):

            for index in range(len(self.sim_list)):
                sim = self.sim_list[index]
                path_dck = os.path.join(self.dir_sim_series, sim, self.filename_dck_template)
                # path_output = os.path.join(self.dir_sim_series, sim, 'out5.txt')
                # self.start_sim(path_dck)  #debug
                if not self.sim_success[index]:
                    # create a new process instance
                    process = multiprocessing.Process(target=self.start_sim,
                                                      args=(path_dck,))  # todo: Schleife auf Basis von processes?
                    processes.append(process)
                    process.start()
                    # start_time = time.time()
                    # while True:
                    #     cpu_percent = psutil.cpu_percent(interval=1)
                    #     if cpu_percent < self.cpu_threshold:
                    #         process.start()
                    #         break
                    #     elif time.time() - start_time > self.timeout:
                    #         sys.exit('Timeout of ' + str(self.timeout) + ' sec reached, program ended.')

                    time.sleep(self.start_time_buffer)
                    start_time = time.time()
                    if len(multiprocessing.active_children()) >= self.multiprocessing_max:
                        process.join()

                    # while len(multiprocessing.active_children()) >= self.multiprocessing_max:
                    #     time.sleep(self.start_time_buffer)
                    # if time.time() - start_time > self.timeout:
                    # sys.exit('Timeout of ' + str(self.timeout) + ' sec reached, program ended.')

            # wait until all simulations are done first, otherwise problems when calculating small series
            start_time = time.time()
            while len(multiprocessing.active_children()) > 0:
                time.sleep(5)
                if time.time() - start_time > self.timeout:
                    sys.exit('Timeout of ' + str(self.timeout) + ' sec reached, program ended.')

            self.check_sim_success()

    def check_sim_success(self):

        for index in range(len(self.sim_list)):
            sim = self.sim_list[index]

            path_output = os.path.join(self.dir_sim_series, sim, 'out5.txt')
            try:
                with open(path_output) as f:
                    reader = csv.reader(f, delimiter="\t")
                    d = list(reader)

                self.sim_success[index] = not len(d) < 8762
            except:  # no file found
                self.sim_success[index] = False
