import sys
# todo: "sys.coinit_flags = 2 " wird hier glaube ich nicht mehr benötigt
# sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED - needed as tkinter has compatibility issues with pywinauto when creating .exe file
import os
import shutil
import functions
import pandas as pd
import re
import psutil
import multiprocessing
import time

from pywinauto.application import Application
from datetime import datetime


class SimulationSeries:
    """ Simulation series class

    A Simulation series is a series of TRNSYS simulations, which are computed in a linear fashion or using
    multiprocessing.
    """

    def __init__(self, current_dir, path_exe, name_excel_file, name_excelsheet, name_base_folder, filename_dck_template,
                 timeout=10*60, cpu_threshold=60):
        """ Initialize simulation series object

        A simulation series is saved in a folder at the same level as the base folder (which contains templates, b17/18
        files, dck file, the Excel input file, weather data, load profiles, ...). The simulation series folder is named
        using a timestamp corresponding to the time of the execution of main.exe. TRSYS has to be installed in order to
        start the simulations.

        Parameters
        ----------
        current_dir : string
            Current directory of the executed main.exe file
        path_exe : string
            Path to the TRNSYS simulation executable file
        name_excel_file : string
            File name of the input Excel file inside the base folder
        filename_dck : string
            File name of the .dck template file
        timeout : int
            Time limitation in sec for a single simulation, before it is canceled and closed #todo(0 means no timeout)
        cpu_threshold : int
            Threshold of the cpu usage that must be undercut to initiate the next simulation
        """

        self.current_dir = current_dir
        self.path_exe = path_exe
        self.filename_excel = name_excel_file
        self.name_excelsheet = name_excelsheet
        self.base_folder_name = name_base_folder
        self.filename_dck = filename_dck_template
        self.timeout = timeout
        self.cpu_threshold = cpu_threshold

        self.sim_list = None    # list of simulation variant names
        self.df_dck = None  # pandas DataFrame with the simulation parameters to be replaced in the .dck Files
        self.b18_series = None  # Series with the .b18 data file names
        self.weather_series = None  # Series with the weather data file names

        current_time = datetime.now().strftime('%d.%m.%Y_%H.%M')  # current time when the main.exe file was executed

        # paths
        self.path_sim_series = os.path.join(self.current_dir, current_time)     # simulation series folder
        self.path_base = os.path.join(self.current_dir, self.base_folder_name)          # base folder
        self.path_excel = os.path.join(self.path_base, self.filename_excel)     # input Excel file

    def import_input_excel(self):
        """ Input Excel file import routine."""

        excel_data = pd.ExcelFile(self.path_excel)  # read input Excel file
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

    def start_sim_series(self):
        """Start simulation series."""

        # copy input Excel file into simulation series folder
        shutil.copy(os.path.join(self.path_base, self.filename_excel), self.path_sim_series)

        for sim in self.sim_list:
            self.create_sim_folder(sim)     # create simulation folder
            path_dck = os.path.join(self.path_sim_series, sim, self.filename_dck)   # dck File path of the current sim

            start_time = time.time()
            while True:
                cpu_percent = psutil.cpu_percent(interval=1)
                if cpu_percent < self.cpu_threshold:
                    self.start_sim(path_dck)    # start simulation
                    break
                elif time.time() - start_time > self.timeout:
                    sys.exit('Timeout of ' + str(self.timeout) + ' sec reached, program ended.')
                time.sleep(5)

    def create_sim_folder(self, sim):
        """Create simulation folder.

        Parameters
        ----------
        sim : string
            name of the simulation variant
        """

        path_sim = os.path.join(self.path_sim_series, sim)  # path of simulation folder
        os.makedirs(path_sim)  # create new directory for simulation
        # shutil.copytree(path_base, path_sim)          # copy all files from base to simulation folder

        # region SOURCE/DESTINATION FILE PATHS FOR COPYING PROCESS

        src_file = [  #todo: dynamisch neu machen
            os.path.join(self.path_base, 'templateDck.dck'),
            os.path.join(self.path_base, 'Lastprofil.txt'),
            os.path.join(self.path_base, 'b18', self.b18_series[sim]),
            os.path.join(self.path_base, 'Wetterdaten', self.weather_series[sim]),
            os.path.join(self.path_base, 'SzenarioAneu.txt'),
            os.path.join(self.path_base, 'Qelww_CHR55025.txt'),
            os.path.join(self.path_base, 'Windetc20190804.txt'),
            os.path.join(self.path_base, 'StrahlungBruck.txt')]

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

    def start_sim(self, path_dck_file):

        path_dck_file = path_dck_file.replace("/", "\\")

        # path_file_raw = r'{}'.format(path_file)     # turn file path into raw string to avoid error messages

        app = Application(backend='uia')
        app.start(self.path_exe)
        app.connect(title="Öffnen", timeout=60 * 30)

        app.Öffnen.FileNameEdit.set_edit_text(path_dck_file)    # insert .dck file name
        Button = app.Öffnen.child_window(title="Öffnen", auto_id="1", control_type="Button").wrapper_object()
        Button.click_input()

        time.sleep(10) # give the simulation some time to start

        # wait until the cpu kernel is no longer needed (simulation has ended)
        app.wait_cpu_usage_lower(threshold=0, timeout=60 * 15)

        # for extra stability
        time.sleep(10)  # give the simulation some time to start
        app.wait_cpu_usage_lower(threshold=0, timeout=60 * 15)

        # print(app.cpu_usage())
        # quit simulation after 10 minutes
        # time.sleep(60 * 10)

        app.kill()  # close window
        time.sleep(10)   # give the simulation some time to quit
        # time.sleep(20)  # give the simulation some time to quit

        # region DELETE REDUNDANT FILES

        # os.remove(path_dck_file[:-3] + 'lst')
        # os.remove(os.path.join(path_sim, 'out11.txt'))
        # os.remove(os.path.join(path_sim, 'out8.txt'))
        # os.remove(os.path.join(path_sim, 'Speicher1_step.out'))

        # endregion

    def start_sim_series_par(self):
        """
#todo start_sim_series aufrufen, anstatt Code doppelt. Ggf. muss die Erstellung der Ordnerstruktur und das Starten der Simulation in 2 Funktionen aufgetrennt werden
        Returns
        -------

        """
        # copy Input Excel file into simulation series folder
        shutil.copy(os.path.join(self.path_base, self.filename_excel), self.path_sim_series)

        path_dck = list()
        for sim in self.sim_list:
            self.create_sim_folder(sim)
            path_dck.append (os.path.join(self.path_sim_series, sim, 'templateDck.dck'))

        # pool_obj = multiprocessing.Pool()
        # pool_obj.map(self.start_sim, path_dck)
        # pool_obj.close()

        for dck in path_dck:
            # create a new process instance
            process = multiprocessing.Process(target=self.start_sim, args=(dck,))

            start_time = time.time()

            while True:
                cpu_percent = psutil.cpu_percent(interval=1)
                if cpu_percent < self.cpu_threshold:
                    # start the process
                    process.start()
                    time.sleep(30)
                    break
                elif time.time() - start_time > self.timeout:
                    sys.exit('Timeout of ' + str(self.timeout) + ' sec reached, program ended.')
                time.sleep(5)

    def start_sim_series_par_fixed_amount(self):

        processes = []

        # copy Input Excel file into simulation series folder
        shutil.copy(os.path.join(self.path_base, self.filename_excel), self.path_sim_series)

        path_dck = list()
        for sim in self.sim_list:
            self.create_sim_folder(sim)
            path_dck.append(os.path.join(self.path_sim_series, sim, 'templateDck.dck'))

        for dck in path_dck:
            # create a new process instance
            process = multiprocessing.Process(target=self.start_sim, args=(dck,))

            start_time = time.time()
            processes.append(process)
            process.start()
            time.sleep(30)
            active_processes = multiprocessing.active_children()
            num_active_processes = len(active_processes)
            start_time = time.time()

            while (num_active_processes >= 4) & (time.time() - start_time < self.timeout):
                active_processes = multiprocessing.active_children()
                num_active_processes = len(active_processes)
                time.sleep(30)
