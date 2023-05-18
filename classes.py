import sys
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED - needed as tkinter has compatibility issues with pywinauto when creating .exe file
import os
import shutil
import functions
import tkinter as tk
import pandas as pd
import re
# import psutil
import multiprocessing

from pywinauto.application import Application
from datetime import datetime
from tkinter import filedialog


class SimulationSeries:
    """ """

    def __init__(self, path_exe, filename_excel):
        """

        Parameters
        ----------
        path_exe :
        """

        current_time = datetime.now().strftime('%d.%m.%Y_%H.%M')  # current time

        # ask directory with base folder 'Basisordner' in it
        root = tk.Tk()
        root.withdraw()
        self.current_dir = filedialog.askdirectory()

        # paths
        self.path_sim_series = os.path.join(self.current_dir, current_time)     # simulation series folder
        self.path_base = os.path.join(self.current_dir, 'Basisordner')          # base folder todo: base folder beschreiben
        self.path_exe = path_exe    # .exe simulation file
        self.filename_excel = filename_excel    # input Excel file name
        self.path_excel = os.path.join(self.path_base, self.filename_excel)     # input Excel file

        os.makedirs(self.path_sim_series)  # create new directory for simulation series

    def import_input_excel(self):
        """
        Returns
        -------
        sim_list : list of strings
            List of simulation names
        weather_series : pandas.Series
            Series with the weather data file name, for each simulation
        df_dck : pandas.DataFrame
            Dataframe with the simulation parameters to be replaced in the .dck File, for each simulation
        b18_series : pandas.Series
            Series with the .b18 data file name, for each simulation

        """

        excel_data = pd.ExcelFile(self.path_excel)  # read input Excel file

        df = excel_data.parse('Simulationsvarianten', index_col=0)  # Excel data as DataFrame

        # transpose data
        df_weather = df[df.index == 'Wetterdaten'].transpose()
        b18_series = df[df.index == 'b18'].transpose()
        df_dck = df[df.index == 'dck'].transpose()

        # convert to series
        self.weather_series = df_weather[1:].squeeze()
        self.b18_series = b18_series[1:].squeeze()

        # use first row as header
        df_dck.columns = df_dck.iloc[0]
        self.df_dck = df_dck[1:]

        # list of simulation variants
        self.sim_list = df.columns[1:].astype(str).tolist()

    def start_sim_series(self):

        # copy Input Excel file into simulation series folder
        shutil.copy(os.path.join(self.path_base, self.filename_excel), self.path_sim_series)

        for sim in self.sim_list:
            self.create_sim_folder(sim)
            path_dck = os.path.join(self.path_sim_series, sim, 'templateDck.dck')
            self.start_sim(path_dck)

    def create_sim_folder(self, sim):

        path_sim = os.path.join(self.path_sim_series, sim)  # path of simulation folder
        os.makedirs(path_sim)  # create new directory for simulation
        # shutil.copytree(path_base, path_sim)          # copy all files from base to simulation folder

        # region SOURCE/DESTINATION FILE PATHS FOR COPYING PROCESS

        src_file = [
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

        # region FIND AND REPLACE PARAMETERS IN .dck AND .b17/b18 FILE

        functions.find_and_replace_param(dst_file[0], r'(@\w+)\s*=\s*([\d.]+)', self.df_dck.loc[sim])
        functions.find_and_replace_param(dst_file[0], r'(\*ASSIGN "b17")', self.df_dck.loc[sim])

        # endregion

    def start_sim(self, path_dck_file):

        path_dck_file = path_dck_file.replace("/", "\\")

        # path_file_raw = r'{}'.format(path_file)     # turn file path into raw string to avoid error messages

        app = Application(backend='uia')
        app.start(self.path_exe)
        app.connect(title="Öffnen", timeout=60 * 15)

        app.Öffnen.FileNameEdit.set_edit_text(path_dck_file)    # insert .dck file name
        Button = app.Öffnen.child_window(title="Öffnen", auto_id="1", control_type="Button").wrapper_object()
        Button.click_input()

        app.wait_cpu_usage_lower(threshold=0, timeout=60 * 15)
        print(app.cpu_usage())
        # app.kill()

        # region DELETE REDUNDANT FILES

        # os.remove(os.path.join(path_sim, 'templateDck.lst'))
        # os.remove(os.path.join(path_sim, 'out11.txt'))
        # os.remove(os.path.join(path_sim, 'out8.txt'))
        # os.remove(os.path.join(path_sim, 'Speicher1_step.out'))

        # endregion

    def start_sim_series_par(self):
        """

        Returns
        -------

        """
        pool_obj = multiprocessing.Pool()
        pool_obj.map(self.start_sim, self.p)
        pool_obj.close()
