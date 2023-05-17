import sys
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED - needed as tkinter has compatibility issues with pywinauto
import os
import shutil
import functions
import tkinter as tk
import pandas as pd
import re
import psutil

from pywinauto.application import Application
from datetime import datetime
from tkinter import filedialog


class SimulationSeries:
    """ """

    def __init__(self, path_exe):
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
        self.path_excel = os.path.join(self.path_base, 'Simulationsvarianten.xlsx')     # input Excel file

        os.makedirs(self.path_sim_series)  # create new directory for simulation series

        # import Excel file
        self.sim_list, self.weather_series, self.df_dck, self.b18_series = self.import_input_excel()

        self.start_sim_series()

    def start_sim_series(self):
        """

        Returns
        -------

        """

        for sim in self.sim_list:

            shutil.copy(os.path.join(self.path_base, 'Simulationsvarianten.xlsx'), self.path_sim_series)  # copy Input Excel file

            # region create simulation folders

            path_sim = os.path.join(self.path_sim_series, sim)  # path of simulation folder
            os.makedirs(path_sim)  # create new directory for simulation
            # shutil.copytree(path_base, path_sim)          # copy all files from base to simulation folder

            # region source/destination file paths for copying process
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
                shutil.copy(src_file[index], dst_file[index])  # copy file

            # replace parameters in .dck File
            functions.find_and_replace_param(dst_file[0], r'(@\w+)\s*=\s*([\d.]+)', self.df_dck.loc[sim])

            # region find and replace .b18/.b17 file name

            with open(dst_file[0], 'r') as file:
                text = file.read()  # read file
                new_text = re.sub(r'(\*ASSIGN "b17")', r'ASSIGN "' + self.b18_series[sim] + '"', text)
            with open(dst_file[0], 'w') as file:
                file.write(new_text)  # overwrite file

            # endregion

            # region find and replace weather data file name

            with open(dst_file[0], 'r') as file:
                text = file.read()  # read file
                new_text = re.sub(r'(\*ASSIGN "tm2")', r'ASSIGN "' + self.weather_series[sim] + '"', text)
            with open(dst_file[0], 'w') as file:
                file.write(new_text)  # overwrite file

            # endregion

            functions.find_and_replace_param(dst_file[0], r'(\*ASSIGN "b17")', self.df_dck.loc[sim])

            # new_text = re.sub(pattern, r'* ASSIGN "replacement"', text)

            # endregion

            # perform simulation
            self.start_sim(os.path.join(path_sim, dst_file[0]))

            # delete redundant files
            # os.remove(os.path.join(path_sim, 'templateDck.lst'))
            # os.remove(os.path.join(path_sim, 'out11.txt'))
            # os.remove(os.path.join(path_sim, 'out8.txt'))
            # os.remove(os.path.join(path_sim, 'Speicher1_step.out'))

    def start_sim(self, path_file):
        """

        Parameters
        ----------
        path_file : string
            Path of the file needed for the simulation (in this case Building.dck)
        """

        path_file = path_file.replace("/", "\\")
        # path_file_raw = r'{}'.format(path_file)     # turn file path into raw string to avoid error messages
        app = Application(backend='uia')
        app.start(self.path_exe)
        app.connect(title="Öffnen", timeout=60 * 15)

        app.Öffnen.FileNameEdit.set_edit_text(path_file)
        Button = app.Öffnen.child_window(title="Öffnen", auto_id="1", control_type="Button").wrapper_object()
        Button.click_input()

        app.wait_cpu_usage_lower(threshold=0, timeout=60 * 15)
        print(app.cpu_usage())
        app.kill()

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
        weather_series = df_weather[1:].squeeze()
        b18_series = b18_series[1:].squeeze()

        # use first row as header
        df_dck.columns = df_dck.iloc[0]
        df_dck = df_dck[1:]

        # list of simulation variants
        sim_list = df.columns[1:].astype(str).tolist()

        return sim_list, weather_series, df_dck, b18_series
