import sys
import os
import shutil
import multiprocessing
import time
import csv
import logging
import glob
import win32com.client
import math
import functions
import openpyxl
import re

import numpy as np
import pandas as pd
import xlwings as xw
import tkinter as tk

from tkinter import filedialog  # explicit import required, as calling from tk.filedialog does not work
from natsort import natsorted
from pywinauto.application import Application
from datetime import datetime


# from line_profiler import LineProfiler
# from tkinter import filedialog


class SimulationSeries:  # todo: Durch Vererbung erweitern, damit auch andere Programme als TRNSYS leichter automatisiert werden können
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
        self.sim_list = None  # list of simulation variant names
        self.sim_success = None  # boolean list, documenting the successful simulation of each simulation variant
        self.df_dck = None  # pandas DataFrame with the simulation parameters to be replaced in the .dck Files
        self.b18_series = None  # pandas series with the .b18 data file names
        self.weather_series = None  # pandas series with the weather data file names

        # Settings
        self.settings = None  # Pandas series with miscellaneous simulation settings
        self.path_exe = None  # path to TRNSYS executable file
        self.name_excelsheet_sim_variants = None  # name of the Excel sheet containing the simulation variants table
        self.filename_dck_template = None  # name of the dck-File template
        self.timeout = None  # if timeout (sec) is reached without starting another simulation, stop program
        self.start_time_buffer = None  # time buffer (sec) between two simulations, for increased stability (optional)
        self.multiprocessing_max = None  # maximum number of simulations that can be calculated simultaneously
        self.autostart_evaluation = False  # start the evaluation routine for the simulation results afterwards if True

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

    def apply_settings(
            self):  # todo automatischer Namenabgleich einführen, am besten mit Prüfung und Meldung bei Unstimmigkeiten
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

            file_list = [self.filename_dck_template, 'Lastprofil.txt', 'SzenarioAneu.txt', 'Qelww_CHR55025.txt',
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
                    self.sim_success[
                        sim_index] = True  # todo: als Übergangslösung wird eine erfolgreiche Simulation vorgetäuscht, um Simulationen wo kritische Daten fehlen überspringen zu können. Langfrisitg soll sim_success aber nur tatsächlich erfolgreiche Simulationen dokumentieren!

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
            app.connect(title="Öffnen", timeout=2)  # self.timeout)

            # insert .dck file path
            app.Öffnen.FileNameEdit.set_edit_text(path_dck_file)

            # press start button
            Button = app.Öffnen.child_window(title="Öffnen", auto_id="1", control_type="Button").wrapper_object()
            Button.click_input()

            # wait for the simulation window to open
            while len(app.windows()) < 1:
                time.sleep(1)

        except Exception:  # TimeoutError:
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
        redundant_file_list = ['out11.txt', 'out8.txt', 'out6.txt', 'out7.txt', 'out10.txt', 'Speicher1_step.out']
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

        while not all(self.sim_success):  # check if any simulation has not been simulated successfully yet

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
                            time.sleep(5)  # pause until number of active simulations drops below maximum
                            if time.time() - start_time > self.timeout:
                                sys.exit('Timeout of ' + str(self.timeout) + ' sec reached, program ended.')
                        time.sleep(5)
                        process.start()  # start process
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

            path_output = os.path.join(self.dir_sim_series, sim, 'out5.txt')  # path of output file
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

    def evaluation(self):

        print('Starting evaluation of simulation results')
        output_folder = os.path.join(self.dir_sim_series, 'evaluation')

        # region NAMES

        trnsys_data_file_name = 'out5.txt'
        cumulative_template_file = os.path.abspath('./Basisordner/Auswertung_Gesamt.xlsx')
        cumulative_output_file = os.path.join(output_folder, 'gesamt.xlsx')
        variant_template_file = './Basisordner/Auswertung_Variante.xlsx'

        raw_data_variant_sheet_name = 'Rohdaten'
        calculation_sheetname = 'Berechn1'
        raw_data_cumulative_sheet_name = 'Rohinputs'
        zone_1_input = 'Zusamm1'
        zone_3_input = 'Zusamm3'
        zone_1_with_output = 'Zone1_Betrieb'
        zone_1_without_output = 'Zone1ges'
        zone_3_with_output = 'Zone3_Betrieb'
        zone_3_without_output = 'Zone3ges'

        # selected_trnsys_columns = ['Period', 'top1', 'top2', 'top3', 'Qventfges', 'qvolgesh', 'qc1', 'qc2', 'qc3', 'pmv1',
        #                            'pmv2', 'pmv3']
        selected_trnsys_columns = ['ta', 'top1', 'top2', 'top3', 'Qventfges', 'qvolgesh', 'qc1', 'qc2', 'qc3',
                                   'qh1', 'qh2',
                                   'qh3', 'pmv1', 'pmv2', 'pmv3', 'ppd1', 'ppd2', 'ppd3', 'clo1', 'clo2', 'clo3',
                                   'met1',
                                   'met2', 'met3']

        var_list_zone1 = ['Period', 'ta', 'tzone1', 'TMSURF_ZONE1', 'relh1', 'vel1', 'pmv1', 'ppd1', 'clo1', 'met1',
                          'work1']
        var_list_zone2 = ['Period', 'ta', 'tzone1.1', 'TMSURF_ZONE1.1', 'relh2', 'vel2', 'pmv2', 'ppd2', 'clo2',
                          'met2', 'work2']
        var_list_zone3 = ['Period', 'ta', 'tzone1.2', 'TMSURF_ZONE1.2', 'relh3', 'vel3', 'pmv3', 'ppd3', 'clo3',
                          'met3', 'work3']
        trnsys_outdoor_temperature = 'ta'

        variant_parameter_file = os.path.join(self.dir_sim_series, self.filename_sim_variants_excel) + '.xlsx'

        # endregion NAMES

        # remove existing files in output folder
        for existing_output in glob.glob(output_folder + '/*.xlsx'):
            os.remove(existing_output)
        os.makedirs(output_folder, exist_ok=True)

        # create copy of evaluation file template
        shutil.copy(cumulative_template_file, cumulative_output_file)

        # read simulation variant parameters
        variant_parameter_df = pd.read_excel(variant_parameter_file, sheet_name='Simulationsvarianten')
        variant_parameter_df.columns = [str(parameter) for parameter in variant_parameter_df.columns]

        # get top level directory list
        variant_folders = next(os.walk(self.dir_sim_series))[1]
        variant_folders.remove('evaluation')
        variant_folders = natsorted(variant_folders)

        # read TRNSYS output and save data
        variant_cnt = 0
        for variant_folder in variant_folders:
            variant_cnt = variant_cnt + 1
            variant_folder_path = os.path.join(self.dir_sim_series, variant_folder)
            variant_file_path = os.path.join(variant_folder_path, trnsys_data_file_name)
            variant_output_file = os.path.join(output_folder, 'variant' + variant_folder + '.xlsx')
            print(variant_file_path)

            if not os.path.exists(variant_file_path):
                print(f'File {variant_file_path} does not exist!')
                # raise ValueError(f'File {variant_file_path} does not exist!')
                continue

            variant_list = variant_parameter_df.columns.to_list()
            variant_list = [str(variant) for variant in variant_list]
            if variant_folder not in variant_list:
                print(f'Did not find {variant_folder} in {variant_parameter_file}')
                # raise ValueError(f'Did not find {variant_folder} in {variant_parameter_file}')
                continue

            # read trnsys output file
            trnsys_df = pd.read_csv(variant_file_path, sep='\s+', skiprows=1, skipfooter=0, engine='python')

            # region todo: ppd1-3 muss im out5.txt vorhanden sein, dann können diese Zeilen entfernt werden
            trnsys_df['ppd1'] = np.nan
            trnsys_df['ppd2'] = np.nan
            trnsys_df['ppd3'] = np.nan
            # endregion

            # todo: ab da etwas chaotisch
            # region DATA
            selected_trnsys_df = trnsys_df[selected_trnsys_columns]
            selected_trnsys_df = selected_trnsys_df.reindex(selected_trnsys_columns, axis=1)

            # Schweiker model for zones 1, 2 & 3
            sm1 = SchweikerDataFrame()
            sm1._df = trnsys_df[var_list_zone1].reindex(var_list_zone1, axis=1)
            sm2 = SchweikerDataFrame()
            sm2._df = trnsys_df[var_list_zone2].reindex(var_list_zone2, axis=1)
            sm3 = SchweikerDataFrame()
            sm3._df = trnsys_df[var_list_zone3].reindex(var_list_zone3, axis=1)

            # adapt column headers
            var_list = ['Period', 'ta', 'tzone', 'TMSURF_ZONE', 'relh', 'vel', 'pmv', 'ppd', 'clo', 'met', 'work']
            sm1.df.columns = var_list
            sm2.df.columns = var_list
            sm3.df.columns = var_list

            # region CREATE DATE COLUMN   #todo: Wird derzeit künstlich erzeugt
            year = 2023  # todo: Jahr derzeit hard coded
            time_increment_profiles = 60
            time = pd.date_range(
                start=str(year) + '-01-01',
                end=str(year + 1) + '-01-01',
                freq=str(time_increment_profiles) + 'min')

            time = time.to_series()

            time_df = pd.DataFrame({
                'Tag': time.dt.day,
                'Monat': time.dt.month,
                'Jahr': time.dt.year,
                'Stunde': time.dt.hour,
                'Minute': time.dt.minute
            })

            time_df = time_df[:-(len(time_df) - len(sm1.df))]  # adapt length

            # endregion

            # insert date columns
            time_df = time_df.reset_index()  # reset index
            sm1._df = pd.concat([time_df, sm1.df], axis=1)
            sm2._df = pd.concat([time_df, sm2.df], axis=1)
            sm3._df = pd.concat([time_df, sm3.df], axis=1)

            # schweiker main
            sm1.schweiker_main()
            sm2.schweiker_main()
            sm3.schweiker_main()

            # combine zones data to one single DataFrame
            # result = pd.concat([sm1.df, sm2.df, sm2.df], axis=1)

            # top1 = (sm1.df.tzone + sm1.df.TMSURF_ZONE) / 2
            # top2 = (sm2.df.tzone + sm2.df.TMSURF_ZONE) / 2
            # top3 = (sm3.df.tzone + sm3.df.TMSURF_ZONE) / 2
            #
            # result = pd.concat([top1, top2, top3], axis=1)
            # result.columns = ['top1', 'top2', 'top3']
            # result['Qventfges'] = ''
            # result['qvolgesh'] = ''
            # result['qc1'] = ''
            # result['qc2'] = ''
            # result['qc3'] = ''

            # region RENAME sm

            # remove redundant columns
            sm1.df.drop(['Tag', 'Monat', 'Jahr', 'Stunde', 'Minute', 'index', 'Period'], axis=1, inplace=True)
            sm2.df.drop(['Tag', 'Monat', 'Jahr', 'Stunde', 'Minute', 'index', 'Period'], axis=1, inplace=True)
            sm3.df.drop(['Tag', 'Monat', 'Jahr', 'Stunde', 'Minute', 'index', 'Period'], axis=1, inplace=True)

            # numerate column names for each zone
            sm1.df.columns = ['schweiker_' + string + '1' for string in sm1.df.columns]
            sm2.df.columns = ['schweiker_' + string + '2' for string in sm2.df.columns]
            sm3.df.columns = ['schweiker_' + string + '3' for string in sm3.df.columns]

            # endregion

            # concatenate output
            result = pd.concat([trnsys_df[['ta', 'top1', 'top2', 'top3', 'tzone1', 'tzone2', 'tzone3', 'Qventfges',
                                           'qvolgesh', 'qc1', 'qc2', 'qc3', 'qh1', 'qh2', 'qh3', 'pmv1', 'pmv2',
                                           'pmv3',
                                           'ppd1', 'ppd2', 'ppd3', 'clo1', 'clo2', 'clo3', 'met1', 'met2', 'met3']],
                                sm1.df, sm2.df, sm3.df], axis=1)

            # sort columns
            result = result[
                ['ta', 'top1', 'top2', 'top3', 'tzone1', 'tzone2', 'tzone3', 'Qventfges', 'qvolgesh', 'qc1',
                 'qc2', 'qc3', 'qh1', 'qh2', 'qh3', 'pmv1', 'pmv2', 'pmv3', 'ppd1', 'ppd2', 'ppd3', 'clo1',
                 'clo2', 'clo3', 'met1', 'met2', 'met3', 'schweiker_pmv1', 'schweiker_pmv2', 'schweiker_pmv3',
                 'schweiker_ppd1', 'schweiker_ppd2', 'schweiker_ppd3', 'schweiker_clo1', 'schweiker_clo2',
                 'schweiker_clo3', 'schweiker_met1', 'schweiker_met2', 'schweiker_met3']]

            # trnsys_df = result  # overwrite with schweiker model data

            # endregion

            # region EXCEL EXPORT

            # copy template and write data in it
            shutil.copy(variant_template_file, variant_output_file)

            wb = xw.Book(variant_output_file)
            ws = wb.sheets[raw_data_variant_sheet_name]
            ws["A2"].options(pd.DataFrame, header=1, index=False, expand='table').value = variant_parameter_df[
                ['File', 'Parameter', variant_folder]]
            ws["B60"].options(pd.DataFrame, header=1, index=False, expand='table').value = result
            wb.save()
            wb.app.quit()

            # update excel to receive cross-referenced values and updates calculations
            # https://stackoverflow.com/questions/40893870/refresh-excel-external-data-with-python
            xlapp = win32com.client.DispatchEx("Excel.Application")
            wb = xlapp.Workbooks.Open(variant_output_file)
            wb.RefreshAll()
            xlapp.CalculateUntilAsyncQueriesDone()
            wb.Save()
            xlapp.Quit()

            # Quit
            xlapp.Quit()

            # copy all zones and all classifications in cumulative file
            # read calculated values from variant and copy in cumulative file
            zone_1_with_df = pd.read_excel(variant_output_file, sheet_name=zone_1_input, usecols=[3], header=None,
                                           nrows=None, skiprows=None)
            zone_1_without_df = pd.read_excel(variant_output_file, sheet_name=zone_1_input, usecols=[2],
                                              header=None,
                                              nrows=None, skiprows=None)
            zone_3_with_df = pd.read_excel(variant_output_file, sheet_name=zone_3_input, usecols=[3], header=None,
                                           nrows=None, skiprows=None)
            zone_3_without_df = pd.read_excel(variant_output_file, sheet_name=zone_3_input, usecols=[2],
                                              header=None,
                                              nrows=None, skiprows=None)

            # create column with all hourly values
            var_list_result_column \
                = ['top1', 'top2', 'top3', 'Qventfges', 'qvolgesh', 'qc1', 'qc2', 'qc3', 'pmv1', 'pmv2', 'pmv3']
            header = pd.DataFrame(var_list_result_column[1:] + ['']).transpose()
            header.columns = var_list_result_column[:-1] + ['']
            # todo: es fehlt ein Tag, deshalb werden 23+2 Werte derzeit übersprungen
            result_column = pd.concat([
                result[var_list_result_column],
                pd.DataFrame(index=['']),
                header],
                axis=0)
            result_column = result_column.drop(result_column.columns[-1], axis=1)
            result_column = result_column.transpose().stack(dropna=False)

            with pd.ExcelWriter(cumulative_output_file, mode="a", engine="openpyxl",
                                if_sheet_exists='overlay') as writer:
                if variant_cnt <= 1:
                    variant_parameter_df[['File', 'Parameter', variant_folder]].to_excel(writer,
                                                                                         sheet_name=raw_data_cumulative_sheet_name,
                                                                                         startrow=1, startcol=0,
                                                                                         index=False)
                else:
                    variant_parameter_df[variant_folder].to_frame().to_excel(writer,
                                                                             sheet_name=raw_data_cumulative_sheet_name,
                                                                             startrow=1, startcol=1 + variant_cnt,
                                                                             index=False)

                zone_1_with_df.to_excel(writer, sheet_name=zone_1_with_output, startrow=1, startcol=6 + variant_cnt,
                                        index=False, header=False)
                zone_1_without_df.to_excel(writer, sheet_name=zone_1_without_output, startrow=1,
                                           startcol=6 + variant_cnt,
                                           index=False, header=False)
                zone_3_with_df.to_excel(writer, sheet_name=zone_3_with_output, startrow=1, startcol=6 + variant_cnt,
                                        index=False, header=False)
                zone_3_without_df.to_excel(writer, sheet_name=zone_3_without_output, startrow=1,
                                           startcol=6 + variant_cnt,
                                           index=False, header=False)
                result_column.to_excel(writer, sheet_name=raw_data_cumulative_sheet_name, startrow=60,
                                       startcol=1 + variant_cnt,
                                       index=False, header=False)
            # print(calculated_variant_df)

            # endregion todo: Als eigene Methode isolieren

        # update cumulative excel
        # https://stackoverflow.com/questions/40893870/refresh-excel-external-data-with-python
        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open(cumulative_output_file)
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        xlapp.Quit()

        # Quit
        xlapp.Quit()


class SchweikerDataFrame:
    """Modified pandas Dataframe for the Schweiker-Model."""

    def __init__(self, *args, **kwargs):
        self._df = pd.DataFrame(*args, **kwargs)

    def __getattr__(self, attr):
        return getattr(self._df, attr)

    @property
    def df(self):
        return self._df

    def read_input_excel(self, sheet_name='Sheet1', skiprows=0):
        # ask simulation variants Excel file path(s)
        root = tk.Tk()
        root.withdraw()
        path_input_file = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")],
                                                     title='Select input Excel file')
        # read Excel data
        excel_data = pd.ExcelFile(path_input_file)
        # convert Excel data into pandas DataFrame
        self._df = pd.DataFrame(excel_data.parse(sheet_name))  # , index_col=0)

        if skiprows > 0:
            self._df.columns = self._df.iloc[skiprows]  # set column headers
            self._df = self._df.drop(range(0, skiprows + 1))  # remove unwanted rows
            self._df = self._df.reset_index(drop=True)

    def interface_TimberBioC(self):

        # remove index from column headers
        for col in self._df.iloc[:, 1:].columns:
            # print(col)
            self._df.columns = self._df.columns.str.replace(col, col[:-1])

        # create table with temporal information
        date_info = pd.to_datetime(self.df[self.df.columns[0]])
        date_df = pd.concat(
            [date_info.dt.year, date_info.dt.month, date_info.dt.day, date_info.dt.hour, date_info.dt.minute], axis=1)
        date_df.columns = ['Jahr', 'Monat', 'Tag', 'Stunde', 'Minute']

        # split zones and concatenate temporal information todo: 1)HARD CODED 2)Dummy Außentemperatur
        dummy = pd.DataFrame(np.random.randint(0, 100, size=(8760, 1)), columns=['Aussentemp'])
        df_z1 = pd.concat([date_df, self._df.iloc[:, 2:8], dummy], axis=1)
        df_z2 = pd.concat([date_df, self._df.iloc[:, 9:15], dummy], axis=1)
        df_z3 = pd.concat([date_df, self._df.iloc[:, 16:22], dummy], axis=1)

        return df_z1, df_z2, df_z3

    def write_output_excel(self):
        # ask output Excel file path
        root = tk.Tk()
        root.withdraw()
        output_path = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")],
                                                 title='Select output Excel file')

        # write into Excel file
        with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", datetime_format="DD.MM.YYYY HH:MM") as writer:

            key = 'output'
            self._df.to_excel(writer, sheet_name=key)
            ws = writer.sheets[key]
            dims = {}
            for row in ws.rows:
                for cell in row:
                    if cell.value:
                        dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
            for col, value in dims.items():
                ws.column_dimensions[col].width = value + 2

    def calcComfort(self):
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

    def schweiker_main(self):

        self._df = functions.calcFloatingAverageTemperature(self.df, values_name='ta', dates_name='index')

        # adapt metabolic rate
        self.df['metAdaptedColumn'] = self.df['met'] - (0.234 * self.Aussentemp_floating_average) / 58.2
        # self.df['metAdaptedColumn'] = self.df['metabolischeRate'] - (0.234 * self.Aussentemp_floating_average) / 58.2
        # determine clothing factor
        self.df['clo'] = 10 ** (-0.172 - 0.000485 * self.df['Aussentemp_floating_average']
                                + 0.0818 * self.df['metAdaptedColumn']
                                - 0.00527 * self.df['Aussentemp_floating_average'] * self.df['metAdaptedColumn'])
        # calculate comfort
        [pmv, ppd] = self.calcComfort()
        self.df['pmv'] = pmv
        self.df['ppd'] = ppd
        df_z1 = self._df

        # export to Excel file
        # self.write_output_excel()

        print('Done')

    # def calcFloatingAverageTemperature(self, values_name='Aussentemp', dates_name='date'):
    #     # todo:  VERALTET, wurde durch eigene standalone funktion ersetzt
    #     floating_alpha = 0.8
    #
    #     if self.df[dates_name].isnull().values.any() or self.df[values_name].isnull().values.any():
    #         raise ValueError('Values are not allowed to be NaN, interpolate if necessary!')
    #
    #     average_name = f'{values_name}_mean'
    #     floating_average_name = f'{values_name}_floating_average'
    #
    #     df = pd.DataFrame()
    #     df[dates_name] = self.df[dates_name].copy()
    #     df[values_name] = self.df[values_name].copy()
    #     df = df.sort_values(dates_name)
    #     df['ymd'] = pd.to_datetime(df[dates_name]).dt.date
    #     mean_df = df.groupby('ymd').mean()
    #     mean_df = mean_df.rename(columns={values_name: average_name})
    #
    #     df = df.merge(mean_df, how='left', on='ymd')
    #     day_counter = 1
    #     next_datapoint_time_delta = pd.Timedelta(0)
    #     # temporary list which consists the index of the first row of each new day
    #     new_day_datapoints = [0]
    #
    #     for i in range(len(df)):
    #         # print(f'row: {i}')
    #
    #         if i != 0:
    #             next_datapoint_time_delta = df.loc[i, 'ymd'] - df.loc[new_day_datapoints[-1], 'ymd']
    #             # print(next_datapoint_time_delta)
    #
    #         if next_datapoint_time_delta >= pd.Timedelta('2D'):
    #             # reset time counter when a time gap happens
    #             new_day_datapoints = [i]
    #             day_counter = 1
    #         elif next_datapoint_time_delta >= pd.Timedelta('1D'):
    #             new_day_datapoints.append(i)
    #             day_counter = day_counter + 1
    #
    #         if day_counter > 8:
    #             df.loc[i, floating_average_name] = (1 - floating_alpha) * df.loc[
    #                 new_day_datapoints[-2], average_name] + floating_alpha * df.loc[
    #                                                    new_day_datapoints[-2], floating_average_name]
    #         elif day_counter > 7:
    #             # DIN 1525251 Formula
    #             df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
    #                 new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name] + 0.5 *
    #                                                 df.loc[new_day_datapoints[-5], average_name] + 0.4 * df.loc[
    #                                                     new_day_datapoints[-6], average_name] + 0.3 * df.loc[
    #                                                     new_day_datapoints[-7], average_name] + 0.2 * df.loc[
    #                                                     new_day_datapoints[-8], average_name]) / 3.8
    #         elif day_counter > 6:
    #             df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
    #                 new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name] + 0.5 *
    #                                                 df.loc[new_day_datapoints[-5], average_name] + 0.4 * df.loc[
    #                                                     new_day_datapoints[-6], average_name] + 0.3 * df.loc[
    #                                                     new_day_datapoints[-7], average_name]) / 3.6
    #         elif day_counter > 5:
    #             df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
    #                 new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name] + 0.5 *
    #                                                 df.loc[new_day_datapoints[-5], average_name] + 0.4 * df.loc[
    #                                                     new_day_datapoints[-6], average_name]) / 3.3
    #         elif day_counter > 4:
    #             df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
    #                 new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name] + 0.5 *
    #                                                 df.loc[new_day_datapoints[-5], average_name]) / 2.9
    #         elif day_counter > 3:
    #             df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
    #                 new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name]) / 2.4
    #         elif day_counter > 2:
    #             df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
    #                 new_day_datapoints[-3], average_name]) / 1.8
    #         elif day_counter > 1:
    #             df.loc[i, floating_average_name] = df.loc[new_day_datapoints[-2], average_name]
    #         elif day_counter <= 1:
    #             df.loc[i, floating_average_name] = df.loc[i, average_name]
    #         else:
    #             raise ValueError('day_counter case is not considered! Fix it!')
    #
    #     self.df['Aussentemp_mean'] = df[average_name]
    #     self.df['Aussentemp_floating_average'] = df[floating_average_name]
