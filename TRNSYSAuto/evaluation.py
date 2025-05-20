import os
import shutil
import math

import TRNSYSAuto.utils as utils
import pandas as pd
import numpy as np


class Evaluation:

    def __init__(self):
        self.date_df = utils.create_date_column(2024)
        self.variant_parameter_df = None
        self.eval_success = None
        self.variant_result_columns = None
        self.zone_1_with_df = None
        self.zone_1_without_df = None
        self.zone_3_with_df = None
        self.zone_3_without_df = None

    def setup_evaluation(self):
        """Set evaluation of simulation series up.

        Setting up the evaluation of a simulation series is only necessary once, continuing the evaluation at a later
        time does not need an additional setup. Doing so anyway results in a reset of the evaluation progress.
        """

        self.logger.info('Setting up evaluation.')

        # create evaluation directory
        os.makedirs(self.path.evaluation_save_dir, exist_ok=True)

        # create copy of cumulative evaluation file template
        shutil.copy(self.path.cumulative_evaluation_template, self.path.cumulative_evaluation_save_file)

        # initialize evaluation success list
        self.eval_success = [False] * len(self.simulations)

        sim_list = [sim.name for sim in self.simulations]

        # initialize tables for cumulative evaluation
        self.variant_result_columns = pd.DataFrame(columns=sim_list)
        self.zone_1_with_df = pd.DataFrame(columns=sim_list)
        self.zone_1_without_df = pd.DataFrame(columns=sim_list)
        self.zone_3_with_df = pd.DataFrame(columns=sim_list)
        self.zone_3_without_df = pd.DataFrame(columns=sim_list)

    def start_evaluation(self):
        """Start evaluation.

        Starts the evaluation process, which includes the evaluation of each individual simulation variant, followed by
        a cumulative evaluation using the combined results of those individual evaluations.
        """

        if not all(self.eval_success):
            # initialize progress bar
            progress = 0
            total = len(self.eval_success) - sum(self.eval_success)
            utils.progress_bar(progress, total)

            # logger entry "start"
            self.logger.info(f'Starting evaluation of {self.filename_sim_variants_excel}.')

            # evaluate variants
            for variant_index, variant_name in enumerate(self.sim_list):

                if not self.eval_success[variant_index]:
                    self.evaluate_variant(variant_name, variant_index)
                    progress += 1
                    utils.progress_bar(progress, total)
                    if progress % 5 == 0:  # save evaluation progress
                        self.save()

            # save progress
            self.save()

        # cumulative evaluation
        self.cumulative_evaluation()

        # logger entry "finish"
        self.logger.info('Evaluation done.')

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
            self.logger.error(f'File {path_variant_file} does not exist!')
            return

        # ...the variant has a corresponding directory
        if variant_name not in self.sim_list:
            self.logger.error(f'Did not find {variant_name} in {self.path.sim_variants_excel}.')
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
        utils.excel_export_variant_evaluation(
            self.sheet_name_variant_input, result, variant_name, save_path_variant_output, self.variant_parameter_df)

        # update excel to receive cross-referenced values and updates calculations
        utils.update_excel_file(save_path_variant_output)

        # create single column with all hourly values, for the cumulative evaluation Excel file
        result_column = utils.to_single_column(result[self.col_headers_result_column])

        # save single column
        self.variant_result_columns[variant_name] = result_column

        self.eval_success[variant_index] = True

        # self.logger.info(f'Finished evaluation of variant {variant_name}.')

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
        utils.progress_bar(progress, total)

        # logger entry "start"
        self.logger.info('Reading variant evaluation files for the cumulative evaluation.')

        for variant_index, variant_name in enumerate(self.sim_list):
            save_path_variant_output = os.path.join(self.path.evaluation_save_dir, variant_name + '.xlsx')

            # read data from variant evaluation excel file, for the cumulative evaluation excel file
            self.zone_1_with_df[variant_name] = read(sheet_name=self.sheet_name_zone_1_input, usecols=[3])
            self.zone_1_without_df[variant_name] = read(sheet_name=self.sheet_name_zone_1_input, usecols=[2])
            self.zone_3_with_df[variant_name] = read(sheet_name=self.sheet_name_zone_3_input, usecols=[3])
            self.zone_3_without_df[variant_name] = read(sheet_name=self.sheet_name_zone_3_input, usecols=[2])

            # update progress bar
            progress += 1
            utils.progress_bar(progress, total)

        # logger entry "export"
        self.logger.info('Exporting cumulative evaluation results.')

        # copy into cumulative evaluation excel file
        self.excel_export_cumulative_evaluation()

        # update cumulative excel
        utils.update_excel_file(self.path.cumulative_evaluation_save_file)

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

        self.logger.info(f'Exported cumulative evaluation successfully to {self.path.cumulative_evaluation_save_file}.')


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

        Foreign code without documentation.
        """

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

         Foreign code without documentation.
         """

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