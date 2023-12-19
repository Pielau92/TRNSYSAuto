import math
import numpy as np
import pandas as pd
import tkinter as tk

from tkinter import filedialog  # explicit import required, as calling from tk.filedialog does not work
from datetime import datetime, date


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
            self._df.columns = self._df.iloc[skiprows]      # set column headers
            self._df = self._df.drop(range(0, skiprows+1))  # remove unwanted rows
            self._df = self._df.reset_index(drop=True)

    def interface_TimberBioC(self):

        # remove index from column headers
        for col in self._df.iloc[:, 1:].columns:
            # print(col)
            self._df.columns = self._df.columns.str.replace(col, col[:-1])

        # create table with temporal information
        date_info = pd.to_datetime(self.df[self.df.columns[0]])
        date_df = pd.concat([date_info.dt.year, date_info.dt.month, date_info.dt.day, date_info.dt.hour, date_info.dt.minute], axis=1)
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

    def calcFloatingAverageTemperature(self, values_name='Aussentemp', dates_name='date'):
        # todo:  VERALTET, wurde durch eigene standalone funktion ersetzt
        floating_alpha = 0.8

        if self.df[dates_name].isnull().values.any() or self.df[values_name].isnull().values.any():
            raise ValueError('Values are not allowed to be NaN, interpolate if necessary!')

        average_name = f'{values_name}_mean'
        floating_average_name = f'{values_name}_floating_average'

        df = pd.DataFrame()
        df[dates_name] = self.df[dates_name].copy()
        df[values_name] = self.df[values_name].copy()
        df = df.sort_values(dates_name)
        df['ymd'] = pd.to_datetime(df[dates_name]).dt.date
        mean_df = df.groupby('ymd').mean()
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

        self.calcFloatingAverageTemperature(values_name='ta', dates_name='index')
        # adapt metabolic rate
        self.df['metAdaptedColumn'] = self.df['met'] - (0.234 * self.Aussentemp_floating_average) / 58.2
        # self.df['metAdaptedColumn'] = self.df['metabolischeRate'] - (0.234 * self.Aussentemp_floating_average) / 58.2
        # determine clothing factor
        self.df['clo'] = 10 ** (-0.172 - 0.000485 * self.df['Aussentemp_floating_average']
                                + 0.0818 * self.df['metAdaptedColumn']
                                - 0.00527 * self.df['Aussentemp_floating_average'] * self.df['metAdaptedColumn'])
        # calculate comfort
        [pmv, ppd] = self.calcComfort()
        self.df['schweiker_pmv'] = pmv
        self.df['schweiker_ppd'] = ppd
        df_z1 = self._df

        # export to Excel file
        # self.write_output_excel()

        print('Done')

# region MATLAB FUNCTION REPLACEMENTS


def isequal(array):
    return all(x == array[0] for x in array)


def datenum(yearColumn, monthColumn, dayColumn, hourColumn, minuteColumn, secondColumn):
    if isinstance(yearColumn, np.int64):
        datenum_single = \
            date.toordinal(datetime(yearColumn, monthColumn, dayColumn, hourColumn, minuteColumn, secondColumn))
        return datenum_single

    elif isinstance(yearColumn, pd.core.series.Series):
        datenum_array = []
        for index in range(len(yearColumn)):
            datenum_array.append(date.toordinal(datetime(yearColumn[index], monthColumn[index], dayColumn[index],
                                                         hourColumn[index], minuteColumn[index], secondColumn[index])))
        return datenum_array


def weeknum(time_array=None, year=None, month=None, day=None):
    if all(x is not None for x in [year, month, day]):
        return date(year, month, day).isocalendar().week
    elif time_array is not None:
        return date(time_array[0], time_array[1], time_array[2]).isocalendar().week
    else:
        print('Wrong Input for weeknum function.')
        return np.nan

# endregion
