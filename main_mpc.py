# -*- coding: utf-8 -*-
"""
Created on Thu Aug  3 13:08:01 2023

@author: Magdalena
"""
import numpy as np
import csv
import os

from pathlib import Path
from statistics import mean


def main():
    # building object
    building = Building(area=158.46,
                        alpha_w=6.5,
                        alpha_s=10.75,
                        k=120.71,
                        cp_tab=23.45,
                        cp_r=54.91,
                        max_heating=13,  # heating - reduced from 6.5 kW to 3.5 kW on 14.11.2019 //Both TOPS: 13 kW
                        max_cooling=-10,  # cooling - both TOPS: -10 kW
                        dt=3600)

    # import data from csv
    # building.import_csv("Test_MPC_Python.csv")

    result = building.optimize()


class Building:
    """Building model class.

    Building model class with a thermally activated building (TAB) component.
    """

    def __init__(self, area, alpha_w, alpha_s, k, cp_tab, cp_r, max_heating, max_cooling, dt=3600):
        """Initialize building model.

        Parameters
        ----------
        area : activated building area [m²]
        alpha_w : heat transfer coefficient (winter) [W/m²K]
        alpha_s : heat transfer coefficient (sommer) [W/m²K]
        k : factor for convection, transition and ventilation losses [W/K]
        cp_tab : absolute heat capacity of TAB component [kWh/K]
        cp_r : absolute heat capacity of room [kWh/K]
        max_heating : maximum heating power [kW]
        max_cooling : maximum cooling power [kW]
        dt : time interval [s] (e.g 3600 = 1 hour)
        """

        # building specific parameters
        self.area = area
        self.alpha_w = int(alpha_w)
        self.alpha_s = int(alpha_s)
        self.k = k
        self.cp_tab = cp_tab
        self.cp_r = cp_r
        self.max_heating = max_heating
        self.max_cooling = max_cooling * -1
        self.dt = dt

        # weather data
        self.ta = None
        self.igs = None
        self.ign = None
        self.time_step_nr = 0

        self.settings = SettingsMPC()

        self.logFile = open("PythonLog.log", "w")

    def read_weather_data(self, path_trnsys_input_file, filename_weather_data='Windetc20190804.txt'):
        """Read weather data for TRNSYS simulation.

        Parameters
        ----------
        path_trnsys_input_file : path to trnsys input file (dck file)
        filename_weather_data : filename of the csv file with weather data

        """

        path_sim_dir = os.path.dirname(path_trnsys_input_file)
        path_weather_data = os.path.join(path_sim_dir, filename_weather_data)
        # path_weather_data = \
        #     Path('C:/Users/pierre/PycharmProjects/BBSR Sommerlicher Komfort/Basisordner/Windetc20190804.txt')

        lines = Path(path_weather_data).read_text().splitlines()
        reader = csv.reader(lines, delimiter='\t')
        # todo: offset und col_index_* in settings verstauen
        offset = 5

        # column index of specific rows
        col_index_ta = 25
        col_index_igs = 26
        col_index_ign = 27

        self.ta, self.igs, self.ign = [], [], []
        for index, row in enumerate(reader):
            if index < offset:
                continue  # apply offset by skipping the first rows

            self.ta.append(float(row[col_index_ta]))
            self.igs.append(float(row[col_index_igs]))
            self.ign.append(float(row[col_index_ign]))

        if not self.dt == 3600:
            self.interpolate_weather_data()

    def interpolate_weather_data(self):
        """Interpolate weather data to right length."""

        self.ta = interpolate(self.ta, 3600 / self.dt)
        self.igs = interpolate(self.igs, 3600 / self.dt)
        self.ign = interpolate(self.ign, 3600 / self.dt)

    def optimize(self):
        """todo"""

        # prediction horizon in terms of time steps, instead of hours
        pred_hor_time_steps = int(self.settings.pred_hor * 3600 / self.dt)
        pred_hor_short_time_steps = int(self.settings.pred_hor_short * 3600 / self.dt)

        index_range = list(range(self.time_step_nr, self.time_step_nr + pred_hor_time_steps))

        Q_solar = self.igs[index_range]  # W/m²
        T_out = self.ta[index_range]  # °C

        dHeat = self.settings.dHeat  # kW
        T_sp = [self.settings.setpoint_temperature] * pred_hor_time_steps  # °C

        Q_heat = np.zeros(pred_hor_time_steps)  # kW
        Q_help_s = np.zeros(pred_hor_short_time_steps)  # kW

        counter = 0
        ChgProgress = 1  # termination criterion optimization - difference between lse_baseline and lse_neu_long
        while counter < self.settings.max_count and ChgProgress >= self.settings.ChgProgTol:

            # baseline calculation
            T_in, T_tab = self.predict(Q_heat, Q_solar, T_out)
            lse_baseline = lse(T_in, T_sp)  # least square error for zero heat input / heat output
            Q_heat_s = self.convert_pred_hor(Q_heat, 'long2short')  # shorten Q_heat

            # loop through hours of prediction horizon
            for i in range(pred_hor_short_time_steps):

                # negative perturbation
                Q_help_s[i] = max(Q_heat_s[i] - dHeat, self.max_cooling)  # limit to minimum cooling power
                Q_help = self.convert_pred_hor(Q_help_s, 'short2long')  # expand Q_help
                T_in, T_tab = self.predict(Q_help, Q_solar, T_out)
                lse_negative = lse(T_in, T_sp)  # least square error, negative perturbation

                # positive perturbation
                Q_help_s[i] = min(Q_heat_s[i] + dHeat, self.max_heating)  # limit to maximum heating power
                Q_help = self.convert_pred_hor(Q_help_s, 'short2long')  # expand Q_help
                T_in, T_tab = self.predict(Q_help, Q_solar, T_out)
                lse_positive = lse(T_in, T_sp)  # least square error, positive perturbation

                # interpretation of the perturbation effect
                match np.argmin([lse_baseline, lse_negative, lse_positive]):
                    case 0:  # baseline calculation has lowest least square error
                        pass
                    case 1:  # negative perturbation has lowest least square error
                        Q_heat_s[i] -= dHeat
                    case 2:  # positive perturbation has lowest least square error
                        Q_heat_s[i] += dHeat

                # limitation that cooling and heating simultaneously in one period is not possible
                min_value, max_value = 0, 0
                if self.settings.season:  # limit to maximum heating power
                    max_value = self.max_heating
                elif not self.settings.season:  # limit to maximum cooling power
                    min_value = self.max_cooling
                Q_heat_s[i] = np.clip(Q_heat_s[i], min_value, max_value)

                Q_help_s[i] = Q_heat_s[i]  # reset of the helping variable to not forget the value

            Q_heat = self.convert_pred_hor(Q_heat_s, 'short2long')  # expand heating vector to prediction horizon
            T_in, T_tab = self.predict(Q_heat, Q_solar, T_out)
            lse_neu_long = lse(T_in, T_sp)  # least square error for final perturbation in this loop run

            # calculate termination criterion: lse from start compared to lse from last perturbation run
            ChgProgress = lse_baseline - lse_neu_long

            # output of final least square error
            # print(counter, ". Durchgang: ", lse_neu_long)

            # loop counter
            counter += 1

        return Q_heat

    def predict(self, Q_heat, Q_solar, T_out):
        """Predict room air temperature and temperature of thermally activated building (TAB) component.

        Parameters
        ----------
        Q_heat : heating/cooling power [kW]
        Q_solar : heat loads from solar radiation [kW]
        T_out : outside air temperature [°C]

        Returns
        -------
        T_in : room air temperature [°C]
        T_tab : temperature of thermally activated building component [°C]

        """

        pred_hor_time_steps = int(self.settings.pred_hor * 3600 / self.dt)

        Q_loss = []  # prediction of convection, transition and ventilation losses [kW]
        Q_tab = []  # thermal heat flow between room and TAB component [kW]
        T_in = list([self.settings.T_start_in])  # prediction of  room temperature [°C]
        T_tab = list([self.settings.T_start_tab])  # temperature TAB [°C]

        alpha = [self.alpha_s, self.alpha_w]  # cooling in summer (season = 0), heating in winter (season = 1)

        for i in range(pred_hor_time_steps):
            Q_loss.append((T_in[i] - T_out[i]) * self.k / 1000)
            Q_tab.append((T_tab[i] - T_in[i]) * (alpha[self.settings.season] / 1000 * self.area))

            T_tab.append((Q_heat[i] - Q_tab[i]) / self.cp_tab * (self.dt / 3600) + T_tab[i])
            T_in.append((Q_tab[i] + Q_solar[i] - Q_loss[i]) / self.cp_r * (self.dt / 3600) + T_in[i])

            # print ("Q_heat[",i,"]:", Q_heat[i], "Q_TAB:", Q_Tab[i], "T_tab:", T_tab[i], "T_in", T_in[i])

        return np.array(T_in[:-1]), np.array(T_tab[:-1])

    # region PREDICTION HORIZON CONVERSION METHODS

    def convert_pred_hor(self, array, mode, dynamic=False):
        """Convert array from one prediction horizon to another.

        Parameters
        ----------
        array : numpy.array
            array to be converted
        mode : str
            conversion mode (from longer to shorter prediction horizon, or from shorter to longer)
        dynamic : bool
            dynamic conversion flag (takes longer, but works with any time step - otherwise only 1 hour or 15 min)
        """

        if dynamic:
            return self._convert_dynamic(array, mode)

        if self.dt not in [3600, 3600 / 4]:
            raise ValueError(
                'Hard coded conversion methods only work with a time step of 1 hour or 15 min. Add additional hard coded'
                ' conversion method of use convert_pred_hor with dynamic=True instead.')

        elif self.dt == 3600:  # time step is 1 hour
            if mode == 'long2short':
                return self._convert_48_16(array)
            elif mode == 'short2long':
                return self._convert_16_48(array)

        elif self.dt == 3600 / 4:  # time step is 15 min
            if mode == 'long2short':
                return self._convert_192_64(array)
            elif mode == 'short2long':
                return self._convert_64_192(array)

    def _convert_dynamic(self, array, mode):
        """Dynamic converter.

        Alternative to the hard coded converters, which converts an array automatically based on the time step. Although
        this method is more versatile, it also needs more resources/time to run. If that is an issue, use hard coded
        converters instead.

        Parameters
        ----------
        array : numpy.array
            array to be converted
        mode : str
            conversion mode ('long2short' or 'short2long')
        """

        def adapt_hour_indices(hour_indices):
            """Adapt hour indices to new time step."""

            if steps_per_hour == 1:
                return hour_indices

            # adapt indices to new time step
            adapted = multiply_nested_list(hour_indices, steps_per_hour)

            # single indices are now multiple indices
            for i, indices in enumerate(adapted):
                if len(indices) == 1:
                    adapted[i] = [indices[0], indices[0] + steps_per_hour]

            return adapted

        def conditional_conversion(indices_to, indices_from):
            """Perform conversion depending on the passed indices.

            Parameters
            ----------
            indices_to : list[int]
                contains start and end of index range, for the receiving array.
            indices_from : list[int]
                contains start and end of index range, for the source array.
            """

            def range2indices(range_list):
                """Turn list with range (e.g. [0,4]) into list of actual values (e.g. [0, 1, 2, 3]."""
                if len(range_list) > 1:
                    return list(range(range_list[0], range_list[1]))
                else:
                    return range_list

            # turn index range into actual index list
            indices_to = range2indices(indices_to)
            indices_from = range2indices(indices_from)

            # length of index lists
            len_to = len(indices_to)
            len_from = len(indices_from)

            # conversion, method depends on the passed index ranges
            if len_from == 1 or len_to == len_from:
                result[indices_to] = array[indices_from]
            elif len_to > len_from:
                # interpolate and floor
                indices_from = np.floor(interpolate(indices_from, int(len_to / len_from))).astype(int).tolist()
                result[indices_to] = array[indices_from]
            elif len_to < len_from:
                # averaging downsampling
                a = list(range(0, len_from, int(len_from / len_to))) + [len_from]
                for i in range(len_to):
                    result[indices_to[i]] = np.mean(array[indices_from[a[i]:a[i + 1]]])

        # region HARD CODED HOUR INDICES FOR CONVERSION

        """follows the following logic:
          long_array[0, 6] = short_array[0, 6]
          long_array[6, 8] = short_array[6]
          ...
          
          If there are more than one time steps per hour, the indices have to be adapted accordingly. See (complicated)
          code further below.
          """

        long = [
            [0, 6],
            [6, 8],
            [8, 10],
            [10, 12],
            [12, 15],
            [15, 18],
            [18, 21],
            [21, 24],
            [24, 30],
            [30, 36],
            [36, 48],
        ]

        short = [
            [0, 6],
            [6],
            [7],
            [8],
            [9],
            [10],
            [11],
            [12],
            [13],
            [14],
            [15],
        ]

        # endregion

        steps_per_hour = int(3600 / self.dt)

        long = adapt_hour_indices(long)
        short = adapt_hour_indices(short)

        if mode == 'long2short':
            result = np.zeros(self.settings.pred_hor_short * steps_per_hour)
            for i in range(len(long)):
                conditional_conversion(short[i], long[i])
        elif mode == 'short2long':
            result = np.zeros(self.settings.pred_hor * steps_per_hour)
            for i in range(len(long)):
                conditional_conversion(long[i], short[i])

        return result

    def _convert_16_48(self, Q_heat_s):
        """Hard coded converter from 16 to 48 values"""
        Q_heat = np.zeros(self.settings.pred_hor)

        Q_heat[0:6] = Q_heat_s[0:6]
        Q_heat[6:8] = Q_heat_s[6]
        Q_heat[8:10] = Q_heat_s[7]
        Q_heat[10:12] = Q_heat_s[8]
        Q_heat[12:15] = Q_heat_s[9]
        Q_heat[15:18] = Q_heat_s[10]
        Q_heat[18:21] = Q_heat_s[11]
        Q_heat[21:24] = Q_heat_s[12]
        Q_heat[24:30] = Q_heat_s[13]
        Q_heat[30:36] = Q_heat_s[14]
        Q_heat[36:48] = Q_heat_s[15]

        return Q_heat

    def _convert_48_16(self, Q_heat):
        """Hard coded converter from 48 to 16 values"""
        Q_heat_s = np.zeros(self.settings.pred_hor_short)

        Q_heat_s[0:6] = Q_heat[0:6]
        Q_heat_s[6] = mean(Q_heat[6:7])
        Q_heat_s[7] = mean(Q_heat[8:9])
        Q_heat_s[8] = mean(Q_heat[10:11])
        Q_heat_s[9] = mean(Q_heat[12:14])
        Q_heat_s[10] = mean(Q_heat[15:17])
        Q_heat_s[11] = mean(Q_heat[18:20])
        Q_heat_s[12] = mean(Q_heat[21:23])
        Q_heat_s[13] = mean(Q_heat[24:29])
        Q_heat_s[14] = mean(Q_heat[30:35])
        Q_heat_s[15] = mean(Q_heat[36:47])

        return Q_heat_s

    def _convert_64_192(self, Q_heat_s):
        """Hard coded converter from 64 to 192 values"""
        Q_heat = np.zeros(int(self.settings.pred_hor * 3600 / self.dt))

        Q_heat[0:24] = Q_heat_s[0:24]
        Q_heat[24:32] = Q_heat_s[[24, 24, 25, 25, 26, 26, 27, 27]]
        Q_heat[32:40] = Q_heat_s[[28, 28, 29, 29, 30, 30, 31, 31]]
        Q_heat[40:48] = Q_heat_s[[32, 32, 33, 33, 34, 34, 35, 35]]
        Q_heat[48:60] = Q_heat_s[[36, 36, 36, 37, 37, 37, 38, 38, 38, 39, 39, 39]]
        Q_heat[60:72] = Q_heat_s[[40, 40, 40, 41, 41, 41, 42, 42, 42, 43, 43, 43]]
        Q_heat[72:84] = Q_heat_s[[44, 44, 44, 45, 45, 45, 46, 46, 46, 47, 47, 47]]
        Q_heat[84:96] = Q_heat_s[[48, 48, 48, 49, 49, 49, 50, 50, 50, 51, 51, 51]]
        Q_heat[96:120] = Q_heat_s[[
            52, 52, 52, 52, 52, 52, 53, 53, 53, 53, 53, 53, 54, 54, 54, 54, 54, 54, 55, 55, 55, 55, 55, 55]]
        Q_heat[120:144] = Q_heat_s[[
            56, 56, 56, 56, 56, 56, 57, 57, 57, 57, 57, 57, 58, 58, 58, 58, 58, 58, 59, 59, 59, 59, 59, 59]]
        Q_heat[144:192] = Q_heat_s[[
            60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 61, 61, 61, 61, 61, 61, 61, 61, 61, 61, 61, 61, 62, 62, 62,
            62, 62, 62, 62, 62, 62, 62, 62, 62, 63, 63, 63, 63, 63, 63, 63, 63, 63, 63, 63, 63]]

        return Q_heat

    def _convert_192_64(self, Q_heat):
        """Hard coded converter from 192 to 64 values"""
        Q_heat_s = np.zeros(int(self.settings.pred_hor_short * 3600 / self.dt))

        Q_heat_s[0:24] = Q_heat[0:24]
        Q_heat_s[24] = np.mean(Q_heat[24:26])
        Q_heat_s[25] = np.mean(Q_heat[26:28])
        Q_heat_s[26] = np.mean(Q_heat[28:30])
        Q_heat_s[27] = np.mean(Q_heat[30:32])
        Q_heat_s[28] = np.mean(Q_heat[32:34])
        Q_heat_s[29] = np.mean(Q_heat[34:36])
        Q_heat_s[30] = np.mean(Q_heat[36:38])
        Q_heat_s[31] = np.mean(Q_heat[38:40])
        Q_heat_s[32] = np.mean(Q_heat[40:42])
        Q_heat_s[33] = np.mean(Q_heat[42:44])
        Q_heat_s[34] = np.mean(Q_heat[44:46])
        Q_heat_s[35] = np.mean(Q_heat[46:48])
        Q_heat_s[36] = np.mean(Q_heat[48:51])
        Q_heat_s[37] = np.mean(Q_heat[51:54])
        Q_heat_s[38] = np.mean(Q_heat[54:57])
        Q_heat_s[39] = np.mean(Q_heat[57:60])
        Q_heat_s[40] = np.mean(Q_heat[60:63])
        Q_heat_s[41] = np.mean(Q_heat[63:66])
        Q_heat_s[42] = np.mean(Q_heat[66:69])
        Q_heat_s[43] = np.mean(Q_heat[69:72])
        Q_heat_s[44] = np.mean(Q_heat[72:75])
        Q_heat_s[45] = np.mean(Q_heat[75:78])
        Q_heat_s[46] = np.mean(Q_heat[78:81])
        Q_heat_s[47] = np.mean(Q_heat[81:84])
        Q_heat_s[48] = np.mean(Q_heat[84:87])
        Q_heat_s[49] = np.mean(Q_heat[87:90])
        Q_heat_s[50] = np.mean(Q_heat[90:93])
        Q_heat_s[51] = np.mean(Q_heat[93:96])
        Q_heat_s[52] = np.mean(Q_heat[96:102])
        Q_heat_s[53] = np.mean(Q_heat[102:108])
        Q_heat_s[54] = np.mean(Q_heat[108:114])
        Q_heat_s[55] = np.mean(Q_heat[114:120])
        Q_heat_s[56] = np.mean(Q_heat[120:126])
        Q_heat_s[57] = np.mean(Q_heat[126:132])
        Q_heat_s[58] = np.mean(Q_heat[132:138])
        Q_heat_s[59] = np.mean(Q_heat[138:144])
        Q_heat_s[60] = np.mean(Q_heat[144:156])
        Q_heat_s[61] = np.mean(Q_heat[156:168])
        Q_heat_s[62] = np.mean(Q_heat[168:180])
        Q_heat_s[63] = np.mean(Q_heat[180:192])

        return Q_heat_s

    # endregion


class SettingsMPC:
    """Class for storing settings of SimulationSeries object."""

    def __init__(self):
        self.pred_hor = 48  # prediction horizon [h]
        self.pred_hor_short = 16  # shortened prediction horizon (to run the program faster) [h]
        self.dHeat = 0.5  # perturbation value [kW]
        self.season = 0  # heating or cooling: heating = 1, cooling = 0
        self.setpoint_temperature = 20

        self.max_count = 500  # max. runs of iteration possible
        self.ChgProgTol = 0.000005  # termination criterion optimization - change in least square error

        # start conditions for optimization
        self.T_start_in = 22  # room temperature [°C]
        self.T_start_tab = 22  # thermally activated building [°C]

        self.header = ["Stunde", "T_out", "Q_solar", "T_sp"]

    def load_settings(self):
        pass

    def save_settings(self):
        pass

    def reset_settings(self):
        pass


def lse(T_in, T_sp):
    """Calculate least square error. todo: Variablennamen inhalts-unabhängig umbenennen (also keiner Temperaturen)"""
    return sum(pow((T_in - T_sp), 2))


def multiply_nested_list(nested_list, factor):
    """Multiply each value inside a nested list of numbers with a factor.

    Parameters
    ----------
    nested_list : list[list[int]]
        nested list of numbers to be multiplied
    factor : int
        factor by which each number is multiplied
    """

    new_nested_list = []
    for list in nested_list:
        new_list = []
        for value in list:
            new_list.append(int(value * factor))
        new_nested_list.append(new_list)

    return new_nested_list


def interpolate(y, factor):
    """Change length of a given array by a certain factor using interpolation.

    Parameters
    ----------
    y : y values of array
    factor : factor by which the length of the array is to be multiplied
    """

    x = range(0, len(y))
    x_interp = [val / factor for val in range(0, int(len(y) * factor))]
    y_interp = np.interp(x_interp, x, y)

    return y_interp
