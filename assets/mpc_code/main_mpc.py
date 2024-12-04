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
from configparser import ConfigParser


def main():
    # building object
    building = Building(area=158.46,
                        alpha_w=6.5,
                        alpha_s=10.75,
                        k=120.71,
                        cp_tab=23.45,
                        cp_r=54.91,
                        max_heating=13000,  # heating - reduced from 6.5 kW to 3.5 kW on 14.11.2019 //Both TOPS: 13 kW
                        max_cooling=-10000,  # cooling - both TOPS: -10 kW
                        dt_trnsys=3600)

    # import data from csv
    # building.import_csv("Test_MPC_Python.csv")

    result = building.optimize()


class Building:
    """Building model class.

    Building model class with a thermally activated building (TAB) component.
    """

    def __init__(self, area, alpha_w, alpha_s, k, cp_tab, cp_r, max_heating, max_cooling, dt_trnsys=3600):
        """Initialize building model.

        Parameters
        ----------
        area : float
            activated building area [m²]
        alpha_w : float
            heat transfer coefficient (winter) [W/m²K]
        alpha_s : float
            heat transfer coefficient (sommer) [W/m²K]
        k : float
            factor for convection, transition and ventilation losses [W/K]
        cp_tab : float
            absolute heat capacity of TAB component [Wh/K]
        cp_r : float
            absolute heat capacity of room [Wh/K]
        max_heating : float
            maximum heating power [W]
        max_cooling : float
            maximum cooling power [W]
        dt_trnsys : int
            time interval of trnsys simulation [s] (e.g 3600 = 1 hour)
        """

        # building specific parameters
        self.area = area
        self.alpha_w = alpha_w
        self.alpha_s = alpha_s
        self.k = k
        self.cp_tab = cp_tab
        self.cp_r = cp_r
        self.max_heating = max_heating
        self.max_cooling = max_cooling * -1
        self.dt_trnsys = dt_trnsys

        # weather data
        self.ta = None  # outside temperature [°C]
        self.igs = None  # global radiation, south [W/m²]
        self.ign = None  # global radiation, north [W/m²]
        self.time_step_nr = 0

        self.settings = SettingsMPC()

        self.path_logFile = "PythonLog.log"

    @property
    def alpha(self):
        """Heat transfer coefficient [W/m²K], depending on the current season."""
        return [self.alpha_s, self.alpha_w][self.settings.season]

    def read_weather_data(self, path_trnsys_input_file, filename_weather_data='Windetc20190804.txt'):
        """Read weather data for TRNSYS simulation.

        Parameters
        ----------
        path_trnsys_input_file : str
            Path to trnsys input file (dck file)
        filename_weather_data : str
            Filename of the csv file with weather data
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

        # save as numpy array
        self.ta = np.array(self.ta)
        self.igs = np.array(self.igs)
        self.ign = np.array(self.ign)

        if not self.dt_trnsys == 3600:
            self.interpolate_weather_data()

    def interpolate_weather_data(self):
        """Interpolate weather data to right length."""

        self.ta = interpolate(self.ta, 3600 / self.dt_trnsys)
        self.igs = interpolate(self.igs, 3600 / self.dt_trnsys)
        self.ign = interpolate(self.ign, 3600 / self.dt_trnsys)

    def optimize(self):
        """todo"""

        convert_Q = self.settings.pred_hor_conversion

        def get_lse(Q, convert=False):

            if convert:
                Q = self.convert_pred_hor(Q, 'short2long')  # expand
            T_in, T_tab = self.predict(Q, Q_solar, T_out)

            return lse(T_in, T_sp)

        def get_indices():
            """Get list of indices from current time step, time step durations (of TRNSYS and prediction) and prediction
            horizon."""

            index_list = list(range(
                self.time_step_nr, self.time_step_nr + pred_hor_time_steps,
                int(self.settings.dt_pred / self.dt_trnsys)))

            time_steps_per_year = int(365 * 24 * 3600 / self.dt_trnsys)

            # if prediction horizon reaches the following year, reset to beginning of year to form a cycle
            for i, val in enumerate(index_list):
                if val < time_steps_per_year:
                    continue
                else:
                    index_list[i] -= time_steps_per_year

            return index_list

        def get_heatpump_costs(Q):

            # coefficient of performance (COP) when heating, energy efficiency ration (EER) when cooling
            f = [self.settings.eer, self.settings.cop][self.settings.season]

            costs = (Q / f) * (3600 / self.dt_trnsys) * cost_pred

            return sum(costs)

        def get_costs(T_R, T_SP, Q, alpha_pos, alpha_neg, beta, gamma):

            result = []
            for (T_R_i, T_SP_i, Q_i) in zip(T_R, T_SP, Q):
                alpha = [alpha_neg, alpha_pos][T_R_i > T_SP_i]    # alpha_pos if T_room > T_setpoint, else alpha_neg
                result.append(alpha * abs(T_R_i - T_SP_i) ** beta + abs(get_heatpump_costs(Q)) ** gamma)

            return result
        
        # prediction horizon in terms of time steps, instead of hours
        pred_hor_time_steps = int(self.settings.pred_hor * 3600 / self.dt_trnsys)
        pred_hor_short_time_steps = int(self.settings.pred_hor_short * 3600 / self.dt_trnsys)

        cost_pred = np.array([0.1] * pred_hor_time_steps)  # [€/kWh] todo DUMMY
        alpha_pos, alpha_neg, beta, gamma = 1, 1, 4, 1.5    # todo DUMMIES

        indices = get_indices()

        Q_solar = self.igs[indices]  # W/m²
        T_out = self.ta[indices]  # °C

        dHeat = self.settings.dHeat  # W
        T_sp = [self.settings.setpoint_temperature] * int(self.settings.pred_hor * 3600 / self.settings.dt_pred)  # °C

        Q_heat = np.zeros(pred_hor_time_steps)  # W

        if convert_Q:
            Q_help = np.zeros(pred_hor_short_time_steps)  # W
        else:
            Q_help = np.zeros(pred_hor_time_steps)  # W

        counter = 0
        ChgProgress = 1
        while counter < self.settings.max_count and ChgProgress >= self.settings.ChgProgTol:

            # baseline calculation
            lse_baseline = get_lse(Q_heat)  # least square error for zero heat input / heat output

            if convert_Q:
                Q_heat = self.convert_pred_hor(Q_heat, 'long2short')  # shorten Q_heat

            # loop through hours of prediction horizon
            for i in range(len(Q_help)):

                # negative perturbation
                Q_help[i] = max(Q_heat[i] - dHeat, self.max_cooling)  # limit to minimum cooling power
                lse_negative = get_lse(Q_help, convert=convert_Q)  # least square error, negative perturbation

                # positive perturbation
                Q_help[i] = min(Q_heat[i] + dHeat, self.max_heating)  # limit to maximum heating power
                lse_positive = get_lse(Q_help, convert=convert_Q)  # least square error, positive perturbation

                # interpretation of perturbation effect
                match np.argmin([lse_baseline, lse_negative, lse_positive]):
                    case 0:  # baseline calculation has lowest least square error
                        pass
                    case 1:  # negative perturbation has lowest least square error
                        Q_heat[i] -= dHeat
                    case 2:  # positive perturbation has lowest least square error
                        Q_heat[i] += dHeat

                # limitation that cooling and heating simultaneously in one period is not possible
                min_value, max_value = 0, 0
                if self.settings.season:
                    max_value = self.max_heating  # limit to maximum heating power
                elif not self.settings.season:
                    min_value = self.max_cooling  # limit to maximum cooling power
                Q_heat[i] = np.clip(Q_heat[i], min_value, max_value)

                Q_help[i] = Q_heat[i]

            if convert_Q:
                Q_heat = self.convert_pred_hor(Q_heat, 'short2long')  # expand back to prediction horizon

            lse_final = get_lse(Q_heat)  # least square error, final perturbation
            ChgProgress = lse_baseline - lse_final
            counter += 1

        return Q_heat

    def predict(self, Q_heat, Q_solar, T_out):
        """Predict room air temperature and temperature of thermally activated building (TAB) component.

        Parameters
        ----------
        Q_heat : list[float]
            Heating/cooling power [W]
        Q_solar : list[float]
            Heat loads from solar radiation [W]
        T_out : list[float]
            Outside air temperature [°C]

        Returns
        -------
        T_in : list[float]
            Room air temperature [°C]
        T_tab : list[float]
            Temperature of thermally activated building component [°C]
        """

        pred_hor_time_steps = int(self.settings.pred_hor * 3600 / self.settings.dt_pred)

        T_in = [self.settings.T_start_in] + [0] * pred_hor_time_steps
        T_tab = [self.settings.T_start_in] + [0] * pred_hor_time_steps

        for t in range(pred_hor_time_steps):
            Q_loss = (T_in[t] - T_out[t]) * self.k  # convection, transition and ventilation losses [W]
            Q_tab = (T_tab[t] - T_in[t]) * self.alpha * self.area   # thermal heat flow between room and TAB [W]

            T_tab[t + 1] = (Q_heat[t] - Q_tab) * (self.dt_trnsys / 3600) / self.cp_tab + T_tab[t]
            T_in[t + 1] = (Q_tab + Q_solar[t] - Q_loss) * (self.dt_trnsys / 3600) / self.cp_r + T_in[t]

        return np.array(T_in[:-1]), np.array(T_tab[:-1])

    # region PREDICTION HORIZON CONVERSION METHODS

    def convert_pred_hor(self, array, mode, dynamic=False):
        """Convert array from one prediction horizon to another.

        Parameters
        ----------
        array : numpy.array
            Array to be converted
        mode : str
            Conversion mode (from longer to shorter prediction horizon, or from shorter to longer)
        dynamic : bool
            Dynamic conversion flag (takes longer, but works with any time step - otherwise only 1 hour or 15 min)
        """

        if dynamic:
            return self._convert_dynamic(array, mode)

        if self.dt_trnsys not in [3600, 3600 / 4]:
            raise ValueError(
                'Hard coded conversion methods only work with a time step of 1 hour or 15 min. Add additional hard coded'
                ' conversion method of use convert_pred_hor with dynamic=True instead.')

        elif self.dt_trnsys == 3600:  # time step is 1 hour
            if mode == 'long2short':
                return self._convert_48_16(array)
            elif mode == 'short2long':
                return self._convert_16_48(array)

        elif self.dt_trnsys == 3600 / 4:  # time step is 15 min
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
            Array to be converted
        mode : str
            Conversion mode ('long2short' or 'short2long')
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
                Contains start and end of index range, for the receiving array.
            indices_from : list[int]
                Contains start and end of index range, for the source array.
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

        steps_per_hour = int(3600 / self.dt_trnsys)

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
        Q_heat = np.zeros(int(self.settings.pred_hor * 3600 / self.dt_trnsys))

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
        Q_heat_s = np.zeros(int(self.settings.pred_hor_short * 3600 / self.dt_trnsys))

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

        self.filename = "settingsMPC.ini"
        self._settings = ConfigParser()
        self._settings.optionxform = str  # keeps capital letters when reading .ini file

        self.dHeat = int()  # perturbation value [W]
        self.dt_pred = int()  # time step of the prediction [s]
        self.max_count = int()  # max. runs of iteration possible
        self.ChgProgTol = float()  # termination criterion optimization - change in least square error
        self.pred_hor = int()  # prediction horizon [h]
        self.pred_hor_conversion = bool()  # perform BOKU conversion of prediction (to run the program faster)
        self.pred_hor_short = int()  # shortened prediction horizon (to run the program faster) [h]
        self.mpc_trigger = int()  # how often the mpc controller is triggered (1=every time step, 4=every 4th, etc.)
        self.cop = float()  # coefficient of performance (COP) of heat pump for heating
        self.eer = float()  # energy efficiency ratio (EER) of heat pump for cooling
        self.cost_optimization = bool()  # cost optimization flag

        # TRNSYS specific simulation parameters
        self.season = 0  # heating or cooling: heating = 1, cooling = 0
        self.setpoint_temperature = 20
        self.T_start_in = 22  # starting room temperature [°C]
        self.T_start_tab = 22  # starting temperature of thermally activated building component [°C]

        self.header = ["Stunde", "T_out", "Q_solar", "T_sp"]

    def load_settings(self, path_dir):
        """Load settings from settings.ini file.

        Parameters
        ----------
        path_dir : str
            Path to directory containing settings.ini file.
        """

        path_settings_file = os.path.join(path_dir, self.filename)
        try:
            self._settings.read(path_settings_file)
        except:
            print(f"Format error in settings file, check {self._save_path}")
            raise SystemExit()

    def apply_settings(self):
        """Apply loaded settings.

        Applies the loaded settings from the settings ini file to the corresponding attributes of the
        SimulationSeries object with the same name.
        """

        def apply_setting():
            """Apply setting value to sim_series.

            Applies the individual settings to the corresponding (name of setting and of class attribute must match).
            Automatically recognizes the type of the setting, based on the type of its corresponding class attribute.
            Raises an error if no corresponding class attribute could be found, or an unsupported type is used (str,
            int, float, bool, list (of strings)).
            """

            if not hasattr(self, setting):
                raise AttributeError(f'Unknown setting "{setting}" in settings.ini file found.')

            attr = getattr(self, setting)

            if isinstance(attr, bool):
                value = self._settings.getboolean(section, setting)
            elif isinstance(attr, str):
                value = self._settings.get(section, setting)
            elif isinstance(attr, int):
                value = self._settings.getint(section, setting)
            elif isinstance(attr, float):
                value = self._settings.getfloat(section, setting)
            elif isinstance(attr, list):
                items = self._settings.get(section, setting).split(',')  # apply comma (,) delimiter
                value = [item.strip() for item in items]  # remove whitespaces at beginning/end of strings
            else:
                raise TypeError(f'Unknown type "{type(attr)}" for setting "{setting}" in settings.ini file. '
                                f'Supported types are string, integer, float, boolean and list (of strings).')

            setattr(self, setting, value)

        for section in self._settings.sections():
            for setting in self._settings.options(section):
                apply_setting()  # save setting value into corresponding class attribute, with the correct datatype


def lse(T_in, T_sp):
    """Calculate least square error. todo: Variablennamen inhalts-unabhängig umbenennen (also keiner Temperaturen)"""
    return sum(pow((T_in - T_sp), 2))


def multiply_nested_list(nested_list, factor):
    """Multiply each value inside a nested list of numbers with a factor.

    Parameters
    ----------
    nested_list : list[list[int]]
        Nested list of numbers to be multiplied
    factor : int
        Factor by which each number is multiplied
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
    y : list
        y values of array
    factor : int
        Factor by which the length of the array is to be multiplied
    """

    x = range(0, len(y))
    x_interp = [val / factor for val in range(0, int(len(y) * factor))]
    y_interp = np.interp(x_interp, x, y)

    return y_interp
