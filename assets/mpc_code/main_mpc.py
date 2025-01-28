# -*- coding: utf-8 -*-
"""
Created on Thu Aug  3 13:08:01 2023

@author: Magdalena
"""
import numpy as np
import scipy.optimize as spo
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

        # electricity price data
        self.price_signal = None  # electricity price [€/kWh]

        self.time_step_nr = 0

        self.settings = SettingsMPC()

        self.path_logFile = "PythonLog.log"

    @property
    def alpha(self):  # todo: zu heat transfer coeff. ändern (Name "alpha" schon in verwendung)
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

    def read_electricity_price_data(self, path_trnsys_input_file,
                                    filename_price_data='EXAA_Day Ahead Preise & CO2-Intensität 2015-2022_1_2019.txt'):

        path_sim_dir = os.path.dirname(path_trnsys_input_file)
        path_price_data = os.path.join(path_sim_dir, filename_price_data)

        lines = Path(path_price_data).read_text().splitlines()
        reader = csv.reader(lines, delimiter='\t')
        # todo: offset und col_index_* in settings verstauen
        offset = 2

        # column index of specific rows
        if self.settings.price_signal == "COST":    # energy price data [€/MWh
            col_index = 2
            factor = 1/1000
        elif self.settings.price_signal == "CO2":   # CO2 emissions [gCO2eq/kWh]
            col_index = 3
            factor = 1

        electricity_price = []
        for index, row in enumerate(reader):
            if index < offset:
                continue  # apply offset by skipping the first rows
            electricity_price.append(float(row[col_index]) * factor)

        # apply offset to remove negative values
        price_signal = np.array(electricity_price) + abs(min(electricity_price))

        # save as numpy array
        self.price_signal = price_signal

    def interpolate_external_data(self):
        """Interpolate weather data to right length."""

        self.ta = interpolate(self.ta, 3600 / self.dt_trnsys)
        self.igs = interpolate(self.igs, 3600 / self.dt_trnsys)
        self.ign = interpolate(self.ign, 3600 / self.dt_trnsys)
        self.price_signal = interpolate(self.price_signal, 3600 / self.dt_trnsys)

    def optimize(self, initial_guess=None):
        """Optimize the heating/cooling power.

         Calculates the optimal heating/cooling power on the basis of the following framework:
            - tries to reach the setpoint temperature within a defined prediction horizon
            - constrained by the maximal heating and cooling power respectively
            - variable energy costs

        Parameters
        ----------
        initial_guess : numpy.ndarray[float]
            Initial guess for the heating/cooling power within the prediction horizon

        Returns
        -------
        Q_heat : numpy.ndarray[float]
            Calculated heating/cooling power, reaching the set point temperature within the prediction horizon
        T_in : numpy.ndarray[float]
            Predicted room temperature, resulting from the computed heating/cooling power
        T_tab : numpy.ndarray[float]
            Predicted TAB temperature, resulting from the computed heating/cooling power
        """

        def get_lse(Q_hc):
            """Get least square error using modified formula."""

            _T_in, _T_tab = self.predict(Q_hc, Q_solar, T_out)
            cost_pred = get_heatpump_costs(Q_hc)
            alpha = get_alpha(_T_in - T_sp)

            return np.sum((alpha * np.abs(_T_in - T_sp) ** self.settings.beta) + cost_pred ** self.settings.gamma)

        def get_alpha(temperature_deviation):
            """Return alpha factor (for the lse calculation) from the deviation of the room temperature from the
            setpoint temperature."""

            alpha_val = []
            for val in temperature_deviation:
                if val >= 0:
                    alpha_val.append(self.settings.alpha_pos)
                elif val < 0:
                    alpha_val.append(self.settings.alpha_neg)

            return alpha_val

        def get_indices():
            """Get list of indices for the current prediction horizon.

            Takes the following into account:
                - prediction horizon
                - current time step
                - time step durations (of TRNSYS and prediction)
                - prediction horizon exceeding the end of the current year
             """

            index_list = list(range(
                self.time_step_nr, self.time_step_nr + pred_hor_time_steps_trnsys,
                int(self.settings.dt_pred / self.dt_trnsys)))

            time_steps_per_year = int(365 * 24 * 3600 / self.dt_trnsys)

            # if prediction horizon reaches the following year, reset to beginning of year to form a cycle
            for _i, val in enumerate(index_list):
                if val < time_steps_per_year:
                    continue
                else:
                    index_list[_i] -= time_steps_per_year

            return index_list

        def get_heatpump_costs(Q):
            """Get energy costs of heat pump in €."""

            if not self.settings.cost_optimization:
                return Q * 0    # no costs

            # coefficient of performance (COP) when heating, energy efficiency ration (EER) when cooling
            f = [self.settings.eer, self.settings.cop][self.settings.season]

            return (np.abs(Q) / f) * (self.dt_trnsys / 3600) * electricity_price

        # prediction horizon in terms of time steps, instead of hours
        pred_hor_time_steps_trnsys = int(self.settings.pred_hor * 3600 / self.dt_trnsys)
        pred_hor_time_steps_pred = int(self.settings.pred_hor * 3600 / self.settings.dt_pred)

        indices = get_indices()  # get right indices of weather and electricity price data

        Q_solar = self.igs[indices]  # solar radiation [W/m²]
        T_out = self.ta[indices]  # outside temperature [°C]
        electricity_price = self.price_signal[indices]  # [€/kWh]
        # electricity_price[electricity_price < 0] = 0  # no electricity price < 0 allowed, otherwise error

        # get initial guess for heating/cooling power [W]
        if initial_guess is None:
            Q_heat = np.zeros(pred_hor_time_steps_pred)
        else:
            Q_heat = initial_guess

        T_sp = [self.settings.setpoint_temperature] * len(Q_heat)  # [°C]

        result = self.scipy_solver(Q_heat, get_lse)

        T_in, T_tab = self.predict(result, Q_solar, T_out)
        costs = get_heatpump_costs(result)

        return result, T_in, T_tab, costs

    def scipy_solver(self, initial_guess, objective_fun):

        # define bounds
        bound = ((
                     [self.max_cooling, 0][self.settings.season],  # minimum heating/cooling power [W]
                     [0, self.max_heating][self.settings.season]  # maximum heating/cooling power [W]
                 ),)
        bounds = bound * len(initial_guess)

        # define options
        options = {'eps': 10}
        if self.settings.set_optimizer_tolerance:
            options.update({'ftol': self.settings.optimizer_tolerance, 'gtol': self.settings.optimizer_tolerance})

        # optimize
        result = spo.minimize(fun=objective_fun, x0=initial_guess, bounds=bounds, options=options)

        result = result.x - result.x % self.settings.heat_pump_mod_step   # apply modulation step size

        return result

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
        T_in : numpy.array[float]
            Room air temperature [°C]
        T_tab : numpy.array[float]
            Temperature of thermally activated building component [°C]
        """

        pred_hor_time_steps = int(self.settings.pred_hor * 3600 / self.settings.dt_pred)

        T_in = [self.settings.T_start_in] + [0] * pred_hor_time_steps
        T_tab = [self.settings.T_start_tab] + [0] * pred_hor_time_steps

        for t in range(pred_hor_time_steps):
            Q_loss = (T_in[t] - T_out[t]) * self.k  # convection, transition and ventilation losses [W]
            Q_tab = (T_tab[t] - T_in[t]) * self.alpha * self.area  # thermal heat flow between room and TAB [W]

            T_tab[t + 1] = (Q_heat[t] - Q_tab) * (self.dt_trnsys / 3600) / self.cp_tab + T_tab[t]
            T_in[t + 1] = (Q_tab + Q_solar[t] - Q_loss) * (self.dt_trnsys / 3600) / self.cp_r + T_in[t]

        return np.array(T_in[:-1]), np.array(T_tab[:-1])


class SettingsMPC:
    """Class for storing settings of SimulationSeries object."""

    def __init__(self):

        self.filename = "settingsMPC.ini"
        self._settings = ConfigParser()
        self._settings.optionxform = str  # keeps capital letters when reading .ini file

        self.dt_pred = int()  # time step of the prediction [s]
        self.optimizer_tolerance = float()  # termination criterion optimization - change in least square error
        self.set_optimizer_tolerance = bool()  # True = pass optimizer_tolerance setting to optimizer - False = ignore
        self.pred_hor = int()  # prediction horizon [h]
        self.mpc_trigger = int()  # how often the mpc controller is triggered (1=every time step, 4=every 4th, etc.)
        self.cop = float()  # coefficient of performance (COP) of heat pump for heating
        self.eer = float()  # energy efficiency ratio (EER) of heat pump for cooling
        self.heat_pump_mod_step = int()  # modulation step size of heat pump [W]
        self.cost_optimization = bool()  # cost optimization flag
        self.price_signal = str()   # cost optimization can either take the energy price or CO2 emission into account

        # least square error calculation constants
        self.alpha_pos = float()  # alpha factor (positive deviation)
        self.alpha_neg = float()  # alpha factor (negative deviation)
        self.beta = float()  # beta exponent
        self.gamma = float()  # gamma exponent

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
