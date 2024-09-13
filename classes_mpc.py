# -*- coding: utf-8 -*-
"""
Created on Thu Feb  6 14:55:04 2020

@author: Magdalena
"""
import numpy as np
import pandas as pd

from statistics import mean
from functions_mpc import lse


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
        dt : time interval [s] (e.g 3600 = 1 hour)
        """

        # building specific parameters
        self.area = area
        self.alpha_w = alpha_w
        self.alpha_s = alpha_s
        self.k = k
        self.cp_tab = cp_tab
        self.cp_r = cp_r
        self.max_heating = max_heating
        self.max_cooling = max_cooling
        self.dt = dt

        self.settings = SettingsMPC()

        self.df = None  # DataFrame with data

    def import_csv(self, filename):
        """Import csv data."""

        self.df = pd.read_csv(filename,
                              encoding="latin1",
                              header=0,
                              names=["Stunde", "T_out", "Q_solar", "T_sp"],  # header
                              index_col=False,
                              delimiter=";",
                              decimal=".")

        self.df["T_sp"] = self.settings.setpoint_temperature

    def optimize(self):
        """todo"""

        dHeat = self.settings.dHeat

        Q_heat = np.zeros(self.settings.pred_hor)
        Q_help_s = np.zeros(self.settings.pred_hor_short)

        counter = 0
        ChgProgress = 1  # termination criterion optimization - difference between lse_baseline and lse_neu_long
        while counter < self.settings.max_count and ChgProgress >= self.settings.ChgProgTol:

            # baseline calculation
            T_in, T_tab = self.predict(Q_heat, self.df["Q_solar"], self.df["T_out"])
            lse_baseline = lse(T_in, self.df["T_sp"])  # least square error for zero heat input / heat output
            Q_heat_s = self.convert_48_16(Q_heat)  # shorten Q_heat

            # loop through hours of prediction horizon
            for i in range(self.settings.pred_hor_short):

                # negative perturbation
                Q_help_s[i] = max(Q_heat_s[i] - dHeat, self.max_cooling)  # limit to minimum cooling power
                Q_help = self.convert_16_48(Q_help_s)  # expand Q_help
                T_in, T_tab = self.predict(Q_help, self.df["Q_solar"], self.df["T_out"])
                lse_negative = lse(T_in, self.df["T_sp"])  # least square error, negative perturbation

                # positive perturbation
                Q_help_s[i] = min(Q_heat_s[i] + dHeat, self.max_heating)  # limit to maximum heating power
                Q_help = self.convert_16_48(Q_help_s)  # expand Q_help
                T_in, T_tab = self.predict(Q_help, self.df["Q_solar"], self.df["T_out"])
                lse_positive = lse(T_in, self.df["T_sp"])  # least square error, positive perturbation

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

            Q_heat = self.convert_16_48(Q_heat_s)  # expand heating vector to prediction horizon
            T_in, T_tab = self.predict(Q_heat, self.df["Q_solar"], self.df["T_out"])
            lse_neu_long = lse(T_in, self.df["T_sp"])  # least square error for final perturbation in this loop run

            # calculate termination criterion: lse from start compared to lse from last perturbation run
            ChgProgress = lse_baseline - lse_neu_long

            # output of final least square error
            print(counter, ". Durchgang: ", lse_neu_long)

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

        Q_loss = []  # prediction of convection, transition and ventilation losses [kW]
        Q_tab = []  # thermal heat flow between room and TAB component [kW]
        T_in = list([self.settings.T_start_in])  # prediction of  room temperature [°C]
        T_tab = list([self.settings.T_start_tab])  # temperature TAB [°C]

        alpha = [self.alpha_s, self.alpha_w]  # cooling in summer (season = 0), heating in winter (season = 1)

        for i in range(self.settings.pred_hor - 1):
            Q_loss.append((T_in[i] - T_out[i]) * self.k / 1000)
            Q_tab.append((T_tab[i] - T_in[i]) * (alpha[self.settings.season] / 1000 * self.area))

            T_tab.append((Q_heat[i] - Q_tab[i]) / self.cp_tab * (self.dt / 3600) + T_tab[i])
            T_in.append((Q_tab[i] + Q_solar[i] - Q_loss[i]) / self.cp_r * (self.dt / 3600) + T_in[i])

            # print ("Q_heat[",i,"]:", Q_heat[i], "Q_TAB:", Q_Tab[i], "T_tab:", T_tab[i], "T_in", T_in[i])
    def convert_16_48(self, Q_heat_s):
        """todo"""
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

    def convert_48_16(self, Q_heat):
        """todo"""
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
