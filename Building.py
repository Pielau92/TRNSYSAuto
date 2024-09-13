# -*- coding: utf-8 -*-
"""
Created on Thu Feb  6 14:55:04 2020

@author: Magdalena
"""
import numpy as np


class Building:
    """Building model class.

    Building model class with a thermally activated building (TAB) component.
    """

    def __init__(self, area, alpha_w, alpha_s, k, cp_tab, cp_r, dt=3600):
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
        self.dt = dt

    def predict(self, Q_heat, Q_solar, T_out, T_start_in, T_start_tab, season, forecast_period):
        """Predict room air temperature and temperature of thermally activated building (TAB) component.

        Parameters
        ----------
        Q_heat : heating/cooling power [kW]
        Q_solar : heat loads from solar radiation [kW]
        T_out : outside air temperature [°C]
        T_start_in : starting room air temperature [°C]
        T_start_tab : starting temperature of thermally activated building component [°C]
        season : heating = 1, cooling = 0
        forecast_period : prediction horizon [h]

        Returns
        -------
        T_in : room air temperature [°C]
        T_tab : temperature of thermally activated building component [°C]

        """

        Q_loss = []  # prediction of convection, transition and ventilation losses [kW]
        Q_tab = []  # thermal heat flow between room and TAB component [kW]
        T_in = list([T_start_in])  # prediction of  room temperature [°C]
        T_tab = list([T_start_tab])  # temperature TAB [°C]

        alpha = [self.alpha_s, self.alpha_w]  # cooling in summer (season = 0), heating in winter (season = 1)

        for i in range(forecast_period - 1):
            Q_loss.append((T_in[i] - T_out[i]) * self.k / 1000)
            Q_tab.append((T_tab[i] - T_in[i]) * (alpha[season] / 1000 * self.area))

            T_tab.append((Q_heat[i] - Q_tab[i]) / self.cp_tab * (self.dt / 3600) + T_tab[i])
            T_in.append((Q_tab[i] + Q_solar[i] - Q_loss[i]) / self.cp_r * (self.dt / 3600) + T_in[i])

            # print ("Q_heat[",i,"]:", Q_heat[i], "Q_TAB:", Q_Tab[i], "T_tab:", T_tab[i], "T_in", T_in[i])

        return np.array(T_in), np.array(T_tab)
