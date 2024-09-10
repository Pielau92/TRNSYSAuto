# -*- coding: utf-8 -*-
"""
Created on Thu Feb  6 14:55:04 2020

@author: Magdalena
"""
import numpy as np


class Building:
    """Building model class."""

    def __init__(self, area, alpha_w, alpha_s, k, cp_tab, cp_r, dt=3600):

        # building specific parameters
        self.area = area        # activated building area [m²]
        self.alpha_w = alpha_w  # alpha value (winter) [W/m²K]
        self.alpha_s = alpha_s  # alpha value(sommer) [W/m²K]
        self.k = k              # factor for convection, transition and ventilation losses [W/K]
        self.cp_tab = cp_tab    # m_TAB * cp_tab [kWh/K]
        self.cp_r = cp_r        # m_R*cp_r [kWh/K]
        self.dt = dt            # time interval [s] (e.g 3600 = 1 hour)

    def calculate(self, Q_heat, Q_solar, T_out, T_start_in, T_start_TAB, season, n):

        # Starting values and definition
        Q_Loss = np.zeros(n)
        Q_Tab = np.zeros(n)
        T_Tab = np.zeros(n)
        T_in = np.zeros(n)

        T_Tab[0] = T_start_TAB
        T_in[0] = T_start_in

        for i in range(n):
            # Prediction convection, transition and ventilation losses [kW]
            Q_Loss[i] = (T_in[i] - T_out[i]) * self.k / 1000

            if season == 1:  # Heating in winter
                # Thermal heat flow from TAB to room [kW]
                # alpha Winter = 6.5 W/m²K
                Q_Tab[i] = (T_Tab[i] - T_in[i]) * (self.alpha_w / 1000 * self.area)

            if season == 0:  # Cooling in summer
                # Thermal heat flow from room to TAB [kW]
                # alpha summer = 10.75 W/m²K
                Q_Tab[i] = (T_Tab[i] - T_in[i]) * (self.alpha_s / 1000 * self.area)

            if i < n - 1:  # to avoid array overflow
                # Temperature TAB [°C]; m_TAB*cp_tab = 23.45 kWh/K
                T_Tab[i + 1] = (Q_heat[i] - Q_Tab[i]) / self.cp_tab * (self.dt / 3600) + T_Tab[i]
                # Prediction Room Temperature [°C]; m_R*cp_r = 54.91 kWh/K
                T_in[i + 1] = (Q_Tab[i] + Q_solar[i] - Q_Loss[i]) / self.cp_r * (self.dt / 3600) + T_in[i]

            # print ("Q_heat[",i,"]:", Q_heat[i], "Q_TAB:", Q_Tab[i], "T_Tab:", T_Tab[i], "T_in", T_in[i])

        return T_in, T_Tab
