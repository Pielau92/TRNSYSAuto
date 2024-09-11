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

    def calculate(self, Q_heat, Q_solar, T_out, T_start_in, T_start_tab, season, n):
        """Calculate todo: spezifizieren

        Parameters
        ----------
        Q_heat :
        Q_solar :
        T_out :
        T_start_in :
        T_start_tab :
        season :
        n :

        Returns
        -------
        T_in :
        T_tab :

        """

        # starting values
        Q_loss = np.zeros(n)    # prediction convection, transition and ventilation losses [kW]
        Q_tab = np.zeros(n)     # thermal heat flow [kW]
        T_in = np.zeros(n)      # prediction room temperature [°C]; m_R * (cp_r = 54.91 kWh/K)
        T_tab = np.zeros(n)     # temperature TAB [°C]; m_TAB *cp_tab = 23.45 kWh/K

        T_tab[0] = T_start_tab
        T_in[0] = T_start_in

        for i in range(n):

            Q_loss[i] = (T_in[i] - T_out[i]) * self.k / 1000

            if season:
                # heating in winter (heat flow from TAB to room)
                Q_tab[i] = (T_tab[i] - T_in[i]) * (self.alpha_w / 1000 * self.area)
            else:
                # cooling in summer (heat flow from room to TAB)
                Q_tab[i] = (T_tab[i] - T_in[i]) * (self.alpha_s / 1000 * self.area)

            if i < n - 1:  # to avoid array overflow
                T_tab[i + 1] = (Q_heat[i] - Q_tab[i]) / self.cp_tab * (self.dt / 3600) + T_tab[i]
                T_in[i + 1] = (Q_tab[i] + Q_solar[i] - Q_loss[i]) / self.cp_r * (self.dt / 3600) + T_in[i]

            # print ("Q_heat[",i,"]:", Q_heat[i], "Q_TAB:", Q_Tab[i], "T_tab:", T_tab[i], "T_in", T_in[i])

        return T_in, T_tab
