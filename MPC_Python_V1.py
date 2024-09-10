# -*- coding: utf-8 -*-
"""
Created on Thu Aug  3 13:08:01 2023

@author: Magdalena
"""

# import matplotlib.pyplot as plt
import pandas as pd
import numpy as np

from Building import Building
from functions_mpc import convert_16_48, convert_48_16, lse

# region Variable & Constant definition
n = 48  # prediction horizont
n_s = 16  # shortened horizont, to run the program faster
dHeat = 0.5  # perturbation value

ChgProgTol = 0.000005  # termination criterion optimization - Change in LSE
ChgProgress = 1  # termination criterion optimization - Difference between LSE_old and LSE

NrIt = 0  # number of iterations - while loop count
max_count = 500  # max. runs of iteration possible

MaxHtg = 13  # max. heating Power - reduced from 6.5 kW to 3.5 kW on 14.11.2019 //Both TOPS: 13 kW
MinHtg = -10  # max. cooling Power - both TOPS: -10 kW

T_start_in = 22  # start conditions for optimization
T_start_TAB = 22  # start conditions for optimization

season = 0  # heating or cooling: heating = 1, cooling = 0

point_temperature = 20

header = ["Stunde", "T_out", "Q_solar", "T_sp"]

# endregion

df = pd.read_csv("Test_MPC_Python.csv",
                 encoding="latin1",
                 header=0,
                 names=header,
                 index_col=False,
                 delimiter=";",
                 decimal=".")

df["T_sp"] = point_temperature  # set point temperatur

Q_heat = np.zeros(n)
Q_heat_s = np.zeros(n_s)
Q_help = np.zeros(n)
Q_help_s = np.zeros(n_s)

# building A
building = Building(area=158.46,
                    alpha_w=6.5,
                    alpha_s=10.75,
                    k=120.71,
                    cp_tab=23.45,
                    cp_r=54.91,
                    dt=3600)

while NrIt < max_count and ChgProgress >= ChgProgTol:

    # BASELINE CALCULATION
    T_in, T_Tab = building.calculate(Q_heat, df["Q_solar"], df["T_out"], T_start_in, T_start_TAB, season, n)
    Lse_0 = lse(T_in, df["T_sp"])  # calculate least square error for zero heat input / heat output

    Q_heat_s = convert_48_16(Q_heat, n_s)   # shorten Q_heat

    # loop to go through the elements of the vector
    for i in range(n_s):

        # region NEGATIVE PERTURBATION

        Q_help_s[i] = Q_heat_s[i] - dHeat  # negative perturbation

        if Q_help_s[i] <= MinHtg:  # limitation to minimum cooling power
            Q_help_s[i] = MinHtg

        Q_help = convert_16_48(Q_help_s, n)     # expand Q_heat

        T_in, T_Tab = building.calculate(Q_help, df["Q_solar"], df["T_out"], T_start_in, T_start_TAB, season, n)
        Lse_m = lse(T_in, df["T_sp"])  # least square error for negative perturbation

        # endregion

        # region POSITIVE PERTURBATION
        Q_help_s[i] = Q_heat_s[i] + dHeat  # positive perturbation

        if Q_help_s[i] >= MaxHtg:  # limitation to maximum heating power
            Q_help_s[i] = MaxHtg

        Q_help = convert_16_48(Q_help_s, n)     # expand Q_heat

        T_in, T_Tab = building.calculate(Q_help, df["Q_solar"], df["T_out"], T_start_in, T_start_TAB, season, n)
        Lse_p = lse(T_in, df["T_sp"])   # least square error for positive perturbation

        # endregion

        # interpretation of the perturbation effect
        Lse_s = [Lse_0, Lse_m, Lse_p]  # vector with LSE possibilities
        Best_Lse_s = min(Lse_s)  # best least square error of the LSE possibilities
        id_Best = Lse_s.index(
            min(Lse_s))  # position of the best least square error of the possibilities: 0 = LSE_0, 1 = LSE_m, 2 = LSE_p

        if id_Best == 0:  # if Best_LSE is LSE_0
            Q_heat_s[i] = Q_heat_s[i]

        if id_Best == 1:  # if Best_LSE is negative perturbation
            Q_heat_s[i] = Q_heat_s[i] - dHeat  # negative perturbation

        if id_Best == 2:  # if Best_LSE is positive perturbation
            Q_heat_s[i] = Q_heat_s[i] + dHeat  # positive perturbation

        # limitations that cooling and heating in one period is not possible
        if season == 0:  # cooling; limitation is max. cooling power
            if Q_heat_s[i] <= MinHtg:  # limitation to minimum cooling power
                Q_heat_s[i] = MinHtg
            if Q_heat_s[i] > 0:  # exclusion of simultaneous heating and cooling in one period
                Q_heat_s[i] = 0

        if season == 1:  # heating; limitation is max. heating power
            if Q_heat_s[i] >= MaxHtg:  # limitation to maximum heating power
                Q_heat_s[i] = MaxHtg
            if Q_heat_s[i] < 0:  # exclusion of simultaneous heating and cooling in one period
                Q_heat_s[i] = 0

        Q_help_s[i] = Q_heat_s[i]  # reset of the helping variable to not forget the value

    Q_heat = convert_16_48(Q_heat_s, n)     # expand heating vector to prediction horizont

    T_in, T_Tab = building.calculate(Q_heat, df["Q_solar"], df["T_out"], T_start_in, T_start_TAB, season, n)

    Lse_neu_long = lse(T_in, df["T_sp"])    # calculate least square error for final perturbation in this loop run
    # calculate termination criterion: least square error from start compared LSE with last perturbation run
    ChgProgress = Lse_0 - Lse_neu_long
    # output of final least square error
    print(NrIt, ". Durchgang: ", Lse_neu_long)
    # loop counter
    NrIt = NrIt + 1

# graphical evaluation
# fig, ax1 = plt.subplots()
# # ax1.plot(df.index,df["T_out"], color = "blue", label = "Außentemperatur")
# # ax1.plot(df.index,df["Q_solar"], color = "orange", label = "Solare Einstrahlung")
# ax1.plot(df.index, df["T_sp"], color="red", label="Solltemperatur")
# ax1.plot(df.index, Q_heat, color="brown", label="Heizleistung")
# ax1.plot(df.index, T_in, color="green", label="Präd. Raumtemperatur")
# ax1.plot(df.index, T_Tab, color="violet", label="TAB-Temperatur")
# ax1.legend(loc="best")
#
# figure_props = {
#     "title": "Test MPC",
#     "ylabel": "Temperatur [°C] / Solar Radiation [kW]",
#     # "ylim": [15,30],
# }
# ax1.set(**figure_props)
