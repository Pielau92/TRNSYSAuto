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

# region VARIABLE & CONSTANT DEFINITION
n = 48  # prediction horizon
n_s = 16  # shortened horizon, to run the program faster
dHeat = 0.5  # perturbation value
season = 0  # heating or cooling: heating = 1, cooling = 0
setpoint_temperature = 20

counter = 0  # while loop counter
max_count = 500  # max. runs of iteration possible
ChgProgTol = 0.000005  # termination criterion optimization - change in least square error
ChgProgress = 1  # termination criterion optimization - difference between lse_baseline and lse_neu_long

# max power [kW]
MaxHtg = 13  # heating - reduced from 6.5 kW to 3.5 kW on 14.11.2019 //Both TOPS: 13 kW
MinHtg = -10  # cooling - both TOPS: -10 kW

# start conditions for optimization
T_start_in = 22     # room temperature [°C]
T_start_tab = 22    # thermally activated building [°C]

# variable initialization
Q_heat = np.zeros(n)
Q_heat_s = np.zeros(n_s)
Q_help = np.zeros(n)
Q_help_s = np.zeros(n_s)
# endregion

# import data from csv
df = pd.read_csv("Test_MPC_Python.csv",
                 encoding="latin1",
                 header=0,
                 names=["Stunde", "T_out", "Q_solar", "T_sp"],  # header
                 index_col=False,
                 delimiter=";",
                 decimal=".")
df["T_sp"] = setpoint_temperature

# building object
building = Building(area=158.46,
                    alpha_w=6.5,
                    alpha_s=10.75,
                    k=120.71,
                    cp_tab=23.45,
                    cp_r=54.91,
                    dt=3600)

while counter < max_count and ChgProgress >= ChgProgTol:

    # baseline calculation
    T_in, T_tab = building.calculate(Q_heat, df["Q_solar"], df["T_out"], T_start_in, T_start_tab, season, n)
    lse_baseline = lse(T_in, df["T_sp"])  # calculate least square error for zero heat input / heat output
    Q_heat_s = convert_48_16(Q_heat, n_s)  # shorten Q_heat

    # loop to go through the elements of the vector
    for i in range(n_s):

        # negative perturbation
        Q_help_s[i] = max(Q_heat_s[i] - dHeat, MinHtg)  # limit to minimum cooling power
        Q_help = convert_16_48(Q_help_s, n)  # expand Q_help
        T_in, T_tab = building.calculate(Q_help, df["Q_solar"], df["T_out"], T_start_in, T_start_tab, season, n)
        lse_negative = lse(T_in, df["T_sp"])  # least square error, negative perturbation

        # positive perturbation
        Q_help_s[i] = min(Q_heat_s[i] + dHeat, MaxHtg)  # limit to maximum heating power
        Q_help = convert_16_48(Q_help_s, n)  # expand Q_help
        T_in, T_tab = building.calculate(Q_help, df["Q_solar"], df["T_out"], T_start_in, T_start_tab, season, n)
        lse_positive = lse(T_in, df["T_sp"])  # least square error, positive perturbation

        # interpretation of the perturbation effect
        match np.argmin([lse_baseline, lse_negative, lse_positive]):
            case 0:  # baseline calculation has lowest least square error
                pass
            case 1:  # negative perturbation has lowest least square error
                Q_heat_s[i] -= dHeat
            case 2:  # positive perturbation has lowest least square error
                Q_heat_s[i] += dHeat

        # limitation that cooling and heating in one period is not possible
        if season:  # heating
            min_value = 0  # no simultaneous heating and cooling in one period
            max_value = MaxHtg  # limit to maximum heating power
        elif not season:  # cooling
            min_value = MinHtg  # limit to maximum cooling power
            max_value = 0  # no simultaneous heating and cooling in one period
        Q_heat_s[i] = np.clip(Q_heat_s[i], min_value, max_value)

        Q_help_s[i] = Q_heat_s[i]  # reset of the helping variable to not forget the value

    Q_heat = convert_16_48(Q_heat_s, n)  # expand heating vector to prediction horizon
    T_in, T_tab = building.calculate(Q_heat, df["Q_solar"], df["T_out"], T_start_in, T_start_tab, season, n)
    lse_neu_long = lse(T_in, df["T_sp"])  # calculate least square error for final perturbation in this loop run

    # calculate termination criterion: least square error from start compared LSE with last perturbation run
    ChgProgress = lse_baseline - lse_neu_long

    # output of final least square error
    print(counter, ". Durchgang: ", lse_neu_long)

    # loop counter
    counter += 1

# graphical evaluation
# fig, ax1 = plt.subplots()
# # ax1.plot(df.index,df["T_out"], color = "blue", label = "Außentemperatur")
# # ax1.plot(df.index,df["Q_solar"], color = "orange", label = "Solare Einstrahlung")
# ax1.plot(df.index, df["T_sp"], color="red", label="Solltemperatur")
# ax1.plot(df.index, Q_heat, color="brown", label="Heizleistung")
# ax1.plot(df.index, T_in, color="green", label="Präd. Raumtemperatur")
# ax1.plot(df.index, T_tab, color="violet", label="TAB-Temperatur")
# ax1.legend(loc="best")
#
# figure_props = {
#     "title": "Test MPC",
#     "ylabel": "Temperatur [°C] / Solar Radiation [kW]",
#     # "ylim": [15,30],
# }
# ax1.set(**figure_props)
