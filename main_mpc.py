# -*- coding: utf-8 -*-
"""
Created on Thu Aug  3 13:08:01 2023

@author: Magdalena
"""

# import matplotlib.pyplot as plt
from classes_mpc import Building

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
building.import_csv("Test_MPC_Python.csv")

result = building.optimize()

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
