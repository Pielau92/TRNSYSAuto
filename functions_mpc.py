# -*- coding: utf-8 -*-
"""
Created on Thu Aug  3 14:22:05 2023

@author: Magdalena
"""

import numpy as np


def Convert_16_48(Q_heat_s, n):
    Q_heat = np.zeros(n)

    for i in range(n):
        if i < 6:
            Q_heat[i] = Q_heat_s[i]
        if i >= 6 and i < 8:
            Q_heat[i] = Q_heat_s[6]
        if i >= 8 and i < 10:
            Q_heat[i] = Q_heat_s[7]
        if i >= 10 and i < 12:
            Q_heat[i] = Q_heat_s[8]
        if i >= 12 and i < 15:
            Q_heat[i] = Q_heat_s[9]
        if i >= 15 and i < 18:
            Q_heat[i] = Q_heat_s[10]
        if i >= 18 and i < 21:
            Q_heat[i] = Q_heat_s[11]
        if i >= 21 and i < 24:
            Q_heat[i] = Q_heat_s[12]
        if i >= 24 and i < 30:
            Q_heat[i] = Q_heat_s[13]
        if i >= 30 and i < 36:
            Q_heat[i] = Q_heat_s[14]
        if i >= 36 and i < 48:
            Q_heat[i] = Q_heat_s[15]
        # print(i, Q_heat[i], Q_heat_s[i])

    return (Q_heat)


def Convert_48_16(Q_heat, n_s):
    Q_heat_s = np.zeros(n_s)

    for i in range(6):
        Q_heat_s[i] = Q_heat[i]

    Q_heat_s[6] = (Q_heat[6] + Q_heat[7]) / 2
    Q_heat_s[7] = (Q_heat[8] + Q_heat[9]) / 2
    Q_heat_s[8] = (Q_heat[10] + Q_heat[11]) / 2

    Q_heat_s[9] = (Q_heat[12] + Q_heat[13] + Q_heat[14]) / 3
    Q_heat_s[10] = (Q_heat[15] + Q_heat[16] + Q_heat[17]) / 3
    Q_heat_s[11] = (Q_heat[18] + Q_heat[19] + Q_heat[20]) / 3
    Q_heat_s[12] = (Q_heat[21] + Q_heat[22] + Q_heat[23]) / 3

    Q_heat_s[13] = (Q_heat[24] + Q_heat[25] + Q_heat[26] + Q_heat[27] + Q_heat[28] + Q_heat[29]) / 6
    Q_heat_s[14] = (Q_heat[30] + Q_heat[31] + Q_heat[32] + Q_heat[33] + Q_heat[34] + Q_heat[35]) / 6

    Q_heat_s[15] = (Q_heat[36] + Q_heat[37] + Q_heat[38] + Q_heat[39] + Q_heat[40] + Q_heat[41] + Q_heat[42] + Q_heat[
        43] + Q_heat[44] + Q_heat[45] + Q_heat[46] + Q_heat[47]) / 12

    return (Q_heat_s)


def LeastSquareError(T_in, T_sp):
    LeastSQE = sum(pow((T_in - T_sp), 2))
    # print ("LSE: ", LeastSQE)
    return (LeastSQE)
