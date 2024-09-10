# -*- coding: utf-8 -*-
"""
Created on Thu Aug  3 14:22:05 2023

@author: Magdalena
"""

import numpy as np
from statistics import mean


def convert_16_48(Q_heat_s, n):
    Q_heat = np.zeros(n)

    for i in range(n):
        if i < 6:
            Q_heat[i] = Q_heat_s[i]
        if 6 <= i < 8:
            Q_heat[i] = Q_heat_s[6]
        if 8 <= i < 10:
            Q_heat[i] = Q_heat_s[7]
        if 10 <= i < 12:
            Q_heat[i] = Q_heat_s[8]
        if 12 <= i < 15:
            Q_heat[i] = Q_heat_s[9]
        if 15 <= i < 18:
            Q_heat[i] = Q_heat_s[10]
        if 18 <= i < 21:
            Q_heat[i] = Q_heat_s[11]
        if 21 <= i < 24:
            Q_heat[i] = Q_heat_s[12]
        if 24 <= i < 30:
            Q_heat[i] = Q_heat_s[13]
        if 30 <= i < 36:
            Q_heat[i] = Q_heat_s[14]
        if 36 <= i < 48:
            Q_heat[i] = Q_heat_s[15]
        # print(i, Q_heat[i], Q_heat_s[i])

    return Q_heat


def convert_48_16(Q_heat, n_s):
    Q_heat_s = np.zeros(n_s)

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


def lse(T_in, T_sp):
    """Calculate least square error."""
    return sum(pow((T_in - T_sp), 2))
