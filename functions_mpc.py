# -*- coding: utf-8 -*-
"""
Created on Thu Aug  3 14:22:05 2023

@author: Magdalena
"""

import numpy as np
from statistics import mean


def convert_16_48(Q_heat_s, n):
    Q_heat = np.zeros(n)

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
