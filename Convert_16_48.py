# -*- coding: utf-8 -*-
"""
Created on Thu Aug  3 14:22:05 2023

@author: Magdalena
"""
import numpy as np

def Convert_16_48 (Q_heat_s, n):
    
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
        #print(i, Q_heat[i], Q_heat_s[i])

    return(Q_heat)
