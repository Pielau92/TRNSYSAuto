# -*- coding: utf-8 -*-
"""
Created on Thu Aug  3 14:22:05 2023

@author: Magdalena
"""


def lse(T_in, T_sp):
    """Calculate least square error. todo: Variablennamen inhaltsunabhängig umbenennen (also keiner Temperaturen)"""
    return sum(pow((T_in - T_sp), 2))
