# -*- coding: utf-8 -*-
"""
Created on Thu Feb  6 14:55:04 2020

@author: Magdalena
"""
import numpy as np

#Function to calculate the building model
def Building(Q_heat, Q_solar, T_out, T_start_in, T_start_TAB, season, n):
    
    #Variable definition
    dt = 3600 #Time interval controller in [s] => 1 h
         
    #BUILDING A
    #Building specific parameter
    Area = 158.46    #Activated Area Building A [m²]
    alpha_w = 6.5    #alpha Winter = 6.5 [W/m²K]
    alpha_s = 10.75  #alpha Sommer = 10.75 [W/m²K]
    k = 120.71       #Factor for convection, transition and ventilation losses [W/K]
    cp_TAB = 23.45   #m_TAB*cp_TAB = 23.45 [kWh/K]
    cp_R = 54.91     #m_R*cp_R = 54.91 [kWh/K]      
        
    #Starting values and definition
    Q_Loss = np.zeros(n)
    Q_Tab = np.zeros(n)
    T_Tab = np.zeros(n)
    T_in = np.zeros(n)

    T_Tab[0] = T_start_TAB;
    T_in[0] = T_start_in;
    
    for i in range(n):
        # Prediction convection, transition and ventilation losses [kW]
        Q_Loss[i] = (T_in[i]-T_out[i])*k/1000
        
        if season == 1: # Heating in winter
            # Thermal heat flow from TAB to room [kW]
            # alpha Winter = 6.5 W/m²K
            Q_Tab[i] = (T_Tab[i]-T_in[i])*(alpha_w/1000*Area)
            
        if season == 0: # Cooling in summer
            # Thermal heat flow from room to TAB [kW]
            # alpha summer = 10.75 W/m²K
            Q_Tab[i] = (T_Tab[i]-T_in[i])*(alpha_s/1000*Area)
 
        if i < n-1:       #to avoid array overflow
            # Temperature TAB [°C]; m_TAB*cp_TAB = 23.45 kWh/K
            T_Tab[i+1] = (Q_heat[i]-Q_Tab[i])/cp_TAB*(dt/3600) + T_Tab[i]
            # Prediction Room Temperature [°C]; m_R*cp_R = 54.91 kWh/K
            T_in[i+1] = (Q_Tab[i]+Q_solar[i]-Q_Loss[i])/cp_R*(dt/3600) + T_in[i]
        
        #print ("Q_heat[",i,"]:", Q_heat[i], "Q_TAB:", Q_Tab[i], "T_Tab:", T_Tab[i], "T_in", T_in[i])

    return(T_in, T_Tab)
        
