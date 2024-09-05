# -*- coding: utf-8 -*-
"""
Created on Thu Aug  3 14:22:05 2023

@author: Magdalena
"""

def LeastSquareError (T_in, T_sp):
   LeastSQE = sum(pow((T_in-T_sp),2))
   #print ("LSE: ", LeastSQE)
   return(LeastSQE)
