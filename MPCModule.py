# Python module for the TRNSYS Type calling Python using CFFI
# Data exchange with TRNSYS uses a dictionary, called TRNData in this file (it is the argument of all functions).
# Data for this module will be in a nested dictionary under the module name,
# i.e. if this file is called "MyScript.py", the inputs will be in TRNData["MyScript"]["inputs"]
# for convenience the module name is saved in thisModule
#
# MKu, 2022-02-15

import numpy
import os

from classes_mpc import Building

thisModule = os.path.splitext(os.path.basename(__file__))[0]

global building


# Initialization: function called at TRNSYS initialization
# ----------------------------------------------------------------------------------------------------------------------
def Initialization(TRNData):

    return


# StartTime: function called at TRNSYS starting time (not an actual time step, initial values should be reported)
# ----------------------------------------------------------------------------------------------------------------------
def StartTime(TRNData):
    global building

    inputs = TRNData[thisModule]["inputs"]

    # TRNSYS input
    building = Building(area=inputs[0],
                        alpha_w=inputs[1],
                        alpha_s=inputs[2],
                        k=inputs[3],
                        cp_tab=inputs[4],
                        cp_r=inputs[5],
                        max_heating=inputs[6],
                        max_cooling=inputs[7],
                        dt=TRNData[thisModule]["simulation time step"])

    building.settings.season = inputs[8]  # heating or cooling: heating = 1, cooling = 0
    building.settings.setpoint_temperature = inputs[9]
    building.settings.T_start_in = inputs[10]  # room temperature [°C]
    building.settings.T_start_tab = inputs[11]  # thermally activated building [°C]

    return


# Iteration: function called at each TRNSYS iteration within a time step
# ----------------------------------------------------------------------------------------------------------------------
def Iteration(TRNData):

    # python output
    TRNData[thisModule]["outputs"][0] = building.optimize()[0]     # first value of Q_heat

    return


# EndOfTimeStep: function called at the end of each time step, after iteration and before moving on to next time step
# ----------------------------------------------------------------------------------------------------------------------
def EndOfTimeStep(TRNData):
    # This model has nothing to do during the end-of-step call

    return


# LastCallOfSimulation: function called at the end of the simulation (once) - outputs are meaningless at this call
# ----------------------------------------------------------------------------------------------------------------------
def LastCallOfSimulation(TRNData):
    # NOTE: TRNSYS performs this call AFTER the executable (the online plotter if there is one) is closed.
    # Python errors in this function will be difficult (or impossible) to diagnose as they will produce no message.
    # A recommended alternative for "end of simulation" actions it to implement them in the EndOfTimeStep() part,
    # within a condition that the last time step has been reached.
    #
    # Example (to be placed in EndOfTimeStep()):
    #
    # stepNo = TRNData[thisModule]["current time step number"]
    # nSteps = TRNData[thisModule]["total number of time steps"]
    # if stepNo == nSteps-1:     # Remember: TRNSYS steps go from 0 to (number of steps - 1)
    #     do stuff that needs to be done only at the end of simulation

    return
