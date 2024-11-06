# Python module for the TRNSYS Type calling Python using CFFI
# Data exchange with TRNSYS uses a dictionary, called TRNData in this file (it is the argument of all functions).
# Data for this module will be in a nested dictionary under the module name,
# i.e. if this file is called "MyScript.py", the inputs will be in TRNData["MyScript"]["inputs"]
# for convenience the module name is saved in thisModule
#
# MKu, 2022-02-15

import numpy
import os

from main_mpc import Building

thisModule = os.path.splitext(os.path.basename(__file__))[0]

global building # building class
global temp     # temporary value holder


# Initialization: function called at TRNSYS initialization
# ----------------------------------------------------------------------------------------------------------------------
def Initialization(TRNData):

    return


# StartTime: function called at TRNSYS starting time (not an actual time step, initial values should be reported)
# ----------------------------------------------------------------------------------------------------------------------
def StartTime(TRNData):
    global building

    inputs = TRNData[thisModule]["inputs"]

    path_trnsys_input_file = TRNData[thisModule]["TRNSYS input file path"]
    path_settings_file = os.path.dirname(path_trnsys_input_file)

    # TRNSYS input
    building = Building(area=inputs[0],
                        alpha_w=inputs[1],
                        alpha_s=inputs[2],
                        k=inputs[3],
                        cp_tab=inputs[4],
                        cp_r=inputs[5],
                        max_heating=inputs[6],
                        max_cooling=inputs[7],
                        dt_trnsys=TRNData[thisModule]["simulation time step"] * 3600,  # same time step like TRNSYS
                        )

    building.settings.load_settings(path_settings_file)
    building.settings.apply_settings()

    building.settings.season = int(inputs[8])  # heating or cooling: heating = 1, cooling = 0
    building.settings.setpoint_temperature = inputs[9]
    building.settings.T_start_in = inputs[10]  # room temperature [°C]
    building.settings.T_start_tab = inputs[11]  # thermally activated building [°C]
    building.settings.dt_pred = 3600
    building.settings.pred_hor_conversion = True

    building.read_weather_data(path_trnsys_input_file)

    # write TRNSYS predefined variables into log file
    for var_name in TRNData[thisModule]:
        building.logFile.write(f'{var_name}:  {str(TRNData[thisModule][var_name])} \n')
    building.logFile.write('\n')

    return


# Iteration: function called at each TRNSYS iteration within a time step
# ----------------------------------------------------------------------------------------------------------------------
def Iteration(TRNData):

    global temp

    inputs = TRNData[thisModule]["inputs"]

    building.time_step_nr = TRNData[thisModule]["current time step number"] - 1

    # "Iteration" triggers every n time steps
    # (mpc_trigger = 1 => every time step, 2 = every second time step, and so on)
    if not building.time_step_nr % building.settings.mpc_trigger:

        # update values
        building.settings.season = int(inputs[8])  # heating or cooling: heating = 1, cooling = 0
        building.settings.setpoint_temperature = inputs[9]
        building.settings.T_start_in = inputs[10]  # room temperature [°C]
        building.settings.T_start_tab = inputs[11]  # thermally activated building [°C]

        temp = building.optimize()[0]     # python output, first value of Q_heat

    TRNData[thisModule]["outputs"][0] = temp

    return


# EndOfTimeStep: function called at the end of each time step, after iteration and before moving on to next time step
# ----------------------------------------------------------------------------------------------------------------------
def EndOfTimeStep(TRNData):

    # log into logfile
    building.logFile.write(f'time: {str(TRNData[thisModule]["time"])}\n')
    building.logFile.write(f'time step: {str(TRNData[thisModule]["current time step number"])}\n\n')

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
