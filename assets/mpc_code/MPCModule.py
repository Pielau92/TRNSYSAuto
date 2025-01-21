# Python module for the TRNSYS Type calling Python using CFFI
# Data exchange with TRNSYS uses a dictionary, called TRNData in this file (it is the argument of all functions).
# Data for this module will be in a nested dictionary under the module name,
# i.e. if this file is called "MyScript.py", the inputs will be in TRNData["MyScript"]["inputs"]
# for convenience the module name is saved in thisModule
#
# MKu, 2022-02-15

import numpy as np
import os

try:
    from main_mpc import Building
except ModuleNotFoundError:
    from assets.mpc_code.main_mpc import Building   # for debugging purposes

thisModule = os.path.splitext(os.path.basename(__file__))[0]

global building  # building class
global Q_heat_start
global value_loggers
global delimiter
global filename_logger

# Initialization: function called at TRNSYS initialization
# ----------------------------------------------------------------------------------------------------------------------
def Initialization(TRNData):
    return TRNData  # usually only empty return statement, but return TRNData for testint with pytest


# StartTime: function called at TRNSYS starting time (not an actual time step, initial values should be reported)
# ----------------------------------------------------------------------------------------------------------------------
def StartTime(TRNData):
    global building
    global value_loggers
    global delimiter
    global filename_logger
    global Q_heat_start

    Q_heat_start = {}

    inputs = TRNData[thisModule]["inputs"]

    path_trnsys_input_file = TRNData[thisModule]["TRNSYS input file path"]
    path_settings_file = os.path.dirname(path_trnsys_input_file)

    # TRNSYS input
    building = Building(area=inputs[0],                 # [W]
                        alpha_w=inputs[1],              # [W/m²K]
                        alpha_s=inputs[2],              # [W/m²K]
                        k=inputs[3],                    # [W/K]
                        cp_tab=inputs[4],               # [Wh/K]
                        cp_r=inputs[5],                 # [Wh/k]
                        max_heating=inputs[6] * 1000,   # [W]
                        max_cooling=inputs[7] * 1000,   # [W]
                        dt_trnsys=TRNData[thisModule]["simulation time step"] * 3600,  # same time step like TRNSYS [s]
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
    building.read_electricity_price_data(path_trnsys_input_file)

    if not building.dt_trnsys == 3600:
        building.interpolate_external_data()

    # write TRNSYS predefined variables into log file
    with open(building.path_logFile, 'w') as f:
        for var_name in TRNData[thisModule]:
            f.write(f'{var_name}:  {str(TRNData[thisModule][var_name])} \n')
        f.write('\n')

    # determine values logger filename
    zone_nr = int(inputs[12])
    filename_logger = f'log_values_zone{str(zone_nr)}.log'

    # header lists
    headers_inputs = ['area_BTA [m²]', 'alpha_w [W/m²K]', 'alpha_s [W/m²K]', 'k_heatloss [W/K]', 'cbta [Wh/K]',
               'cr [Wh/K]', 'qheizmax [kW]', 'qkuehlmax [kW]', 'heizperiode [bool]', 'theizsollminideal [°C]',
               'tzone [°C]', 'tnodeo [°C]', 'Zone']
    headers_time_steps = ['index', 'heizperiode [bool]', 'theizsollminideal [°C]',
               'tzone [°C]', 'tnodeo [°C]', 'Zone', 'Qheat [kW]']

    # add delimiter
    delimiter = '\t'
    headers_inputs = delimiter.join(headers_inputs)
    headers_time_steps = delimiter.join(headers_time_steps)

    # write into values logger...
    with open(filename_logger, 'w') as f:
        f.write(f'{headers_inputs}\n')   # ...header of input values from TRNSYS
        for value in inputs:    # ...input values from TRNSYS
            f.write(f'{round(value, 2)}{delimiter}'.replace('.', ','))
        f.write(f'\n\n{headers_time_steps}')    # ...header of time step values

    return TRNData  # usually only empty return statement, but return TRNData for testint with pytest


# Iteration: function called at each TRNSYS iteration within a time step
# ----------------------------------------------------------------------------------------------------------------------
def Iteration(TRNData):
    global Q_heat_start

    inputs = TRNData[thisModule]["inputs"]
    zone_nr = int(inputs[12])

    building.time_step_nr = TRNData[thisModule]["current time step number"] - 1

    # region FOR DEBUGGING PURPOSES

    # skip optimization algorithm until this time step, return dummy value instead
    skip_to = 365 * 24 * 3600 / building.dt_trnsys - 100  # skip to those last iterations
    skip_to = None  

    if skip_to and building.time_step_nr < (365*24*3600/building.dt_trnsys - skip_to):
        TRNData[thisModule]["outputs"][0] = 10
        return TRNData  # usually only empty return statement, but return TRNData for testint with pytest

    # endregion

    # "Iteration" triggers every n time steps
    # (mpc_trigger = 1 => every time step, 2 = every second time step, and so on)
    if not building.time_step_nr % building.settings.mpc_trigger:
        # update values
        building.settings.season = int(inputs[8])  # heating or cooling: heating = 1, cooling = 0
        building.settings.setpoint_temperature = inputs[9]
        building.settings.T_start_in = inputs[10]  # room temperature [°C]
        building.settings.T_start_tab = inputs[11]  # thermally activated building [°C]

        if str(zone_nr) in Q_heat_start.keys():
            Q_heat, T_in, T_tab = building.optimize(Q_heat_start[str(zone_nr)])  # python output
        else:
            Q_heat, T_in, T_tab = building.optimize()  # python output

    Q_heat_start[str(zone_nr)] = \
        np.append(Q_heat[1:], Q_heat[-1])  # predicted heating power as starting point in next iteration
    TRNData[thisModule]["outputs"][0] = Q_heat[0] / 1000    # output is first value of Q_heat, kW

    # write to values logger
    log_outputs = [building.settings.season, building.settings.setpoint_temperature, inputs[10], inputs[11], zone_nr,
     TRNData[thisModule]["outputs"][0]]
    filename_logger = f'log_values_zone{str(zone_nr)}.log'
    with open(filename_logger, 'a') as f:
        # f.write(f'\n{row}')
        f.write(f'\n{building.time_step_nr}')
        for log_output in log_outputs:
            f.write(f'{delimiter}{log_output}'.replace('.', ','))

    return TRNData  # usually only empty return statement, but return TRNData for testint with pytest


# EndOfTimeStep: function called at the end of each time step, after iteration and before moving on to next time step
# ----------------------------------------------------------------------------------------------------------------------
def EndOfTimeStep(TRNData):

    return TRNData  # usually only empty return statement, but return TRNData for testint with pytest


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

    return TRNData  # usually only empty return statement, but return TRNData for testint with pytest
