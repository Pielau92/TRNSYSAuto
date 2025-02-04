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
    from assets.mpc_code.main_mpc import Building  # for debugging purposes

thisModule = os.path.splitext(os.path.basename(__file__))[0]

global building
global delimiter
global filename_logger
global Q_hc
global costs
global emissions


# Initialization: function called at TRNSYS initialization
# ----------------------------------------------------------------------------------------------------------------------
def Initialization(TRNData):
    return TRNData  # usually only empty return statement, but return TRNData for testing with pytest


# StartTime: function called at TRNSYS starting time (not an actual time step, initial values should be reported)
# ----------------------------------------------------------------------------------------------------------------------
def StartTime(TRNData):
    global building
    global delimiter
    global filename_logger
    global Q_hc
    global costs
    global emissions

    inputs = TRNData[thisModule]["inputs"]
    path_trnsys_input_file = TRNData[thisModule]["TRNSYS input file path"]
    path_settings_file = os.path.dirname(path_trnsys_input_file)
    Q_hc = {}
    costs = {}
    emissions = {}

    # TRNSYS input into building object
    building = Building(area=inputs[0],         # [W]
                        alpha_w=inputs[1],      # [W/m²K]
                        alpha_s=inputs[2],      # [W/m²K]
                        k=inputs[3],            # [W/K]
                        cp_tab=inputs[4],       # [Wh/K]
                        cp_r=inputs[5],         # [Wh/k]
                        max_heating=inputs[6],  # [W]
                        max_cooling=inputs[7],  # [W]
                        dt_trnsys=TRNData[thisModule]["simulation time step"] * 3600,  # same time step like TRNSYS [s]
                        )

    # load settings
    building.settings.load_settings(path_settings_file)
    building.settings.apply_settings()

    building.settings.season = int(inputs[8])           # heating or cooling: heating = 1, cooling = 0
    building.settings.setpoint_temperature = inputs[9]
    building.settings.T_start_in = inputs[10]           # room temperature [°C]
    building.settings.T_start_tab = inputs[11]          # thermally activated building component temperature [°C]
    building.settings.dt_pred = 3600                    # time step during prediction/optimization

    # read external data sources
    building.read_weather_data(path_trnsys_input_file)
    building.read_electricity_price_data(path_trnsys_input_file)

    # interpolate external data if necessary
    if not building.dt_trnsys == 3600:
        building.interpolate_external_data()

    building.get_price_signal()

    # write TRNSYS predefined variables into log file
    with open(building.path_logFile, 'w') as f:
        for var_name in TRNData[thisModule]:
            f.write(f'{var_name}:  {str(TRNData[thisModule][var_name])} \n')
        f.write('\n')

    filename_logger = f'log_values_zone{str(int(inputs[12]))}.log'

    # header lists
    headers_inputs \
        = ['area_BTA [m²]', 'alpha_w [W/m²K]', 'alpha_s [W/m²K]', 'k_heatloss [W/K]', 'cbta [Wh/K]', 'cr [Wh/K]',
           'qheizmax [W]', 'qkuehlmax [W]', 'heizperiode [bool]', 'tbtasoll [°C]', 'tzone [°C]', 'tnodeo [°C]', 'Zone']

    headers_time_steps = ['index', 'heizperiode [bool]', 'tbtasoll [°C]', 'tzone [°C]', 'tnodeo [°C]', 'Zone',
                          'Qheat [W]', 'Costs [€]', 'Emissions [gCO2eq]']

    # add delimiter
    delimiter = '\t'
    headers_inputs = delimiter.join(headers_inputs)
    headers_time_steps = delimiter.join(headers_time_steps)

    # write into values logger...
    with open(filename_logger, 'w') as f:
        f.write(f'{headers_inputs}\n')  # ...header of input values from TRNSYS
        for value in inputs:  # ...input values from TRNSYS
            f.write(f'{round(value, 2)}{delimiter}'.replace('.', ','))
        f.write(f'\n\n{headers_time_steps}')  # ...header of simulation variable values

    return TRNData  # usually only empty return statement, but return TRNData for testing with pytest


# Iteration: function called at each TRNSYS iteration within a time step
# ----------------------------------------------------------------------------------------------------------------------
def Iteration(TRNData):
    inputs = TRNData[thisModule]["inputs"]
    zone_nr = str(int(inputs[12]))

    building.time_step_nr = TRNData[thisModule]["current time step number"] - 1

    # region FOR DEBUGGING PURPOSES

    # skip any heating zone that is not zone 1
    if not zone_nr == "1":
        TRNData[thisModule]["outputs"][0] = 1
        return TRNData  # usually only empty return statement, but return TRNData for testing with pytest

    skip_hours = False

    if skip_hours:
        start_hour = 4500
        stop_hour = 4900
        skip_condition = not \
            (start_hour * 3600 / building.dt_trnsys) < building.time_step_nr < (stop_hour * 3600 / building.dt_trnsys)

        # skip optimization algorithm, return dummy value instead
        if skip_condition:
            TRNData[thisModule]["outputs"][0] = 10
            return TRNData  # usually only empty return statement, but return TRNData for testing with pytest

    # endregion

    # "Iteration" triggers every n time steps (mpc_trigger = 1 => every time step, 2 = every second time step etc.)
    remainder = building.time_step_nr % building.settings.mpc_trigger
    if not remainder:

        # update values
        building.settings.season = int(inputs[8])  # heating or cooling: heating = 1, cooling = 0
        building.settings.setpoint_temperature = inputs[9]
        building.settings.T_start_in = inputs[10]  # room temperature [°C]
        building.settings.T_start_tab = inputs[11]  # thermally activated building [°C]

        # pass initial guess from last iteration, if available
        if zone_nr in Q_hc.keys():
            initial_guess = np.append(Q_hc[zone_nr][1:], Q_hc[zone_nr][-1])
            Q_hc[zone_nr], T_in, T_tab, costs[zone_nr], emissions[zone_nr] = building.optimize(initial_guess)
        else:
            Q_hc[zone_nr], T_in, T_tab, costs[zone_nr], emissions[zone_nr] = building.optimize()

    # output first value of Q_heat [W] (or if the mpc controller does not trigger in this iteration: take value from
    # output of last time it was triggered
    TRNData[thisModule]["outputs"][0] = Q_hc[zone_nr][remainder]

    # write to values logger
    log_outputs = [building.settings.season, building.settings.setpoint_temperature, inputs[10], inputs[11],
                   int(zone_nr), TRNData[thisModule]["outputs"][0], costs[zone_nr][remainder],
                   emissions[zone_nr][remainder]]
    filename_logger = f'log_values_zone{zone_nr}.log'
    with open(filename_logger, 'a') as f:
        # f.write(f'\n{row}')
        f.write(f'\n{building.time_step_nr}')
        for log_output in log_outputs:
            f.write(f'{delimiter}{round(log_output, 2)}'.replace('.', ','))

    return TRNData  # usually only empty return statement, but return TRNData for testing with pytest


# EndOfTimeStep: function called at the end of each time step, after iteration and before moving on to next time step
# ----------------------------------------------------------------------------------------------------------------------
def EndOfTimeStep(TRNData):
    return TRNData  # usually only empty return statement, but return TRNData for testing with pytest


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

    return TRNData  # usually only empty return statement, but return TRNData for testing with pytest
