import sys
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED - needed as tkinter has compatibility issues with pywinauto

import classes

sim_series = classes.SimulationSeries('C:\TRNSYS18\Exe\TrnEXE64.exe', 'Simulationsvarianten.xlsx')

sim_series.import_input_excel()
sim_series.start_sim_series()
# sim_series.start_sim_series_par()

