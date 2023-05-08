import pandas as pd
import re
from pywinauto import application   #, keyboard, Desktop
#   import pywinauto


def import_input_excel(path):
    """

    Parameters
    ----------
    path : str
        Path of the input Escel file

    Returns
    -------
    sim_list : list of strings
        List of simulation names
    weather_series : pandas.Series
        Series with the weather data file name, for each simulation
    df_dck : pandas.DataFrame
        Dataframe with the simulation parameters to be replaced in the .dck File, for each simulation
    b18_series : pandas.Series
        Series with the .b18 data file name, for each simulation

    """

    excel_data = pd.ExcelFile(path)  # read input Excel file

    df = excel_data.parse('Simulationsvarianten', index_col=0)  # Excel data as DataFrame

    # transpose data
    df_weather = df[df.index == 'Wetterdaten'].transpose()
    b18_series = df[df.index == 'b18'].transpose()
    df_dck = df[df.index == 'dck'].transpose()

    # convert to series
    weather_series = df_weather[1:].squeeze()
    b18_series = b18_series[1:].squeeze()

    # use first row as header
    df_dck.columns = df_dck.iloc[0]
    df_dck = df_dck[1:]

    # list of simulation variants
    sim_list = df.columns[1:].astype(str).tolist()

    return sim_list, weather_series, df_dck, b18_series


def find_and_replace_param(path_file, pattern, parameters):
    """

    Parameters
    ----------
    pattern :
    path_file :
    parameters :

    Returns
    -------

    """

    def replace_param(match, parameters):
        """

        Parameters
        ----------
        match :
        parameters :

        Returns
        -------

        """
        for name, value in parameters.items():
            if match.group(1).casefold() == '@' + name.casefold():
                return match.group(1)[1:] + '=' + str(value)

    with open(path_file, 'r') as file:
        text = file.read()  # read file

        # find, then replace parameter values
        text = re.sub(pattern=pattern, repl=lambda match: replace_param(match, parameters), string=text)

    with open(path_file, 'w') as file:
        file.write(text)  # overwrite file


def start_sim(path_exe, path_file):
    """

    Parameters
    ----------
    path_exe : string
        Path of the .exe file of the simulation application to be started
    path_file : string
        Path of the file needed for the simulation (in this case Building.dck)
    """

    path_file = path_file.replace("/", "\\")
    # path_file_raw = r'{}'.format(path_file)     # turn file path into raw string to avoid error messages
    app = application.Application()
    app.start(path_exe)

    app_window = app["Öffnen"]  # Hier muss der Titel des Fensters eingesetzt werden
    app_window.wait('ready')
    app_window.set_focus()
    app_window.FileNameEdit.set_edit_text(path_file)    #"C:\Users\pierre\PycharmProjects\TimberBioC\Basisordner\Building.dck"
    app_window.set_focus()
    app_window.type_keys("{ENTER}")
