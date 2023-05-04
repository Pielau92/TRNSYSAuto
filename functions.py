import pandas as pd
import re


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
        Series with the weather data file name for each simulation
    parameters_df : pandas.DataFrame
        DataFrame with parameters and their respective value, which are used to overwrite the text files

    """

    excel_data = pd.ExcelFile(path)  # read input Excel file

    df = excel_data.parse('Simulationsvarianten', index_col=0)  # Excel data as DataFrame

    sim_list = df.index.astype(str).tolist()
    weather_series = df.Wetterdaten
    parameters_df = df.iloc[:, 1:]

    return sim_list, weather_series, parameters_df


def find_and_replace_param(path_file, parameters):
    """

    Parameters
    ----------
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
                return match.group(1) + '=' + str(value)  # todo: Soll der @ bleiben oder beim Tausch gelöscht werden?

    with open(path_file, 'r') as file:
        text = file.read()  # read file

        # find, then replace parameter values
        text = re.sub(r'(@\w+)\s*=\s*([\d.]+)', lambda match: replace_param(match, parameters), text)

    with open(path_file, 'w') as file:
        file.write(text)  # overwrite file
