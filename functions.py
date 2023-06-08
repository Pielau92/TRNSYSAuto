import re
import pandas as pd

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
        if match.group(1).casefold() == '@' + name.casefold():  # find case insensitive
            return match.group(1)[1:] + '=' + str(value)


def find_and_replace_param(path_file, pattern, parameters):
    """

    Parameters
    ----------
    path_file : path of the .txt file
    pattern :
    parameters :

    Returns
    -------

    """

    # read file
    with open(path_file, 'r') as file:
        text = file.read()

        # find, then replace parameter values
        text = re.sub(pattern=pattern, repl=lambda match: replace_param(match, parameters), string=text)

    # overwrite file
    with open(path_file, 'w') as file:
        file.write(text)

def import_settings_excel(path_settings_excel='C:/Users/pierre/PycharmProjects/TimberBioC/Basisordner/Einstellungen.xlsx'):

    excel_data = pd.ExcelFile(path_settings_excel)  # read input Excel file
    df = excel_data.parse('Einstellungen', index_col=0)  # Excel data as DataFrame
    return df.Wert
