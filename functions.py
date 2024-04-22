import win32com.client
import pandas as pd
import xlwings as xw
import tkinter as tk
import re
import pickle

from tkinter import filedialog  # explicit import required, as calling from tk.filedialog does not work properly


def replace_parameter_value(match, parameters):
    """Replace parameter value and remove marking from re.Match object.

    Parameters
    ----------
    match : re.Match
        re.Match object from the re.sub() method.
    parameters : pandas.Series
        Series with parameter values to be replaced, the parameter name has to be in the row index name.
    """
    for name, value in parameters.items():
        if match.group(1).casefold() == '@' + name.casefold():  # find, case-insensitive
            return match.group(1)[1:] + '=' + str(value)


def find_and_replace_parameter_values(path_file, pattern, parameters):
    """Find and replace parameter values within a .txt file.

    Finds parameter values within a .txt file, replaces them and overwrites the .txt file. For this, the following must
    apply:
    -   The parameter values must follow a marked parameter name, which follows a specific pattern in order to be found
    -   The replacement value has to be passed in "replacements" as a pandas.Series, where the row indizes match the
        name of the marked parameter

    Example:
        # .txt file containing "@Parameter1=1; Parameter2=2; @Parameter3=3."
        path_file = 'C:\\Users\\JohnDoe\\Desktop\\text.txt'

        # @ at the start, followed by text, '=' in the middle and a number at the end
        pattern = r'(@\w+)\s*=\s*([\d.]+)'

        replacements = pd.Series({
            'Parameter1': 4,
            'Parameter2': 5,
            'Parameter3': 6
        })

        find_and_replace_parameter_values(path_file, pattern, replacements)

        # .txt file now shows "Parameter1=4; Parameter2=2; Parameter3=6."

    Parameters
    ----------
    path_file : str
        Path of the .txt file.
    pattern : str
        Pattern which marks text to be replaced.
    parameters : pandas.Series
        Series with parameter values to be replaced, the parameter name has to be in the index name.
    """

    # read file
    with open(path_file, 'r') as file:
        text = file.read()

        # find, then replace text values
        text = re.sub(pattern=pattern,
                      repl=lambda match:  # variable "match" is produced by re.sub() itself
                      replace_parameter_value(match, parameters),  # apply replace_text() method on each matching string
                      string=text)

    # overwrite file
    with open(path_file, 'w') as file:
        file.write(text)


def find_and_replace(path_file, pattern, replacement):
    """Find string in text and replace by another string.

    Reads text from a .txt file, searches for a defined string, replaces it by another string and overwrites the .txt
    file.

    Parameters
    ----------
    path_file : str
        Path of the .txt file.
    pattern : str
        Pattern which marks text to be replaced.
    replacement : str
        Text to be inserted instead.
    """

    with open(path_file, 'r') as file:
        text = file.read()  # read file
        new_text = re.sub(pattern=pattern, repl=replacement, string=text)
    with open(path_file, 'w') as file:
        file.write(new_text)  # overwrite file


def update_excel_file(path_excel_file):
    """Update calculations in excel file.

    https://stackoverflow.com/questions/40893870/refresh-excel-external-data-with-python
    """

    xlapp = win32com.client.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open(path_excel_file)
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    wb.Save()
    xlapp.Quit()


def excel_export_variant_evaluation(sheet_name_variant_input, result, variant_folder, variant_output_file,
                                    variant_parameter_df):
    """Output routine for variant excel file."""

    wb = xw.Book(variant_output_file)
    ws = wb.sheets[sheet_name_variant_input]
    ws["A2"].options(pd.DataFrame, header=1, index=False, expand='table').value = variant_parameter_df[
        ['File', 'Parameter', variant_folder]]
    ws["B60"].options(pd.DataFrame, header=1, index=False, expand='table').value = result
    wb.save()
    wb.app.quit()


def to_single_column(df_input):
    """Turn pandas DataFrame into one single column.

    All columns of df_input are stacked over each other with one free row inbetween and the column headers on top.

    Example:
        df = pd.DataFrame({
        'col1': [1, 2, 3],
        'col2': [4, 5, 6],
        'col3': [7, 8, 9]
        })

        result = functions.to_single_column(df)
        print(result)
        """

    var_list_result_column = df_input.columns
    header = pd.DataFrame(var_list_result_column[1:] + ['']).transpose()
    header.columns = var_list_result_column[:-1] + ['']
    single_column = pd.concat([
        df_input[var_list_result_column],
        pd.DataFrame(index=['']),
        header],
        axis=0)
    # single_column = single_column.drop(single_column.columns[-1], axis=1)
    single_column = single_column.transpose().stack(dropna=False)
    return single_column


def progress_bar(progress, total):
    percent = 100 * (progress / float(total))
    bar = '█' * int(percent) + '-' * (100 - int(percent))
    print(f"\r|{bar}| {percent:.2f}%", end="\r")


def load(savefile_path):
    with open(savefile_path, 'rb') as file:
        return pickle.load(file)


def ask_filenames():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilenames()


def ask_filename():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename()


def ask_dir():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askdirectory()
