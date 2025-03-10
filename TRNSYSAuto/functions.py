import os
import win32com.client
import pandas as pd
# import xlwings as xw
import tkinter as tk
import re
import pickle
import shutil

from tkinter import filedialog  # explicit import required, as calling from tk.filedialog does not work properly


def replace_parameter_values(path_file, parameters):
    """Find and replace parameter values within a .txt file.

    Finds parameter values within a .txt file, replaces them and overwrites the .txt file. For this, the following must
    apply:
    -   A parameter means in this case a name/value pair inside a txt file - specifically: a word followed by "=" and a
        number (there may be any amount of white spaces or tabs between the word and "=" and between "=" and the number)
    -   The replacement values have to be passed in "parameters", where the row indizes must match the name of the
        parameter whose number value is to be replaced

    Example:
        # .txt file containing "Parameter1=1; Parameter2=2; @Parameter3=3."
        path_file = 'C:\\Users\\JohnDoe\\Desktop\\text.txt'

        parameters = pd.Series({'Parameter1': 4, 'Parameter2': 5, 'Parameter3': 6})

        replace_parameter_value(path_file, parameters)

        # .txt file now shows "Parameter1=4; Parameter2=5; Parameter3=6."

    Parameters
    ----------
    path_file : str
        Path of the .txt file.
    parameters : pandas.Series
        Series with parameter values to be replaced, the parameter name has to be in the index name.
    """

    def replacer(match):
        parameter = match.group(1)  # parameter name
        misc = match.group(3)  # miscellaneous characters after the number value (typically comments)

        if parameter in parameters.index:
            return f"{parameter} = {parameters[parameter]} {misc}"  # replace, if parameter name matches
        else:
            return match.group(0)  # return unchanged text

    with open(path_file, 'r') as file:
        text = file.read()

    pattern = re.compile(r'^(\w*)'  # word at beginning of line
                         + r'[\s\t]*=[\s\t]*'  # equal sign (=), with any number of white spaces/tabs before and after
                         + r'(\d+\.?\d*)'  # number, which any number of decimal digits using '.' as delimiter
                         + r'(.*)$',  # any characters, until the end of the line is reached (typically comments)
                         re.MULTILINE)

    new_text = pattern.sub(replacer, text)

    # overwrite file
    with open(path_file, 'w') as file:
        file.write(new_text)


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

    # wb = xw.Book(variant_output_file)
    # ws = wb.sheets[sheet_name_variant_input]
    # ws["A2"].options(pd.DataFrame, header=1, index=False, expand='table').value = variant_parameter_df[
    #     ['File', 'Parameter', variant_folder]]
    # ws["B60"].options(pd.DataFrame, header=1, index=False, expand='table').value = result
    # wb.save()
    # wb.app.quit()
    pass


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
    header.index = ['header']
    single_column = pd.concat([
        df_input[var_list_result_column],
        pd.DataFrame(index=['']),
        header],
        axis=0)
    # single_column = single_column.drop(single_column.columns[-1], axis=1)
    single_column = single_column.transpose().stack(future_stack=True)  # , dropna=False)
    return single_column


def progress_bar(progress, total):
    percent = 100 * (progress / float(total))
    bar = '█' * int(percent) + '-' * (100 - int(percent))
    print(f"\r|{bar}| {percent:.2f}%", end="\r")


def load(savefile_path):
    with open(savefile_path, 'rb') as file:
        return pickle.load(file)


def ask_filenames(initialdir=None):
    root = tk.Tk()
    root.withdraw()

    if initialdir:
        filenames = filedialog.askopenfilenames(initialdir=initialdir)
    else:
        filenames = filedialog.askopenfilenames()

    return [path.replace("/", "\\") for path in filenames]


def ask_filename(initialdir=None):
    root = tk.Tk()
    root.withdraw()

    if initialdir:
        filename = filedialog.askopenfilename(initialdir=initialdir)
    else:
        filename = filedialog.askopenfilename()

    return filename.replace("/", "\\")


def ask_dir():
    """CAUTION: Has compatibility issues with pywinauto (explorer does not open to ask directory location), look in
    main.py for a fix."""
    root = tk.Tk()
    root.withdraw()
    return filedialog.askdirectory()


def create_date_column(year, time_increment_profiles=60):
    date = pd.date_range(
        start=str(year) + '-01-01',
        end=str(year + 1) + '-01-01',
        freq=str(time_increment_profiles) + 'min')

    date = date.to_series()

    date_df = pd.DataFrame({
        'Tag': date.dt.day,
        'Monat': date.dt.month,
        'Jahr': date.dt.year,
        'Stunde': date.dt.hour,
        'Minute': date.dt.minute
    })
    date_df.reset_index(inplace=True)  # reset index

    return date_df


def set_env_and_paths(conda_venv_name):
    """ Set environment variables and paths before launching TRNEXE using Calling Python From TRNSYS.

    Translation and adaptation of Calling Python From TRNSYS batch file "RunTrnsysStudioWithCondaEnvironment.bat", which
    is used to run the TRNSYS Studio and use the Python (CFFI) interface with a miniconda environment.

    Parameters
    ----------
    conda_venv_name : str
        Name of the conda virtual environment (venv) to be used, should have at least cffi and numpy installed.
    """

    # set required environment variables for the conda environment to be found and used by the TRNSYS Python Interface
    # add directory with python to the path (to the front of the path!)
    username = os.getlogin()  # get os username
    os.environ["PATH"] = f"C:\\Users\\{username}\\miniconda3\\condabin;" \
                         + os.environ["PATH"]
    os.environ["PATH"] = f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name};" \
                         + os.environ["PATH"]
    os.environ["PATH"] = f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}\\bin;" \
                         + os.environ["PATH"]
    os.environ["PATH"] = f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}\\Library\\mingw-w64\\bin;" \
                         + os.environ["PATH"]
    os.environ["PATH"] = f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}\\Library\\bin;" \
                         + os.environ["PATH"]
    os.environ["PATH"] = f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}\\Library\\usr\\bin;" \
                         + os.environ["PATH"]
    os.environ["PATH"] = f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}\\Scripts;" \
                         + os.environ["PATH"]

    # Set PYTHONHOME to the same directory
    os.environ["PYTHONHOME"] = f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}"

    # set PYTHONPATH to the site-packages directory (which is within your environment\Lib)
    os.environ["PYTHONPATH"] = f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}\\Lib\\site-packages"


def copy_files(source_path, destination_path):
    """Copy file(s) from source path(s) to destination path(s).

    The number of source paths must match the number of destination paths.

    Parameters
    ----------
    source_path : str | list(str)
        path (or list of paths) of file to be copied.
    destination_path : str | list(str)
        path (or list of paths) where source file is to be copied.

    Returns
    -------
    files_not_found list
        Returns the path of the file that caused the function to raise a FileNotFoundError exception, returns None on
        success.
    """

    files_not_found = []
    if isinstance(source_path, str) & isinstance(destination_path, str):
        source_path = [source_path]
        destination_path = [destination_path]

    for file_index in range(len(source_path)):
        try:
            shutil.copy(source_path[file_index], destination_path[file_index])
        except FileNotFoundError:
            files_not_found.append(source_path[file_index])

    if files_not_found:
        return files_not_found
