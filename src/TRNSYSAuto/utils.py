import sys, os
import shutil
import win32com.client

import pandas as pd
import tkinter as tk

from tkinter import filedialog  # explicit import required, as calling from tk.filedialog does not work properly
from pandas import Series, DataFrame
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

from trnsys_simulation.utils import *


def get_root_dir() -> str:
    """Get root directory path.

    :return: root directory path
    """

    if getattr(sys, 'frozen', False):  # if program is run from an executable .exe file
        return parent_dir(path=sys.executable, levels=2)
    else:  # if program is run from IDE or command window
        return parent_dir(path=__file__, levels=3)


def update_excel_file(path_excel_file: str) -> None:
    """Update calculations in Excel file.

    https://stackoverflow.com/questions/40893870/refresh-excel-external-data-with-python

    :param str path_excel_file: path to Excel file
    """

    xlapp = win32com.client.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open(path_excel_file)
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    wb.Save()
    xlapp.Quit()


def to_single_column(df_input: DataFrame) -> Series:
    """Turn pandas DataFrame into one single column.

    All columns of df_input are stacked over each other with one free row inbetween and the column headers on top.

    Example:
        df = pd.DataFrame({
        'col1': [1, 2, 3],
        'col2': [4, 5, 6],
        'col3': [7, 8, 9]
        })

        result = utils.to_single_column(df)
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


def ask_filenames(initialdir: str = None) -> list[str]:
    """Open explorer window and ask for a single or multiple files, return path(s) to file(s).

    :param str initialdir: path to directory initially shown when opening the explorer window
    :return: list of paths to files
    """
    root = tk.Tk()
    root.withdraw()

    if initialdir:
        filenames = filedialog.askopenfilenames(initialdir=initialdir)
    else:
        filenames = filedialog.askopenfilenames()

    return [path.replace("/", "\\") for path in filenames]


def ask_filename(initialdir: str = None) -> str:
    """Open explorer window and ask for a single file, return path to file.

    :param str initialdir: path to directory initially shown when opening the explorer window
    :return: path to file
    """
    root = tk.Tk()
    root.withdraw()

    if initialdir:
        filename = filedialog.askopenfilename(initialdir=initialdir)
    else:
        filename = filedialog.askopenfilename()

    return filename.replace("/", "\\")


def ask_dir(initialdir: str = None) -> str:
    """Open explorer window and ask for a single directory, return path to directory.

    CAUTION: Has compatibility issues with pywinauto (explorer does not open to ask directory location), look in main.py
    for a fix.

    :param str initialdir: path to directory initially shown when opening the explorer window
    :return: path to directory
    """

    root = tk.Tk()
    root.withdraw()

    return filedialog.askdirectory(initialdir=initialdir)


def create_date_column(year: int, time_increment_profiles: int = 60) -> DataFrame:
    """Create date column as a DataFrame, for a specific year and time increment.

    :param int year: year
    :param int time_increment_profiles: time increment, in minutes
    :return: DataFrame with date column
    """

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


def set_env_and_paths(conda_venv_name: str) -> None:
    """Set required environment variables for the conda environment to be found and used by the TRNSYS Python Interface,
    before launching TRNEXE using Calling Python From TRNSYS.

    Translation and adaptation of Calling Python From TRNSYS batch file "RunTrnsysStudioWithCondaEnvironment.bat", which
    is used to run the TRNSYS Studio and use the Python (CFFI) interface with a miniconda environment. The virtual
    environment should have at least cffi and numpy installed.

    :param str conda_venv_name: name of the conda virtual environment
    """

    username = os.getlogin()  # get operating system username

    # list of paths to be added
    add_paths = [
        f"C:\\Users\\{username}\\miniconda3\\condabin",
        f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}",
        f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}\\bin",
        f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}\\Library\\mingw-w64\\bin",
        f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}\\Library\\bin",
        f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}\\Library\\usr\\bin",
        f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}\\Scripts"
    ]

    # add directory with python to the path (to the front of the path!)
    add_paths.reverse()
    paths = ';'.join(add_paths)
    os.environ["PATH"] = paths + os.environ["PATH"]

    # Set PYTHONHOME to the same directory
    os.environ["PYTHONHOME"] = f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}"

    # set PYTHONPATH to the site-packages directory (which is within your environment\Lib)
    os.environ["PYTHONPATH"] = f"C:\\Users\\{username}\\miniconda3\\envs\\{conda_venv_name}\\Lib\\site-packages"


def copy_files(source_path: str | list[str], destination_path: str | list[str]) -> list[str] | None:
    """Copy file(s) from source path(s) to destination path(s).

    The number of source paths must match the number of destination paths.

    :param str | list[str] source_path: path (or list of paths) of file to be copied
    :param str | list[str] destination_path: path (or list of paths) where source file is to be copied
    :return: list of paths to files that raised a FileNotFoundError exception, returns None on success
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


def logical_or(boolean_lists: list[list[bool]]) -> list[bool]:
    """Performs an element-wise logical or condition on all boolean lists passed inside boolean_lists.

    :param list[list[bool]] boolean_lists: List of boolean lists
    :return: boolean list after element-wise logical or operation
    """

    # check if all lists have the same length
    lengths = [len(boolean_list) for boolean_list in boolean_lists]
    if not lengths[:-1] == lengths[1:]:
        raise ValueError('All boolean lists must have the same length.')

    return [any(values) for values in zip(*boolean_lists)]


def cell_insert_series_to_excel(data: pd.Series, path: str, sheet_name: str, start_cell: str):
    """Insert pandas Series into specific cell, inside a specific Excel sheet.

    :param pd.Series data: data to be inserted into the Excel file
    :param str path: path to excel file
    :param str sheet_name: sheet name inside Excel file
    :param str start_cell: starting cell inside Excel sheet for insertion
    """

    col_letters, start_row = coordinate_from_string(start_cell)
    start_col = column_index_from_string(col_letters)
    with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        data.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=start_row - 1,
            startcol=start_col - 1,
            header=False,
            index=False,
        )
