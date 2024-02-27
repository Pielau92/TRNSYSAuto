import win32com.client
import numpy as np
import pandas as pd
import xlwings as xw
from datetime import datetime, date
import re

import functions


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


def calcFloatingAverageTemperature(df_input, values_name='Aussentemp', dates_name='date'):
    floating_alpha = 0.8

    if df_input[dates_name].isnull().values.any() or df_input[values_name].isnull().values.any():
        raise ValueError('Values are not allowed to be NaN, interpolate if necessary!')

    average_name = f'{values_name}_mean'
    floating_average_name = f'{values_name}_floating_average'

    df = pd.DataFrame()
    df[dates_name] = df_input[dates_name].copy()
    df[values_name] = df_input[values_name].copy()
    df = df.sort_values(dates_name)
    df['ymd'] = pd.to_datetime(df[dates_name]).dt.date
    mean_df = df.groupby('ymd').mean(numeric_only=False)
    mean_df = mean_df.rename(columns={values_name: average_name})

    df = df.merge(mean_df, how='left', on='ymd')
    day_counter = 1
    next_datapoint_time_delta = pd.Timedelta(0)
    # temporary list which consists the index of the first row of each new day
    new_day_datapoints = [0]

    for i in range(len(df)):
        # print(f'row: {i}')

        if i != 0:
            next_datapoint_time_delta = df.loc[i, 'ymd'] - df.loc[new_day_datapoints[-1], 'ymd']
            # print(next_datapoint_time_delta)

        if next_datapoint_time_delta >= pd.Timedelta('2D'):
            # reset time counter when a time gap happens
            new_day_datapoints = [i]
            day_counter = 1
        elif next_datapoint_time_delta >= pd.Timedelta('1D'):
            new_day_datapoints.append(i)
            day_counter = day_counter + 1

        if day_counter > 8:
            df.loc[i, floating_average_name] = (1 - floating_alpha) * df.loc[
                new_day_datapoints[-2], average_name] + floating_alpha * df.loc[
                                                   new_day_datapoints[-2], floating_average_name]
        elif day_counter > 7:
            # DIN 1525251 Formula
            df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
                new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name] + 0.5 *
                                                df.loc[new_day_datapoints[-5], average_name] + 0.4 * df.loc[
                                                    new_day_datapoints[-6], average_name] + 0.3 * df.loc[
                                                    new_day_datapoints[-7], average_name] + 0.2 * df.loc[
                                                    new_day_datapoints[-8], average_name]) / 3.8
        elif day_counter > 6:
            df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
                new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name] + 0.5 *
                                                df.loc[new_day_datapoints[-5], average_name] + 0.4 * df.loc[
                                                    new_day_datapoints[-6], average_name] + 0.3 * df.loc[
                                                    new_day_datapoints[-7], average_name]) / 3.6
        elif day_counter > 5:
            df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
                new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name] + 0.5 *
                                                df.loc[new_day_datapoints[-5], average_name] + 0.4 * df.loc[
                                                    new_day_datapoints[-6], average_name]) / 3.3
        elif day_counter > 4:
            df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
                new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name] + 0.5 *
                                                df.loc[new_day_datapoints[-5], average_name]) / 2.9
        elif day_counter > 3:
            df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
                new_day_datapoints[-3], average_name] + 0.6 * df.loc[new_day_datapoints[-4], average_name]) / 2.4
        elif day_counter > 2:
            df.loc[i, floating_average_name] = (df.loc[new_day_datapoints[-2], average_name] + 0.8 * df.loc[
                new_day_datapoints[-3], average_name]) / 1.8
        elif day_counter > 1:
            df.loc[i, floating_average_name] = df.loc[new_day_datapoints[-2], average_name]
        elif day_counter <= 1:
            df.loc[i, floating_average_name] = df.loc[i, average_name]
        else:
            raise ValueError('day_counter case is not considered! Fix it!')

    df_input['Aussentemp_mean'] = df[average_name]
    df_input['Aussentemp_floating_average'] = df[floating_average_name]
    return df_input


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
    """Output routine for variant excel file.

    todo: Mit pandas Methode ins Excel schreiben, wie beim cumulative evaluation file
    """
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

    """
    var_list_result_column = df_input.columns
    header = pd.DataFrame(var_list_result_column[1:] + ['']).transpose()
    header.columns = var_list_result_column[:-1] + ['']
    single_column = pd.concat([
        df_input[var_list_result_column],
        pd.DataFrame(index=['']),
        header],
        axis=0)
    single_column = single_column.drop(single_column.columns[-1], axis=1)
    single_column = single_column.transpose().stack(dropna=False)
    return single_column

# region BIN


# def read_excel(file, sheet_name=0, usecols=None, nrows=None, skiprows=None):
#     df = pd.read_excel(file, sheet_name=sheet_name, usecols=usecols, nrows=nrows, skiprows=skiprows)
#     return df
#
#
# def read_multi_header(file, sheet_name=0, index_col=None, usecols=None, to_row=None, header_top=0, header_bottom=1):
#     # if type(header) == list:
#     #    largest_header = header.pop(header.index(max(header)))
#     # else:
#     #    raise ValueError('given header is not a list')
#
#     if to_row is None:
#         nrows = None
#     else:
#         nrows = to_row - header_bottom
#
#     # get data and lowest header row
#     df = pd.read_excel(file,
#                        sheet_name=sheet_name,
#                        index_col=index_col,
#                        # header=0,
#                        usecols=usecols,
#                        nrows=nrows,
#                        skiprows=header_bottom - 1,
#                        parse_dates=False)
#     # print(df)
#
#     #
#     index = pd.read_excel(file,
#                           sheet_name=sheet_name,
#                           index_col=index_col,
#                           header=None,
#                           skiprows=header_top - 1,
#                           nrows=header_bottom - header_top + 1,
#                           usecols=usecols,
#                           parse_dates=False)
#     # print(index)
#     index = index.fillna(method='ffill', axis=1)
#     df.columns = pd.MultiIndex.from_arrays(index.values)
#     # print(df)
#
#
# def convert_time(data_frame_of_file, conversion, date_format=None):
#     if conversion == "unix":
#         data_frame_of_file.index = pd.to_datetime(data_frame_of_file.index, unit='s')
#     elif conversion == "datetime":
#         if format == None:
#             data_frame_of_file.index = pd.to_datetime(data_frame_of_file.index, errors='coerce')
#         else:
#             data_frame_of_file.index = pd.to_datetime(data_frame_of_file.index, format=date_format)
#
#     elif conversion == 'split':
#         data_frame_of_file = data_frame_of_file.reset_index()
#         date_series = combine64(years=data_frame_of_file[data_frame_of_file.columns[0]],
#                                 months=data_frame_of_file[data_frame_of_file.columns[1]],
#                                 days=data_frame_of_file[data_frame_of_file.columns[2]],
#                                 hours=data_frame_of_file[data_frame_of_file.columns[3]],
#                                 minutes=data_frame_of_file[data_frame_of_file.columns[4]])
#         data_frame_of_file = data_frame_of_file.assign(date=date_series)
#         data_frame_of_file.set_index('date', inplace=True)
#
#     elif conversion == 'two':
#         data_frame_of_file = data_frame_of_file.reset_index(drop=True)
#         datetime_string_column = data_frame_of_file[data_frame_of_file.columns[0]] + ' ' + data_frame_of_file[
#             data_frame_of_file.columns[1]]
#
#         if format == None:
#             date_series = pd.to_datetime(datetime_string_column, errors='coerce')
#         else:
#             date_series = pd.to_datetime(datetime_string_column, format=date_format)
#
#         data_frame_of_file = data_frame_of_file.assign(date=date_series)
#         data_frame_of_file.set_index('date', inplace=True)
#     return data_frame_of_file
#
#
# def combine64(years, months=1, days=1, weeks=None, hours=None, minutes=None,
#               seconds=None, milliseconds=None, microseconds=None, nanoseconds=None):
#     years = np.asarray(years) - 1970
#     months = np.asarray(months) - 1
#     days = np.asarray(days) - 1
#     types = ('<M8[Y]', '<m8[M]', '<m8[D]', '<m8[W]', '<m8[h]',
#              '<m8[m]', '<m8[s]', '<m8[ms]', '<m8[us]', '<m8[ns]')
#     vals = (years, months, days, weeks, hours, minutes, seconds,
#             milliseconds, microseconds, nanoseconds)
#     return sum(np.asarray(v, dtype=t) for t, v in zip(types, vals)
#                if v is not None)
#
#
# def get_filename_without_extension(name):
#     name = name.replace("\\", '/')
#     return ".".join((name.split("/")[len(name.split("/")) - 1]).split(".")[:-1])
#
#
# def datenum(yearColumn, monthColumn, dayColumn, hourColumn, minuteColumn, secondColumn):
#     if isinstance(yearColumn, np.int64):
#         datenum_single = \
#             date.toordinal(datetime(yearColumn, monthColumn, dayColumn, hourColumn, minuteColumn, secondColumn))
#         return datenum_single
#
#     elif isinstance(yearColumn, pd.core.series.Series):
#         datenum_array = []
#         for index in range(len(yearColumn)):
#             datenum_array.append(date.toordinal(datetime(yearColumn[index], monthColumn[index], dayColumn[index],
#                                                          hourColumn[index], minuteColumn[index], secondColumn[index])))
#         return datenum_array
#
#
# def weeknum(time_array=None, year=None, month=None, day=None):
#     if all(x is not None for x in [year, month, day]):
#         return date(year, month, day).isocalendar().week
#     elif time_array is not None:
#         return date(time_array[0], time_array[1], time_array[2]).isocalendar().week
#     else:
#         print('Wrong Input for weeknum function.')
#         return np.nan
#
#
# def isequal(array):
#     """Check if all values in input array are identical"""
#     return all(x == array[0] for x in array)

# endregion
