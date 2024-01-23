import numpy as np
import pandas as pd
import openpyxl
import glob
import os
import shutil
import win32com.client
# from tkinter import filedialog
import schweiker_model

from natsort import natsorted
# from line_profiler import LineProfiler
import xlwings as xw


def main(trnsys_folder, filename_sim_variants_excel):
    print('Starting evaluation of simulation results')
    output_folder = os.path.join(trnsys_folder, 'evaluation')

    # region NAMES

    trnsys_data_file_name = 'out5.txt'
    cumulative_template_file = os.path.abspath('./Basisordner/Auswertung_Gesamt.xlsx')
    cumulative_output_file = os.path.join(output_folder, 'gesamt.xlsx')
    variant_template_file = './Basisordner/Auswertung_Variante.xlsx'

    raw_data_variant_sheet_name = 'Rohdaten'
    calculation_sheetname = 'Berechn1'
    raw_data_cumulative_sheet_name = 'Rohinputs'
    zone_1_input = 'Zusamm1'
    zone_3_input = 'Zusamm3'
    zone_1_with_output = 'Zone1_Betrieb'
    zone_1_without_output = 'Zone1ges'
    zone_3_with_output = 'Zone3_Betrieb'
    zone_3_without_output = 'Zone3ges'

    # selected_trnsys_columns = ['Period', 'top1', 'top2', 'top3', 'Qventfges', 'qvolgesh', 'qc1', 'qc2', 'qc3', 'pmv1',
    #                            'pmv2', 'pmv3']
    selected_trnsys_columns = ['ta', 'top1', 'top2', 'top3', 'Qventfges', 'qvolgesh', 'qc1', 'qc2', 'qc3', 'qh1', 'qh2',
                               'qh3', 'pmv1', 'pmv2', 'pmv3', 'ppd1', 'ppd2', 'ppd3', 'clo1', 'clo2', 'clo3', 'met1',
                               'met2', 'met3']

    var_list_zone1 = ['Period', 'ta', 'tzone1', 'TMSURF_ZONE1', 'relh1', 'vel1', 'pmv1', 'ppd1', 'clo1', 'met1', 'work1']
    var_list_zone2 = ['Period', 'ta', 'tzone1.1', 'TMSURF_ZONE1.1', 'relh2', 'vel2', 'pmv2', 'ppd2', 'clo2', 'met2', 'work2']
    var_list_zone3 = ['Period', 'ta', 'tzone1.2', 'TMSURF_ZONE1.2', 'relh3', 'vel3', 'pmv3', 'ppd3', 'clo3', 'met3', 'work3']
    trnsys_outdoor_temperature = 'ta'

    variant_parameter_file = os.path.join(trnsys_folder, filename_sim_variants_excel) + '.xlsx'

    # endregion NAMES

    # remove existing files in output folder
    for existing_output in glob.glob(output_folder + '/*.xlsx'):
        os.remove(existing_output)
    os.makedirs(output_folder, exist_ok=True)

    # create copy of evaluation file template
    shutil.copy(cumulative_template_file, cumulative_output_file)

    # read simulation variant parameters
    variant_parameter_df = pd.read_excel(variant_parameter_file, sheet_name='Simulationsvarianten')

    # get top level directory list
    variant_folders = next(os.walk(trnsys_folder))[1]
    variant_folders.remove('evaluation')
    variant_folders = natsorted(variant_folders)
    
    # read TRNSYS output and save data
    variant_cnt = 0
    for variant_folder in variant_folders:
        variant_cnt = variant_cnt + 1
        variant_folder_path = os.path.join(trnsys_folder, variant_folder)
        variant_file_path = os.path.join(variant_folder_path, trnsys_data_file_name)
        variant_output_file = os.path.join(output_folder, 'variant' + variant_folder + '.xlsx')
        print(variant_file_path)
        if not os.path.exists(variant_file_path):
            raise ValueError(f'File {variant_file_path} does not exist!')

        if variant_folder not in variant_parameter_df.columns:
            raise ValueError(f'Did not find {variant_folder} in {variant_parameter_file}')

        # read trnsys output file
        trnsys_df = pd.read_csv(variant_file_path, sep='\s+', skiprows=1, skipfooter=23, engine='python')

        # region todo: ppd1-3 muss im out5.txt vorhanden sein, dann können diese Zeilen entfernt werden
        trnsys_df['ppd1'] = np.nan
        trnsys_df['ppd2'] = np.nan
        trnsys_df['ppd3'] = np.nan
        # endregion

        # todo: ab da etwas chaotisch
        # region DATA
        selected_trnsys_df = trnsys_df[selected_trnsys_columns]
        selected_trnsys_df = selected_trnsys_df.reindex(selected_trnsys_columns, axis=1)

        # Schweiker model for zones 1, 2 & 3
        sm1 = schweiker_model.SchweikerDataFrame()
        sm1._df = trnsys_df[var_list_zone1].reindex(var_list_zone1, axis=1)
        sm2 = schweiker_model.SchweikerDataFrame()
        sm2._df = trnsys_df[var_list_zone2].reindex(var_list_zone2, axis=1)
        sm3 = schweiker_model.SchweikerDataFrame()
        sm3._df = trnsys_df[var_list_zone3].reindex(var_list_zone3, axis=1)

        # adapt column headers
        var_list = ['Period', 'ta', 'tzone', 'TMSURF_ZONE', 'relh', 'vel', 'pmv', 'ppd', 'clo', 'met', 'work']
        sm1.df.columns = var_list
        sm2.df.columns = var_list
        sm3.df.columns = var_list

        # region CREATE DATE COLUMN   #todo: Wird derzeit künstlich erzeugt
        year = 2023  # todo: Jahr derzeit hard coded
        time_increment_profiles = 60
        time = pd.date_range(
            start=str(year) + '-01-01',
            end=str(year + 1) + '-01-01',
            freq=str(time_increment_profiles) + 'min')

        time = time.to_series()

        time_df = pd.DataFrame({
            'Tag': time.dt.day,
            'Monat': time.dt.month,
            'Jahr': time.dt.year,
            'Stunde': time.dt.hour,
            'Minute': time.dt.minute
        })

        time_df = time_df[:-(len(time_df) - len(sm1.df))]  # adapt length

        # endregion

        # insert date columns
        time_df = time_df.reset_index()  # reset index
        sm1._df = pd.concat([time_df, sm1.df], axis=1)
        sm2._df = pd.concat([time_df, sm2.df], axis=1)
        sm3._df = pd.concat([time_df, sm3.df], axis=1)

        # schweiker main
        sm1.schweiker_main()
        sm2.schweiker_main()
        sm3.schweiker_main()

        # combine zones data to one single DataFrame
        # result = pd.concat([sm1.df, sm2.df, sm2.df], axis=1)

        # top1 = (sm1.df.tzone + sm1.df.TMSURF_ZONE) / 2
        # top2 = (sm2.df.tzone + sm2.df.TMSURF_ZONE) / 2
        # top3 = (sm3.df.tzone + sm3.df.TMSURF_ZONE) / 2
        #
        # result = pd.concat([top1, top2, top3], axis=1)
        # result.columns = ['top1', 'top2', 'top3']
        # result['Qventfges'] = ''
        # result['qvolgesh'] = ''
        # result['qc1'] = ''
        # result['qc2'] = ''
        # result['qc3'] = ''

        # region RENAME sm

        # remove redundant columns
        sm1.df.drop(['Tag', 'Monat', 'Jahr', 'Stunde', 'Minute', 'index', 'Period'], axis=1, inplace=True)
        sm2.df.drop(['Tag', 'Monat', 'Jahr', 'Stunde', 'Minute', 'index', 'Period'], axis=1, inplace=True)
        sm3.df.drop(['Tag', 'Monat', 'Jahr', 'Stunde', 'Minute', 'index', 'Period'], axis=1, inplace=True)

        # numerate column names for each zone
        sm1.df.columns = ['schweiker_' + string + '1' for string in sm1.df.columns]
        sm2.df.columns = ['schweiker_' + string + '2' for string in sm2.df.columns]
        sm3.df.columns = ['schweiker_' + string + '3' for string in sm3.df.columns]

        # endregion

        # concatenate output
        result = pd.concat([trnsys_df[['ta', 'top1', 'top2', 'top3', 'tzone1', 'tzone2', 'tzone3', 'Qventfges',
                                       'qvolgesh', 'qc1', 'qc2','qc3', 'qh1', 'qh2', 'qh3', 'pmv1', 'pmv2', 'pmv3',
                                       'ppd1', 'ppd2', 'ppd3', 'clo1', 'clo2', 'clo3', 'met1', 'met2', 'met3']],
                            sm1.df, sm2.df, sm3.df], axis=1)

        # sort columns
        result = result[['ta', 'top1', 'top2', 'top3', 'tzone1', 'tzone2', 'tzone3', 'Qventfges', 'qvolgesh', 'qc1',
                         'qc2', 'qc3', 'qh1', 'qh2', 'qh3', 'pmv1', 'pmv2', 'pmv3', 'ppd1', 'ppd2','ppd3', 'clo1',
                         'clo2', 'clo3', 'met1', 'met2', 'met3', 'schweiker_pmv1', 'schweiker_pmv2', 'schweiker_pmv3',
                         'schweiker_ppd1', 'schweiker_ppd2', 'schweiker_ppd3', 'schweiker_clo1', 'schweiker_clo2',
                         'schweiker_clo3', 'schweiker_met1', 'schweiker_met2', 'schweiker_met3']]

        # trnsys_df = result  # overwrite with schweiker model data

        # endregion

        # region EXCEL EXPORT

        # copy template and write data in it
        shutil.copy(variant_template_file, variant_output_file)

        wb = xw.Book(variant_output_file)
        ws = wb.sheets[raw_data_variant_sheet_name]
        ws["A2"].options(pd.DataFrame, header=1, index=False, expand='table').value = variant_parameter_df[
            ['File', 'Parameter', variant_folder]]
        ws["B60"].options(pd.DataFrame, header=1, index=False, expand='table').value = result
        wb.save()
        wb.app.quit()

        # update excel to receive cross-referenced values and updates calculations
        # https://stackoverflow.com/questions/40893870/refresh-excel-external-data-with-python
        xlapp = win32com.client.DispatchEx("Excel.Application")
        wb = xlapp.Workbooks.Open(variant_output_file)
        wb.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        wb.Save()
        xlapp.Quit()

        # Quit
        xlapp.Quit()

        # copy all zones and all classifications in cumulative file
        # read calculated values from variant and copy in cumulative file
        zone_1_with_df = pd.read_excel(variant_output_file, sheet_name=zone_1_input, usecols=[3], header=None,
                                       nrows=None, skiprows=None)
        zone_1_without_df = pd.read_excel(variant_output_file, sheet_name=zone_1_input, usecols=[2], header=None,
                                          nrows=None, skiprows=None)
        zone_3_with_df = pd.read_excel(variant_output_file, sheet_name=zone_3_input, usecols=[3], header=None,
                                       nrows=None, skiprows=None)
        zone_3_without_df = pd.read_excel(variant_output_file, sheet_name=zone_3_input, usecols=[2], header=None,
                                          nrows=None, skiprows=None)

        # create column with all hourly values
        var_list_result_column\
            = ['top1', 'top2', 'top3', 'Qventfges', 'qvolgesh', 'qc1', 'qc2', 'qc3', 'pmv1', 'pmv2', 'pmv3']
        header = pd.DataFrame(var_list_result_column[1:] + ['']).transpose()
        header.columns = var_list_result_column[:-1] + ['']
        # todo: es fehlt ein Tag, deshalb werden 23+2 Werte derzeit übersprungen
        result_column = pd.concat([
            result[var_list_result_column],
            pd.DataFrame(index=['']*24),
            header],
            axis=0)
        result_column = result_column.drop(result_column.columns[-1], axis=1)
        result_column = result_column.transpose().stack(dropna=False)

        with pd.ExcelWriter(cumulative_output_file, mode="a", engine="openpyxl", if_sheet_exists='overlay') as writer:
            if variant_cnt <= 1:
                variant_parameter_df[['File', 'Parameter', variant_folder]].to_excel(writer,
                                                                                     sheet_name=raw_data_cumulative_sheet_name,
                                                                                     startrow=1, startcol=0,
                                                                                     index=False)
            else:
                variant_parameter_df[variant_folder].to_frame().to_excel(writer,
                                                                         sheet_name=raw_data_cumulative_sheet_name,
                                                                         startrow=1, startcol=1 + variant_cnt,
                                                                         index=False)

            zone_1_with_df.to_excel(writer, sheet_name=zone_1_with_output, startrow=1, startcol=6 + variant_cnt,
                                    index=False, header=False)
            zone_1_without_df.to_excel(writer, sheet_name=zone_1_without_output, startrow=1, startcol=6 + variant_cnt,
                                       index=False, header=False)
            zone_3_with_df.to_excel(writer, sheet_name=zone_3_with_output, startrow=1, startcol=6 + variant_cnt,
                                    index=False, header=False)
            zone_3_without_df.to_excel(writer, sheet_name=zone_3_without_output, startrow=1, startcol=6 + variant_cnt,
                                       index=False, header=False)
            result_column.to_excel(writer, sheet_name=raw_data_cumulative_sheet_name, startrow=60, startcol=1+variant_cnt,
                                index=False, header=False)
        # print(calculated_variant_df)

        # endregion todo: Als eigene Methode isolieren

    # update cumulative excel
    # https://stackoverflow.com/questions/40893870/refresh-excel-external-data-with-python
    xlapp = win32com.client.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open(cumulative_output_file)
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    wb.Save()
    xlapp.Quit()

    # Quit
    xlapp.Quit()


def read_excel(file, sheet_name=0, usecols=None, nrows=None, skiprows=None):
    df = pd.read_excel(file, sheet_name=sheet_name, usecols=usecols, nrows=nrows, skiprows=skiprows)
    return df


def read_multi_header(file, sheet_name=0, index_col=None, usecols=None, to_row=None, header_top=0, header_bottom=1):
    # if type(header) == list:
    #    largest_header = header.pop(header.index(max(header)))
    # else:
    #    raise ValueError('given header is not a list')

    if to_row is None:
        nrows = None
    else:
        nrows = to_row - header_bottom

    # get data and lowest header row
    df = pd.read_excel(file,
                       sheet_name=sheet_name,
                       index_col=index_col,
                       # header=0,
                       usecols=usecols,
                       nrows=nrows,
                       skiprows=header_bottom - 1,
                       parse_dates=False)
    # print(df)

    #
    index = pd.read_excel(file,
                          sheet_name=sheet_name,
                          index_col=index_col,
                          header=None,
                          skiprows=header_top - 1,
                          nrows=header_bottom - header_top + 1,
                          usecols=usecols,
                          parse_dates=False)
    # print(index)
    index = index.fillna(method='ffill', axis=1)
    df.columns = pd.MultiIndex.from_arrays(index.values)
    # print(df)


def convert_time(data_frame_of_file, conversion, date_format=None):
    if conversion == "unix":
        data_frame_of_file.index = pd.to_datetime(data_frame_of_file.index, unit='s')
    elif conversion == "datetime":
        if format == None:
            data_frame_of_file.index = pd.to_datetime(data_frame_of_file.index, errors='coerce')
        else:
            data_frame_of_file.index = pd.to_datetime(data_frame_of_file.index, format=date_format)

    elif conversion == 'split':
        data_frame_of_file = data_frame_of_file.reset_index()
        date_series = combine64(years=data_frame_of_file[data_frame_of_file.columns[0]],
                                months=data_frame_of_file[data_frame_of_file.columns[1]],
                                days=data_frame_of_file[data_frame_of_file.columns[2]],
                                hours=data_frame_of_file[data_frame_of_file.columns[3]],
                                minutes=data_frame_of_file[data_frame_of_file.columns[4]])
        data_frame_of_file = data_frame_of_file.assign(date=date_series)
        data_frame_of_file.set_index('date', inplace=True)

    elif conversion == 'two':
        data_frame_of_file = data_frame_of_file.reset_index(drop=True)
        datetime_string_column = data_frame_of_file[data_frame_of_file.columns[0]] + ' ' + data_frame_of_file[
            data_frame_of_file.columns[1]]

        if format == None:
            date_series = pd.to_datetime(datetime_string_column, errors='coerce')
        else:
            date_series = pd.to_datetime(datetime_string_column, format=date_format)

        data_frame_of_file = data_frame_of_file.assign(date=date_series)
        data_frame_of_file.set_index('date', inplace=True)
    return data_frame_of_file


def combine64(years, months=1, days=1, weeks=None, hours=None, minutes=None,
              seconds=None, milliseconds=None, microseconds=None, nanoseconds=None):
    years = np.asarray(years) - 1970
    months = np.asarray(months) - 1
    days = np.asarray(days) - 1
    types = ('<M8[Y]', '<m8[M]', '<m8[D]', '<m8[W]', '<m8[h]',
             '<m8[m]', '<m8[s]', '<m8[ms]', '<m8[us]', '<m8[ns]')
    vals = (years, months, days, weeks, hours, minutes, seconds,
            milliseconds, microseconds, nanoseconds)
    return sum(np.asarray(v, dtype=t) for t, v in zip(types, vals)
               if v is not None)


def get_filename_without_extension(name):
    name = name.replace("\\", '/')
    return ".".join((name.split("/")[len(name.split("/")) - 1]).split(".")[:-1])


def calcFloatingAverageTemperature(df, values_name='Aussentemp', dates_name='date'):
    floating_alpha = 0.8

    if df[dates_name].isnull().values.any() or df[values_name].isnull().values.any():
        raise ValueError('Values are not allowed to be NaN, interpolate if necessary!')

    average_name = f'{values_name}_mean'
    floating_average_name = f'{values_name}_floating_average'

    df = pd.DataFrame()
    df[dates_name] = df[dates_name].copy()
    df[values_name] = df[values_name].copy()
    df = df.sort_values(dates_name)
    df['ymd'] = pd.to_datetime(df[dates_name]).dt.date
    mean_df = df.groupby('ymd').mean()
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

    df['Aussentemp_mean'] = df[average_name]
    df['Aussentemp_floating_average'] = df[floating_average_name]


# to acknoledge functions on the bottom of the file, this is necessary
if __name__ == '__main__':
    # lp = LineProfiler()
    # lp_wrapper = lp(main)
    # lp.run('main()')
    # lp.print_stats()
    main()
