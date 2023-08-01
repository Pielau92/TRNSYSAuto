import numpy as np
import pandas as pd
import openpyxl
import glob
import os
import shutil
import win32com.client
# from tkinter import filedialog

from natsort import natsorted
# from line_profiler import LineProfiler
import xlwings as xw


def main(trnsys_folder, filename_sim_variants_excel):
    print('Starting evaluation of simulation results')
    # trnsys_folder = filedialog.askdirectory()
    # trnsys_folder = './23.05.2023_18.22/'
    trnsys_data_file_name = 'out5.txt'
    cumulative_template_file = './Basisordner/Auswertung_Gesamt.xlsx'
    variant_template_file = './Basisordner/Auswertung_Variante.xlsx'
    # output_folder = './out/'
    output_folder = os.path.join(trnsys_folder, 'evaluation')
    raw_data_variant_sheet_name = 'Rohdaten'
    calculation_sheetname = 'Berechn1'
    raw_data_cumulative_sheet_name = 'Rohinputs'
    zone_1_input = 'Zusamm1'
    zone_3_input = 'Zusamm3'
    zone_1_with_output = 'Zone1_mit'
    zone_1_without_output = 'Zone1_ohne'
    zone_3_with_output = 'Zone3_mit'
    zone_3_without_output = 'Zone3_ohne'

    selected_trnsys_columns = ['Period', 'top1', 'top2', 'top3', 'Qventfges', 'qvolgesh', 'qc1', 'qc2', 'qc3', 'pmv1', 'pmv2', 'pmv3']
    trnsys_outdoor_temperature = 'ta'


    # logic starts here - DO NOT CHANGE ANYTHING BELOW UNLESS YOU KNOW WHAT YOU ARE DOING #

    # if not trnsys_folder.endswith('/'):
    #     trnsys_folder = trnsys_folder + '/'
    # if not output_folder.endswith('/'):
    #     output_folder = output_folder + '/'

    for existing_output in glob.glob(output_folder + '/*.xlsx'):
        os.remove(existing_output)
    os.makedirs(output_folder, exist_ok=True)

    cumulative_template_file = os.path.abspath(cumulative_template_file)
    # cumulative_output_file = os.path.abspath(f"{output_folder}gesamt{os.path.splitext(os.path.basename(cumulative_template_file))[1]}")
    cumulative_output_file = os.path.join(output_folder,'gesamt.xlsx')
    shutil.copy(cumulative_template_file, cumulative_output_file)

    variant_parameter_file = os.path.join(trnsys_folder, filename_sim_variants_excel) + '.xlsx'
    variant_parameter_df = pd.read_excel(variant_parameter_file, sheet_name='Simulationsvarianten')


    # get top level directories
    variant_folders = next(os.walk(trnsys_folder))[1]
    variant_folders.remove('evaluation')
    variant_folders = natsorted(variant_folders)
    variant_cnt = 0
    # iterate through trnsys variant folders
    for variant_folder in variant_folders:
        variant_cnt = variant_cnt + 1
        variant_folder_path = os.path.join(trnsys_folder, variant_folder)
        variant_file_path = os.path.join(variant_folder_path, trnsys_data_file_name)
        print(variant_file_path)
        if not os.path.exists(variant_file_path):
            raise ValueError(f'File {variant_file_path} does not exist!')

        if variant_folder not in variant_parameter_df.columns:
            raise ValueError(f'Did not find {variant_folder} in {variant_parameter_file}')


        # read trnsys output file
        trnsys_df = pd.read_csv(variant_file_path, sep='\s+', skiprows=1, skipfooter=23, engine='python')

        selected_trnsys_df = trnsys_df[selected_trnsys_columns]
        selected_trnsys_df = selected_trnsys_df.reindex(selected_trnsys_columns, axis=1)

        #print(selected_trnsys_df)

        # copy template and write data in it
        # variant_output_file = os.path.abspath(f"{output_folder}variant_{variant_folder}{os.path.splitext(os.path.basename(variant_template_file))[1]}")
        variant_output_file = os.path.join(output_folder, 'variant' + variant_folder + '.xlsx')

        shutil.copy(variant_template_file, variant_output_file)

        wb = xw.Book(variant_output_file)
        ws = wb.sheets[raw_data_variant_sheet_name]
        ws["A2"].options(pd.DataFrame, header=1, index=False, expand='table').value = variant_parameter_df[['File', 'Parameter', variant_folder]]
        ws["B60"].options(pd.DataFrame, header=1, index=False, expand='table').value = selected_trnsys_df
        ws = wb.sheets[calculation_sheetname]
        ws["C40"].options(pd.DataFrame, header=False, index=False, expand='table').value = trnsys_df[trnsys_outdoor_temperature].to_frame()
        wb.save()
        wb.app.quit()
        #with pd.ExcelWriter(variant_output_file, mode="a", engine="xlwings", if_sheet_exists='overlay') as writer:
        #    """workBook = writer.book
        #    try:
        #        workBook.remove(workBook[raw_data_variant_sheet_name])
        #    except:
        #        print("Worksheet does not exist")
        #    finally:"""
            #variant_parameter_df[['File', 'Parameter', variant_folder]].to_excel(writer, sheet_name=raw_data_variant_sheet_name, startrow=1, startcol=0, index=False)
            #selected_trnsys_df.to_excel(writer, sheet_name=raw_data_variant_sheet_name, startrow=59, startcol=1, index=False)
            #trnsys_df[trnsys_outdoor_temperature].to_frame().to_excel(writer, sheet_name=calculation_sheetname, startrow=39, startcol=2, index=False, header=False)

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
        zone_1_with_df = pd.read_excel(variant_output_file, sheet_name=zone_1_input, usecols=[3], header=None, nrows=None, skiprows=None)
        zone_1_without_df = pd.read_excel(variant_output_file, sheet_name=zone_1_input, usecols=[2], header=None,  nrows=None, skiprows=None)
        zone_3_with_df = pd.read_excel(variant_output_file, sheet_name=zone_3_input, usecols=[3], header=None,  nrows=None, skiprows=None)
        zone_3_without_df = pd.read_excel(variant_output_file, sheet_name=zone_3_input, usecols=[2], header=None,  nrows=None, skiprows=None)

        with pd.ExcelWriter(cumulative_output_file, mode="a", engine="openpyxl", if_sheet_exists='overlay') as writer:
            if variant_cnt <= 1:
                variant_parameter_df[['File', 'Parameter', variant_folder]].to_excel(writer, sheet_name=raw_data_cumulative_sheet_name, startrow=1, startcol=0, index=False)
            else:
                variant_parameter_df[variant_folder].to_frame().to_excel(writer, sheet_name=raw_data_cumulative_sheet_name, startrow=1, startcol=1+variant_cnt, index=False)

            zone_1_with_df.to_excel(writer, sheet_name=zone_1_with_output, startrow=1, startcol=6+variant_cnt, index=False, header=False)
            zone_1_without_df.to_excel(writer, sheet_name=zone_1_without_output, startrow=1, startcol=6+variant_cnt, index=False, header=False)
            zone_3_with_df.to_excel(writer, sheet_name=zone_3_with_output, startrow=1, startcol=6+variant_cnt, index=False, header=False)
            zone_3_without_df.to_excel(writer, sheet_name=zone_3_without_output, startrow=1, startcol=6+variant_cnt, index=False, header=False)
        #print(calculated_variant_df)


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







    #sheet_name = "Zusamm"
    #usecols = "B:R"
    #to_row = 21
    #skiprows = 2#lambda x: x in [0, 1]
    #header = [0,1]
    #perceived_temperatures = read_multi_header(gesamt_file, index_col=0, sheet_name=sheet_name, usecols=usecols, to_row=to_row, header_top=3, header_bottom=5)
    #fanger = read_multi_header(gesamt_file, index_col=0, sheet_name=sheet_name, usecols=usecols, to_row=48, header_top=35, header_bottom=36)

def read_excel(file, sheet_name=0, usecols=None, nrows=None, skiprows=None):
    df = pd.read_excel(file, sheet_name=sheet_name, usecols=usecols, nrows=nrows, skiprows=skiprows)
    return df

def read_multi_header(file, sheet_name=0, index_col=None, usecols=None, to_row=None, header_top=0, header_bottom=1):

    #if type(header) == list:
    #    largest_header = header.pop(header.index(max(header)))
    #else:
    #    raise ValueError('given header is not a list')

    if to_row is None:
        nrows=None
    else:
        nrows = to_row - header_bottom

    # get data and lowest header row
    df = pd.read_excel(file,
                       sheet_name=sheet_name,
                       index_col=index_col,
                       #header=0,
                       usecols=usecols,
                       nrows=nrows,
                       skiprows=header_bottom-1,
                       parse_dates=False)
    #print(df)

    #
    index = pd.read_excel(file,
                          sheet_name=sheet_name,
                          index_col=index_col,
                          header=None,
                          skiprows=header_top-1,
                          nrows=header_bottom-header_top+1,
                          usecols=usecols,
                          parse_dates=False)
    #print(index)
    index = index.fillna(method='ffill', axis=1)
    df.columns = pd.MultiIndex.from_arrays(index.values)
    #print(df)


def convert_time(data_frame_of_file, conversion, date_format=None):
    if conversion == "unix":
        data_frame_of_file.index = pd.to_datetime(data_frame_of_file.index, unit='s')
    elif conversion == "datetime":
        if format == None:
            data_frame_of_file.index = pd.to_datetime(data_frame_of_file.index, errors='coerce')
        else:
            data_frame_of_file.index = pd.to_datetime(data_frame_of_file.index,format=date_format)

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
        datetime_string_column = data_frame_of_file[data_frame_of_file.columns[0]] + ' ' + data_frame_of_file[data_frame_of_file.columns[1]]

        if format == None:
            date_series = pd.to_datetime(datetime_string_column, errors='coerce')
        else:
            date_series = pd.to_datetime(datetime_string_column,format=date_format)

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


# to acknoledge functions on the bottom of the file, this is necessary
if __name__ == '__main__':
    #lp = LineProfiler()
    #lp_wrapper = lp(main)
    #lp.run('main()')
    #lp.print_stats()
    main()

