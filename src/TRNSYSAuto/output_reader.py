import os

import pandas as pd

from utils import cell_insert_series_to_excel


def read_out5(path: str) -> pd.DataFrame:
    """Read TRNSYS output file (out5.txt).

    - identify header row -> first row with the value "Period" in the first column
    - identify value rows -> row has a plus sign followed by an integer in the first column (e.g. '+1')
    - convert value type to float
    - remove white spaces from column headers
    
    :param str path: path to out5.txt file
    :return: DataFrame with data
    """

    # read data
    df = pd.read_csv(path, sep='\t')

    # identify header row and set as headers
    idx = df.index[df.iloc[:, 0].eq("Period")][0]  # find header row
    df.columns = df.iloc[idx].astype(str)  # set header row as column headers
    df = df.drop(index=idx).reset_index(drop=True)  # delete header row from data

    # identify value rows and only keep those
    df = df[
        df.iloc[:, 0]  # search in first column
        .astype(str).str.contains(r'\+\d+', na=False)  # search for pattern "+integer" (e.g. "+1")
    ].reset_index(drop=True)  # reset row index
    df = df.iloc[:, 1:]  # delete first column

    # convert value type to float
    df = df.apply(pd.to_numeric, errors='coerce').astype(float)

    # remove white spaces from column headers
    df.columns = [col_name.strip() for col_name in df.columns]

    return df


if __name__ == '__main__':
    path_csv = os.path.join(
        'C:/Users/pierre/Documents/TRNSYSAuto/36h test/1',
        'out5.txt'
    )

    data = read_out5(path=path_csv)

    print(data.head())

    path_xlsx = os.path.join(
        'C:/Users/pierre/Documents/TRNSYSAuto/36h test/1',
        'Bewertung_Klein_ZQ3Demo_test.xlsx'
    )

    cell_insert_series_to_excel(
        data=data['ta'],
        path=path_xlsx,
        sheet_name='Bewertung',
        start_cell='W6',
    )

    ...
