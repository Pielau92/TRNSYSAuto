import re

import pandas as pd

from dataclasses import dataclass, field
from pandas import DataFrame


@dataclass
class SimParameters:
    """Data container for simulation parameters used to overwrite simulation file templates."""

    dck: dict | None  # parameters to be overwritten inside .dck file
    mpc_settings: dict | None  # parameters to be overwritten inside mpc controller settings file
    b18: str  # b17/b18 file name, inside .dck file
    weather: str  # weather data file name, inside .dck file
    mpc_enabled: bool  # if True, use TRNSYS/Python coupling to access MPC controller from inside TRNSYS simulation


class ExcelData:
    """Data container for data read from simulation variant Excel file."""

    def __init__(self, path_excel: str, sheet_name: str):
        self.path_excel: str = path_excel  # Excel file path
        self.sheet_name: str = sheet_name  # name of Excel sheet with simulation variants information

        self.raw_excel_df: DataFrame = self.import_excel()
        self.excel_df: DataFrame = self.transform_excel_data()
        self.parameters: dict[SimParameters] = self.get_sim_params()

    def import_excel(self) -> DataFrame:
        """Import simulation variants Excel file.

        :return: DataFrame with raw dataset.
        """

        # read simulation variants Excel file
        excel_data = pd.ExcelFile(self.path_excel)

        # convert Excel data into pandas DataFrame
        df = excel_data.parse(self.sheet_name, index_col=0)

        # make sure all column headers are string
        df.columns = [str(parameter) for parameter in df.columns]

        return df

    def transform_excel_data(self) -> DataFrame:
        """Transform raw Excel dataset and return as DataFrame.

        Separates simulation parameters by target file (columns) for each simulation variant (rows).

        :return: DataFrame with transformed dataset.
        """

        # conversion function, from DataFrame to dict with values converted to a specific type
        convert = lambda df, name, dtype: {variant: dtype(df[variant][name]) for variant in df.columns[1:]}

        excel_data = {
            'dck':
                self.raw_excel_df[self.raw_excel_df.index == 'dck'].set_index('Parameter').to_dict(),
            'mpc_settings':
                self.raw_excel_df[self.raw_excel_df.index == 'mpc_settings'].set_index('Parameter').to_dict(),
            'weather':
                convert(self.raw_excel_df, 'Wetterdaten', str),
            'b18':
                convert(self.raw_excel_df, 'b18', str),
            'mpc_enabled':
                convert(self.raw_excel_df, 'mpc_enabled', bool),
        }

        return DataFrame(excel_data)

    def get_sim_params(self) -> dict[SimParameters]:
        """Get simulation parameters from Excel dataset."""

        def manage_empty_entries(entry):
            if isinstance(entry, dict):  # if
                entry = {key: item for key, item in entry.items()
                         if not str(item) == 'nan'}  # remove "nan" values from dict

            if len(entry) == 0 or entry == 'nan':  # if empty dict/str, set None
                entry = None

            return entry

        # replace empty cells with None
        data = self.excel_df.map(manage_empty_entries)

        # convert into dict[SimParameters]
        data = data.transpose().to_dict()
        data = {key: SimParameters(**data[key]) for key, item in data.items()}

        return data


@dataclass
class B18Data:
    path_b18: str
    ref_areas: list[float] = field(default_factory=list)

    def read_ref_areas(self):
        """Read reference area of each zone defined inside the b17/b18 file."""

        # open file
        with open(self.path_b18, 'r') as file:
            lines = file.readlines()

        progress = 0
        # find zones definition row
        for row_n, line in enumerate(lines):
            if line == '*  Z o n e s\n':
                progress += row_n
                break

        # extract zone names from row
        zones = lines[progress + 2].split()[1:]

        for zone in zones:
            # find zone definition start
            for row_n, line in enumerate(lines[progress:]):
                if line == f'*  Z o n e  {zone}  /  A i r n o d e  {zone}\n':
                    progress += row_n
                    break

            # find reference area definition of zone
            for row_n, line in enumerate(lines[progress:]):
                if ' REFAREA= ' in line:
                    progress += row_n
                    break

            # extract reference area value
            match = re.search(r' REFAREA\s*=\s*([\d.]+)', lines[progress])
            self.ref_areas.append(float(match.group(1)))
