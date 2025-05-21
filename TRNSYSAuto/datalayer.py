import pandas as pd

from dataclasses import dataclass
from pandas import DataFrame


@dataclass
class ExcelData:
    excel_df: DataFrame = None
    parameters: dict = None

    def import_excel(self, path: str, sheet_name: str) -> None:
        """Import simulation variants Excel file."""

        # read simulation variants Excel file
        excel_data = pd.ExcelFile(path)

        # convert Excel data into pandas DataFrame
        df = excel_data.parse(sheet_name, index_col=0)

        # make sure all column headers are string
        df.columns = [str(parameter) for parameter in df.columns]

        self.excel_df = df

    def get_sim_params(self) -> None:
        """Get simulation parameters from imported Excel data and save as dictionary."""

        # separate data by target file
        excel_dict = {}
        excel_dict['dck'] = self.excel_df[self.excel_df.index == 'dck'].set_index('Parameter').to_dict()
        excel_dict['mpc'] = self.excel_df[self.excel_df.index == 'mpc'].set_index('Parameter').to_dict()
        excel_dict['weather'] = \
            {variant: str(self.excel_df[variant]['Wetterdaten']) for variant in self.excel_df.columns[1:]}
        excel_dict['b18'] = {variant: str(self.excel_df[variant]['b18']) for variant in self.excel_df.columns[1:]}

        # restructure dictionary (variant wise)
        variants = self.excel_df.columns[1:]
        parameters = {}
        for variant in variants:
            parameters[variant] = {}
            for target in excel_dict.keys():
                value = excel_dict[target][variant]
                if len(value) == 0:  # if empty dict/str, set None
                    value = None
                parameters[variant][target] = value

        self.parameters = parameters


@dataclass
class SimParameters:
    dck: dict | None
    mpc: dict | None
    b18: str
    weather: str
