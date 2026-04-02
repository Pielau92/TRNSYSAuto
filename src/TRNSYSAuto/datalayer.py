import pandas as pd

from pandas import DataFrame

from trnsys_simulation.datalayer import SimParameters


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
            'mpc':
                convert(self.raw_excel_df, 'mpc', str),
        }

        return DataFrame(excel_data)

    def get_sim_params(self) -> dict[SimParameters]:
        """Get simulation parameters from Excel dataset."""

        def manage_empty_entries(entry):
            if isinstance(entry, dict):  # if
                entry = {key: item for key, item in entry.items()
                         if not str(item) == 'nan'}  # remove "nan" values from dict

            if isinstance(entry, bool):
                return entry

            if len(entry) == 0 or entry == 'nan':  # if empty dict/str, set None
                entry = None

            return entry

        # replace empty cells with None
        data = self.excel_df.map(manage_empty_entries)

        # convert into dict[SimParameters]
        data = data.transpose().to_dict()
        data = {key: SimParameters(**data[key]) for key, item in data.items()}

        return data
