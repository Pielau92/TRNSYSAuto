import os
import multiprocessing
from configparser import ConfigParser


class Settings:
    """Class for storing settings of SimulationSeries object."""

    def __init__(self, sim_series):
        self.sim_series = sim_series
        self._save_path = sim_series.path.settings
        self._settings = ConfigParser()

    def load_settings(self):
        try:
            self._settings.read(self._save_path)
        except:
            print("Format error in settings file, check settings.ini")
            raise SystemExit()

    def apply_settings(self):
        """Apply imported settings to SimulationSeries object.

        Applies the imported settings from the settings Excel file to the corresponding attributes of the
        SimulationSeries object with the same name.
        """

        def apply_setting():
            """Apply setting value to sim_series.

            Applies the individual settings to the corresponding (name of setting and of class attribute must match).
            Automatically recognizes the type of the setting, based on the type of its corresponding class attribute.
            Raises an error if no corresponding class attribute could be found, or an unsupported type is used (str,
            int, float, bool, list (of strings)).
            """

            if not hasattr(self.sim_series, setting):
                raise AttributeError(f'Unknown setting "{setting}" in settings Excel file found.')

            attr = getattr(self.sim_series, setting)

            if isinstance(attr, str):
                value = self._settings.get(section, setting)
            elif isinstance(attr, int):
                value = self._settings.getint(section, setting)
            elif isinstance(attr, float):
                value = self._settings.getfloat(section, setting)
            elif isinstance(attr, bool):
                value = self._settings.getboolean(section, setting)
            elif isinstance(attr, list):
                items = self._settings.get(section, setting).split(',')  # apply comma (,) delimiter
                value = [item.strip() for item in items]  # remove whitespaces at beginning/end of strings
            else:
                raise TypeError(f'Unknown type "{type(attr)}" for setting "{setting}" in settings.ini file. '
                                f'Supported types are string, integer, float, boolean and list (of strings).')

            setattr(self.sim_series, setting, value)

        for section in self._settings.sections():
            for setting in self._settings.options(section):
                apply_setting()  # save setting value into corresponding class attribute, with the correct datatype

        if self.sim_series.multiprocessing_max == 'auto':
            self.sim_series.multiprocessing_max = multiprocessing.cpu_count()
        else:
            self.sim_series.multiprocessing_max = int(self.sim_series.multiprocessing_max)

    def save_settings(self):
        pass
        # with open(self._save_path, 'w') as f:
        #     self._settings.write(f)

    def reset_settings(self):
        pass


class PathSettings:
    """todo: docstring schreiben + zu PathConfig umbenennen (settings stehen dem User zur Verfügung, paths aber nicht
        deshalb config)"""

    def __init__(self, sim_series, path_original_sim_variants_excel):
        self.sim_series = sim_series

        # original simulation variants within base directory
        self.original_sim_variants_excel = path_original_sim_variants_excel

    @property
    def root(self):
        """Path to root directory."""
        return os.path.dirname(os.getcwd())

    @property
    def data_dir(self, dir_name='data'):
        """Path to data directory (contains input directory and results directory)."""
        return os.path.join(self.root, dir_name)

    @property
    def input_dir(self, dir_name='input'):
        """Path to input directory (optional storage location for simulation variants Excel files, default initial
        directory when asking to select a simulation variants Excel file)."""
        return os.path.join(self.data_dir, dir_name)

    @property
    def results_dir(self, dir_name='results'):
        """Path to results directory (contains all simulation series folders, containing in turn all simulation and
        evaluation results)."""
        return os.path.join(self.data_dir, dir_name)

    @property
    def assets_dir(self, dir_name='assets'):
        """Path to assets directory (contains all files directly needed by TRNSYS)."""
        return os.path.join(self.root, dir_name)

    @property
    def settings(self, filename='settings.ini'):
        """Path to settings.ini file."""
        return os.path.join(self.root, filename)

    @property
    def sim_series_dir(self):
        """Path to simulation series directory."""
        return os.path.join(self.results_dir, self.sim_series.dirname_sim_series)

    @property
    def sim_variants_excel(self):
        """Path to simulation series Excel file, copied from the base directory "Basisordner"."""
        return os.path.join(self.sim_series_dir, self.sim_series.filename_sim_variants_excel) + '.xlsx'

    @property
    def logfile(self):
        """Path to logfile."""
        return os.path.join(self.sim_series_dir, self.sim_series.filename_logger)

    @property
    def evaluation_save_dir(self, dir_name='evaluation'):
        """Path to directory, where evaluation results are saved."""
        return os.path.join(self.sim_series_dir, dir_name)

    @property
    def cumulative_evaluation_save_file(self, filename='gesamt.xlsx'):
        """Path to cumulative evaluation file."""
        return os.path.join(self.evaluation_save_dir, filename)

    @property
    def cumulative_evaluation_template(self, filename='Auswertung_Gesamt.xlsx'):
        """Path to cumulative evaluation template file."""
        return os.path.join(self.assets_dir, filename)

    @property
    def variant_evaluation_template(self, filename='Auswertung_Variante.xlsx'):
        """Path to variant evaluation template file."""
        return os.path.join(self.assets_dir, filename)

    @property
    def savefile(self):
        """Path to savefile where the SimulationSeries object (and the simulation/evaluation progress) is saved."""
        return os.path.join(self.sim_series_dir, self.sim_series.filename_savefile)
