import os


class Settings:
    """Class for storing settings of SimulationSeries object."""
    def __init__(self, sim_series):
        self.sim_series = sim_series


class PathSettings:
    """todo: docstring schreiben"""

    def __init__(self, sim_series, path_original_sim_variants_excel):
        self.sim_series = sim_series

        # original simulation variants within base directory
        self.original_sim_variants_excel = path_original_sim_variants_excel

    @property
    def sim_series_dir(self):
        """Path to simulation series directory."""
        return os.path.abspath(self.sim_series.dirname_sim_series)

    @property
    def base_dir(self, base_name='Basisordner'):
        """Path to base directory "Basisordner"."""
        return os.path.abspath(base_name)

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
        return os.path.join(self.base_dir, filename)

    @property
    def variant_evaluation_template(self, filename='Auswertung_Variante.xlsx'):
        """Path to variant evaluation template file."""
        return os.path.join(self.base_dir, filename)

    @property
    def savefile(self):
        """Path to savefile where the SimulationSeries object (and the simulation/evaluation progress) is saved."""
        return os.path.join(self.sim_series_dir, self.sim_series.filename_savefile)