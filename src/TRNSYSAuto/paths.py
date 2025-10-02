from os.path import join, expanduser
from dataclasses import dataclass

from configs import Configs


@dataclass
class Paths:
    _configs: Configs
    root: str  # path to root directory
    config: str  # path to configuration ini file
    original_sim_variants_excel: str  # path to original simulation variants Excel file

    results_dir: str = join(expanduser('~'), 'documents', 'TRNSYSAuto')  # path to results output directory

    @property
    def configs(self) -> str:
        """Path to configs.ini file."""
        return join(self.root, 'configs.ini')

    @property
    def sim_series_dir(self) -> str:
        """Path to simulation series directory."""
        return join(self.results_dir, self._configs.runtime.dirname_sim_series)

    @property
    def logfile(self) -> str:
        """Path to logfile."""
        return join(self.sim_series_dir, self._configs.filenames.logger)

    @property
    def savefile(self):
        """Path to savefile where the SimulationSeries object (and the simulation/evaluation progress) is saved."""
        return join(self.sim_series_dir, self._configs.filenames.savefile)

    @property
    def data_dir(self):
        """Path to data directory (contains input directory and results directory)."""
        return join(self.root, 'data')

    @property
    def input_dir(self):
        """Path to input directory (optional storage location for simulation variants Excel files, default initial
        directory when asking to select a simulation variants Excel file)."""
        return join(self.data_dir, 'input')

    @property
    def assets_dir(self):
        """Path to assets directory (contains all files directly needed by TRNSYS)."""
        return join(self.root, 'assets')

    @property
    def sim_variants_excel(self):
        """Path to simulation series Excel file copy."""
        return join(self.sim_series_dir, self._configs.runtime.filename_sim_variants_excel) + '.xlsx'

    @property
    def evaluation_save_dir(self):
        """Path to directory, where evaluation results are saved."""
        return join(self.sim_series_dir, 'evaluation')

    @property
    def cumulative_evaluation_save_file(self):
        """Path to cumulative evaluation file."""
        return join(self.evaluation_save_dir, 'gesamt.xlsx')

    @property
    def cumulative_evaluation_template(self):
        """Path to cumulative evaluation template file."""
        return join(self.assets_dir, 'Auswertung_Gesamt.xlsx')

    @property
    def variant_evaluation_template(self):
        """Path to variant evaluation template file."""
        return join(self.assets_dir, 'Auswertung_Variante.xlsx')
