import configparser

from os.path import join, expanduser
from typing import Type, TypeVar, List
from dataclasses import dataclass, fields

T = TypeVar("T")  # type placeholder

@dataclass
class Configs:
    """Dataclass that contains static (imported from .ini file) and runtime (set automatically at runtime)
    configurations."""

    @dataclass
    class General:
        path_exe: str  # path to TRNSYS executable file
        timeout: int  # if timeout is reached without starting another simulation, stop program [s]
        start_time_buffer: int  # time buffer between two simulations, for increased stability [s] (optional)
        multiprocessing_max: int  # maximum number of simulations performed simultaneously
        multiprocessing_autodetect: bool  # if true, override multiprocessing_max with number of cpu cores
        eval_save_interval: int  # the evaluation progress is saved after each save interval
        conda_venv_name: str  # name of the conda virtual environment (venv) to be used

    @dataclass
    class Filenames:
        dck_template: str
        logger: str
        trnsys_output: str
        savefile: str
        redundant: list[str]
        templates: list[str]
        templates_assets: list[str]

    @dataclass
    class SheetNames:
        """Excel sheet names"""

        sim_variants: str
        variant_input: str
        calculation: str
        cumulative_input: str
        zone_1_input: str
        zone_3_input: str
        zone_1_with_operating_time: str
        zone_1_without_operating_time: str
        zone_3_with_operating_time: str
        zone_3_without_operating_time: str

    @dataclass
    class ColumnHeaders:
        zone1: list[str]
        zone2: list[str]
        zone3: list[str]
        result_column: list[str]
        trnsys_output: list[str]
        sim_variant: list[str]

    @dataclass
    class Runtime:
        """Contains configurations set at runtime."""
        execution_time: str
        filename_sim_variants_excel: str

        @property
        def dirname_sim_series(self) -> str:
            return f'{self.filename_sim_variants_excel}_{self.execution_time}'

    # save each configuration section in own object
    general: General
    filenames: Filenames
    sheetnames: SheetNames
    col_headers: ColumnHeaders
    runtime: Runtime = None

    """Mapping between section names used to load configurations from .ini file. Sections that are not mapped here will
       not be loaded from the .ini file and have to be filled elsewhere, which is recommended for runtime
       configurations for example. The mapping is structured as a dictionary, where:
            key:    section name inside Configs dataclass
            value:  section name inside .ini configs file
    """

    load_mapping = {
        'general': 'General',
        'filenames': 'Filenames',
        'sheetnames': 'Excel sheet names',
        'col_headers': 'Column headers',
    }


def load_from_ini(path: str) -> Configs:
    """Load all settings from ini file.

    :param str path: path to settings ini file
    :return: Settings dataclass instance, containing all loaded settings values
    """

    mapping = Configs.load_mapping

    kwargs = {}
    for field in fields(Configs):
        if field.name in mapping.keys():
            kwargs[field.name] = get_ini_section(path, mapping[field.name], field.type)
        else:
            continue  # if section is not mapped, do not load from .ini file

    return Configs(**kwargs)


def get_ini_section(path: str, section: str, cls: Type[T]) -> T:
    """Load configuration section from ini file and return as dataclass instance.

    For each attribute defined in the passed dataclass, a corresponding key value pair has to be present inside the
    specified section of the ini file. The value of each corresponding key value pair is collected and then passed to
    the dataclass constructor, to return an instance of said class.

    :param str path: path to ini file
    :param str section: name of the section to be loaded used in the ini file
    :param Type[T] cls: class (actual class, not an instance of it) whose instance is used to save configurations
    :return: class instance
    """

    # read ini file
    config = configparser.ConfigParser()
    config.optionxform = str  # keep capital letters
    config.read(path)

    # get section, if it exists
    if section not in config:
        raise ValueError(f'Section "{section}" not found in {path}')
    cfg_section = config[section]

    kwargs = {}
    for field in fields(cls):
        name = field.name
        typ = field.type

        if field.name not in cfg_section:
            raise ValueError(f'Missing key "{field.name}" in [{section}]')

        raw = cfg_section[name]

        # type conversion
        if typ == int:
            value = int(raw)
        elif typ == float:
            value = float(raw)
        elif typ == str:
            value = raw
        elif typ in {List[str], list[str]}:
            value = [x.strip() for x in raw.split(',') if x.strip()]
        elif typ == bool:
            if raw.lower() in {'true', '1', 'yes', 'on'}:
                value = True
            elif raw.lower() in {'false', '0', 'no', 'off'}:
                value = False
            else:
                raise ValueError(f"Invalid boolean value: {raw}")
        else:
            raise TypeError(f'Unsupported field type {typ} with value "{raw}" in section "{section}" inside {path}.')

        kwargs[name] = value

    return cls(**kwargs)


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
