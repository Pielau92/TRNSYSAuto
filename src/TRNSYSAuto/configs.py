from dataclasses import dataclass

from simple_config_manager.configs import _Configs


@dataclass
class General:
    _section_name = 'General'

    path_exe: str  # path to TRNSYS executable file
    multiprocessing_max: int  # maximum number of simulations performed simultaneously
    multiprocessing_autodetect: bool  # if true, override multiprocessing_max with number of cpu cores
    eval_save_interval: int  # the evaluation progress is saved after each save interval
    conda_venv_name: str  # name of the conda virtual environment (venv) to be used


@dataclass
class Filenames:
    _section_name = 'Filenames'

    dck_template: str
    logger: str
    trnsys_output: str
    savefile: str
    mpc_configs: str
    redundant: list[str]
    templates: list[str]
    templates_assets: list[str]


@dataclass
class SheetNames:
    """Excel sheet names"""
    _section_name = 'Excel sheet names'

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
    _section_name = 'Column headers'

    zone1: list[str]
    zone2: list[str]
    zone3: list[str]
    result_column: list[str]
    trnsys_output: list[str]
    sim_variant: list[str]


@dataclass
class Time:
    _section_name = 'Time'

    timeout_sim: int  # if timeout is reached without starting another simulation, stop whole program [s]
    timeout_open_dck_window: int  # if timeout is reached without opening dck selection window, stop defective simulation [s]
    timeout_open_sim_window: int  # if timeout is reached without opening simulation window, stop defective simulation [s]
    buffer_sim_start: int  # time buffer between two simulations, for increased stability [s]


@dataclass
class Runtime:
    """Contains configurations set at runtime."""

    execution_time: str
    filename_sim_variants_excel: str
    dirname_sim_series: str


@dataclass(init=False)  # take __init__ method from _Configs class, keeps section fields from being required arguments
class Configs(_Configs):  # inherit from _Configs parent class

    general: General
    filenames: Filenames
    sheetnames: SheetNames
    col_headers: ColumnHeaders
    time: Time
    runtime: Runtime = None
