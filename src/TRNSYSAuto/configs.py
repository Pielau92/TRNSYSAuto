from dataclasses import dataclass

from simple_config_manager.configs import _Configs


@dataclass
class General:
    _section_name = 'General'

    path_exe: str  # path to TRNSYS executable file
    multiprocessing_max: int  # maximum number of simulations performed simultaneously
    multiprocessing_autodetect: bool  # if true, override multiprocessing_max with number of cpu cores
    conda_venv_name: str  # name of the conda virtual environment (venv) to be used
    sheet_name_sim_variants: str  # name of sheet inside Excel input file containing the simulation definitions


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
    windetc: str


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
    time: Time
    runtime: Runtime = None
