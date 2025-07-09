import configparser

from typing import Type, TypeVar, List
from dataclasses import fields
from config.configs import Configs

T = TypeVar("T")  # type placeholder

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
