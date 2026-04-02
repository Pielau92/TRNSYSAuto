import re

from dataclasses import dataclass, field


@dataclass
class SimParameters:
    """Data container for simulation parameters used to overwrite simulation file templates."""

    dck: dict | None  # parameters to be overwritten inside .dck file
    mpc_settings: dict | None  # parameters to be overwritten inside mpc controller settings file
    b18: str  # b17/b18 file name, inside .dck file
    weather: str  # weather data file name, inside .dck file
    mpc: str  # auxiliary .dck file name for the mpc controller code


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
