import os
import re
import csv
import logging

from dataclasses import dataclass
from pywinauto import Application

from trnsys_simulation.datalayer import B18Data, SimParameters
from trnsys_simulation.utils import parent_dir, delete_files, find_and_replace, replace_parameter_values


class Simulation:
    @dataclass
    class Paths:
        root: str
        parent: str
        dck: str
        exe: str
        mpc: str

    @dataclass
    class Configs:

        @dataclass
        class Filenames:
            dck_template: str
            trnsys_output: str
            mpc_configs: str
            redundant: list[str]

        path_exe: str
        timeout_sim: int
        timeout_open_dck_window: int
        timeout_open_sim_window: int
        buffer_sim_start: int

        filenames: Filenames

    def __init__(self, path_dir: str, path_exe: str, params: SimParameters, configs: Configs, logger: logging.Logger):

        self.configs = configs

        self.path = Simulation.Paths(
            root=path_dir,
            parent=parent_dir(path_dir),
            dck=os.path.join(path_dir, self.configs.filenames.dck_template),
            exe=path_exe,
            mpc=os.path.join(path_dir, self.configs.filenames.mpc_configs),
        )

        self.name = os.path.basename(self.path.root)  # name of simulation
        self.params = params
        self.logger = logger
        self.b18_data = B18Data(path_b18=os.path.join(self.path.root, params.b18))

        self.success: bool = False  # True, if simulated successfully
        self.ignore: bool = False  # if True, do not simulate

    @property
    def sim_hours(self) -> int:
        """Number of simulated hours."""

        params = {key.lower(): value for key, value in self.params.dck.items()}

        return params.get('stop', 8760) - params.get('start', 0)

    def setup(self):
        """Setup simulation.

        Overwrites content of simulation files (like templates or configuration files) according to the simulation
        definition.
        """

        self._overwrite_dck_file_parameters()
        self._overwrite_floor_area()
        self._overwrite_mpc_settings_parameters()

    def delete_redundant_files(self):
        """Delete redundant files after Simulation, to free storage space."""

        redundant_file_paths = [os.path.join(self.path.root, file) for file in self.configs.filenames.redundant]
        delete_files(redundant_file_paths)

    def check_success(self) -> bool:
        """Check if simulation was calculated successfully, based on the TRNSYS output file(s).

        :return: success flag as boolean
        """
        # path of output file
        path_output = os.path.join(self.path.parent, self.name, self.configs.filenames.trnsys_output)

        try:
            with open(path_output) as f:
                data = list(csv.reader(f, delimiter="\t"))

            # simulation was successful, if hourly data is complete (8760 entries)
            self.success = not len(data) < self.sim_hours + 2
        except FileNotFoundError:  # no file found
            self.success = False

        return self.success

    def start_application(self) -> Application | None:
        """Start simulation application and return Application object.

        Performs all necessary steps to start the simulation application. If an error occurs, log error message and
        return None instead.

        :return: Application object
        """

        def open_dck_window():
            app.connect(title="Öffnen", timeout=self.configs.timeout_open_dck_window)
            app.Öffnen.wait('visible')
            app.Öffnen.set_focus()

        def insert_dck_path():
            app.Öffnen.FileNameEdit.set_edit_text(self.path.dck)

        def press_button():
            Button = app.Öffnen.child_window(title="Öffnen", auto_id="1",
                                             control_type="Button").wrapper_object()
            Button.click_input()

        def wait_for_start():
            app.Öffnen.wait_not('visible', timeout=self.configs.timeout_open_sim_window)

        # start application
        app = Application(backend='uia')
        app.start(self.path.exe)

        try:

            open_dck_window()  # open .dck file selection window
            insert_dck_path()  # insert .dck file path
            press_button()  # press start button
            wait_for_start()  # wait for the simulation window to open

            return app

        except Exception as e:
            app.kill()  # close window
            return None

    def _overwrite_dck_file_parameters(self):
        """Overwrite parameters inside .dck File.

        Overwrites the parameters inside the .dck File, according to the corresponding simulation description in
        the simulation variants Excel file.
        """

        # replace weather data file name inside .dck file
        find_and_replace(
            self.path.dck, pattern=r'ASSIGN\s+"[^\.]*\.tm2"', replacement=r'ASSIGN "' + self.params.weather + '"')

        # replace .b17/.b18 file name inside .dck file
        find_and_replace(
            self.path.dck, pattern=r'ASSIGN\s+"[^\.]*\.b(17|18)"', replacement=r'ASSIGN "' + self.params.b18 + '"')

        # replace parameter values
        if self.params.dck:
            replace_parameter_values(self.path.dck, self.params.dck, mark=True)

        # enable/disable python coupling
        find_and_replace(
            self.path.dck, pattern=r'INCLUDE\s+"[^\.]*\.dck"', replacement=r'INCLUDE "' + self.params.mpc + '"')

    def _overwrite_mpc_settings_parameters(self):
        """Overwrite parameters inside settingsMPC.ini File.

        Overwrites the parameters inside the .dck File, according to the corresponding simulation description in
        the simulation variants Excel file.
        """

        if self.params.mpc_settings:
            replace_parameter_values(self.path.mpc, self.params.mpc_settings)

    def _overwrite_floor_area(self):
        """Read floor areas from b17/18 file and overwrite floor area values inside dck file."""

        def replacer(match):
            return (f"{match.group(1)} "  # parameter name
                    f"= {ref_area}")  # reference area

        # read floor areas
        self.b18_data.read_ref_areas()

        # read content of dck file
        with open(self.path.dck, 'r') as file:
            dck_content = file.read()

        # create new content for dck file by overwriting floor area values
        new_dck_content = dck_content
        for zone, ref_area in enumerate(self.b18_data.ref_areas):
            pattern = re.compile(rf'^(Anutz{zone + 1})'  # "Anutz" followed by zone number (e.g Anutz1, Anutz99, ...)
                                 + r'[\s\t]*=[\s\t]*'  # equal sign with any number of white spaces/tabs before or after
                                 + r'(.*)$',  # any characters until  end of line is reached (typically comments)
                                 re.MULTILINE)

            new_dck_content = pattern.sub(replacer, new_dck_content)

        # overwrite dck file
        with open(self.path.dck, 'w') as file:
            file.write(new_dck_content)
