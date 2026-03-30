import csv
import logging
import multiprocessing
import os
import re
import time
from dataclasses import dataclass

from pywinauto import Application

from trnsys_simulation.utils import parent_dir, delete_files, find_and_replace, replace_parameter_values
from trnsys_simulation.datalayer import B18Data, SimParameters


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

        # find simulation start parameter, if set by user, otherwise start = 0 (start of year)
        start = 0
        for key in list(self.params.dck.keys()):
            if str.lower(key) == 'start':
                start = self.params.dck[key]
                break

        # find simulation stop parameter, if set by user, otherwise stop = 8760 (end of year)
        stop = 8760
        for key in list(self.params.dck.keys()):
            if str.lower(key) == 'stop':
                stop = self.params.dck[key]
                break

        return stop - start

    def setup(self):
        """Setup simulation.

        Overwrites content of simulation files (like templates or configuration files) according to the simulation
        definition.
        """

        self._overwrite_dck_file_parameters()
        self._overwrite_floor_area()
        self._overwrite_mpc_settings_parameters()

    def simulate(self, lock: multiprocessing.Lock = None):
        """Start simulation.

        Starts a TRNSYS simulation and uses a specified dck-file as input. If multiprocessing is used, a lock is passed
        which ensures no other simulation starts until a specific point is reached. In this case, the lock is released
        as soon as the TRNSYS simulation window opens.
        Optionally, start_time_buffer acts as a time buffer before releasing the lock (see self.configs).

        :param multiprocessing.Lock lock:  lock object from the multiprocessing module.
        """

        app = self._start_application()

        if not app:  # if an exception/error occurs, abort (and release any lock so the next simulation may start)

            if lock is not None:
                lock.release()
            return

        # add a time buffer before releasing the lock, which delays the next simulation
        time.sleep(self.configs.buffer_sim_start)
        if lock is not None:
            lock.release()

        window_title = 'TRNSYS: ' + self.path.dck
        window_title = window_title.replace('documents', 'Documents')  # workaround, as search is case-sensitive

        success_message = app.window(title=window_title)  # .window(control_type="Text")
        try:
            success_message.wait('visible', timeout=self.configs.timeout_sim)
        except TimeoutError:
            pass  # goes ahead and closes window after time out

        app.kill()  # close window
        time.sleep(5)

        # delete redundant files
        path_sim = os.path.dirname(self.path.dck)
        redundant_file_paths = [os.path.join(path_sim, file) for file in self.configs.filenames.redundant]
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

    def _start_application(self) -> Application | None:
        """Start simulation application and return Application object.

        Performs all necessary steps to start the simulation application. If an error occurs, log error message and
        return None instead.

        :return: Application object
        """

        # start application
        app = Application(backend='uia')
        app.start(self.path.exe)

        # open .dck file selection window
        try:
            app.connect(title="Öffnen", timeout=self.configs.timeout_open_dck_window)
            app.Öffnen.wait('visible')
            app.Öffnen.set_focus()
        except Exception as e:  # if error occurs, abort
            msg = f'{e} error occurred while opening .dck file selection window for simulation {self.name}.'

            if isinstance(e, TimeoutError):
                msg += f' Timeout is set to {self.configs.timeout_open_dck_window} sec.'

            self.logger.error(msg=msg)  # log error message
            app.kill()  # close window
            return None

        # insert .dck file path
        try:
            app.Öffnen.FileNameEdit.set_edit_text(self.path.dck)
        except Exception as e:
            self.logger.error(f'{e} error occurred while inserting .dck file path for simulation {self.name}.')
            app.kill()  # close window
            return None

        # press start button
        try:
            Button = app.Öffnen.child_window(title="Öffnen", auto_id="1", control_type="Button").wrapper_object()
            Button.click_input()
        except Exception as e:
            self.logger.error(f'{e} error occurred while pressing confirmation button of .dck file selection window for'
                              f' simulation {self.name}.')
            app.kill()  # close window
            return None

        # wait for the simulation window to open
        try:
            app.Öffnen.wait_not('visible', timeout=self.configs.timeout_open_sim_window)
        except Exception as e:  # if error occurs, abort
            msg = f'{e} error occurred while waiting for simulation window to open for simulation {self.name}.'

            if isinstance(e, TimeoutError):
                msg += f' Timeout is set to {self.configs.timeout_open_sim_window} sec.'

            self.logger.error(msg=msg)  # log error message
            app.kill()  # close window
            return None

        return app

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
