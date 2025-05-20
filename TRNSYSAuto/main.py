# region FIX askdirectory window does not open
"""There are compatibility issues between filedialog.askdirectory() and pywinauto, which cause the askdirectory window
 not to open. To fix this, use the following lines. This must happen before importing pywinauto and tkinter."""
# import sys
# import warnings

# deactivate warnings as workaround for higher stability, but it is not optimal as other warnings are also suppressed
# warnings.simplefilter("ignore", UserWarning)

# sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED
# endregion

import multiprocessing
import TRNSYSAuto.utils as utils

from TRNSYSAuto.gui import gui

"""For some unknown reason (when producing an exe-File), main is also affected by multiprocessing (therefore
directory/file is asked multiple times), freeze_support() prevents this."""
multiprocessing.freeze_support()


def main():
    root_dir = utils.get_root_dir()

    # start GUI
    gui(root_dir)


if __name__ == '__main__':
    main()
