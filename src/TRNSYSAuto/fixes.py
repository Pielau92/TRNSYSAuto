import sys
import warnings


def fix_askdirectory_pywinauto_compatibility() -> None:
    """Fix compatibility issues between filedialog.askdirectory() and pywinauto.

    Due to compatibility issues between filedialog.askdirectory() and pywinauto, the askdirectory window may  not open.
    To fix this, run this function before importing pywinauto or tkinter."""

    # deactivate warnings as workaround for higher stability (not optimal though, as other warnings are also suppressed)
    warnings.simplefilter("ignore", UserWarning)

    sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED


def fix_multiprocessing_pyinstaller_compatibility() -> None:
    """Fix compatibility issues between pyinstaller and multiprocessing.

    When creating an executable file with pyinstaller, the multiprocessing module causes issues such as:
     - issues with GUI (e.g. opening explorer multople times)
     - trying to multiprocess tasks it is not supposed to

    See also: https://pyinstaller.org/en/stable/common-issues-and-pitfalls.html#multi-processing
    """

    if getattr(sys, 'frozen', False):  # if program is run from an executable .exe file
        from multiprocessing import freeze_support
        freeze_support()
