import TRNSYSAuto.fixes as fixes

fixes.fix_askdirectory_pywinauto_compatibility()

from TRNSYSAuto.utils import get_root_dir
from TRNSYSAuto.gui import gui

fixes.fix_multiprocessing_pyinstaller_compatibility()


def main():
    root_dir = get_root_dir()

    # start GUI
    gui(root_dir)


if __name__ == '__main__':
    main()
