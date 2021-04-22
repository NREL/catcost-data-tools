# -*- coding: utf-8 -*-
"""
@author: kvanalls
"""

import os
import sys
import logging


# def _append_run_path():
    # if getattr(sys, 'frozen', False):

        # pathlist = []

        # If the application is run as a bundle, the pyInstaller bootloader
        # extends the sys module by a flag frozen=True and sets the app
        # path into variable _MEIPASS'.
        # pathlist.append(sys._MEIPASS)

        # the application exe path
        # _main_app_path = os.path.dirname(sys.executable)
        # pathlist.append(_main_app_path)

        # append to system path enviroment
        # os.environ["PATH"] += os.pathsep + os.pathsep.join(pathlist)

    # logging.error("current PATH: %s", os.environ['PATH'])


if __name__ == '__main__':
    from catcost_data_tools import catcost_data_tools_main

    # _append_run_path()
    catcost_data_tools_main.main()
