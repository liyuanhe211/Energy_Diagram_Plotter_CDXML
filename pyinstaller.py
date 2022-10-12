# -*- coding: utf-8 -*-
__author__ = 'LiYuanhe'

import sys
import os
import math
import copy
import shutil
import re
import time
import random
import subprocess
from collections import OrderedDict

# import pathlib
# parent_path = str(pathlib.Path(__file__).parent.resolve())
# sys.path.insert(0,parent_path)

from Python_Lib.My_Lib_Stock import *

import PyInstaller.__main__

main_py_file = 'Draw_Energy_Diagram_XML.py'
generated_exe_name = "__Energy Diagram Plotter CDXML 3.5.exe"
icon = r"UI\Draw_Energy_Diagram_Icon.ico"
include_all_folder_contents = []
include_folders = ["UI", "Python_Lib","Examples",r"C:\Anaconda3\Lib\site-packages\setuptools"]
include_files = ["Draw_Energy_Diagram_XML.bat",
                 "__matplotlib_DPI_Manual_Setting.txt"]
delete_files = ["Qt5WebEngineCore.dll",
                "mkl_avx512.1.dll",
                "mkl_avx.1.dll",
                "mkl_mc3.1.dll",
                "mkl_avx2.1.dll",
                "mkl_mc.1.dll",
                "mkl_tbb_thread.1.dll",
                "mkl_sequential.1.dll",
                "mkl_vml_avx.1.dll",
                "mkl_vml_mc.1.dll",
                "mkl_vml_avx2.1.dll",
                "mkl_vml_mc3.1.dll",
                "mkl_vml_mc2.1.dll",
                "mkl_vml_avx512.1.dll",
                "mkl_vml_def.1.dll",
                "mkl_vml_cmpt.1.dll"]


PyInstaller.__main__.run([
    main_py_file,
    "--icon",icon, '-y'
])


def copy_folder(src, dst):
    """

    :param src:
    :param dst: dst will *contain* src folder
    :return:
    """
    target = os.path.realpath(os.path.join(dst, filename_class(src).name))
    if os.path.isdir(target):
        # input('Confirm delete: '+target+" >>>")
        try:
            shutil.rmtree(target)
            print("Deleting:", target)
        except Exception:
            print("Delete Failed:", target)
            return None
    print("Copying:", src, 'to', dst)
    shutil.copytree(src, target)


generated_folder_name = os.path.join('dist',filename_class(main_py_file).name_stem)


for file in include_files:
    print(f"Copying {file} to {generated_folder_name}")
    shutil.copy(file,generated_folder_name)

for folder in include_folders:
    copy_folder(folder, generated_folder_name)

for file in delete_files:
    file = os.path.join(generated_folder_name,file)
    if os.path.isfile(file):
        print(f"Deleting {file}")
        os.remove(file)
    else:
        print(f"File to remove not exist: {file}")

for folder in include_all_folder_contents:
    target = os.path.realpath(os.path.join(generated_folder_name, filename_class(folder).name))
    for current_object in os.listdir(folder):
        current_object = os.path.join(folder, current_object)
        if os.path.isfile(current_object):
            shutil.copy(current_object, generated_folder_name)
        else:
            copy_folder(current_object, generated_folder_name)

shutil.move(os.path.join(generated_folder_name,filename_class(main_py_file).name_stem+'.exe'),
            os.path.join(generated_folder_name,generated_exe_name))

open_explorer_and_select(os.path.realpath(generated_folder_name))