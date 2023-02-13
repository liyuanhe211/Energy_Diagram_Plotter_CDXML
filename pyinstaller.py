# -*- coding: utf-8 -*-
__author__ = 'LiYuanhe'

import os
import shutil

# import pathlib
# parent_path = str(pathlib.Path(__file__).parent.resolve())
# sys.path.insert(0,parent_path)

from Python_Lib.My_Lib_Stock import *

import PyInstaller.__main__

main_py_file = 'Draw_Energy_Diagram_XML.py'
generated_exe_name = "__Energy Diagram Plotter CDXML 3.5.exe"
icon = r"UI\Draw_Energy_Diagram_Icon.ico"
include_all_folder_contents = []
include_folders = ["UI", "Examples"]
include_files = ["__matplotlib_DPI_Manual_Setting.txt"]
delete_files = []

generated_folder_name = filename_class(main_py_file).name_stem
if not os.path.exists(generated_folder_name):    
    os.mkdir(generated_folder_name)

PyInstaller.__main__.run([
    main_py_file,
    '-i', icon, '-n', generated_exe_name, '-y', '-F', '-w', '-p', '.\\Python_Lib', '--clean'
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


for file in include_files:
    print(f"Copying {file} to {generated_folder_name}")
    shutil.copy(file, generated_folder_name)
shutil.copy(os.path.join('dist', generated_exe_name), generated_folder_name)

for folder in include_folders:
    copy_folder(folder, generated_folder_name)

for file in delete_files:
    file = os.path.join(generated_folder_name, file)
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

open_explorer_and_select(os.path.realpath(generated_folder_name))
