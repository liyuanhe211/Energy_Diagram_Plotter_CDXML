# -*- coding: utf-8 -*-
__author__ = 'LiYuanhe'

#
import os
from PyQt6 import QtGui, QtCore, QtWidgets, uic
from PyQt6.QtWidgets import QApplication
import platform

# import matplotlib
# matplotlib.use("QtAgg")
# from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as MpFigureCanvas
# from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as MpNavToolBar
# from matplotlib import pyplot
# import matplotlib.patches as patches
# from matplotlib.figure import Figure as MpFigure
# from matplotlib import pylab

import sys
import os
import math
import copy
import shutil
import re
import time
import random
import traceback
import pathlib

Python_Lib_path = str(pathlib.Path(__file__).parent.resolve())
sys.path.append(Python_Lib_path)
from My_Lib_Stock import *


#
# def set_Windows_scaling_factor_env_var():
#
#     # Sometimes, the scaling factor of PyQt is different from the Windows system scaling factor, reason unknown
#     # For example, on a 4K screen sets to 250% scaling on Windows, PyQt reads a default 300% scaling,
#     # causing everything to be too large, this function is to determine the ratio of the real DPI and the PyQt DPI
#
#     import platform
#     if platform.system() == 'Windows':
#         import ctypes
#         try:
#             import win32api
#             MDT_EFFECTIVE_DPI = 0
#             monitor = win32api.EnumDisplayMonitors()[0]
#             dpiX,dpiY = ctypes.c_uint(),ctypes.c_uint()
#             ctypes.windll.shcore.GetDpiForMonitor(monitor[0].handle,MDT_EFFECTIVE_DPI,ctypes.byref(dpiX),ctypes.byref(dpiY))
#             DPI_ratio_for_monitor = (dpiX.value+dpiY.value)/2/96
#         except Exception as e:
#             traceback.print_exc()
#             print(e)
#             DPI_ratio_for_monitor = 0
#
#         DPI_ratio_for_device = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
#         PyQt_scaling_ratio = QApplication.primaryScreen().devicePixelRatio()
#         print(f"Windows 10 High-DPI debug:",end=' ')
#         Windows_DPI_ratio = DPI_ratio_for_monitor if DPI_ratio_for_monitor else DPI_ratio_for_device
#         if DPI_ratio_for_monitor:
#             print("Using monitor DPI.")
#             ratio_of_ratio = DPI_ratio_for_monitor / PyQt_scaling_ratio
#         else:
#             print("Using device DPI.")
#             ratio_of_ratio = DPI_ratio_for_device / PyQt_scaling_ratio
#
#         if ratio_of_ratio>1.05 or ratio_of_ratio<0.95:
#             use_ratio = "{:.2f}".format(ratio_of_ratio)
#             print(f"{DPI_ratio_for_monitor=}, {DPI_ratio_for_device=}, {PyQt_scaling_ratio=}")
#             print(f"Using GUI high-DPI ratio: {use_ratio}")
#             print("----------------------------------------------------------------------------")
#             os.environ["QT_SCALE_FACTOR"] = use_ratio
#         else:
#             print("Ratio of ratio near 1. Not scaling.")
#
#         return Windows_DPI_ratio,PyQt_scaling_ratio
#
#
def get_matplotlib_DPI_setting(Windows_DPI_ratio):
    matplotlib_DPI_setting = 60
    if platform.system() == 'Windows':
        matplotlib_DPI_setting = 60 / Windows_DPI_ratio
    if os.path.isfile("__matplotlib_DPI_Manual_Setting.txt"):
        matplotlib_DPI_manual_setting = open("__matplotlib_DPI_Manual_Setting.txt").read()
        if is_int(matplotlib_DPI_manual_setting):
            matplotlib_DPI_setting = matplotlib_DPI_manual_setting
    else:
        with open("__matplotlib_DPI_Manual_Setting.txt", 'w') as matplotlib_DPI_Manual_Setting_file:
            matplotlib_DPI_Manual_Setting_file.write("")
    matplotlib_DPI_setting = int(matplotlib_DPI_setting)
    print(
        f"\nMatplotlib DPI: {matplotlib_DPI_setting}. \n"
        f"Set an appropriate integer in __matplotlib_DPI_Manual_Setting.txt if the preview size doesn't match the output.\n")

    return matplotlib_DPI_setting


def get_open_directories():
    if not Qt.QApplication.instance():
        Qt.QApplication(sys.argv)

    file_dialog = Qt.QFileDialog()
    file_dialog.setFileMode(Qt.QFileDialog.DirectoryOnly)
    file_dialog.setOption(Qt.QFileDialog.DontUseNativeDialog, True)
    file_view = file_dialog.findChild(Qt.QListView, 'listView')

    # to make it possible to select multiple directories:
    if file_view:
        file_view.setSelectionMode(Qt.QAbstractItemView.MultiSelection)
    f_tree_view = file_dialog.findChild(Qt.QTreeView)
    if f_tree_view:
        f_tree_view.setSelectionMode(Qt.QAbstractItemView.MultiSelection)

    if file_dialog.exec():
        return file_dialog.selectedFiles()

    return []


class Qt_Widget_Common_Functions:
    closing = QtCore.pyqtSignal()

    def center_the_widget(self, activate_window=True):
        frame_geometry = self.frameGeometry()
        screen_center = QtGui.QGuiApplication.primaryScreen().availableGeometry().center()
        frame_geometry.moveCenter(screen_center)
        self.move(frame_geometry.topLeft())
        if activate_window:
            self.window().activateWindow()

    def closeEvent(self, event: QtGui.QCloseEvent):
        # print("Window {} closed".format(self))
        self.closing.emit()
        if hasattr(super(), "closeEvent"):
            return super().closeEvent(event)

    def open_config_file(self):

        config_file = os.path.join(filename_class(sys.argv[0]).path, 'Config.ini')

        config_file_failure = False
        if not os.path.isfile(config_file):
            config_file_failure = True
        else:
            with open(config_file) as self.config_File:
                try:
                    self.config = eval(self.config_File.read())
                except Exception as e:
                    traceback.print_exc()
                    print(e)
                    config_file_failure = True

        if config_file_failure:
            with open(config_file, 'w') as self.config_File:
                self.config_File.write('{}')

        with open(config_file) as self.config_File:
            self.config = eval(self.config_File.read())

    def load_config(self, key, absence_return=""):
        if key in self.config:
            return self.config[key]
        else:
            self.config[key] = absence_return
            self.save_config()
            return absence_return

    def save_config(self):
        config_file = os.path.join(filename_class(sys.argv[0]).path, 'Config.ini')
        with open(config_file, 'w') as self.config_File:
            self.config_File.write(repr(self.config))


class Drag_Drop_TextEdit(QtWidgets.QTextEdit):
    drop_accepted_signal = QtCore.pyqtSignal(list)

    def __init__(self):
        super(self.__class__, self).__init__()
        self.setText(" Drop Area")
        self.setAcceptDrops(True)

        font = Qt.QFont()
        font.setFamily("arial")
        font.setPointSize(13)
        self.setFont(font)

        self.setAlignment(Qt.Qt.AlignCenter)

    def dropEvent(self, event):
        if event.mimeData().urls():
            event.accept()
            self.drop_accepted_signal.emit([x.toLocalFile() for x in event.mimeData().urls()])
            self.reset_dropEvent(event)

    def reset_dropEvent(self, event):
        mimeData = Qt.QMimeData()
        mimeData.setText("")
        dummyEvent = Qt.QDropEvent(event.posF(), event.possibleActions(),
                                   mimeData, event.mouseButtons(), event.keyboardModifiers())

        super(self.__class__, self).dropEvent(dummyEvent)


def default_signal_for_connection(signal):
    if isinstance(signal, QtWidgets.QPushButton) or isinstance(signal, QtWidgets.QToolButton) or isinstance(signal, QtWidgets.QRadioButton) or isinstance(
            signal, QtWidgets.QCheckBox):
        signal = signal.clicked
    elif isinstance(signal, QtWidgets.QLineEdit):
        signal = signal.textChanged
    elif isinstance(signal, QtWidgets.QDoubleSpinBox) or isinstance(signal, QtWidgets.QSpinBox):
        signal = signal.valueChanged
    return signal


def disconnect_all(signal, slot):
    signal = default_signal_for_connection(signal)
    marker = False
    while not marker:
        try:
            signal.disconnect(slot)
        except Exception as e:  # TODO: determine what's the specific exception?
            # traceback.print_exc()
            # print(e)
            marker = True


def connect_once(signal, slot):
    signal = default_signal_for_connection(signal)
    disconnect_all(signal, slot)
    signal.connect(slot)


def build_fileDialog_filter(allowed_appendix: list, tags=()):
    """

    :param allowed_appendix: a list of list, each group shows together [[xlsx,log,out],[txt,com,gjf]]
    :param tags: list, tag for each group, default ""
    :return: a compiled filter ready for Qt.getOpenFileNames or other similar functions
            e.g. "Input File (*.gjf *.inp *.com *.sdf *.xyz)\n Output File (*.out *.log *.xlsx *.txt)"
    """

    if not tags:
        tags = [""] * len(allowed_appendix)
    else:
        assert len(tags) == len(allowed_appendix)

    ret = ""
    for count, appendix_group in enumerate(allowed_appendix):
        ret += tags[count].strip()
        ret += "(*."
        ret += ' *.'.join(appendix_group)
        ret += ')'
        if count + 1 != len(allowed_appendix):
            ret += '\n'

    return ret


def alert_UI(message="", title="", parent=None):
    # 旧版本的alert UI定义是alert_UI(parent=None，message="")
    if not isinstance(message, str) and isinstance(title, str) and parent is None:
        parent, message, title = message, title, ""
    elif not isinstance(message, str) and isinstance(title, str) and isinstance(parent, str):
        parent, message, title = message, title, parent
    print(message)
    if not Qt.QApplication.instance():
        Qt.QApplication(sys.argv)
    if not title:
        title = message
    Qt.QMessageBox.critical(parent, title, message)


def warning_UI(message="", parent=None):
    # 旧版本的alert UI定义是alert_UI(parent=None，message="")
    if not isinstance(message, str):
        message, parent = parent, message
    print(message)
    if not Qt.QApplication.instance():
        Qt.QApplication(sys.argv)
    Qt.QMessageBox.warning(parent, message, message)


def information_UI(message="", parent=None):
    # 旧版本的alert UI定义是alert_UI(parent=None，message="")

    if not isinstance(message, str):
        message, parent = parent, message
    print(message)
    if not Qt.QApplication.instance():
        Qt.QApplication(sys.argv)
    Qt.QMessageBox.information(parent, message, message)


def wait_confirmation_UI(parent=None, message=""):
    if not Qt.QApplication.instance():
        Qt.QApplication(sys.argv)
    button = Qt.QMessageBox.warning(parent, message, message, Qt.QMessageBox.Ok | Qt.QMessageBox.Cancel)
    if button == Qt.QMessageBox.Ok:
        return True
    else:
        return False


def get_open_file_UI(parent, start_path: str, allowed_appendix, title="No Title", tags=(), single=False):
    """

    :param parent
    :param start_path:
    :param allowed_appendix: same as function (build_fileDialog_filter)
            but allow single str "txt" or single list ['txt','gjf'] as input, list of list is not necessary
    :param title:
    :param tags:
    :param single:
    :return: a list of files if not single, a single filepath if single
    """

    if not Qt.QApplication.instance():
        Qt.QApplication(sys.argv)

    if isinstance(allowed_appendix, str):  # single str
        allowed_appendix = [[allowed_appendix]]
    if [x for x in allowed_appendix if isinstance(x, str)]:  # single list not list of list
        allowed_appendix = [allowed_appendix]

    filename_filter_string = build_fileDialog_filter(allowed_appendix, tags)

    if single:
        ret = Qt.QFileDialog.getOpenFileName(parent, title, start_path, filename_filter_string)
        if ret:  # 上面返回 ('E:/My_Program/Python_Lib/elements_dict.txt', '(*.txt)')
            ret = ret[0]
    else:
        ret = Qt.QFileDialog.getOpenFileNames(parent, title, start_path, filename_filter_string)
        if ret:  # 上面返回 (['E:/My_Program/Python_Lib/elements_dict.txt'], '(*.txt)')
            ret = ret[0]

    return ret


def show_pixmap(image_filename, graphicsView_object):
    # must call widget.show() holding the graphicsView, otherwise the View.size() will get a wrong (100,30) value
    if os.path.isfile(image_filename):
        pixmap = Qt.QPixmap()
        pixmap.load(image_filename)

        print(graphicsView_object.size())

        if pixmap.width() > graphicsView_object.width() or pixmap.height() > graphicsView_object.height():
            pixmap = pixmap.scaled(graphicsView_object.size(), Qt.Qt.KeepAspectRatio, Qt.Qt.SmoothTransformation)
    else:
        pixmap = Qt.QPixmap()

    graphicsPixmapItem = Qt.QGraphicsPixmapItem(pixmap)
    graphicsScene = Qt.QGraphicsScene()
    graphicsScene.addItem(graphicsPixmapItem)
    graphicsView_object.setScene(graphicsScene)


def update_UI():
    QtCore.QCoreApplication.processEvents()


def exit_UI():
    QtCore.QCoreApplication.instance().quit()


def clear_layout(layout):
    while not layout.isEmpty():
        layout.itemAt(0).widget().deleteLater()
        layout.removeItem(layout.itemAt(0))


def add_list_to_layout(layout, list_of_item):
    for item in list_of_item:
        if isinstance(item, Qt.QWidget):
            layout.addWidget(item)
        if isinstance(item, Qt.QLayout):
            layout.addLayout(item)


def pyqt_ui_compile(filename):
    # 允许将.ui文件放在命名为UI的文件夹下，或程序目录下，但只输入文件名，而不必输入“UI/”

    if filename[:3] in ['UI\\', 'UI/']:
        filename = filename[3:]

    ui_filename = filename_class(filename).replace_append_to('ui')
    # print(os.path.abspath(ui_filename))
    if not os.path.isfile(ui_filename):
        ui_filename = 'UI/' + ui_filename
    modify_log_filename = filename_class(ui_filename).replace_append_to('txt')
    py_file = filename_class(ui_filename).replace_append_to('py')

    modify_time = ""
    if os.path.isfile(modify_log_filename):
        with open(modify_log_filename) as modify_log_file:
            modify_time = modify_log_file.read()

    if modify_time != str(int(os.path.getmtime(ui_filename))):
        print("GUI MODIFIED:", ui_filename)
        with open(modify_log_filename, 'w') as modify_log_file:
            modify_log_file.write(str(int(os.path.getmtime(ui_filename))))

        ui_File_Compile = open(py_file, 'w')
        uic.compileUi(ui_filename, ui_File_Compile)
        ui_File_Compile.close()
        with open(py_file, encoding='gbk') as ui_File_Compile_object:
            ui_File_Compile_content = ui_File_Compile_object.read()
        with open(py_file, 'w', encoding='utf-8') as ui_File_Compile_object:
            ui_File_Compile_object.write(ui_File_Compile_content)


def wait_messageBox(message, title="Please Wait..."):
    if not Qt.QApplication.instance():
        Qt.QApplication(sys.argv)

    message_box = Qt.QMessageBox()
    message_box.setWindowTitle(title)
    message_box.setText(message)

    return message_box
