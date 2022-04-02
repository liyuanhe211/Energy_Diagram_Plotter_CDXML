# -*- coding: utf-8 -*-
__author__ = 'LiYuanhe'
#
import os
from PyQt5.Qt import QApplication
# if QApplication.desktop().screenGeometry().width()>2000:
os.environ["QT_SCALE_FACTOR"] = "0.85"
from PyQt5 import Qt
from PyQt5 import uic

QApplication.setAttribute(Qt.Qt.AA_EnableHighDpiScaling, True)

import matplotlib

matplotlib.use("Qt5Agg")
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as MpFigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as MpNavToolBar
import matplotlib.pyplot as MpPyplot
import matplotlib.patches as patches
from .adjustText import *

import sys
import os
import math
import copy
import shutil
import re
import time
import random

from .My_Lib_Stock import *


def get_open_directories():
    import os
    from PyQt5.Qt import QApplication

    os.environ["QT_SCALE_FACTOR"] = "0.9"
    from PyQt5 import Qt
    QApplication.setAttribute(Qt.Qt.AA_EnableHighDpiScaling, True)

    if not Qt.QApplication.instance():
        Application = Qt.QApplication(sys.argv)

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
        paths = file_dialog.selectedFiles()

    return paths



class Qt_Widget_Common_Functions():
    closing = Qt.pyqtSignal()

    def center_the_widget(self, activate_window=True):
        frame_geometry = self.frameGeometry()
        screen_center = Qt.QDesktopWidget().availableGeometry().center()
        frame_geometry.moveCenter(screen_center)
        self.move(frame_geometry.topLeft())
        if activate_window:
            self.window().activateWindow()

    def closeEvent(self, event: Qt.QCloseEvent):
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
            with open(config_file, 'r') as self.config_File:
                try:
                    self.config = eval(self.config_File.read())
                except:
                    config_file_failure = True

        if config_file_failure:
            with open(config_file, 'w') as self.config_File:
                self.config_File.write('{}')

        with open(config_file, 'r') as self.config_File:
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


class Drag_Drop_TextEdit(Qt.QTextEdit):
    drop_accepted_signal = Qt.pyqtSignal(list)

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


def disconnect_all(signal, slot):
    if isinstance(signal, Qt.QPushButton) or isinstance(signal, Qt.QToolButton) or isinstance(signal, Qt.QRadioButton) or isinstance(signal, Qt.QCheckBox):
        signal = signal.clicked
    elif isinstance(signal, Qt.QLineEdit):
        signal = signal.textChanged
    elif isinstance(signal, Qt.QDoubleSpinBox) or isinstance(signal, Qt.QSpinBox):
        signal = signal.valueChanged
    marker = False
    while not marker:
        try:
            signal.disconnect(slot)
        except:
            marker = True

def connect_once(signal, slot):
    if isinstance(signal, Qt.QPushButton) or isinstance(signal, Qt.QToolButton) or isinstance(signal, Qt.QRadioButton) or isinstance(signal, Qt.QCheckBox):
        signal = signal.clicked
    elif isinstance(signal, Qt.QLineEdit):
        signal = signal.textChanged
    elif isinstance(signal, Qt.QDoubleSpinBox) or isinstance(signal, Qt.QSpinBox):
        signal = signal.valueChanged
    disconnect_all(signal, slot)
    signal.connect(slot)



def build_fileDialog_filter(allowed_appendix: list, tags=[]):
    '''

    :param allowed_appendix: a list of list, each group shows together [[xlsx,log,out],[txt,com,gjf]]
    :param note: list, tag for each group, default ""
    :return: a compiled filter ready for Qt.getOpenFileNames or other similar functions
            e.g. "Input File (*.gjf *.inp *.com *.sdf *.xyz)\n Output File (*.out *.log *.xlsx *.txt)"
    '''

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
    if not isinstance(message, str) and isinstance(title, str) and parent == None:
        parent, message, title = message, title, ""
    elif not isinstance(message, str) and isinstance(title, str) and isinstance(parent, str):
        parent, message, title = message, title, parent
    print(message)
    if not Qt.QApplication.instance():
        Application = Qt.QApplication(sys.argv)
    if not title:
        title = message
    Qt.QMessageBox.critical(parent, title, message)



def warning_UI(message="", parent=None):
    # 旧版本的alert UI定义是alert_UI(parent=None，message="")
    if not isinstance(message, str):
        message, parent = parent, message
    print(message)
    if not Qt.QApplication.instance():
        Application = Qt.QApplication(sys.argv)
    Qt.QMessageBox.warning(parent, message, message)


def information_UI(message="", parent=None):
    # 旧版本的alert UI定义是alert_UI(parent=None，message="")

    if not isinstance(message, str):
        message, parent = parent, message
    print(message)
    if not Qt.QApplication.instance():
        Application = Qt.QApplication(sys.argv)
    Qt.QMessageBox.information(parent, message, message)


def wait_confirmation_UI(parent=None, message=""):
    if not Qt.QApplication.instance():
        Application = Qt.QApplication(sys.argv)
    button = Qt.QMessageBox.warning(parent, message, message, Qt.QMessageBox.Ok | Qt.QMessageBox.Cancel)
    if button == Qt.QMessageBox.Ok:
        return True
    else:
        return False


def get_open_file_UI(parent, start_path: str, allowed_appendix, title="No Title", tags=[], single=False):
    '''

    :param start_path:
    :param allowed_appendix: same as function (build_fileDialog_filter)
            but allow single str "txt" or single list ['txt','gjf'] as input, list of list is not necessary
    :param title:
    :param tags:
    :param single:
    :return: a list of files if not single, a single filepath if single
    '''

    if not Qt.QApplication.instance():
        Application = Qt.QApplication(sys.argv)

    if isinstance(allowed_appendix, str):  # single str
        allowed_appendix = [[allowed_appendix]]
    if [x for x in allowed_appendix if isinstance(x, str)]:  # single list not list of list
        allowed_appendix = [allowed_appendix]

    filter = build_fileDialog_filter(allowed_appendix, tags)

    if single:
        ret = Qt.QFileDialog.getOpenFileName(parent, title, start_path, filter)
        if ret:  # 上面返回 ('E:/My_Program/Python_Lib/elements_dict.txt', '(*.txt)')
            ret = ret[0]
    else:
        ret = Qt.QFileDialog.getOpenFileNames(parent, title, start_path, filter)
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
    Qt.QCoreApplication.processEvents()


def exit_UI():
    Qt.QCoreApplication.instance().quit()


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
        with open(py_file, encoding='ANSI') as ui_File_Compile_object:
            ui_File_Compile_content = ui_File_Compile_object.read()
        with open(py_file, 'w', encoding='utf-8') as ui_File_Compile_object:
            ui_File_Compile_object.write(ui_File_Compile_content)


def wait_messageBox(message, title="Please Wait..."):
    if not Qt.QApplication.instance():
        Application = Qt.QApplication(sys.argv)

    message_box = Qt.QMessageBox()
    # message_box.setAttribute(Qt.Qt.WA_DeleteOnClose)
    message_box.setWindowTitle(title)
    message_box.setText(message)

    return message_box


#
# def update_check(online_file_url, current_version, reminded_version=-1):
#     '''
#     A function for any PyQt program, that will check online webpage to remind the user to update
#     The webpage should contain paragraph starting with each version number
#     each line should be started by a page_version control, like <Ver15> <Ver12-15>
#     to remind whether the update should be read by the script (Ver15 means upward compatible, this information can only be read by version 15 decipher or higher)
#     例如：
#         <VER1>20180415
#         <VER1>支持由Office365产生的Excel文件
#         <VER1>在界面中增加哪些选项可直接预览的提示
#     :param online_file_url:
#     :param current_version:
#     :param reminded_version: 哪个版本之后不再提示
#     :return: text of update
#
#     '''
#
#     import collections
#
#     info_version = 1
#     ret = urlopen_inf_retry(online_file_url, prettify=False, retry_limit=1).splitlines()
#     content = []
#     for line in ret:
#         re_ret = re.findall(r'<VER(\d+\-*\d*)>(.+)', line)
#         if re_ret:
#             line_version = re_ret[0][0].split('-')
#             if len(line_version) == 1:
#                 line_version = [int(line_version[0]), float('inf')]
#             else:
#                 line_version = [int(line_version[0]), int(line_version[1])]
#             if line_version[0] <= info_version <= line_version[1]:
#                 content.append(re_ret[0][1])
#
#     content = split_list(content, lambda x: re.findall(r'\[VER(\d+)\]', x), include_separator=True)
#     content_dict = collections.OrderedDict()
#     download_dict = {}
#     for version in content:
#         re_ret = re.findall(r'\[VER(\d+)\](.+)', version[0])
#         if re_ret and is_int(re_ret[0][0]):
#             content_dict[int(re_ret[0][0])] = version[1:]
#             download_dict[int(re_ret[0][0])] = re_ret[0][1]
#
#     content_dict = collections.OrderedDict(sorted(content_dict.items(), reverse=True, key=lambda x: x[0]))
#     ret = ""
#     for key, value in content_dict.items():
#         if key > int(current_version) and key > int(reminded_version):
#             ret += '\n'.join(value) + '\n'
#         elif key <= int(reminded_version):
#             print('Version', reminded_version, 'Skipped.')
#     if ret.strip():
#         return (key, download_dict[key], ret.strip())
#     else:
#         return None
#
#
# def alert_update(online_file_url, current_version, reminded_version_filename):
#     if not os.path.isfile(reminded_version_filename):
#         reminded_version = -1
#     else:
#         with open(reminded_version_filename) as reminded_version_file:
#             reminded_version = int(reminded_version_file.read())
#
#     ret_value = update_check(online_file_url, current_version, reminded_version)
#
#     if ret_value:
#         version, download_url, text = ret_value
#
#         if not Qt.QApplication.instance():
#             Application = Qt.QApplication(sys.argv)
#         ret = Qt.QMessageBox.critical(None, "New Version Available",
#                                       "已有新版本，更新功能:\n---------------------------------\n" + text + '\n---------------------------------\n是否打开下载页面？（若点击No to All 将不再提示此版本）',
#                                       Qt.QMessageBox.Ok | Qt.QMessageBox.Ignore | Qt.QMessageBox.NoToAll)
#         update_UI()
#         if ret == Qt.QMessageBox.Ok:
#             open_tab(download_url)
#         if ret == Qt.QMessageBox.NoToAll:
#             with open(reminded_version_filename, 'w') as reminded_version_file:
#                 reminded_version_file.write(str(version))
#
#
# def alert_update_thread(online_file_url, current_version, reminded_version_filename):
#     import threading
#     if not os.path.isfile(reminded_version_filename):
#         reminded_version = -1
#     else:
#         with open(reminded_version_filename) as reminded_version_file:
#             reminded_version = int(reminded_version_file.read())
#
#     update_thread = threading.Thread(target=alert_update, args=[online_file_url, current_version, reminded_version_filename])
#     update_thread.start()