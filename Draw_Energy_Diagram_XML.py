# -*- coding: utf-8 -*-
__author__ = 'LiYuanhe'

# 自动防止重叠（改成非贪心）
# Num tag 在同一行时Tag的加粗
# 支持MECP画点
# 支持纵轴切掉一部分

from Python_Lib.My_Lib_PyQt import *

import sys
import os
import math
import copy
import shutil
import re
import time
import subprocess
from openpyxl import load_workbook
import random
from datetime import datetime
import ctypes


number_font_size = 17
tag_font_size = 17

matplotlib.rcParams.update({'font.family': 'Times New Roman'})

temp_folder = os.path.join(filename_class(sys.argv[0]).path,'TEMP')
if not os.path.isdir(temp_folder):
    os.mkdir(temp_folder)

# possible_ChemDraw_program_locations =\
# [r"C:\Program Files (x86)\PerkinElmerInformatics\ChemOffice2017\ChemDraw\ChemDraw.exe",
# r"C:\Program Files (x86)\PerkinElmerInformatics\ChemOffice2016\ChemDraw\ChemDraw.exe",
# r"C:\Program Files (x86)\CambridgeSoft\ChemOffice2015\ChemDraw\ChemDraw.exe",
# r"C:\Program Files (x86)\CambridgeSoft\ChemOffice2014\ChemDraw\ChemDraw.exe",
# r"C:\Program Files (x86)\CambridgeSoft\ChemOffice2013\ChemDraw\ChemDraw.exe"]


if __name__ == '__main__':
    pyqt_ui_compile('Draw_Energy_Diagram_UI_XML.py')
    from UI.Draw_Energy_Diagram_UI_XML import Ui_Draw_Energy_Diagram_Form

    APPID = 'LYH.DrawEnergyDiagram.3.4'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APPID)

    Application = Qt.QApplication(sys.argv)
    Application.setWindowIcon(Qt.QIcon('UI/Draw_Energy_Diagram_Icon.png'))


# doc header
default_document = '''<?xml version="1.0" encoding="UTF-8" ?>
<!DOCTYPE CDXML SYSTEM "http://www.cambridgesoft.com/xml/cdxml.dtd" >
<CDXML
 CreationProgram="ChemDraw 17.0.0.206"
 Name="temp.cdxml"
 BoundingBox="226.92 96.70 607.08 337.20"
 WindowPosition="0 0"
 WindowSize="-2147483648 -2147483648"
 WindowIsZoomed="yes"
 LabelFont="3"
 LabelSize="10"
 LabelFace="96"
 CaptionFont="4"
 CaptionSize="10"
 HashSpacing="2.49"
 MarginWidth="1.59"
 LineWidth="0.60"
 BoldWidth="2.01"
 BondLength="14.40"
 BondSpacing="18"
 PrintMargins="99.21 70.87 99.21 70.87"
 color="0"
 bgcolor="1"
>'''

# font table, id is in the ChemDraw program, cannot be changed, ignore the change font function
default_fonts = '''<fonttable>
<font id="3" charset="iso-8859-1" name="Arial"/>
<font id="4" charset="iso-8859-1" name="Times New Roman"/>'''
default_fonts_end = '''</fonttable>'''

fonts_template = '''<font id="[Font_ID]" charset="[Font_CharSet]" name="[Font_Name]"/>'''

# color table, r.g.b with a maximum of one
colors_list = '''<colortable>
<color r="1" g="1" b="1"/>
<color r="0" g="0" b="0"/>
<color r="1" g="0" b="0"/>
<color r="1" g="1" b="0"/>
<color r="0" g="1" b="0"/>
<color r="0" g="1" b="1"/>
<color r="0" g="0" b="1"/>
<color r="1" g="0" b="1"/>'''.splitlines()
default_colors_end = '''</colortable>'''

# 似乎它的index是从colortable从2开始数的
colors_translate = {'b': 8, 'g': 6, 'r': 4, 'c': 7, 'm': 9, 'y': 5, 'k': 3, 'w': 2}


# add a color, return the proper color index. If the color doesn't exist in the color table, add it, then return the index
def add_color(r, g=None, b=None):
    '''
    
    :param colors_list: 
    :param color_translate: 
    :param r: in 0~1 float
    :param g: 
    :param b: 
    :return:  a index, pointing to the newly built color
    '''
    global colors_translate
    global colors_list

    if r.replace('TAG',"") in colors_translate:
        if g!=None and b!=None:
            alert_UI("Add_color function error. Report bug.",'Add_color function error')
        return colors_translate[r]
    if r.startswith('#'):
        # in the format of #DDEEFF
        if g != None and b != None:
            alert_UI("Add_color function error. Report bug.", 'Add_color function error')
        r=r.replace('#','0x')

        #XXXXXX --> (0.XX, 0.XX, 0.XX)
        b=(eval(r)%0x100)/0x100
        g=(int(eval(r)/0x100)%0x100)/0x100
        r=(int(eval(r)/0x10000))/0x100

    if (r,g,b) in colors_translate:
        return colors_translate[(r,g,b)]

    color_template = '''<color r="[R_Value]" g="[G_Value]" b="[B_Value]"/>'''  # in 0~1 float

    insert_line = color_template.replace("[R_Value]", str(r)).replace("[G_Value]", str(g)).replace("[B_Value]", str(b))
    colors_list.append(insert_line)
    colors_translate[(r, g, b)] = len(colors_list)
    return len(colors_list)


default_document_end = '''</CDXML>'''

# doesn't matter, 3*3 page is large enough
page_template = '''<page
 id="[ID]"
 BoundingBox="0 0 [Page_Height_Pixel] [Page_Width_Pixel]"
 HeaderPosition="36"
 FooterPosition="36"
 PrintTrimMarks="yes"
 HeightPages="[Page_Height_Page_Count]"
 WidthPages="[Page_Width_Page_Count]"
>'''

def add_page(page_height_count,page_width_count):
    height = 1930.48/3*page_height_count
    width = 1360.76/3*page_width_count
    return page_template.replace("[Page_Height_Pixel]",str(height))\
.replace("[Page_Width_Pixel]",str(width))\
.replace("[Page_Height_Page_Count]",str(page_height_count))\
.replace('[ID]', str(random.randint(1000000, 2000000))) \
.replace("[Page_Width_Page_Count]",str(page_width_count))


page_template_end = '''</page>'''

# add fragment only to bond connections, do not add it to graphic and texts
fragment_template = '''<fragment
 id="45600"
 BoundingBox="0 0 0 0"
 Z="28"
>'''

# add a group, with integral-yes, then the user cannot move each elements before de-frooze it
group_template = '''<group
 id="453612182"
 BoundingBox="156.80 154.98 531.20 261.66"
 Z="[Z]"
 Integral="yes"
>'''

# the Z determines which element will cover which one, and which bond is broken if overlapped
def add_group(Z):
    group_template = '''<group
 id="453612182"
 BoundingBox="156.80 154.98 531.20 261.66"
 Z="[Z]"
 Integral="yes"
>'''
    return group_template.replace('[Z]',str(Z))

group_template_end = "</group>"

# add an atom, then a bond can be made between atoms
def add_node(id, X, Y, Z):
    node_template = '''<n
 id="[ID]"
 p="[X] [Y]"
 Z="[Z]"
 AS="N"
/>'''
    return node_template.replace('[ID]', str(id)).replace('[X]', str(X)).replace('[Y]', str(Y)).replace('[Z]', str(Z))

# add a bold, horizonal bond between two atoms,
def add_state_Bond(color_index, begin_index, Z, end_index=None):
    state_line_template = '''<b
 id="[ID]"
 Z="[Z]"
 color="[Color_index]"
 B="[Begin_index]"
 E="[End_index]"
 Display="Bold"
 BS="N"
/>'''

    bond_id = str(random.randint(2000000,3000000))

    if end_index == None:
        end_index = begin_index + 1

    return state_line_template.replace('[Color_index]', str(color_index)) \
        .replace('[Begin_index]', str(begin_index)) \
        .replace('[ID]', str(bond_id)) \
        .replace('[End_index]', str(end_index)) \
        .replace('[Z]', str(Z))


# add a dashed bond between two bonded states
def add_dash_link_Bond(color_index, begin_index, Z, end_index=None):
    link_line_dash_template = '''<b
 id="[ID]"
 Z="[Z]"
 color="[Color_index]"
 B="[Begin_index]"
 E="[End_index]"
 Display="Dash"
 BS="N"
/>'''

    if end_index == None:
        end_index = begin_index + 1

    return link_line_dash_template.replace('[Color_index]', str(color_index)) \
        .replace('[Begin_index]', str(begin_index)) \
        .replace('[ID]', str(random.randint(3000000, 4000000))) \
        .replace('[End_index]', str(end_index)) \
        .replace('[Z]', str(Z))

# add a solid bond between two bonded states
def add_solid_link_Bond(color_index, begin_index, Z, end_index=None):
    link_line_single_template = '''<b
 id="[ID]"
 Z="[Z]"
 color="[Color_index]"
 B="[Begin_index]"
 E="[End_index]"
 BS="N"
/>'''
    if end_index == None:
        end_index = begin_index + 1

    return link_line_single_template.replace('[Color_index]', str(color_index)) \
        .replace('[Begin_index]', str(begin_index)) \
        .replace('[End_index]', str(end_index)) \
        .replace('[ID]', str(random.randint(4000000, 5000000))) \
        .replace('[Z]', str(Z))



def add_text(text, X, Y, Z, font_index, size, color_index, face_index,vertical=False,right_align=False):
    '''
    add text, at x, y, Z
    :param text:
    :param X:
    :param Y:
    :param Z:
    :param font_index:
    :param size:
    :param color_index:
    :param face_index:
    :param vertical: for Y axis label
    :param right_align: for Y axis number
    :return:
    '''

    if text == None:
        return ""
    text = str(text)
    if text.strip()=="":
        return ""

    #似乎当text只有一个字母的时候总是向右偏一点
    if len(text)==1:
        X-=3

    faces = {'Bold': 1, "Normal": 0}

    if face_index in faces:
        face_index = faces[face_index]

    text_template = '''<t
 id="[ID]"
 p="[X] [Y]"
 BoundingBox="0 0 110 110"
 Z="[Z]"
 CaptionJustification="Center"
 Justification="Center"
 [vertical]
 LineHeight="auto"
><s font="[Font_index]" size="[Size]" color="[Color_index]" face="[Face_index]">[Text]</s></t>'''

    if vertical:
        text_template = text_template.replace('[vertical]','RotationAngle="17694720"')
    else:
        text_template = text_template.replace('[vertical]\n','')
    if right_align:
        text_template = text_template.replace('"Center"\n','"Right"\n')

    return text_template.replace('[X]', str(X)) \
        .replace('[Y]', str(Y)) \
        .replace('[Z]', str(Z)) \
        .replace('[Font_index]', str(font_index)) \
        .replace('[Size]', str(size)) \
        .replace('[ID]', str(random.randint(5000000, 6000000))) \
        .replace('[Color_index]', str(color_index)) \
        .replace('[Face_index]', str(face_index)) \
        .replace('[Text]', text)

# add an ancher point, also serve as a water mark
def test(start_x, start_y, Z=0):
    # a water mark pointing right up
    x2 = start_x
    x1=x2+1.2
    y2 = start_y
    y1=y2-1.2

    # x2,x1 = x1,x2
    # y2,y1 = y1,y2

    template = '''<graphic
 id="453611"
 BoundingBox="[x1] [y1] [x2] [y2]"
 Z="[Z]"
 GraphicType="Orbital"
 OrbitalType="lobeFilled"
/>'''
    return template.replace('[x1]',str(x1))\
.replace('[x2]',str(x2))\
.replace('[y1]',str(y1))\
.replace('[y2]',str(y2))\
.replace('[Z]',str(Z))


# add graphic line of states (bold), and dashed/non-dashed links
def add_graphic_line(x1,x2,y1,y2,Z,color_index=3,dash=False,bold=False,width=0.6):
    arrow_id = str(random.randint(6000000,7000000))

    # these two elements must exist in order to be able to freeze it using Integral group
    graphic_line_template='''<graphic
 id="'''+str(random.randint(7000000,8000000))+'''"
 SupersededBy="'''+arrow_id+'''"
 BoundingBox="481.66 244.50 306 244.50"
 Z="[Z]"
 [Line_type]
 GraphicType="Line"
/>

<arrow
 id="'''+arrow_id+'''"
 BoundingBox="226 296.43 423 297.56"
 Z="[Z]"
 FillType="None"
 ArrowheadType="Solid"
 color="[Color_Index]"
 Head3D="[x2] [y2] 0"
 LineWidth="[Width]"
 Tail3D="[x1] [y1] 0"
 Center3D="538.50 445.75 0"
 MajorAxisEnd3D="735.50 445.75 0"
 MinorAxisEnd3D="538.50 642.75 0"
 [Line_type]
/>'''

    if bold==False and dash==False:
        graphic_line_template = graphic_line_template.replace(' [Line_type]\n',"")
    elif bold:
        graphic_line_template = graphic_line_template.replace('[Line_type]', 'LineType="Bold"')
    elif dash:
        graphic_line_template = graphic_line_template.replace('[Line_type]', 'LineType="Dashed"')


    return graphic_line_template.replace('[x2]',str(x2))\
.replace('[x1]',str(x1))\
.replace('[y1]',str(y1))\
.replace('[y2]',str(y2))\
.replace('[Z]',str(Z))\
.replace('[Color_Index]',str(color_index))\
.replace('[Width]',str(width))\

fragment_template_end = '''</fragment>'''


class State_Line:
    def __init__(self, state, x_center, span=0.6, style='-', color='k', width=4):
        # print(state)
        # if color==None:
        #     color = 'k'

        self.tag = state[0]
        self.energy = state[1]

        self.x_center = x_center
        self.span = span
        self.width = width
        self.style = style
        self.color = color

        self.x_start = self.x_center - self.span / 2
        self.x_end = self.x_center + self.span / 2

        self.x = (self.x_start, self.x_end)

        self.y = (self.energy, self.energy)

        self.annotate = True  # annotate(energy label) required

    def __hash__(self):
        return hash((self.x, self.y, self.span, self.style, self.width))

    def __eq__(self, other):
        return hash(self) == hash(other)


class Connecting_Line:
    def __init__(self, state_line1: State_Line, state_line2: State_Line, style='-', color='k', width=1):
        if state_line1.x_center > state_line2.x_center:  # swap 1 & 2 if reversed order
            state_line1, state_line2 = state_line2, state_line1

        # if color==None:
        #     color = 'k'

        self.state_line1 = state_line1
        self.state_line2 = state_line2

        # the connecting line will intrude the state line due to line width
        self.x_start = self.state_line1.x_end + 0.02
        self.x_end = self.state_line2.x_start - 0.01

        self.y_start = self.state_line1.energy
        self.y_end = self.state_line2.energy

        self.x = (self.x_start, self.x_end)
        self.y = (self.y_start, self.y_end)

        self.style = style
        self.color = color
        self.width = width

        self.annotate = False  # not annotate required

    def __hash__(self):
        return hash((self.x_start, self.x_end, self.y_start, self.y_end, self.style, self.width))

    def __eq__(self, other):
        return hash(self) == hash(other)


class Collision_Box:
    def __init__(self, xcenter=None, xlength=None, xleft=None, xright=None,
                 ycenter=None, yheight=None, ybottom=None, ytop=None):

        self.xleft = xleft
        self.xright = xright
        self.ybottom = ybottom
        self.ytop = ytop

        if xcenter != None and xlength != None:
            self.xleft = xcenter - xlength / 2
            self.xright = xcenter + xlength / 2

        if ycenter != None and yheight != None:
            self.ybottom = ycenter - yheight / 2
            self.ytop = ycenter + yheight / 2
        if ybottom != None and yheight != None:
            self.ybottom = ybottom
            self.ytop = ybottom + yheight

    def collide(self, other):
        if other.xleft > self.xright:
            return False
        if other.xright < self.xleft:
            return False
        if other.ytop < self.ybottom:
            return False
        if other.ybottom > self.ytop:
            return False

        return True


class MpWidget_Energy_Diagram(Qt.QWidget):
    def __init__(self, parent=None, y=[]):
        super(MpWidget_Energy_Diagram, self).__init__()
        self.setParent(parent)

        self.dpi = 24
        self.fig = MpPyplot.figure(figsize=(2, 2), dpi=self.dpi, )

        self.diagram_subplot = MpPyplot.subplot(1, 1, 1)
        self.fig.subplots_adjust(wspace=0.12, left=0.04, right=0.98)
        self.canvas = MpFigureCanvas(self.fig)
        self.canvas.setParent(self)
        self.diagram_subplot.clear()
        self.diagram_subplot.plot(range(len(y)), y, 'r')
        self.canvas.draw()

        self.mpl_toolbar = MpNavToolBar(self.canvas, self)
        self.mpl_toolbar.setIconSize(Qt.QSize(18, 18))
        self.mpl_toolbar.layout().setSpacing(6)

        self.hLayout = Qt.QHBoxLayout()
        self.hLayout.addWidget(self.mpl_toolbar)

        self.vLayout = Qt.QVBoxLayout()
        self.vLayout.addWidget(self.canvas)
        self.vLayout.addLayout(self.hLayout)
        self.vLayout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(self.vLayout)
        self.diagram_data = []

        self.fluctuation_warned = False

        self.diagram_limit = (-1, 0)

        self.occupied_spaces = []  # list of previous collision boxes

        self.initiate()
        self.printed_states=[]

    def diagram_Update(self, diagram_states, x_span, y_span, state_line_span, color='k', num_with_tag=False):
        '''
        :param diagram_states: the Energy contains Tag and Energy like: [('S', 0), ('TS1', 183), ('IM1', 75), ('P', 24)]
        :param style: set style of line, default None, meaning 'k'
        :return:
        '''

        # if color==None:
        #     color = 'k'

        if 'TAG' in color:
            return None

        if diagram_states not in self.printed_states:
            print(diagram_states)
            self.printed_states.append(diagram_states)

        self.x_span = x_span
        self.y_span = y_span

        self.collision_boxes = []
        self.annotate_objects = []
        self.xs_for_adjust_text = []
        self.ys_for_adjust_text = []
        # verify data
        invalid_data = [x[1] for x in diagram_states if not (isinstance(x[1], float) or isinstance(x[1], int) or x[1] == None)]
        if invalid_data:
            Qt.QMessageBox.critical(self, 'Invalid Data Found.', 'Invalid Data Found:\n' + str(invalid_data) + '\n'
                                                                                                               'Only numbers was allowed in the data section.\n'
                                                                                                               'Remove formula by Ctrl+Alt+V if used.\n'
                                                                                                               'Program Terminating...',
                                    Qt.QMessageBox.Abort)

        # print(self.x_span)
        self.diagram_subplot.set_xlim(*self.x_span)
        self.diagram_subplot.set_ylim(*self.y_span)

        self.canvas.draw()

        # self.diagram_subplot.clear()
        self.diagram_data = copy.deepcopy(diagram_states)
        self.current_states = []

        for count, state in enumerate(diagram_states):
            if state[1] != None:  # 不能用 if not state[1]
                state_line = State_Line(state, count, color=color, span=state_line_span)
                self.draw_line(state_line, num_with_tag=num_with_tag)
                self.current_states.append(state_line)

        self.paths.append(self.current_states)

        if 'tag' not in color.lower():
            for count in range(len(self.current_states) - 1):
                connecting_line = Connecting_Line(self.current_states[count], self.current_states[count + 1], color=color)
                self.draw_line(connecting_line)
                self.connecting_lines.append(connecting_line)

        self.diagram_subplot.tick_params(axis='x', labelbottom='off')
        self.diagram_subplot.tick_params(axis='y')

        if os.path.isfile("Energy_Diagram_Y_Axis_Text.txt"):
            with open('Energy_Diagram_Y_Axis_Text.txt') as text_file:
                y_axis_text = text_file.readline()
        else:
            with open('Energy_Diagram_Y_Axis_Text.txt', 'w') as text_file:
                y_axis_text = "Solvated Free Energy (kJ/mol)"
                text_file.write(y_axis_text)

        MpPyplot.ylabel(y_axis_text,
                        fontsize='xx-large',
                        weight='normal')
        MpPyplot.subplots_adjust(left=0.12, right=0.93, top=0.95)

        self.canvas.draw()

    def initiate(self):
        self.diagram_subplot.clear()

        self.paths = []
        self.connecting_lines = []

        self.diagram_subplot.tick_params(axis='x', labelbottom='off')
        self.diagram_subplot.tick_params(axis='y',
                                         labelsize='xx-large')

        if os.path.isfile("Energy_Diagram_Y_Axis_Text.txt"):
            with open('Energy_Diagram_Y_Axis_Text.txt') as text_file:
                y_axis_text = text_file.readline()
        else:
            with open('Energy_Diagram_Y_Axis_Text.txt', 'w') as text_file:
                y_axis_text = "Solvated Free Energy (kJ/mol)"
                text_file.write(y_axis_text)

        MpPyplot.ylabel(y_axis_text,
                        fontsize='xx-large',
                        weight='bold')
        MpPyplot.xticks(np.arange(5),[])

        MpPyplot.subplots_adjust(left=0.12, right=0.93, top=0.95)

    def max_steps(self):
        return max([len(x) for x in self.paths])

    def all_states(self):
        return sum(self.paths, [])

    def draw_line(self, line_object: State_Line, num_with_tag=False):

        assume_figure_length = 60  # assume the figure span 60 characters
        assume_figure_height = 20  # assume the figure span 20 lines

        renderer = MpPyplot.gca().get_figure().canvas.get_renderer()  # for BBox
        fig = plt.gcf()
        size = fig.get_size_inches() * fig.dpi

        # print(size)

        if line_object in sum(self.paths, []) + self.connecting_lines:
            return None

        if isinstance(line_object, State_Line):
            self.xs_for_adjust_text += [line_object.x[0], line_object.x[1]]
            self.ys_for_adjust_text += [line_object.y[0], line_object.y[1]]



        if 'tag' not in line_object.color.lower():  # independent tag 用color做了标记
            self.diagram_subplot.plot(line_object.x,
                                      line_object.y,
                                      line_object.style,
                                      color=line_object.color,
                                      lw=line_object.width)

        if 'tag' not in line_object.color.lower():  # independent tag 用color做了标记
            if line_object.annotate:
                text = "{:.1f}".format(line_object.energy)
                if num_with_tag:  # number和tag一起显示
                    if hasattr(line_object, 'tag') and line_object.tag:
                        text = str(line_object.tag) + ' ' + text
                energy_annotate = self.diagram_subplot.annotate(text,
                                                                (line_object.x_center, line_object.energy + 3),
                                                                fontsize=number_font_size,
                                                                verticalalignment='bottom',
                                                                horizontalalignment='center',
                                                                color=line_object.color,weight='normal')
                self.annotate_objects.append(energy_annotate)
                # energy_annotate.draggable()

        if not num_with_tag or 'tag' in line_object.color.lower():  # 如果要求num和tag一起显示，这里就用不着了；但是如果是用tag标记的independent tag，仍然从这里显示
            if hasattr(line_object, 'tag') and line_object.tag:
                tag = self.diagram_subplot.annotate(line_object.tag,
                                                    (line_object.x_center, line_object.energy),
                                                    verticalalignment='top',
                                                    horizontalalignment='center',
                                                    xytext=(0, -8), textcoords='offset points',
                                                    fontsize=tag_font_size,
                                                    color=line_object.color.lower().replace("tag", ""),  # independent tag 用color做了标记
                                                    weight='bold')

                # tag.draggable()
                self.annotate_objects.append(tag)


class myWidget(Ui_Draw_Energy_Diagram_Form, Qt.QWidget, Qt_Widget_Common_Functions):
    def __init__(self):
        super(self.__class__, self).__init__()
        self.setupUi(self)

        self.energy_diagram = MpWidget_Energy_Diagram()
        self.energy_diagram.setSizePolicy(Qt.QSizePolicy.Preferred, Qt.QSizePolicy.Preferred)
        # self.drag_drop_textEdit = Drag_Drop_TextEdit()
        # self.data_tableWidget.hide()
        # self.horizontalLayout.setStretch(0, 1)
        self.verticalLayout.insertWidget(1,self.energy_diagram, 1)
        # self.verticalLayout.insertWidget(4, self.drag_drop_textEdit, 2)

        self.open_config_file()

        connect_once(self.load_file_pushButton, self.choose_xlsx)
        connect_once(self.update_file_pushButton, self.update_xlsx)
        # connect_once(self.drag_drop_textEdit.drop_accepted_signal, self.xlsx_dropped)
        connect_once(self.x_lower_limit_spinBox, self.x_limits_change)
        connect_once(self.x_upper_limit_spinBox, self.x_limits_change)
        connect_once(self.Y_lower_limit_spinBox, self.y_limits_change)
        connect_once(self.Y_upper_limit_spinBox, self.y_limits_change)
        connect_once(self.num_with_tag_checkBox, self.update_xlsx)
        connect_once(self.save_cdx_pushButton, self.save_cdxml)
        connect_once(self.Y_label_lineEdit.textChanged,self.store_Y_label_to_file)
        connect_once(self.Y_tick_auto_checkBox.clicked,self.toggle_auto_Y_tick)
        connect_once(self.save_cdx_pushButton_2,self.save_cdx_pushButton.click)

        self.dash_solid_button_group = Qt.QButtonGroup()
        self.dash_solid_button_group.addButton(self.dash_line_radioButton)
        self.dash_solid_button_group.addButton(self.solid_line_radioButton)

        self.line_bond_button_group = Qt.QButtonGroup()
        self.line_bond_button_group.addButton(self.use_lines_radioButton)
        self.line_bond_button_group.addButton(self.use_bonds_radioButton)

        self.state_precision_spinBox.setValue(self.load_config('State Precision Digits', 1))
        self.state_line_span_doubleSpinBox.setValue(self.load_config('State Line Span', 0.6))
        self.Auto_collision_avoidance_checkBox.setChecked(self.load_config('Auto_collision_avoidance', True))
        self.allow_overlap_states_checkBox.setChecked(self.load_config('allow_overlap_states_checkBox', False))
        self.num_with_tag_checkBox.setChecked(self.load_config('Num with Tag', False))
        self.solid_line_radioButton.setChecked(self.load_config('Use Solid Link', True))
        self.dash_line_radioButton.setChecked(not self.load_config('Use Solid Link', True))
        self.use_lines_radioButton.setChecked(self.load_config('Use Lines', True))
        self.use_bonds_radioButton.setChecked(not self.load_config('Use Lines', True))
        self.Y_tick_auto_checkBox.setChecked(self.load_config('Y_tick_auto_checkBox', True))
        self.Y_label_precision_spinBox.setValue(self.load_config('Y_label_precision_spinBox', 0))
        self.Y_label_lineEdit.setText(self.load_config('Y_label_lineEdit', "Solvated Free Energy (kJ/mol)"))

        self.use_temp_file_checkBox.setChecked(self.load_config('use_temp_file_checkBox', True))
        self.with_MECP_checkBox.hide()

        self.X_tick_checkBox.setChecked(self.load_config('X_tick_checkBox',True))
        self.Y_tick_checkBox.setChecked(self.load_config('Y_tick_checkBox', True))

        self.resize(self.load_config('Draw_energy_diagram_window_width',1011),
                     self.load_config('Draw_energy_diagram_window_height',580))



        self.x_limits_changed = False
        self.y_limits_changed = False

        self.toggle_auto_Y_tick()

        self.center_the_widget()
        update_UI()

        # if os.path.isfile('Draw_Energy_Diagram_XML_CheckUpdate.exe'):
        #     subprocess.Popen(['Draw_Energy_Diagram_XML_CheckUpdate.exe'])
        # else:
        #     subprocess.Popen(['python','Draw_Energy_Diagram_XML_CheckUpdate.py'])

    def toggle_auto_Y_tick(self):
        self.Y_tick_doubleSpinBox.setEnabled(not self.Y_tick_auto_checkBox.isChecked())
        self.Y_number_doubleSpinBox.setEnabled(not self.Y_tick_auto_checkBox.isChecked())
        self.Y_label_precision_spinBox.setEnabled(not self.Y_tick_auto_checkBox.isChecked())

    def store_Y_label_to_file(self):
        with open(os.path.join(filename_class(sys.argv[0]).path,'Energy_Diagram_Y_Axis_Text.txt'),'w') as Y_axis_text_file:
            Y_axis_text_file.write(self.Y_label_lineEdit.text())


    def x_limits_change(self):
        if self.x_upper_limit_spinBox.value() > self.x_lower_limit_spinBox.value():
            self.x_limits_changed = True
        else:
            self.x_limits_changed = False
        self.update_xlsx()


    def y_limits_change(self):
        if self.Y_upper_limit_spinBox.value() > self.Y_lower_limit_spinBox.value():
            self.y_limits_changed = True
        else:
            self.y_limits_changed = False
        self.update_xlsx()


    def choose_xlsx(self):
        # self.xlsx_dropped(r"C:\Users\LiYuanhe\Desktop\temp.xlsx")
        # self.save_cdxml()
        xlsx_filename = Qt.QFileDialog.getOpenFileName(self, 'Choose Input XLSX file', self.load_config('Energy Diagram Last Path'),
                                                       filter="xlsx File (*.xlsx)")
        if xlsx_filename:
            xlsx_filename = xlsx_filename[0]
            self.xlsx_dropped(xlsx_filename)

    def update_xlsx(self):
        if self.xlsx_file_lineEdit.text():
            self.energy_diagram.initiate()
            self.xlsx_dropped(self.xlsx_file_lineEdit.text(), called_from_update=True)

    def xlsx_dropped(self, xlsx_filename, called_from_update=False):
        '''
            read xlsx data. The first column can be the format setting, e.g. 'r'
        '''

        # 防止因为改变x, y limit 触发多次limit changed. 在函数末尾重新 connect

        disconnect_all(self.x_lower_limit_spinBox, self.x_limits_change)
        disconnect_all(self.x_upper_limit_spinBox, self.x_limits_change)
        disconnect_all(self.Y_lower_limit_spinBox, self.y_limits_change)
        disconnect_all(self.Y_upper_limit_spinBox, self.y_limits_change)

        if isinstance(xlsx_filename, list):  # drop area 给出的是一个list，可以接受多个文件
            xlsx_filename = xlsx_filename[0]

        if filename_class(xlsx_filename).append != 'xlsx':
            return None

        self.energy_diagram.initiate()

        self.config['Energy Diagram Last Path'] = filename_class(xlsx_filename).path
        self.save_config()

        self.xlsx_file_lineEdit.setText(xlsx_filename)

        self.xlsx_filename = xlsx_filename
        self.load_xlsx(self.xlsx_filename)
        # self.table_update()

        if [route[0][1] for route in self.routes if not (isinstance(route[0][1], float) or isinstance(route[0][1], int) or route[0][1] == None)]:  # 存在格式设定
            self.color_of_lines = [route[0][1] for route in self.routes]
            # save something like [None,'r']
        else:
            self.color_of_lines = []
        # print(self.routes)
        self.x_span = [-1, max([len(route) for route in self.routes])-1]
        if not called_from_update:
            self.x_limits_changed = False
            self.x_lower_limit_spinBox.setValue(self.x_span[0]+2)
            self.x_upper_limit_spinBox.setValue(self.x_span[1])
            # if not self.color_of_lines:
            #     self.x_span[1] += 1
        else:
            if self.x_limits_changed:
                self.x_span = [self.x_lower_limit_spinBox.value()-2, self.x_upper_limit_spinBox.value()]

        self.x_upper_limit_spinBox.setMinimum(self.x_lower_limit_spinBox.value()+1)
        self.x_lower_limit_spinBox.setMaximum(self.x_upper_limit_spinBox.value() - 1)

        get_y_range = [state[1] for state in sum(self.routes,[])]
        get_y_range = [x for x in get_y_range if is_float(x)]
        min_energy = min(get_y_range)
        max_energy = max(get_y_range)
        y_span = max_energy - min_energy
        self.y_span = [min_energy - y_span * 0.3, max_energy + y_span * 0.5]
        if not called_from_update:
            self.y_limits_changed = False
            self.Y_lower_limit_spinBox.setValue(self.y_span[0])
            self.Y_upper_limit_spinBox.setValue(self.y_span[1])
        else:
            if self.y_limits_changed:
                self.y_span = [self.Y_lower_limit_spinBox.value(), self.Y_upper_limit_spinBox.value()]


        self.paths_for_cdx_drawing = []
        self.colors_for_cdx_drawing = []

        for count, route in enumerate(self.routes):
            if self.color_of_lines:  # have format setting
                self.energy_diagram.diagram_Update(route[1:], x_span=self.x_span, y_span=self.y_span, color=self.color_of_lines[count],
                                                   num_with_tag=self.num_with_tag_checkBox.isChecked(),
                                                   state_line_span=self.state_line_span_doubleSpinBox.value())
                self.paths_for_cdx_drawing.append(route[1:])
                self.colors_for_cdx_drawing.append(self.color_of_lines[count])

            else:  # doesn't have format setting
                self.energy_diagram.diagram_Update(route, x_span=self.x_span, y_span=self.y_span, num_with_tag=self.num_with_tag_checkBox.checked(),
                                                   state_line_span=self.state_line_span_doubleSpinBox.value())
                self.paths_for_cdx_drawing.append(route)
                self.colors_for_cdx_drawing.append('k')

        connect_once(self.x_lower_limit_spinBox, self.x_limits_change)
        connect_once(self.x_upper_limit_spinBox, self.x_limits_change)
        connect_once(self.Y_lower_limit_spinBox, self.y_limits_change)
        connect_once(self.Y_upper_limit_spinBox, self.y_limits_change)

        self.center_the_widget()
        self.show()

    def save_cdxml(self):

        # 保存各种设置
        self.config['State Precision Digits'] = self.state_precision_spinBox.value()
        self.config['State Line Span'] = self.state_line_span_doubleSpinBox.value()
        self.config['Auto_collision_avoidance']=self.Auto_collision_avoidance_checkBox.isChecked()
        self.config['allow_overlap_states_checkBox']=self.allow_overlap_states_checkBox.isChecked()
        self.config['Num with Tag'] = self.num_with_tag_checkBox.isChecked()

        self.config['Use Solid Link'] = self.solid_line_radioButton.isChecked()
        self.config['Use Lines'] = self.use_lines_radioButton.isChecked()
        # print(self.config['Use Lines'])

        self.config['X_tick_checkBox'] = self.X_tick_checkBox.isChecked()
        self.config['Y_tick_checkBox'] = self.Y_tick_checkBox.isChecked()
        self.config['Y_tick_auto_checkBox'] = self.Y_tick_auto_checkBox.isChecked()
        self.config['Y_label_precision_spinBox'] = self.Y_label_precision_spinBox.value()
        self.config['Y_label_lineEdit'] = self.Y_label_lineEdit.text()

        self.config['Draw_energy_diagram_window_width']= self.width()
        self.config['Draw_energy_diagram_window_height']=self.height()

        self.config['use_temp_file_checkBox']=self.use_temp_file_checkBox.isChecked()

        self.save_config()

        if self.xlsx_file_lineEdit.text() == "":
            self.choose_xlsx()

        fonts_list = [default_fonts, default_fonts_end]

        #获得Energy diagram 与ChemDraw 文档的坐标对应关系

        x_range = self.energy_diagram.diagram_subplot.get_xlim()  # the x coordinate range
        y_range = self.energy_diagram.diagram_subplot.get_ylim()  # the y coordinate range
        width = self.energy_diagram.canvas.width()  # height of diagram in pixels
        height = self.energy_diagram.canvas.height()  # height of diagram in pixels

        x_offset = 200  # spaces left on the left in cdxml
        y_offset = 100  # spaces left on the top in cdxml
        factor = 0.6  # factor from the pixel in matplotlib to chemdraw

        text_line_height = 10  # how many ChemDraw pixels are one line of text took


        #两个轴的Matplotlib-->ChemDraw转换函数
        x_mapping = lambda x: width / (x_range[1] - x_range[0]) * x * factor+x_offset  # map function coordinate to pixel
        y_mapping = lambda y: height / (y_range[1] - y_range[0]) * (y_range[1]-y) * factor+y_offset  # map function coordinate to pixel


        # Z轴编号从多少开始，无所谓，足够大就行
        z_start = 50


        # 确定Excel第一行的Independent text该相当于放在哪个坐标
        x_length = max([len(x) for x_count,x in enumerate(self.paths_for_cdx_drawing) if 'TAG' not in self.colors_for_cdx_drawing[x_count]])
        maximum_ys = [-float('inf') for x in range(x_length)]
        minimum_ys = [float('inf') for x in range(x_length)]
        for route_count,route_content in enumerate(self.paths_for_cdx_drawing):
            if 'TAG' not in self.colors_for_cdx_drawing[route_count]:
                for column_count, item in enumerate(route_content):
                    if item[1]==None:
                        continue
                    maximum_ys[column_count] = max(maximum_ys[column_count],item[1])
                    minimum_ys[column_count] = min(minimum_ys[column_count], item[1])

        # independent tag wants to be at "two lines" below the state line, to stay below the minimum tag
        tags_y_position = [y_mapping(minimum_ys[i])+text_line_height*2+2 for i in range(x_length)]

        state_line_span = self.state_line_span_doubleSpinBox.value()  # the length of each state line, 1 is full
        num_with_tag = self.num_with_tag_checkBox.isChecked()

        nodes = []  # a list of notes, for counting ID and counting Z
        nodes_id_start = 2000 #足够大就可以了
        state_lines = []  # for Z counting
        link_lines = []  # for Z counting
        texts = []  # for Z counting
        current_texts = [] # for text of current route
        avoidance = [] # remember each of previous collision box (x,y_start,y_end) for Greedy collision avoidance
        font_index = 4 #目前没法换其他的font
        size_of_tag = 10
        line_width_avoidance = 3 # how many ChemDraw pixels are needed to put a text on top a line
        size_of_number = 10
        remember_drawn_states = [] # list of tuples (x_count, state_tuple[1]) to avoid draw one line exactly on top of another
        face_of_tag = 'Bold'
        face_of_number = 'Normal'

        #现在的Z轴该用多少
        def current_z():
            return z_start + len(nodes) + len(state_lines) + len(link_lines) + len(texts)

        def get_avoidance_text_position(avoidance, current_x, current_y):
            '''
            # using greedy algorithm to do text avoidance
            :param avoidance: all the obstacles, including texts and lines, (x,y_start,y_end)
            :param current_x: x pixel
            :param current_y: y pixel
            :return: a pixel value of variable current y
            '''

            #需要多少空间才能塞得下
            required_clearance = text_line_height-1

            # 按Y其实坐标排序当前所有的障碍
            extract = sorted([(i[1],i[2]) for i in avoidance if i[0] == current_x], key=lambda x:x[0])

            # 如果障碍之间无法放下文本，则将他们合并起来变成一个大的障碍
            for i in range(len(extract)-2,-1,-1):
                if extract[i][1]>=extract[i+1][0]:
                    extract[i] = (extract[i][0],extract[i+1][1])
                    extract.pop(i+1)

            #检查区域是否都不重叠
            if not all([extract[i][1]<extract[i+1][0] for count in range(len(extract)-1)]):
                alert_UI("get_avoidance_text_position function error. Report Bug.")

            #没障碍直接返回（似乎其实不可能发生？）
            if not extract:
                return current_y

            # 从障碍范围变成允许范围
            allowed_ranges = []
            for count,i in enumerate(extract):
                if count==0:
                    allowed_ranges.append((-float('inf'),i[0]))
                else:
                    allowed_ranges.append((extract[count-1][1],i[0]))
                if i is extract[-1]:
                    allowed_ranges.append((i[1],float('inf')))

            allowed_ranges = [x for x in allowed_ranges if x[1]-x[0]>required_clearance]


            #每个允许范围到标签想要位置的距离
            distances = []
            for i in allowed_ranges:
                #如果有允许范围直接包住了标签想要的位置，则直接返回它
                if i[0]<=current_y and i[1] >= current_y:
                    if current_y-i[0]>required_clearance:
                        return current_y
                    else:
                        # 标签内无法直接放下，需要在范围内稍微移动一点
                        return i[0]+required_clearance

                # 否则给出距离是多少
                elif i[1]<=current_y:
                    distances.append(current_y-i[1])
                elif i[0]>=current_y:
                    distances.append(i[0]-current_y)

            # 插到距离最小的地方
            insert = distances.index(min(distances))
            if allowed_ranges[insert][1]<=current_y:
                # print('via1')
                return allowed_ranges[insert][1]
            else:
                # print('via2')
                return allowed_ranges[insert][0]+required_clearance

        def add_state(use_bond, x_count, state_tuple, state_line_span, remember_drawn_states, avoidance,
                      nodes_id_start, nodes,
                      fragment_list,text_to_write,
                      color, font_index=4, size_of_tag=10, size_of_number=10, face_of_tag='Bold', face_of_number='Normal', num_with_tag=False,
                      is_tag_line=False):

            # if state_tuple[1]==None:
                # print(1)

            duplicate = False


            # independent tag line 不划线，只加tag
            if not is_tag_line:
                if use_bond:
                    bond_nodes = []

                    # bond 对应两个node
                    node = add_node(nodes_id_start + len(nodes),
                                    x_mapping(x_count - state_line_span / 2),
                                    y_mapping(state_tuple[1]),
                                    Z=current_z())
                    nodes.append(node)
                    fragment_list.append(node)

                    node = add_node(nodes_id_start + len(nodes),
                                    x_mapping(x_count + state_line_span / 2),
                                    y_mapping(state_tuple[1]),
                                    Z=current_z())
                    nodes.append(node)
                    fragment_list.append(node)


                    # Check whether user allowed overlap, or the state is not drawn
                    if (x_count, state_tuple[1]) not in remember_drawn_states or self.allow_overlap_states_checkBox.isChecked():
                        state_line = add_state_Bond(add_color(color), nodes_id_start + len(nodes) - 2, Z=current_z())
                        bond_nodes.append(nodes_id_start + len(nodes) - 2)
                        bond_nodes.append(nodes_id_start + len(nodes) - 1)
                        state_lines.append(state_line)
                        fragment_list.append(state_line)
                        remember_drawn_states.append((x_count, state_tuple[1]))
                        avoidance.append((x_count, y_mapping(state_tuple[1]) - 2, y_mapping(state_tuple[1]) + 1))
                    else:
                        duplicate = True
                else:
                    x1 = x_mapping(x_count - state_line_span / 2)
                    x2 = x_mapping(x_count + state_line_span / 2)
                    y1 = y_mapping(state_tuple[1])
                    y2 = y1
                    if (x_count, state_tuple[1]) not in remember_drawn_states or self.allow_overlap_states_checkBox.isChecked():
                        state_line = add_graphic_line(x1,x2,y1,y2,current_z(),
                                                      add_color(color),
                                                      bold=True)
                        state_lines.append(state_line)
                        fragment_list.append(state_line)
                        remember_drawn_states.append((x_count, state_tuple[1]))
                        avoidance.append((x_count, y1-2,y1+1))
                    else:
                        duplicate = True

            text_tuple = (state_tuple,avoidance,x_count,font_index,size_of_number,color,duplicate,is_tag_line,num_with_tag)

            text_to_write.append(text_tuple)
            # print(text_to_write)

            if not is_tag_line:
                if not use_bond:
                    return ((x1, x2, y1, y2),duplicate)
                else:
                    return  (bond_nodes,duplicate)

        # 所有的text得统一放到最后画，上面记录上所有这个函数需要的信息
        def write_one_text(state_tuple,avoidance,x_count,font_index,size_of_number,color,duplicate,is_tag_line,num_with_tag):
            decimal = self.state_precision_spinBox.value()
            if num_with_tag:
                text= (str(state_tuple[0])+'  ' if state_tuple[0] else "") + "{:.[decimal]f}".replace('[decimal]',str(int(decimal))).format(state_tuple[1])

                # 画在线上方
                y_position = y_mapping(state_tuple[1]) - line_width_avoidance

                if is_tag_line:
                    text=state_tuple[0]
                    # 统一记录了independent tag需要呆的坐标（比最小的低一行）
                    y_position = tags_y_position[x_count]
                if self.Auto_collision_avoidance_checkBox.isChecked():
                    y_position = get_avoidance_text_position(avoidance, x_count, y_position)
                    # print(y_position)
                    # print(text)
                text_1 = add_text(text,
                                  x_mapping(x_count), y_position ,
                                  font_index=font_index,
                                  size=size_of_number,
                                  color_index=add_color(color),
                                  face_index='Normal' if not is_tag_line else "Bold",
                                  Z=current_z()+10000)
                if text and text.strip():
                    avoidance.append((x_count, y_position - text_line_height, y_position + 1))
                text_2 = ""

            else:
                tag_y_position = y_mapping(state_tuple[1]) + text_line_height
                if is_tag_line:
                    tag_y_position = tags_y_position[x_count]
                if self.Auto_collision_avoidance_checkBox.isChecked():
                    tag_y_position = get_avoidance_text_position(avoidance, x_count, tag_y_position)

                text_1 = add_text(state_tuple[0],
                                  x_mapping(x_count), tag_y_position,
                                  font_index=font_index,
                                  size=size_of_tag,
                                  color_index=add_color(color),
                                  face_index="Bold",
                                  Z=current_z()+10000)
                if isinstance(state_tuple[0],str) and state_tuple[0].strip():
                    avoidance.append((x_count, tag_y_position - text_line_height, tag_y_position + 1))

                number_y_position = y_mapping(state_tuple[1]) - line_width_avoidance
                if self.Auto_collision_avoidance_checkBox.isChecked():
                    number_y_position = get_avoidance_text_position(avoidance, x_count, number_y_position)


                text_2 = add_text("{:.[decimal]f}".replace('[decimal]',str(int(decimal))).format(state_tuple[1]),
                                  x_mapping(x_count), number_y_position,
                                  font_index=font_index,
                                  size=size_of_number,
                                  color_index=add_color(color),
                                  face_index='Normal',
                                  Z=current_z()+10000)


                if is_tag_line or duplicate:
                    text_2=""
                else:
                    avoidance.append((x_count, number_y_position - text_line_height, number_y_position + 1))

            if text_1:
                texts.append(text_1)
                current_texts.append(text_1)
            if text_2:
                texts.append(text_2)
                current_texts.append(text_2)

        def route(states, id_start,text_to_write,current_texts, nodes, color, font_index, size_of_tag, size_of_number, face_of_tag, face_of_number, num_with_tag):
            '''

            :param states: [('Sub', -60.), (None, 80.8), ('TS', 0), (None, None), ('IM1', 86.6), ('Prod', -46.2)]
                          the tag and energy pair of each states, the Value is None if empty
            :return: A fragment, contain all nodes, state_lines and linkers of a single route (single line in xlsx) 
            '''

            node_used = []

            is_tag_line = 'TAG' in color
            color = color.replace('TAG', "")

            fragment_list = []


            # 记录是否有用line记录的上一个态，用于决定是否加link
            last_state = None
            last_duplicate=None
            # 记录是否有用bond记录的上一个态
            has_node = False
            last_bond_nodes_duplicated = None
            if self.use_lines_radioButton.isChecked():
                for count, state in enumerate(states):

                    if state[1]==None:
                        continue

                    add_state_ret = add_state(False,count, state, state_line_span,remember_drawn_states,avoidance,
                              nodes_id_start, nodes,
                              fragment_list,text_to_write,
                              color, font_index=font_index, size_of_number=size_of_number, size_of_tag=size_of_tag, face_of_tag=face_of_tag,
                              face_of_number=face_of_number, num_with_tag=num_with_tag, is_tag_line=is_tag_line)

                    if add_state_ret!=None:
                        current_state, current_duplicate = add_state_ret

                        # 判定是不是之前有一个态（注意不能用是否是第一个态来确定，因为某些曲线不是从x=0开始的
                        if last_state != None and not is_tag_line and last_duplicate!=None and not all([last_duplicate,current_duplicate]):
                            x2, dump, y2, dump = current_state
                            dump, x1, dump, y1 = last_state
                            if self.dash_line_radioButton.isChecked():
                                link = add_graphic_line(x1,x2,y1,y2,
                                                        Z=current_z(),color_index=add_color(color),dash=True)
                            else:
                                link = add_graphic_line(x1, x2, y1, y2,
                                                        Z=current_z(), color_index=add_color(color))
                            link_lines.append(link)
                            fragment_list.append(link)

                        last_state = current_state
                        last_duplicate = current_duplicate

            else:
                fragment_list = [fragment_template]
                for count, state in enumerate(states):
                    if state[1] ==None:
                        continue

                    add_bond_state_ret = add_state(True,count, state, state_line_span,remember_drawn_states,avoidance,
                              nodes_id_start, nodes,
                              fragment_list,text_to_write,
                              color, font_index=font_index, size_of_number=size_of_number, size_of_tag=size_of_tag, face_of_tag=face_of_tag,
                              face_of_number=face_of_number, num_with_tag=num_with_tag, is_tag_line=is_tag_line)

                    if add_bond_state_ret!=None:
                        bond_nodes, bond_nodes_duplicated = add_bond_state_ret

                        if not is_tag_line:
                            if bond_nodes:
                                node_used+=bond_nodes

                        if has_node and not is_tag_line and last_bond_nodes_duplicated!=None and not all([last_bond_nodes_duplicated,bond_nodes_duplicated]):
                            if self.dash_line_radioButton.isChecked():
                                add_link_connection = add_dash_link_Bond
                            else:
                                add_link_connection = add_solid_link_Bond
                            link_start_node =nodes_id_start + len(nodes) - 3
                            link = add_link_connection(add_color(color), link_start_node, current_z())
                            node_used.append(link_start_node)
                            node_used.append(link_start_node+1)

                            link_lines.append(link)
                            fragment_list.append(link)

                        has_node = True

                        last_bond_nodes_duplicated=bond_nodes_duplicated

            def sort_key(x:str):
                # 整理一下，好与chemdraw标准文件比对（其实没用，无大碍）
                if x.startswith('<b'):
                    return 2
                if x.startswith('<n'):
                    return 1
                if x.startswith('<t'):
                    return 3
                if x.startswith('<a'):
                    return 4
                return 0

            fragment_list.sort(key=sort_key)

            for i in range(len(fragment_list)-1,-1,-1):
                # 删除没有用过的，没有连接某根键的node，否则chemdraw会用红的错误圈标出来重合的空节点
                re_ret = re.findall('''^<n\n id=\\"(\d+?)\\"''',fragment_list[i])
                if re_ret:
                    id = int(re_ret[0])
                    if id not in node_used:
                        fragment_list.pop(i)

            # 空 fragment 不返回，否则会在打开时报一个警报
            ret = []
            if any([x.strip() for x in fragment_list[1:]]) or (fragment_list and fragment_list[0].startswith("<graphic")):
                if self.use_bonds_radioButton.isChecked():
                    fragment_list.append(fragment_template_end)
                ret += [add_group(current_z())]+fragment_list+[group_template_end]

            return ret

        fragments = []
        text_to_write = []

        for route_count, route_content in enumerate(self.paths_for_cdx_drawing):
            fragments += route(route_content,
                               nodes_id_start,
                               text_to_write,
                               current_texts,
                               nodes,
                               self.colors_for_cdx_drawing[route_count],
                               font_index=4,
                               size_of_tag=size_of_tag, size_of_number=size_of_number, face_of_tag=face_of_tag, face_of_number=face_of_number,
                               num_with_tag=num_with_tag)

        maximum_x_count = max([x[2] for x in text_to_write])
        for i in range(maximum_x_count+1):
            #调整顺序，让贪心更合理 [1,2,3,4,5,6,7,8] --> [4,5,3,6,2,7,1,8]
            texts_of_this_x = [x for x in text_to_write if x[2]==i and x[6]==False]
            texts_of_this_x.sort(key = lambda x:x[0][1])# 按Y轴位置排序
            # 从中位数开始往两边写,中位数只有一个的时候归前面
            temp1 = list(reversed(texts_of_this_x[:math.ceil(len(texts_of_this_x)/2)]))
            temp2 = texts_of_this_x[math.ceil(len(texts_of_this_x)/2):]
            sequence_to_write = [None]*len(texts_of_this_x)
            sequence_to_write[::2] = temp1
            sequence_to_write[1::2] = temp2
            # 把tag放在最后
            tag_line = [x for x in sequence_to_write if x[7]==True]
            if tag_line:
                sequence_to_write.remove(tag_line[0])
                sequence_to_write.append(tag_line[0])

            for text_function_tuple in sequence_to_write:
                write_one_text(*text_function_tuple)

        fragments += current_texts

        def automatic_Y_axis_label_distance(y_range):
            distance = y_range[1]-y_range[0]
            #截取至整10，整2或整5，取7个以上,14个以下
            sep=10**(math.floor(math.log10(distance))-1)
            count = int(distance/sep) # should between 10 and 100
            if 10<=count<=15:
                pass
            elif 16<=count<=35:
                sep*=2
            elif 35<=count<=70:
                sep*=5
            elif 71<=count<=99:
                sep*=10
            else:
                alert_UI('automatic_Y_axis_distance Function Error. Report bug.','automatic_Y_axis_distance Function Error.')

            return sep

        # draw_frame
        def add_frame(x_range, y_range,
                      y_axis_annotate_text, y_axis_annotate_offset=-30,
                      y_label_offset=-4,
                      color_of_text='k',
                      tick_length=2.2,
                      Z=200):
            # add this to the fragment

            ret = []
            #四个角的位置
            x_start_pixel = x_mapping(x_range[0])
            x_end_pixel = x_mapping(x_range[1])
            y_start_pixel = y_mapping(y_range[0])
            y_end_pixel = y_mapping(y_range[1])

            #四条边
            ret.append(add_graphic_line(x_start_pixel, x_end_pixel, y_start_pixel, y_start_pixel, current_z()))
            ret.append(add_graphic_line(x_start_pixel, x_end_pixel, y_end_pixel, y_end_pixel, current_z()))
            ret.append(add_graphic_line(x_start_pixel, x_start_pixel, y_start_pixel, y_end_pixel, current_z()))
            ret.append(add_graphic_line(x_end_pixel, x_end_pixel, y_start_pixel, y_end_pixel, current_z()))

            ret.append(test(x_start_pixel, y_start_pixel))

            # X tick
            if self.X_tick_checkBox.isChecked():
                for x_mark_position in range(int(math.floor(x_range[0]) + 1), int(math.floor(x_range[1])+1)):
                    ret.append(add_graphic_line(x_mapping(x_mark_position), x_mapping(x_mark_position),
                                                y_mapping(y_range[0]), y_mapping(y_range[0]) - tick_length, current_z()))

            if self.Y_tick_auto_checkBox.isChecked() or \
                    self.Y_tick_doubleSpinBox.value()==0 or \
                    self.Y_number_doubleSpinBox.value()==0:

                #默认值，tick按函数取，label是tick的一半多，precision看末位差别
                y_tick_distance = automatic_Y_axis_label_distance(y_range)
                y_label_distance = y_tick_distance*2
                y_label_precision = 0 if math.log10(y_tick_distance)>=0 else -math.floor((math.log10(y_tick_distance)))
            else:
                y_tick_distance = self.Y_tick_doubleSpinBox.value()
                y_label_distance = self.Y_number_doubleSpinBox.value()
                y_label_precision = self.Y_label_precision_spinBox.value()


            # 从0开始数，标整数
            for y_mark_count in range(int(math.floor(y_range[0] / y_tick_distance) + 1), int(math.floor(y_range[1] / y_tick_distance)+1)):
                y_mark_position = y_mark_count * y_tick_distance
                #太接近轴就不标了
                if abs(y_mark_position-y_range[0])/(y_range[1]-y_range[0])<0.02 or\
                    abs(y_mark_position - y_range[1]) / (y_range[1] - y_range[0]) < 0.02:
                    continue
                ret.append(add_graphic_line(x_mapping(x_range[0]), x_mapping(x_range[0]) + tick_length,
                                            y_mapping(y_mark_position), y_mapping(y_mark_position), current_z()))
            for y_mark_count in range(int(math.floor(y_range[0] / y_label_distance) + 1), int(math.floor(y_range[1] / y_label_distance) + 1)):
                y_mark_position = y_mark_count * y_label_distance
                # 太接近轴就不标了
                if abs(y_mark_position - y_range[0]) / (y_range[1] - y_range[0]) < 0.02 or \
                        abs(y_mark_position - y_range[1]) / (y_range[1] - y_range[0]) < 0.02:
                    continue
                ret.append(add_text("{:.[digit]f}".replace('[digit]', str(y_label_precision)) \
                                    .format(y_mark_position),
                                    x_mapping(x_range[0]) + y_label_offset,
                                    y_mapping(y_mark_position)+size_of_number/2-1,
                                    Z=current_z()+10000,
                                    font_index=font_index, size=size_of_number, color_index=add_color(color_of_text),
                                    face_index=face_of_number, right_align=True))
            ret.append(group_template_end)
            ret.append(add_text(y_axis_annotate_text, x_mapping(x_range[0]) + y_axis_annotate_offset,
                                y_mapping((y_range[0] + y_range[1]) / 2),
                                Z=current_z()+10000,
                                font_index=font_index, size=size_of_number, color_index=add_color(color_of_text),
                                face_index=face_of_tag, vertical=True
                                ))
            return [add_group(current_z())]+ret

        y_axis_text = self.Y_label_lineEdit.text()

        fragments+=add_frame(x_range,y_range,y_axis_text)

        output_list = [default_document] + colors_list + [default_colors_end] + \
                      fonts_list + [add_page(3,3)] + fragments + [page_template_end] + [
            default_document_end]

        save_successful = False
        if self.use_temp_file_checkBox.isChecked():
            output_file = os.path.join(temp_folder, 'temp_' + str(int(time.time() * 100000)) + '.cdxml')
        else:
            output_file = filename_class(self.xlsx_file_lineEdit.text()).replace_append_to('cdxml')

        #如果已存在，提示替换
        if os.path.isfile(output_file):
            button = Qt.QMessageBox.warning(None,'File exist',"File already exist, confirm replacing:",Qt.QMessageBox.Ok|Qt.QMessageBox.Cancel)
            if button == Qt.QMessageBox.Cancel:
                output_file = os.path.join(temp_folder, 'temp_' + str(int(time.time() * 100000)) + '.cdxml')

        while not save_successful:
            try:
                with open(output_file, 'w') as cdxml_output:
                    for i in output_list:
                        cdxml_output.write(i)
                        cdxml_output.write('\n\n')
                save_successful=True
            except PermissionError:
                alert_UI("Permission Error\n    File may in use. Close this file in ChemDraw.\n    Or you don't have permission to write to this directory, give the program administrator privilege.\nThen click OK.","Permission Error")

        open_explorer_and_select(output_file)

    def load_xlsx(self, filename):
        worksheet = load_workbook(filename=filename).worksheets[0]

        xlsx_list = []
        for row in worksheet.rows:
            row_list = []
            for cell in row:
                if cell.value=="":
                    row_list.append(None)
                else:
                    row_list.append(cell.value)
            xlsx_list.append(row_list)


        # 第一行可以有独立的tag，也可以没有。如果第一行全部是数字，或第一行的颜色用“TAG”标注，认为是tag
        has_independent_tag=False
        #可以用tag标记
        if isinstance(xlsx_list[0][0],str) and "tag" in xlsx_list[0][0].lower():
            has_independent_tag = True
        #有任何一个非空str
        if any([isinstance(x,str) for x in xlsx_list[0][1:]]):
            has_independent_tag = True
        #全是空的（用来兼容原来的文件）
        if all([x==None for x in xlsx_list[0][1:]]):
            has_independent_tag = True

        if has_independent_tag:
            self.independent_tags = xlsx_list[0]
            xlsx_list = xlsx_list[1:]
        else:
            self.independent_tags = [None for dump in range(max([len(i) for i in xlsx_list]))]


        if self.independent_tags[0]==None:
            self.independent_tags[0]='k'

        # split to data part and Tag part
        self.xlsx_data = []

        for count, row in enumerate(xlsx_list):
            if not [x for x in row if x != None]:  # All None sequence; No remove_blank was used is because all 0 line might exist
                self.xlsx_data.append(xlsx_list[:count])
                self.xlsx_data.append(xlsx_list[count + 1:])
                break
        else:  # 没有任何标签
            self.xlsx_data.append(xlsx_list)
            self.xlsx_data.append([])

        for row in self.xlsx_data[0]:
            if row[0]==None:
                row[0]='k'

        # rule out number of Tag is more than Data
        if len(self.xlsx_data[0]) < len(self.xlsx_data[1]):
            alert_UI('You have '+str(len(self.xlsx_data[1]))+' tag lines and '+str(len(self.xlsx_data[0])) +' data lines in the Excel input file.\nThere are more tag lines than needed. Correct this mistake.',"Tag More Than Data Error.")
            exit()


        self.xlsx_data[1] = [[(str(y) if (y is not None) else y) for y in x] for x in self.xlsx_data[1]]

        # supplement data and tag according to dimension of data
        self.column_count = max([len(x) for x in self.xlsx_data[0]])
        self.row_count = len(self.xlsx_data[0])

        self.xlsx_data[1] += [[None] * self.column_count] * (self.row_count - len(self.xlsx_data[1]))

        for row in range(self.row_count):
            self.xlsx_data[0][row] += [None] * (self.column_count - len(self.xlsx_data[0][row]))
            self.xlsx_data[1][row] += [None] * (self.column_count - len(self.xlsx_data[1][row]))

        self.routes = [list(zip(self.xlsx_data[1][row], self.xlsx_data[0][row])) for row in range(self.row_count)]

        # Add a route of indepenedent tags at the last
        # get each minimum of a column of datas, to determine the position of independent tag
        minimums = []
        for column_count in range(1, len(self.xlsx_data[0][0])):
            column = [row[column_count] for row in self.xlsx_data[0]]
            column = [x for x in column if isinstance(x, float)]
            if column:
                column_min = min(column) - (max(column) - min(column)) * 0.5
            else:
                column_min = -100
            minimums.append(column_min)

        self.routes.append([(None, self.independent_tags[0] + "TAG")] + [(x, minimums[count]) for count, x in enumerate(self.independent_tags[1:])])

        # print(self.routes)

    # def table_update(self):
    #     self.data_tableWidget.setColumnCount(self.column_count)
    #     self.data_tableWidget.setRowCount(self.row_count * 3 - 1)
    #     for row, route in enumerate(self.routes):
    #         for column, state in enumerate(route):
    #             text1 = Qt.QLabel(
    #                 '''<p align="center"><span style="font-size:10pt; font-weight:600;">''' + (str(state[0]) if state[0] != None else "") + '''</span></p>''')
    #
    #             if state[1] == None:
    #                 format_num = ""
    #             else:
    #                 try:
    #                     format_num = "{:.1f}".format(state[1])
    #                 except:
    #                     format_num = str(state[1])
    #
    #             text2 = Qt.QLabel('''<p align="center"><span style="font-size:10pt;">''' + format_num + '''</span></p>''')
    #
    #             self.data_tableWidget.setCellWidget(row * 3, column, text1)
    #             self.data_tableWidget.setCellWidget(row * 3 + 1, column, text2)
    #
    #     self.resize_table()

    # def resize_table(self):
    #     # equal-the widths
    #     self.show()
    #     total_width = self.data_tableWidget.size().width() - 2
    #
    #     if total_width / self.column_count > 45:
    #         for column in range(self.column_count):
    #             self.data_tableWidget.setColumnWidth(column, total_width / self.column_count)
    #     else:
    #         for column in range(self.column_count):
    #             self.data_tableWidget.setColumnWidth(column, 45)


if __name__ == '__main__':
    my_Qt_Program = myWidget()

    my_Qt_Program.show()
    sys.exit(Application.exec_())
