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

sys.path.append('D:\My_Program\Python_Lib')
sys.path.append('E:\My_Program\Python_Lib')

# ALL Numbers in SI if not mentioned
R = 8.3144648
k_B = 1.3806503E-23
N_A = 6.02214179E23
c = 299792458
h = 6.62606896E-34
pi = math.pi

# units
Hartree__kcal_mol = 627.51
Hartree__KJ_mol = 2625.49962
Hartree__J = 4.359744575E-18
Hartree__cm_1 = 219474.6363

kcal__kJ = 4.184
atm__Pa = 101325

bohr__m = 5.2917721092E-11
bohr__A = 0.52917721092
amu__kg = 1.660539040E-27

month = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]


def listdir(filename: str):
    # listdir of a file or a folder, return a list, contain the absolute path of the
    if os.path.isfile(filename):
        path = filename_class(filename[0]).path
        return [os.path.join(path, x) for x in os.listdir(path)]
    elif os.path.isdir(filename):
        path = filename
        return [os.path.join(path, x) for x in os.listdir(path)]


def nCr(n, r):
    f = math.factorial
    return f(n) / f(r) / f(n - r)


def reverse(string):
    l = list(string)
    l.reverse()
    return "".join(l)


def rreplace(string, old, new, count=None):
    """string right replace"""
    string = str(string)
    r = reverse(string)
    if count is None:
        count = -1
    r = r.replace(reverse(old), reverse(new), count)
    return type(string)(reverse(r))


def open_explorer_and_select(file_path):
    import subprocess
    open_explorer_command = r'explorer /select,"' + str(file_path).replace('/', '\\') + '"'
    subprocess.Popen(open_explorer_command)

def merge_dicts(*dict_args):
    """
    Given any number of dicts, shallow copy and merge into a new dict,
    precedence goes to key value pairs in latter dicts.
    """
    result = {}
    for dictionary in dict_args:
        result.update(dictionary)
    return result


def list_or(input_list):
    # input [a,b,c] return a or b or c
    ret = False
    for i in input_list:
        ret = ret or i
    return ret


def list_and(input_list):
    # input [a,b,c] return a and b and c
    ret = True
    for i in input_list:
        ret = ret and i
    return ret


import os


def addToClipBoard(text):
    import pyperclip
    pyperclip.copy(text)

#
# def rreplace(original_string, search, replace):
#     return replace.join(original_string.rsplit(search, 1))


class filename_class:
    def __init__(self, fullpath):
        fullpath = fullpath.replace('\\', '/')
        self.depth = fullpath.count('/')
        self.re_path_temp = re.match(r".+/", fullpath)
        if self.re_path_temp:
            self.path = self.re_path_temp.group(0)  # 包括最后的斜杠
        else:
            self.path = ""
        self.name = fullpath[len(self.path):]
        if self.name.rfind('.') != -1:
            self.name_stem = self.name[:self.name.rfind('.')]  # not including "."
            self.append = self.name[len(self.name_stem) - len(self.name) + 1:]
        else:
            self.name_stem = self.name
            self.append = ""

        self.only_remove_append = self.path + self.name_stem  # not including "."

    def replace_append_to(self, new_append):
        return self.only_remove_append + '.' + new_append


def replace_append(filepath, new_append):
    return filename_class(filepath).replace_append_to(new_append)


def get_dict_value(dict, key):
    return dict[key] if key in dict else ""


def remove_key_from_dict(dict, key):
    if key in dict:
        dict.pop(key)
    return dict


def safe_read_dict(dict, key, default_value=""):
    if key in dict:
        return dict[key]
    else:
        return default_value


def secure_print(*object_to_print):
    # print some character will cause UnicodeEncodeError,
    # if the message is not necessarily printed, use this function will just print nothing and aviod the error

    try:
        print(*object_to_print)
    except:
        print("Print function error. Print of information omitted.")


def get_print_str(*object_to_print, sep=" "):
    ret = ""
    for object in object_to_print:
        try:
            ret += str(object) + sep
        except:
            print("get_print_str Error...")

    return ret


def read_last_n_lines_fast(file, n_lines):
    return read_last_n_char_fast(file, '\n', n_lines)


# def read_last_n_char_fast(file, char, n):
#     '''
#     a fast method to read the last n appearence of a specific character, and return one multi-line decoded string
#     :param file:
#     :param n:
#     :return:
#     '''
#     char = char.encode()
#     import os
#     current_line = 0
#     with open(file, 'rb') as f:
#         f.seek(-2, os.SEEK_END)
#         while True:
#             if f.read(1) == char:
#                 current_line += 1
#                 if current_line == n:
#                     last_lines = f.read().decode()
#                     break
#             f.seek(-2, os.SEEK_CUR)
#
#     return last_lines

def read_last_n_char_fast(file, char, n):
    """
    a fast method to read the last n appearance of a specific character, and return one multi-line decoded string
    if there is less than n matches, the whole file will be returned
    :param file:
    :param n:
    :return:
    """
    char = char.encode('utf-8')
    import mmap
    with open(file,"r+b") as f:
        # memory-map the file, size 0 means whole file
        m = mmap.mmap(f.fileno(), 0)
                              # prot argument is *nix only
        current_cut = m.rfind(char)
        count=1
        while count<n:
            current_cut = m.rfind(char,0,current_cut)
            if current_cut == -1:
                count = n
                current_cut = 0
            count+=1
        m.seek(current_cut)
        return m.read().decode()
#
#
# file = r"D:\Gaussian\LXQ_Rh_Carbene\Confsearch\190\temp\190_solvated__confsearch_P1__cpptraj__nosol_nobox_04_17__last__xtbopt_traj__Pieced___135.xtbopt_traj.xyz"
# print(read_last_n_char_fast(file,"\n",100))

def split_list_by_item(input_list: list, separator, lower_case_match=False, include_separator=False, include_empty=False):
    return split_list(input_list, separator, lower_case_match, include_separator, include_empty)


def split_list(input_list: list, separator, lower_case_match=False, include_separator=False, include_separator_after=False, include_empty=False):
    '''

    :param input_list:
    :param separator: a separator, either a str or function. If it's a function, it should take a str as input, and return
    :param lower_case_match:
    :param include_separator:
    :param include_empty:
    :return:
    '''
    ret = []
    temp = []

    if include_separator or include_separator_after:
        assert not (include_separator and include_separator_after), 'include_separator and include_separator_after can not be True at the same time'

    for item in input_list:

        split_here_bool = False
        if callable(separator):
            split_here_bool = separator(item)
        elif isinstance(item, str) and item == separator:
            split_here_bool = True
        elif lower_case_match and isinstance(item, str) and item.lower() == separator.lower():
            split_here_bool = True

        if split_here_bool:
            if include_separator_after:
                temp.append(item)
            ret.append(temp)
            temp = []
            if include_separator:
                temp.append(item)
        else:
            temp.append(item)
    ret.append(temp)

    if not include_empty:
        ret = [x for x in ret if x]

    return ret


def get_appropriate_ticks(ranges, num_tick_limit=(4, 6), accept_cloest_out_of_range=True):
    '''
    a function to get the desired ticks, e.g. for 1.2342 - 1.58493, with a tick_limit of (4,8),
    the tick should be (1.25,1.30,1.35,1.40,1.45,1.50,1.55)
    :param ranges: a 2-tuple, upper limit and lower limit
    :param num_tick_limit: the maximum and minimum amount of ticks
    :param accept_cloest_out_of_range: if cloest, out-of-range answer is accepted if in-range answer is not possible
    :return: a (lower-limit, upper-limit, spacing) tuple
    if no appropriate choice is possible, and accept_cloest_out_of_range = False, return [ranges[1],ranges[0],ranges[1]-ranges[0]]
    '''
    # the ticking should be either ending in 5 or 2 or 0
    if ranges[1] < ranges[0]:
        ranges = reversed(ranges)

    assert all(x > 0 for x in num_tick_limit)
    assert ranges[0] != ranges[1]

    span = abs(ranges[1] - ranges[0])
    mid_limit = sum(num_tick_limit) / 2
    ideal_distance = span / mid_limit
    ideal_distance_log = int(math.log(ideal_distance, 10))
    test_distance = 10 ** ideal_distance_log
    test_distances = [test_distance * x for x in (0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50, 100)]
    tick_start_points = [math.ceil(ranges[0] / x) * x for x in test_distances]
    test_tick_count = []
    for count, test_distance in enumerate(test_distances):
        tick_start_point = tick_start_points[count]
        test_tick_count.append(int((ranges[1] - tick_start_point) / test_distance))
    acceptable_counts = [x for x in test_tick_count if num_tick_limit[0] <= x <= num_tick_limit[1]]
    if not acceptable_counts and not accept_cloest_out_of_range:
        return [ranges[0], ranges[1], ranges[1] - ranges[0]]
    elif not acceptable_counts:
        acceptable_count_differences = [abs(x - mid_limit) for x in test_tick_count]
        optimal_index = acceptable_count_differences.index(min(acceptable_count_differences))
        optimal_distance = test_distances[optimal_index]
    else:
        acceptable_count_differences = [abs(x - mid_limit) for x in acceptable_counts]
        optimal_count = acceptable_counts[acceptable_count_differences.index(min(acceptable_count_differences))]
        optimal_index = test_tick_count.index(optimal_count)
        optimal_distance = test_distances[optimal_index]
    optimal_start_point = tick_start_points[optimal_index]
    optimal_end_point = int((ranges[1] - optimal_start_point) / optimal_distance) * optimal_distance + optimal_start_point
    return [optimal_start_point, optimal_end_point, optimal_distance]


def get_input_with_while_cycle(break_condition=lambda x: not x.strip(), input_prompt="", strip_quote=True, backup_file=None):
    '''
    get multiple line of input, terminate with a condition, return the accepted lines
    :param break_condition: give a function, when it is met, the while loop is terminated.
    :param input_prompt: will print this every line
    :param backup_file: a file-like object (created by "open()") which will store the inputs for backup
    :return: list of accepted lines
    '''

    ret = []
    while True:
        input_line = input(input_prompt)
        if backup_file:
            backup_file.write(input_line)
            backup_file.write('\n')
        if strip_quote:
            input_line = input_line.strip().strip('"')
        if break_condition(input_line):
            break
        else:
            ret.append(input_line)
    return ret


def PolygonArea(corners):
    n = len(corners)  # of corners
    area = 0.0
    for i in range(n):
        j = (i + 1) % n
        area += corners[i][0] * corners[j][1]
        area -= corners[j][0] * corners[i][1]
    area = abs(area) / 2.0
    return area


def remove_special_chr_from_str(input_str):
    '''
    A function for fuzzy search "3-propyl-N'-ethylcarbodiim"-->"propylnethylcarbodiim"
    :param input_str:
    :return:
    '''
    import string
    ret = ''.join(ch for ch in input_str if ch not in string.punctuation + string.whitespace + string.digits).lower()
    if not ret:
        ret = ''.join(ch for ch in input_str if ch not in string.punctuation + string.whitespace).lower()
    else:
        ret = input_str
    return ret


def get_unused_filename(input_filename, replace_hash=True, use_proper_filename=True):
    '''
    verify whether the filename is already exist, if it is, a filename like filename_01.append; filename_02.append will be returned.
    maximum 99 files can be generated
    :param input_filename:
    :return: a filename
    '''

    input_filename = os.path.realpath(input_filename)

    if use_proper_filename:
        input_filename = proper_filename(input_filename, replace_hash=replace_hash)

    if not os.path.isfile(input_filename) and not os.path.isdir(input_filename):
        # 是新的
        return input_filename
    else:
        if os.path.isfile(input_filename):
            no_append = filename_class(input_filename).only_remove_append
            append = filename_class(input_filename).append
        else:
            no_append = input_filename
            append = ""

        number = 1
        ret = no_append + "_" + '{:0>2}'.format(number) + (('.' + append) if append else "")
        while os.path.isfile(ret) or os.path.isdir(ret):
            number += 1
            if number == 9999:
                Qt.QMessageBox.critical(None, "YOU HAVE 9999 INPUT FILE?!", "AND YOU DON'T CLEAN IT?!",
                                        Qt.QMessageBox.Ok)
                break
            ret = no_append + "_" + '{:0>2}'.format(number) + (('.' + append) if append else "")

        return ret


optimization_timer_u3yc24t04389y09sryc09384yn098 = 0 # this wired name is to avoid collision with other files
def optimization_timer(position_label=""):
    '''
    A simple function to record the time to current operation, and print the eclapsed time till then
    '''

    # return None # comment this out to activate this function

    global optimization_timer_u3yc24t04389y09sryc09384yn098
    if optimization_timer_u3yc24t04389y09sryc09384yn098==0:
        optimization_timer_u3yc24t04389y09sryc09384yn098 = time.time()
    else:
        delta = time.time() - optimization_timer_u3yc24t04389y09sryc09384yn098
        optimization_timer_u3yc24t04389y09sryc09384yn098 = time.time()
        print("————————————",position_label,int(delta*1000))


def remove_duplicate(input_list: list, access_function=None):
    ret = []
    index = []
    for i in input_list:
        if callable(access_function):
            if access_function(i) not in index:
                index.append(access_function(i))
                ret.append(i)
        else:
            if i not in index:
                index.append(i)
                ret.append(i)

    return ret


def remove_blank(input_list: list):
    return [x for x in input_list if x]


def cas_wrapper(input: str, strict=False, correction=False):
    '''
    Match or not:
                    Partial     strict
    111-11-5           Yes        Yes     normal match
    111-11-1           No!         No     not match CAS number with wrong check digit
    111-11-5aa         Yes        Yes     match if other non [digit,'-'] concatenate with it

    # partial (will print a warning)
    111-11             Yes                partial match
    111-11-aa          Yes                same for partial match
    111-11aa           Yes                same for partial match
    111-1

    # not match with other number concatenate with it (prevent phone-number match)
    111-11-523          No
    0111-11-5           No

    # wrong format
    111-112-3           No
    111-119             No
    12345678-12-2       No

    :param input:
    :param strict: match complete or partial
    :param correction:允许纠正验证位错误
    :return: completed CAS number, if not find or the check digit not match the initial input, return '',
    '''

    prefix = r"(^|[^\d-])"  # prevent 0111-11-1
    base = r"([1-9]\d{1,7}-\d{2})"  # matches 111-11
    postfix = r"(\-\d)"
    closure_complete = r"($|[^\d-])"  # matches 111-11-1, prevent 111-11-123
    closure_partial = r"|(\-($|[^\d-]))|($|[^\d-])"  # matches 111-11-, 111-11

    re_complete = ''.join([prefix, "(", base, postfix, ")", closure_complete])
    re_partial = ''.join([prefix, base, '((', postfix, closure_complete, ')', closure_partial, ')'])

    find_complete = re.findall(re_complete, input)  # match complete 128-38-2-->128-38
    find_partial = re.findall(re_partial, input)  # match the former digits of 128-38-2-->128-38

    if strict:
        if len(find_complete) > 1:  # 找到多个结果
            print('\n\n\nMultiple CAS match.', input, '\n\n\n')

        if not find_complete:
            return ""

        find_complete = find_complete[0][1]
        find_partial = find_partial[0][1]

    else:
        if len(find_partial) > 1:  # 找到多个结果
            print('\n\n\nMultiple CAS match.', input, '\n\n\n')

        if not find_partial:
            return ""

        find_partial = find_partial[0][1]

        if find_complete:
            find_complete = find_complete[0][1]
        else:
            find_complete = ""

    # 计算验证位
    only_digit = list(reversed([int(dig) for dig in find_partial if dig.isdigit()]))
    check_digit = sum([only_digit[i] * (i + 1) for i in range(len(only_digit))]) % 10

    ret = find_partial + '-' + str(check_digit)

    if find_complete and ret == find_complete:
        return ret

    else:
        if strict:  # 如果 strict，不满足检验直接跳出检测，返回空
            return ""

    if not find_complete:
        print('CAS Wrapper Doubt! Find:', repr(ret), '. Complete wrapper:', repr(find_complete), '. Original:',
              repr(input))
        return ret

    if find_complete and ret != find_complete:
        print('CAS Wrapper Disagree! Find:', repr(ret), '. Complete wrapper:', repr(find_complete), '. Original:',
              repr(input))
        if correction:  # 允许纠正错误的验证位
            return ret
        else:
            return ""

    if not find_partial:
        return ""


def transpose_2d_list(list_input):
    return list(map(list, zip(*list_input)))


def filename_filter(input_filename, including_appendix=True, path_as_filename=False, replace_hash=True):
    return proper_filename(input_filename, including_appendix, path_as_filename, replace_hash)


def is_float(input_str):
    # 确定字符串可以转换为float
    try:
        float(input_str)
        return True
    except:
        return False


def is_int(input_str):
    if not is_float(input_str):
        return False
    num = float(input_str)
    if int(input_str) == num:
        return True
    else:
        return False


def proper_filename(input_filename, including_appendix=True, path_as_filename=False, replace_hash=True, replace_dot=True, replace_space=True):
    '''

    :param input_filename:
    :param including_appendix:
    :param path_as_filename: 是否将路径转换为文件名(/home/gauuser/file.txt --> __home__gauuser__file.txt )
    :return:
    '''
    if path_as_filename:
        path = ""
        filename_stem = filename_class(input_filename).only_remove_append
    else:
        path = filename_class(input_filename).path
        filename_stem = filename_class(input_filename).name_stem
    append = filename_class(input_filename).append

    # remove illegal characters of filename
    forbidden_chr = "<>:\"'/\\|?*-\n. "
    if not replace_hash:
        forbidden_chr = forbidden_chr.replace('-', '')
    if not replace_space:
        forbidden_chr = forbidden_chr.replace(' ', '')
    if not replace_dot:
        forbidden_chr = forbidden_chr.replace('.', '')
    for chr in forbidden_chr:
        filename_stem = filename_stem.replace(chr, '__')

    if append:
        if including_appendix:
            ret = filename_stem + '.' + append
        else:
            ret = filename_stem + '_' + append
    else:
        ret = filename_stem

    while "____" in ret:
        ret = ret.replace('____', "__")

    return os.path.join(path, ret)


def same_length_2d_list(input_2D_list, fill=""):
    '''
    read a list of list, and fill the sub_list to the same length
    :return: 
    '''
    max_column_count = max([len(x) for x in input_2D_list])
    ret = [x + [""] * (max_column_count - len(x)) for x in input_2D_list]
    return ret


def find_within_braket(input_str, get_last_one=False):
    '''
    Get all text within braket
    :param input_str: "123123127941[12313[123]112313]adaf[123]
    :return: [12313[123]112313][123]
    '''
    in_braket = 0
    ret = ""

    last_one_start = []  # 记录每一个最外括号起始位置
    last_one_end = -1  # 记录最后一个最外括号终止位置

    for count, char in enumerate(input_str):
        if char == '[':
            if in_braket == 0:
                last_one_start.append(count)
            in_braket += 1
        if char == ']':
            in_braket -= 1
            if in_braket == 0:
                last_one_end = count
        if char == ']' or in_braket > 0:
            ret += char

    if get_last_one:
        if last_one_start and last_one_end != -1:
            qualified_starts = [x for x in last_one_start if x < last_one_end]
            if qualified_starts:
                return input_str[qualified_starts[-1]:last_one_end + 1]
            else:
                return ""

    return ret


def phrase_range_selection(input_str, by_index=True):
    '''
    Input a range like 1,5,7-9; output a list by index [0,4,6,7,8]; if not index [1,5,7,8,9]
    :param input_str:
    :param by_index: the index will be 1 less than what's inputed
    :return:
    '''

    if not input_str.strip():
        return []

    input_list = input_str.replace(',', ' ').split(' ')
    choices = copy.deepcopy(input_list)

    for choice in input_list:
        if '-' in choice:
            choices.remove(choice)
            if not re.findall('\d+\-\d+', choice):
                print("Invalid")
                return None

            start, end = choice.split('-')
            choices += [str(x) for x in range(int(start), int(end) + 1)]
    if by_index:
        choices = sorted(list(set([int(x) - 1 for x in choices if '-' not in x])))
    else:
        choices = sorted(list(set([int(x) for x in choices if '-' not in x])))
    return choices


def filename_from_url(url):
    forbidden_chr = "<>:\"/\\|?*-"
    if 'http://' in url:
        ret = re.findall(r"http://(.+)", url)[0]
    else:
        ret = url
    for chr in forbidden_chr:
        ret = ret.replace(chr, '___')
    ret = 'Download/' + ret
    return ret

elements_dict = {0:"X",89:'Ac',47:'Ag',13:'Al',95:'Am',18:'Ar',33:'As',85:'At',79:'Au',5:'B',56:'Ba',4:'Be',107:'Bh',83:'Bi',97:'Bk',35:'Br',6:'C',20:'Ca',48:'Cd',58:'Ce',98:'Cf',17:'Cl',96:'Cm',112:'Cn',27:'Co',24:'Cr',55:'Cs',29:'Cu',105:'Db',110:'Ds',66:'Dy',68:'Er',99:'Es',63:'Eu',9:'F',26:'Fe',114:'Fl',100:'Fm',87:'Fr',31:'Ga',64:'Gd',32:'Ge',1:'H',2:'He',72:'Hf',80:'Hg',67:'Ho',108:'Hs',53:'I',49:'In',77:'Ir',19:'K',36:'Kr',57:'La',3:'Li',103:'Lr',71:'Lu',116:'Lv',101:'Md',12:'Mg',25:'Mn',42:'Mo',109:'Mt',7:'N',11:'Na',41:'Nb',60:'Nd',10:'Ne',28:'Ni',102:'No',93:'Np',8:'O',76:'Os',15:'P',91:'Pa',82:'Pb',46:'Pd',61:'Pm',84:'Po',59:'Pr',78:'Pt',94:'Pu',88:'Ra',37:'Rb',75:'Re',104:'Rf',111:'Rg',45:'Rh',86:'Rn',44:'Ru',16:'S',51:'Sb',21:'Sc',34:'Se',106:'Sg',14:'Si',62:'Sm',50:'Sn',38:'Sr',73:'Ta',65:'Tb',43:'Tc',52:'Te',90:'Th',22:'Ti',81:'Tl',69:'Tm',92:'U',118:'Uuo',115:'Uup',117:'Uus',113:'Uut',23:'V',74:'W',54:'Xe',39:'Y',70:'Yb',30:'Zn',40:'Zr'}

num_to_element_dict = elements_dict
temp1 = {value: key for key, value in elements_dict.items()}
temp2 = {value.lower(): key for key, value in elements_dict.items()}
temp3 = {value.upper(): key for key, value in elements_dict.items()}

element_to_num_dict = {key: value for key, value in list(temp1.items()) + list(temp2.items()) + list(temp3.items())}


def chr_is_chinese(char):
    return 0x4e00 <= ord(char) <= 0x9fa5


def has_chinese_char(string):
    for char in string:
        if chr_is_chinese(char):
            return True

    return False


def has_only_alphabat(string):
    for char in string.lower():
        if not ord('a') <= ord(char) <= ord('z'):
            return False
    return True


from threading import Thread
import functools


def mytimeout(timeout):
    def deco(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            res = [Exception('function [%s] timeout [%s seconds] exceeded!' % (func.__name__, timeout))]

            def newFunc():
                try:
                    res[0] = func(*args, **kwargs)
                except Exception as e:
                    res[0] = e

            t = Thread(target=newFunc)
            t.daemon = True
            try:
                t.start()
                t.join(timeout)
            except Exception as je:
                print('error starting thread')
                raise je
            ret = res[0]
            if isinstance(ret, BaseException):
                print("Timeout in mytimeout decorator", func)
            return ret

        return wrapper

    return deco
