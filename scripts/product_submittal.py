# product_submittal.py
import xlwings as xw
import pandas as pd
from scripts.DWG import DWG
from rename import rename
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfObject
from datetime import date
import os
"""
author: Sage Gendron
Master file to concatenate submittal drawings and spec sheets based on
unique drawing codes. Takes user input to index available drawings, asks
submittal refining questions and returns a full submittal file in a base
directory.
"""
# IMMUTABLE GLOBAL VARIABLES USED FOR EASE IN UPDATING; THIS IS NOT BEST PRACTICE
ordered_alphabet: list[str] = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R',
                               'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH',
                               'AI', 'AJ', 'AK', 'AL', 'AM', 'AN']

spec_loc: str = r'C:\Estimating\Specification Pages'
cover_page_loc: str = r'C:\Estimating\Submittal\Template Cover Page.pdf'

# locate directory for drawings based on type to speed up walk
dir_all: str = r'C:\Estimating\CAD Drawings\Kits'
dir_0: str = r'C:\Estimating\CAD Drawings\Kits\NO CV'
dir_2: str = r'C:\Estimating\CAD Drawings\Kits\2-WAY'
dir_3: str = r'C:\Estimating\CAD Drawings\Kits\3-WAY'
dir_l: str = r'C:\Estimating\CAD Drawings\Kits\LARGE SIZE'
dir_sc: str = r'C:\Estimating\CAD Drawings\Kits\Stacked Coil Kits'

job_name_cell: str = 'C19'
rep_name_cell: str = 'C23'

# BLOCK OF GLOBAL VARIABLES FOR controls()
v243_act: dict[str, str] = {'ON/OFF, FC': 'ME4430_4530.pdf', 'ON/OFF, FO': 'ME4430_4530.pdf',
                            'MOD, FLP': 'ME4140_4240_4340.pdf', 'MOD, FC': 'ME4940.pdf', 'MOD, FO': 'ME4840.pdf'}
v321_act: dict[str, str] = {'ON/OFF, FC': 'ME4430_4530.pdf', 'ON/OFF, FO': 'ME4430_4530.pdf',
                            'MOD, FLP': 'ME4140_4240_4340.pdf', 'MOD, FC': 'ME4840.pdf', 'MOD, FO': 'ME4940.pdf'}
v411_act: dict[str, dict[str, str]] = {
    'ON/OFF, FC': {'1"': 'ME5430_ME5630_ME5830.pdf', '1-1/2"': 'ME5430_ME5630_ME5830.pdf',
                   '2"': 'ME5440_5640_5840.pdf', '3"': 'ME5000_ON_SN.pdf', '4"': 'ME5000_ON_SN.pdf'},
    'ON/OFF, FO': {'1"': 'ME5430_ME5630_ME5830.pdf', '1-1/2"': 'ME5430_ME5630_ME5830.pdf',
                   '2"': '5440', '3"': '5850-ON', '4"': 'ME5000_ON_SN.pdf'},
    'MOD, FLP': {'1"': 'ME5130-ME5330.pdf', '1-1/2"': 'ME5130-ME5330.pdf', '2"': 'ME5140-ME5340.pdf',
                 '3"': 'ME5150-ME5350.pdf', '4"': 'ME5150-ME5350.pdf'},
    'MOD, FC': {'1"': 'ME5430_ME5630_ME5830.pdf', '1-1/2"': 'ME5430_ME5630_ME5830.pdf',
                '2"': 'ME5440_5640_5840.pdf', '3"': 'ME5000_ON_SN.pdf', '4"': 'ME5000_ON_SN.pdf'},
    'MOD, FO': {'1"': 'ME5430_ME5630_ME5830.pdf', '1-1/2"': 'ME5430_ME5630_ME5830.pdf',
                '2"': 'ME5440_5640_5840.pdf', '3"': 'ME5000_ON_SN.pdf', '4"': 'ME5000_ON_SN.pdf'}
    }
ninetwo_act: dict[str, str] = {'ON/OFF, FC': 'ME4430_4530.pdf', 'ON/OFF, FO': 'ME4430_4530.pdf',
                               'MOD, FLP': 'VA-7482-8002-RA.pdf', 'MOD, FC': 'ME4940.pdf', 'MOD, FO': 'ME4840.pdf'}
eightfive_act: dict[str, str] = {'MOD, FLP': 'VA9310-HGA-2.pdf', 'MOD, FC': 'VA9208-GGA-2.pdf',
                                 'MOD, FO': 'VA9208-GGA-2.pdf'}

cv_list: list[str] = ['+V2', '+V3', '+V4', '+BE', '+JCI']
picv_list: list[str] = ['91', '92', '93', '85', '94']


def generate_submittal():
    """
    Master submittal generation function. Pulls information required from schedule, creates DWG objects, calls functions
    to location pdf drawings, calls functions to select spec sheets, then calls a function to concatenate/save them to
    the submittal file.

    :returns: None - calls fx concat to write pdf file
    """
    wb: xw.Book = xw.Book.caller()

    # take dwg column as pandas dataframe object and send to list
    df: pd.DataFrame = pd.read_excel(wb.fullname, sheet_name='SCHEDULE', header=0, usecols='B:C,E:F,H:I,K,S,T,X,Z,AB',
                                     skiprows=31, nrows=1000)
    df.dropna(thresh=5, inplace=True)

    # pull pd.Series out of DataFrame as lists to iterate over
    eq_qty: list[int] = df['qty'].values.tolist()
    gpm: list[float] = df['gpm'].values.tolist()
    pkg_key: list[str] = df['pkg_key'].values.tolist()
    size: list[str] = df['pipe_size'].values.tolist()
    pipe_type: list[str] = df['pipe_type'].values.tolist()
    coil_type: list[str] = df['coil_type'].values.tolist()
    cv_pn: list[str] = df['cv_pn'].values.tolist()
    cv_size: list[str] = df['cv_size_type'].values.tolist()
    act_signal: list[str] = df['act_signal'].values.tolist()
    sch_dwg: list[str] = df['dwg'].values.tolist()
    ss: list[str] = df['ss'].values.tolist()

    # instantiate persistent variables
    sch_dict: dict[str, DWG] = {}
    list_initial: list[str] = []
    list_initial_lg: list[str] = []
    list_controls: list[str] = []

    # loop through all kits
    i: int
    for i in range(len(pkg_key)):
        # skip if row quantity is 0
        if eq_qty[i] == 0:
            continue
        pkg: str = pkg_key[i]
        # skip if package key empty/NaN
        if type(pkg) in (float, None):
            continue
        # skip if drawing empty/NaN
        if type(sch_dwg[i]) in (float, None):
            continue
        # skip if the package key has already been included in submittal
        if pkg in sch_dict.keys():
            continue

        # instantiate DWG object differently if cv included by HCI or cv by other
        if type(cv_size[i]) in (float, None):
            # create DWG object for each dwg type/pkg key
            dwg: DWG = DWG(sch_dwg[i], pkg, cv_pn[i], 'TBD', act_signal[i])
        else:
            # create DWG object for each dwg type/pkg key
            dwg: DWG = DWG(sch_dwg[i], pkg, cv_pn[i], cv_size[i].split()[0], act_signal[i])

        # if compact kit, mark DWG object as compact
        try:
            if size[i] == '3/4"' and type(gpm[i]) not in (None, str) and 0 < gpm[i] <= 3.1:
                dwg.setcompact()
        except TypeError:
            raise Exception('Please ensure GPM fields are blank or filled with numeric values. Save. Try again.')

        # if SS kit, mark DWG object as SS
        if ss[i] == 'YES':
            dwg.setss()
        # if size greater than 2", mark DWG object as large
        if size[i] in ('2 1/2"', '2-1/2"', '3"', '4"', '5"', '6"', '8"', '10"', '12"'):
            dwg.setlarge()
        # if sweat connections required, mark DWG object as sweat
        if pipe_type[i] == 'SWT':
            dwg.setsweat()
        # if press connections required, mark DWG object as press
        elif pipe_type[i] == 'PRESS' or coil_type[i] == 'PRESS':
            dwg.setpress()
        # if & in drawing, mark DWG object as stacked
        if '&' in dwg.name:
            dwg.setstacked()
        # if an additional isolation valve is required, mark DWG object as such
        if '+TI' in dwg.parts()[1] or '+DI' in dwg.parts()[1] or '+I' in dwg.parts()[1]:
            dwg.setiso()

        # add completed DWG object to dictionary with pkg_key as key
        sch_dict[pkg] = dwg

    # get filepaths to all drawings to be included in submittals
    sch_dict = find_dwg(sch_dict)

    # iterate through all pkg_key : dwg pairs in dictionary to identify spec sheets
    pkg: str
    dwg: DWG
    for pkg, dwg in sch_dict.items():
        # split dwg filename by '-' and '&'
        parts: list[str] = dwg.parts()

        # if no coil drawing, remove reference as it doesn't affect submittal
        if parts[0] == 'O':
            parts.pop(0)

        # if stacked coil
        if dwg.stacked:
            # if large x large
            if parts[0].startswith('L') and parts[2].startswith('L'):
                list_initial_lg = lst_lg(parts[:2], list_initial_lg)
                list_initial_lg = lst_lg(parts[2:4], list_initial_lg)
            # else if large x small
            elif parts[0].startswith('L') and not parts[2].startswith('L'):
                list_initial_lg = lst_lg(parts[:2], list_initial_lg)
                if dwg.press:
                    list_initial = lst_sp(parts[2:4], dwg.isolation, list_initial)
                elif dwg.ss:
                    list_initial = lst_ss(parts[2:4], dwg.compact, dwg.sweat, dwg.isolation, list_initial)
                else:
                    list_initial = lst_base(parts[2:4], dwg.compact, dwg.sweat, dwg.isolation, list_initial)
            # else small x small
            else:
                if dwg.press:
                    list_initial = lst_sp(parts[:2], dwg.isolation, list_initial)
                    list_initial = lst_sp(parts[2:4], dwg.isolation, list_initial)
                elif dwg.ss:
                    list_initial = lst_ss(parts[:2], dwg.compact, dwg.sweat, dwg.isolation, list_initial)
                    list_initial = lst_ss(parts[2:4], dwg.compact, dwg.sweat, dwg.isolation, list_initial)
                else:
                    list_initial = lst_base(parts[:2], dwg.compact, dwg.sweat, dwg.isolation, list_initial)
                    list_initial = lst_base(parts[2:4], dwg.compact, dwg.sweat, dwg.isolation, list_initial)

        # if large size kit
        elif parts[0].startswith('L'):
            list_initial_lg = lst_lg(parts, list_initial_lg)
            # list_acc = lst2_L(parts, list_acc)  # commented out as not including accessory specs at this time

        # if small size kit
        else:
            if dwg.press:
                list_initial = lst_sp(parts, dwg.isolation, list_initial)
            elif dwg.ss:
                list_initial = lst_ss(parts, dwg.compact, dwg.sweat, dwg.isolation, list_initial)
            else:
                list_initial = lst_base(parts, dwg.compact, dwg.sweat, dwg.isolation, list_initial)
            # list_acc = lst2(parts, list_acc)  # commented out as not including accessory specs at this time

        # if dwg object flagged to include controls
        if type(dwg.ctrl_model) not in (None, float):
            temp_controls = controls(parts[1], dwg.ctrl_model, dwg.ctrl_size, dwg.act_signal)

            # de-duplicate control valve spec pages
            i: int
            for i in range(len(temp_controls)):
                if temp_controls[i] not in list_controls:
                    list_controls.append(temp_controls[i])

    # concatenate all lists into a single list for pdf merging
    list_final: list[str] = list_initial + list_initial_lg + list_controls

    concat(wb, sch_dict, list_final)


def find_dwg(sch_dict):
    """
    Generate filepaths for drawings in dwg_list

    :param dict sch_dict: dict of deduplicated drawings from schedule as keys, package keys as values
    :return: sch_dict (:py:class:'dict') - dict of deduplicated drawings from schedule as keys, package keys as values
    :rtype: dict
    """
    # walk through folders to find drawing, append file path to list
    pkg: str
    dwg: DWG
    for pkg, dwg in sch_dict.items():
        name: str = dwg.name

        # check to ensure filetype is included from dwg column on schedule
        if name[-4:] != '.pdf':
            name = f"{name}.pdf"

        # if the first characters are 'O-'
        if name.startswith('O-'):
            name = name[2:]

        # loop through subfolders and files therein to find drawing location
        path: str
        subdir: list[str]
        files: list[str]
        # if no cv kit
        if name[0] in 'YNIO':
            for path, subdir, files in os.walk(dir_0):
                for file in files:
                    if file == dwg.name:
                        dwg.fpath = os.path.join(path, file)
        # if two-way kit
        elif name[0] == '2' and '&' not in name:
            for path, subdir, files in os.walk(dir_2):
                for file in files:
                    if file == dwg.name:
                        dwg.fpath = os.path.join(path, file)
        # if three-way kit
        elif name[0] == '3' and '&' not in name:
            for path, subdir, files in os.walk(dir_3):
                for file in files:
                    if file == dwg.name:
                        dwg.fpath = os.path.join(path, file)
        # if large-size kit
        elif name[0] == 'L' and '&' not in name:
            for path, subdir, files in os.walk(dir_l):
                for file in files:
                    if file == name:
                        dwg.fpath = os.path.join(path, file)
        # if stacked kit
        elif '&' in name:
            for path, subdir, files in os.walk(dir_sc):
                for file in files:
                    if file == name:
                        dwg.fpath = os.path.join(path, file)
        # otherwise walk all directories to look for the dwg code
        else:
            for path, subdir, files in os.walk(dir_all):
                for file in files:
                    if file == dwg.name:
                        dwg.fpath = os.path.join(path, file)

    return sch_dict


def lst_lg(dwg_parts, list_initial_lg):
    """
    Parses through large drawings to identify spec pages required. Always includes a butterfly valve spec sheet.

    :param list dwg_parts: smart kit code (drawing name) split by hyphens
    :param list list_initial_lg: list of spec sheets (literal strings) identified for this large-sized kit
    :return: list_initial_L (:py:class:'list') - list of spec sheets (literal strings) identified for this large-sized kit
    :rtype: list
    """
    y: str = 'TS-LF.pdf'
    i: str = 'BFV.pdf'
    b: str = 'TB-LF.pdf'
    g: str = 'TGV-Flanged.pdf'
    a: str = 'TA-W.pdf'
    h: str = 'HF.pdf'
    # loop through first portion of drawing code (ie L2YOBO) to grab spec sheet filenames
    char: str
    for char in dwg_parts[0]:
        if char == 'Y' and y not in list_initial_lg:
            list_initial_lg.append(y)
        elif char == 'I' and i not in list_initial_lg:
            list_initial_lg.append(i)
        elif char == 'B' and b not in list_initial_lg:
            list_initial_lg.append(b)
        elif char == 'G' and g not in list_initial_lg:
            list_initial_lg.append(g)
        elif char == 'A' and a not in list_initial_lg:
            list_initial_lg.append(a)
    # check to ensure butterfly valve spec included if not already
    if i not in list_initial_lg:
        list_initial_lg.append(i)
    # check for flex hoses
    if dwg_parts[1][1] == 'H' and h not in list_initial_lg:
        list_initial_lg.append(h)
    return list_initial_lg


def lst_base(dwg_parts, is_comp, is_swt, ex_iso, list_initial):
    """
    author: Sage Gendron \n
    Loops through characters in the first half of the smart kit code to come up with a list of spec sheets required.

    :param list dwg_parts: smart kit code (drawing name) split by hyphens
    :param bool is_comp: is this kit compact?
    :param bool is_swt: does this kit have sweat end(s)? (primarily for loose isolation valves)
    :param bool ex_iso: does this kit require external isolation valves? (vs. integral isolation)
    :param list list_initial: list of spec sheets (literal strings) identified for this typical kit type
    :returns: list_initial (:py:class:'list') - list of spec sheets (literal strings) identified for this typical kit type
    :rtype: list
    """
    y: str = 'TY.pdf'
    y_c: str = 'TY1.pdf'
    i_s: str = 'JF-100SG.pdf'
    i_t: list[str] = ['JF-100GMxF.pdf', 'JF-100TG.pdf']
    u: str = 'TU.pdf'
    b: str = 'TB.pdf'
    g: str = 'TGV.pdf'
    a: str = 'TA.pdf'
    a_c: str = 'TA1.pdf'
    n: str = 'NT.pdf'
    h: str = 'HC.pdf'
    # loop through first portion of drawing code (ie 2YUBO) to grab spec sheet filenames
    char: str
    for char in dwg_parts[0]:
        if char == 'Y':
            # check for compact
            if is_comp and y_c not in list_initial:
                list_initial.append(y_c)
            elif not is_comp and y not in list_initial:
                list_initial.append(y)
        elif char == 'I':
            # check for sweat
            if is_swt and i_s not in list_initial:
                list_initial.append(i_s)
            # check for threaded
            elif i_t[0] not in list_initial and i_t[1] not in list_initial:
                list_initial.extend(i_t)
        elif char == 'U' and u not in list_initial:
            list_initial.append(u)
        elif char == 'B' and b not in list_initial:
            list_initial.append(b)
        elif char == 'G' and g not in list_initial:
            list_initial.append(g)
        elif char == 'A':
            # check for compact
            if is_comp and a_c not in list_initial:
                list_initial.append(a_c)
            elif not is_comp and a not in list_initial:
                list_initial.append(a)
        elif char == 'N' and n not in list_initial:
            list_initial.append(n)
    # check for flex hoses
    if dwg_parts[1][2] == 'H' and h not in list_initial:
        list_initial.append(h)
    # check for external ball valves
    if ex_iso:
        # check for sweat
        if is_swt and i_s not in list_initial:
            list_initial.append(i_s)
        # check for threaded
        elif i_t[0] not in list_initial and i_t[1] not in list_initial:
            list_initial.extend(i_t)
    return list_initial


def lst_sp(dwg_parts, ex_iso, list_initial):
    """
    author: Sage Gendron \n
    Loops through characters in the first half of the smart kit code to come up with a list of spec sheets required.
    Specifically called if press is one of the ends to the kit.

    :param list dwg_parts: smart kit code (drawing name) split by hyphens
    :param bool ex_iso: does this kit require external isolation valves? (vs. integral isolation)
    :param list list_initial: list of spec sheets (literal strings) identified for this press kit type
    :return: list_initial (:py:class:'list') - list of spec sheets (literal strings) identified for this press kit type
    :rtype: list
    """
    y: str = 'TY-P.pdf'
    u: str = 'TU-P.pdf'
    b: str = 'TB-P.pdf'
    g: str = 'TGV.pdf'
    a: str = 'TA-P.pdf'
    n: str = 'NT-P.pdf'
    i: str = '100_PXP.pdf'
    h: str = 'HC.pdf'
    # loop through first portion of drawing code (ie 2YUBO) to grab spec sheet filenames
    char: str
    for char in dwg_parts[0]:
        if char == 'Y' and y not in list_initial:
            list_initial.append(y)
        elif char == 'U' and u not in list_initial:
            list_initial.append(u)
        elif char == 'B' and b not in list_initial:
            list_initial.append(b)
        elif char == 'G' and g not in list_initial:
            list_initial.append(g)
        elif char == 'A' and a not in list_initial:
            list_initial.append(a)
        elif char == 'N' and n not in list_initial:
            list_initial.append(n)
        elif char == 'I' and i not in list_initial:
            list_initial.append(i)
    # check for flex hoses
    if dwg_parts[1][2] == 'H' and h not in list_initial:
        list_initial.append(h)
    # check for external ball valves
    if ex_iso and i not in list_initial:
        list_initial.append(i)
    return list_initial


def lst_ss(dwg_parts, is_comp, is_swt, ex_iso, list_initial):
    """
    author: Sage Gendron \n
    Grab spec sheet names (literal strings) as required based on components in first half of the smart kit code.
    Specifically called if the kit is flagged SS in the schedule.

    :param list dwg_parts: smart kit code (drawing name) split by hyphens
    :param bool is_comp: is this kit compact?
    :param bool is_swt: does this kit have sweat end(s)? (primarily for loose isolation valves)
    :param bool ex_iso: does this kit require external isolation valves? (vs. integral isolation)
    :param list list_initial: list of spec sheets (literal strings) identified for this SS trim kit type
    :return: list_initial (:py:class:'list') - list of spec sheets (literal strings) identified for this SS trim kit type
    :rtype: list
    """
    y: str = 'TY-SS.pdf'
    y_c: str = 'TY1-SS.pdf'
    i_s: str = 'JF-100SG.pdf'
    i_t: list[str] = ['JF-100GMxF.pdf', 'JF-100TG.pdf']
    u: str = 'TU.pdf'
    b: str = 'TB-SS.pdf'
    g: str = 'TGV.pdf'
    a: str = 'TA-SS.pdf'
    a_c: str = 'TA1-SS.pdf'
    n: str = 'NT-SS.pdf'
    h: str = 'HC.pdf'
    # loop through first portion of drawing code (ie 2YUBO) to grab spec sheet filenames
    char: str
    for char in dwg_parts[0]:
        if char == 'Y':
            # check for compact
            if is_comp and y_c not in list_initial:
                list_initial.append(y_c)
            elif not is_comp and y not in list_initial:
                list_initial.append(y)
        elif char == 'I':
            # check for sweat
            if is_swt and i_s not in list_initial:
                list_initial.append(i_s)
            elif i_t[0] not in list_initial and i_t[1] not in list_initial:
                list_initial.extend(i_t)
        elif char == 'U' and u not in list_initial:
            list_initial.append(u)
        elif char == 'B' and b not in list_initial:
            list_initial.append(b)
        elif char == 'G' and g not in list_initial:
            list_initial.append(g)
        elif char == 'A':
            # check for compact
            if is_comp and a_c not in list_initial:
                list_initial.append(a_c)
            elif not is_comp and a not in list_initial:
                list_initial.append(a)
        elif char == 'N' and n not in list_initial:
            list_initial.append(n)
    # check for flex hoses
    if dwg_parts[1][2] == 'H' and h not in list_initial:
        list_initial.append(h)
    # check for external ball valves
    if ex_iso:
        # check for sweat
        if is_swt and i_s not in list_initial:
            list_initial.append(i_s)
        # check for threaded
        elif i_t[0] not in list_initial and i_t[1] not in list_initial:
            list_initial.extend(i_t)
    return list_initial


def controls(dwg_suffix, cv_pn, cv_size, act_signal):
    """
    author: Sage Gendron
    Identify if a control valve or PICV called for by drawing name. If called out, add correct control valve type and
    actuator for that control valve based on actuator signal and control valve part number indicated on schedule.

    :param str dwg_suffix: accessory/alt/add portion of the drawing code
    :param str cv_pn: control valve part number from cv_pn cell in this particular row from schedule
    :param str cv_size: control valve size from cv_size_type cell in this particular row from schedule
    :param str act_signal: actuator signal from act_signal cell in this particular row from schedule
    :return: list_controls (:py:class:'list') - a list of literal strings indicating spec sheet names for controls parts
    :rtype: list
    """
    cv_check: bool = False
    picv_check: bool = False
    list_controls: list[str] = []

    # loop through cv_list and picv_list to identify if a cv or picv indicated by dwg_suffix
    i: int = 0
    while i < len(cv_list) and cv_check is False and picv_check is False:
        if cv_list[i] in dwg_suffix:
            cv_check = True
        if picv_list[i] in dwg_suffix:
            picv_check = True
        i += 1

    # if a control valve is found called out in dwg_suffix (cv_list)
    if cv_check:
        if cv_pn.startswith('V243'):
            list_controls.extend(['V243.pdf', v243_act[act_signal]])
        elif cv_pn.startswith('V321'):
            list_controls.extend(['V321.pdf', v321_act[act_signal]])
        elif cv_pn.startswith('VE'):
            list_controls.extend(['V411_V431.pdf', v411_act[act_signal][cv_size]])
    # if a PICV is found called out in dwg_suffix (picv_list)
    elif picv_check:
        if cv_pn.startswith('T92'):
            list_controls.extend(['PICV-92.pdf', ninetwo_act[act_signal]])
        elif cv_pn.startswith('T85'):
            list_controls.extend(['PICV-85.pdf', eightfive_act[act_signal]])
        elif cv_pn.startswith('T94'):
            list_controls.extend(['PICV-94FA.pdf', 'PICV-94F Actuator.pdf'])

    return list_controls


def concat(wb, dwg_dict, list_final):
    """
    author: Sage Gendron \n
    Creates the pdf container via pdf_cover_page(), then grabs all required drawing filepaths, then appends spec sheets
    from all the list_initial/list_controls spec sheet lists. Finally, writes the pdf file to the active folder.

    :param xlwings.Book wb: the project file generate_submittal is being called from
    :param dict dwg_dict: dictionary of dwgs as keys and values: [package key, dwg filepath]
    :param list list_final: list of accessory filepaths in order of occurrence
    :return: n/a - file is written
    """
    # name file
    target: str = rename(wb, 'SUBMITTAL', 'pdf')

    # generate cover page using the schedule and create the submittal pdf merger object with (1) pdf contained
    merger, acro_form = pdf_cover_page(wb)

    # instantiate list for ordered drawings
    dwg_list: list[None] = [None] * 40

    # loop through dwg_dict to sort dwgs by package key
    pkg: str
    dwg: DWG
    for pkg, dwg in dwg_dict.items():
        # handle pkg_keys greater than Z (AA-AN)
        if len(pkg) > 1:
            i: int = ord(pkg[1]) - ord('A') + 26
        # deduct the numeric difference between the acting pkg_key and capital 'A' (if A-Z)
        else:
            i: int = ord(pkg) - ord('A')
        # setting the filepath indexed into a list in package key order
        try:
            if dwg.fpath not in dwg_list:
                dwg_list[i] = dwg.fpath
        # if somehow the index didn't get created or >40, append the drawing filepath to the end of the list
        except IndexError:
            if dwg.fpath not in dwg_list:
                dwg_list.append(dwg.fpath)

    # add pdf dwg filepath to merger in order of package keys, skipping any letters with no unique dwg
    dwg: str
    for dwg in dwg_list:
        if dwg is None:
            continue
        merger.addpages(PdfReader(dwg).pages)

    # add accessories in order they were added to list
    pdf: str
    for pdf in list_final:
        merger.addpages(PdfReader(f"{spec_loc}\\{pdf}").pages)

    merger.trailer.Root.AcroForm = acro_form

    # write pdf file to folder
    merger.write(target)


def pdf_cover_page(wb_sch):
    """
    Grabs and fills out the submittal cover page with information from the schedule.

    :param xlwings.Book wb_sch: the project file generate_submittal is being called from
    :return:
        - merger (:)
        - acro_form (:pdfrw:class:PdfWriter:)
    :rtype: ( , PdfWriter)
    """
    ANNOT_KEY = '/Annots'
    ANNOT_FIELD_KEY = '/T'
    ANNOT_VAL_KEY = '/V'
    ANNOT_RECT_KEY = '/Rect'
    SUBTYPE_KEY = '/Subtype'
    WIDGET_SUBTYPE_KEY = '/Widget'

    # pull out quote number from filepath
    quote_number: str = wb_sch.fullname.split('//')[-1].split('_')[-1][:-5]

    # dict pertaining to fields developed in Adobe
    data_dict: dict[str, str] = {
        'Job Name': wb_sch.sheets['SCHEDULE'].range(job_name_cell).value,
        'Buy Sell Rep': wb_sch.sheets['SCHEDULE'].range(rep_name_cell).value,
        'Date': date.today().strftime('%m-%d-%Y'),
        'Quote Number': quote_number
    }

    # instantiate pdfrw object for field fill out
    template_pdf: PdfReader = PdfReader(cover_page_loc)

    # iterating in a loop of 1
    i: int = 0
    for page in template_pdf.pages:
        if i >= 1:
            break
        annotations = page[ANNOT_KEY]
        for annotation in annotations:
            if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
                if annotation[ANNOT_FIELD_KEY]:
                    key = annotation[ANNOT_FIELD_KEY][1:-1]
                    if key in data_dict.keys():
                        # if we include check boxes, this will have to be changed
                        if type(data_dict[key]) == bool:
                            pass
                        else:
                            annotation.update(
                                PdfDict(V='{}'.format(data_dict[key])))
                            annotation.update(PdfDict(AP=''))
                            # annotation.update(PdfDict(Ff=1))  ## multi-line doesn't have support for read-only text
                            i += 1
    # makes the fields appear
    template_pdf.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject('true')))
    acro_form = template_pdf.Root.AcroForm

    # write pdf to file
    merger: PdfWriter = PdfWriter()
    merger.addpages(template_pdf.pages)

    return merger, acro_form
