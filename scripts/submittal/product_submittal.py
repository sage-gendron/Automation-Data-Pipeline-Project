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
Master file to concatenate submittal drawings and spec sheets based on unique drawing codes. Takes user input to index 
available drawings, asks submittal refining questions and returns a full submittal file in a base directory.
"""
# IMMUTABLE GLOBAL VARIABLES USED FOR EASE IN UPDATING; THIS IS NOT BEST PRACTICE
ordered_keys: list[str] = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R',
                           'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI',
                           'AJ', 'AK', 'AL', 'AM', 'AN']

spec_loc: str = r'C:\Estimating\Specification Pages'
cover_page_loc: str = r'C:\Estimating\Submittal\Template Cover Page.pdf'

# locate directory for drawings based on type to speed up walk
dir_all: str = r'C:\Estimating\CAD Drawings\Kits'
dir_1: str = r'C:\Estimating\CAD Drawings\Kits\TYPE 1'
dir_2: str = r'C:\Estimating\CAD Drawings\Kits\TYPE 2'
dir_3: str = r'C:\Estimating\CAD Drawings\Kits\TYPE 3'
dir_l: str = r'C:\Estimating\CAD Drawings\Kits\LARGE SIZE'

job_name_cell: str = 'C19'
co_name_cell: str = 'C23'

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


def build_dwgs(df):
    """
    www

    :param pd.DataFrame df:
    :return: sch_dict -
    :rtype: dict
    """
    sch_dict: dict[str, DWG] = {}

    # loop through all rows on schedule
    i: int
    for index, row in df.iterrows():
        # skip if row quantity is 0
        if row['qty'] == 0:
            continue
        # skip if package key empty/NaN or if package key has already been included in submittal
        if type(pkg := row['pkg_key']) in (float, None) or pkg in sch_dict:
            continue
        # skip if drawing empty/NaN
        if type(dwg := row['dwg']) in (float, None):
            continue


        # instantiate DWG object differently if control part by us or by other
        if type(ctrl_size := row['control_size_type']) in (float, None):
            dwg: DWG = DWG(dwg, pkg, row['control_pt'], 'TBD', row['Signal'])
        else:
            dwg: DWG = DWG(dwg, pkg, row['control_pt'], ctrl_size.split()[0], row['Signal'])

        # if marked as small, mark DWG object as small
        try:
            if size := row['size'] == 'SMALL' and type(gpm := row['gpm']) not in (None, str) and 0 < gpm <= 3.1:
                dwg.set_sm()
        except TypeError:
            raise Exception('Please ensure GPM fields are blank or filled with numeric values. Save. Try again.')

        # if marked as large, mark DWG object as large
        if size == 'LARGE':
            dwg.set_lg()
        # if special case 1 is required, mark DWG as such
        if row['sp_case_1'] == 'YES':
            dwg.set_sp_case_1()
        # if special case 2 is required, mark DWG as such
        elif row['sys_type'] == 'SP_CASE_2' or row['conn_type'] == 'SP_CASE_2':
            dwg.set_sp_case_2()

        # add completed DWG object to dictionary with pkg_key as key
        sch_dict[pkg] = dwg

    return sch_dict


def generate_submittal():
    """
    Master submittal generation function. Pulls information required from schedule, creates DWG objects, calls functions
    to location pdf drawings, calls functions to select spec sheets, then calls a function to concatenate/save them to
    the submittal file.

    :return: None - calls fx concat to write pdf file
    """
    # instantiate Book instance to interact with Excel
    wb: xw.Book = xw.Book.caller()

    # take dwg column as pandas dataframe object and send to list
    df: pd.DataFrame = pd.read_excel(wb.fullname, sheet_name='SCHEDULE', header=0, usecols='B:C,E:F,H:I,K,S,T,X,Z,AB',
                                     skiprows=31, nrows=1000)
    df.dropna(thresh=5, inplace=True)

    # build dwg objects and map to package keys
    sch_dict = build_dwgs(df)

    # get filepaths to all drawings to be included in submittals
    sch_dict = find_dwgs(sch_dict)

    # # pull pd.Series out of DataFrame as lists to iterate over
    # row_qty: list[int] = df['qty'].values.tolist()
    # gpm: list[float] = df['gpm'].values.tolist()
    # pkg_key: list[str] = df['pkg_key'].values.tolist()
    # size: list[str] = df['size'].values.tolist()
    # sys_type: list[str] = df['sys_type'].values.tolist()
    # conn_type: list[str] = df['conn_type'].values.tolist()
    # control_pt: list[str] = df['control_pt'].values.tolist()
    # control_size: list[str] = df['control_size_type'].values.tolist()
    # act_signal: list[str] = df['act_signal'].values.tolist()
    # sch_dwg: list[str] = df['dwg'].values.tolist()
    # sp_case_1: list[str] = df['sp_case_1'].values.tolist()

    # instantiate persistent variables
    spec_list: list[str] = []
    spec_list_lg: list[str] = []
    controls_list: list[str] = []

    # iterate through all pkg_key : dwg pairs in dictionary to identify spec sheets
    pkg: str
    dwg: DWG
    for dwg in sch_dict.values():
        # split dwg filename by '-' and '&'
        parts: list[str] = dwg.parts()

        # large size doesn't have special cases
        if dwg.lg:
            spec_list_lg = lst_lg(parts, spec_list_lg)

        else:
            if dwg.sp_case_2:
                spec_list = sp_case_2(parts, spec_list)
            elif dwg.sp_case_1:
                spec_list = sp_case_1(parts, dwg.sm, spec_list)
            else:
                spec_list = lst_base(parts, dwg.sm, spec_list)

        # if dwg object flagged to include controls
        if type(dwg.ctrl_model) not in (None, float):
            controls_list = controls(parts[1], dwg.ctrl_model, dwg.ctrl_size, dwg.ctrl_signal)

    # concatenate all lists into a single list for pdf merging
    final_spec_list: list[str] = spec_list + spec_list_lg + controls_list

    concat(wb, sch_dict, final_spec_list)


def find_dwgs(sch_dict):
    """
    Generate filepaths for drawings in sch_dict

    :param dict sch_dict: dict of deduplicated drawings from schedule as values, package keys as keys
    :return: sch_dict - dict of deduplicated drawings from schedule as values, package keys as keys
    :rtype: dict
    """
    # walk through folders to find drawing, append file path to list
    dwg: DWG
    for dwg in sch_dict.values():
        name: str = dwg.name

        # check to ensure filetype is included from dwg column on schedule to counteract manually added drawing codes
        if name[-4:] != '.pdf':
            name = f"{name}.pdf"

        # loop through subfolders and files therein to find drawing location
        path: str
        subdir: list[str]
        files: list[str]
        # if no cv kit
        if name[0] not in '23L':
            for path, subdir, files in os.walk(dir_1):
                for file in files:
                    if file == dwg.name:
                        dwg.fpath = os.path.join(path, file)
        # if type 2 drawing
        elif name[0] == '2':
            for path, subdir, files in os.walk(dir_2):
                for file in files:
                    if file == dwg.name:
                        dwg.fpath = os.path.join(path, file)
        # if type 3 drawing
        elif name[0] == '3':
            for path, subdir, files in os.walk(dir_3):
                for file in files:
                    if file == dwg.name:
                        dwg.fpath = os.path.join(path, file)
        # if large-size drawing
        elif name[0] == 'L':
            for path, subdir, files in os.walk(dir_l):
                for file in files:
                    if file == name:
                        dwg.fpath = os.path.join(path, file)
        # otherwise walk all directories to look for the drawing
        else:
            for path, subdir, files in os.walk(dir_all):
                for file in files:
                    if file == dwg.name:
                        dwg.fpath = os.path.join(path, file)

    return sch_dict


def lst_lg(dwg_parts, spec_list_lg):
    """
    Parses through large drawings to identify spec pages required. Always includes a butterfly valve spec sheet.

    :param list dwg_parts: smart kit code (drawing name) split by hyphens
    :param list spec_list_lg: list of spec sheets (literal strings) identified for this large-sized kit
    :return: spec_list_lg - list of spec sheets (literal strings) identified for this large-sized kit
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
        if char == 'Y' and y not in spec_list_lg:
            spec_list_lg.append(y)
        elif char == 'I' and i not in spec_list_lg:
            spec_list_lg.append(i)
        elif char == 'B' and b not in spec_list_lg:
            spec_list_lg.append(b)
        elif char == 'G' and g not in spec_list_lg:
            spec_list_lg.append(g)
        elif char == 'A' and a not in spec_list_lg:
            spec_list_lg.append(a)
    # check to ensure butterfly valve spec included if not already
    if i not in spec_list_lg:
        spec_list_lg.append(i)
    # check for flex hoses
    if dwg_parts[1][1] == 'H' and h not in spec_list_lg:
        spec_list_lg.append(h)
    return spec_list_lg


def lst_base(dwg_parts, is_sm, spec_list):
    """
    Loops through characters in the first half of the smart kit code to come up with a list of spec sheets required.

    :param list dwg_parts: smart kit code (drawing name) split by hyphens
    :param bool is_sm: is this kit compact?
    :param list spec_list: list of spec sheets (literal strings) identified for this typical kit type
    :returns: list_initial - list of spec sheets (literal strings) identified for this typical kit type
    :rtype: list
    """
    y: str = 'TY.pdf'
    y_c: str = 'TY1.pdf'
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
            # check for small size
            if is_sm and y_c not in spec_list:
                spec_list.append(y_c)
            elif not is_sm and y not in spec_list:
                spec_list.append(y)
        elif char == 'U' and u not in spec_list:
            spec_list.append(u)
        elif char == 'B' and b not in spec_list:
            spec_list.append(b)
        elif char == 'G' and g not in spec_list:
            spec_list.append(g)
        elif char == 'A':
            # check for small size
            if is_sm and a_c not in spec_list:
                spec_list.append(a_c)
            elif not is_sm and a not in spec_list:
                spec_list.append(a)
        elif char == 'N' and n not in spec_list:
            spec_list.append(n)
    # check for flex hoses
    if dwg_parts[1][2] == 'H' and h not in spec_list:
        spec_list.append(h)

    return spec_list


def sp_case_2(dwg_parts, spec_list):
    """
    Loops through characters in the first half of the smart kit code to come up with a list of spec sheets required.
    Specifically called if press is one of the ends to the kit.

    :param list dwg_parts: smart kit code (drawing name) split by hyphens
    :param list spec_list: list of spec sheets (literal strings) identified for this press kit type
    :return: spec_list - list of spec sheets (literal strings) identified for this press kit type
    :rtype: list
    """
    y: str = 'TY-P.pdf'
    u: str = 'TU-P.pdf'
    b: str = 'TB-P.pdf'
    g: str = 'TGV.pdf'
    a: str = 'TA-P.pdf'
    n: str = 'NT-P.pdf'
    h: str = 'HC.pdf'
    # loop through first portion of drawing code (ie 2YUBO) to grab spec sheet filenames
    char: str
    for char in dwg_parts[0]:
        if char == 'Y' and y not in spec_list:
            spec_list.append(y)
        elif char == 'U' and u not in spec_list:
            spec_list.append(u)
        elif char == 'B' and b not in spec_list:
            spec_list.append(b)
        elif char == 'G' and g not in spec_list:
            spec_list.append(g)
        elif char == 'A' and a not in spec_list:
            spec_list.append(a)
        elif char == 'N' and n not in spec_list:
            spec_list.append(n)
    # check for flex hoses
    if dwg_parts[1][2] == 'H' and h not in spec_list:
        spec_list.append(h)
    return spec_list


def sp_case_1(dwg_parts, is_sm, spec_list):
    """
    Grab spec sheet names (literal strings) as required based on components in first half of the smart kit code.
    Specifically called if the kit is flagged SS in the schedule.

    :param list dwg_parts: smart kit code (drawing name) split by hyphens
    :param bool is_sm: is this kit compact?
    :param list spec_list: list of spec sheets (literal strings) identified for this SS trim kit type
    :return: spec_list - list of spec sheets (literal strings) identified for this SS trim kit type
    :rtype: list
    """
    y: str = 'TY-SS.pdf'
    y_c: str = 'TY1-SS.pdf'
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
            # check for small size
            if is_sm and y_c not in spec_list:
                spec_list.append(y_c)
            elif not is_sm and y not in spec_list:
                spec_list.append(y)
        elif char == 'U' and u not in spec_list:
            spec_list.append(u)
        elif char == 'B' and b not in spec_list:
            spec_list.append(b)
        elif char == 'G' and g not in spec_list:
            spec_list.append(g)
        elif char == 'A':
            # check for small size
            if is_sm and a_c not in spec_list:
                spec_list.append(a_c)
            elif not is_sm and a not in spec_list:
                spec_list.append(a)
        elif char == 'N' and n not in spec_list:
            spec_list.append(n)
    # check for flex hoses
    if dwg_parts[1][2] == 'H' and h not in spec_list:
        spec_list.append(h)
    return spec_list


def controls(dwg_suffix, cv_pn, cv_size, act_signal):
    """
    Identify if control type 1 or 2 called for by drawing name. If called out, add correct control part for that control
    type based on signal and control part number indicated one engineered schedule.

    :param str dwg_suffix: accessory/alt/add portion of the drawing code
    :param str cv_pn: control valve part number from cv_pn cell in this particular row from schedule
    :param str cv_size: control valve size from cv_size_type cell in this particular row from schedule
    :param str act_signal: actuator signal from act_signal cell in this particular row from schedule
    :return: controls_list - a list of literal strings indicating spec sheet names for controls parts
    :rtype: list
    """
    controls_list: list[str] = []

    # if a control valve is found called out in dwg_suffix (cv_list)
    if '+CTRL_1' in dwg_suffix:
        if cv_pn.startswith('V243'):
            controls_list.extend(['V243.pdf', v243_act[act_signal]])
        elif cv_pn.startswith('V321'):
            controls_list.extend(['V321.pdf', v321_act[act_signal]])
        elif cv_pn.startswith('VE'):
            controls_list.extend(['V411_V431.pdf', v411_act[act_signal][cv_size]])
    # if a PICV is found called out in dwg_suffix (picv_list)
    elif '+CTRL_2' in dwg_suffix:
        if cv_pn.startswith('T92'):
            controls_list.extend(['PICV-92.pdf', ninetwo_act[act_signal]])
        elif cv_pn.startswith('T85'):
            controls_list.extend(['PICV-85.pdf', eightfive_act[act_signal]])
        elif cv_pn.startswith('T94'):
            controls_list.extend(['PICV-94FA.pdf', 'PICV-94F Actuator.pdf'])

    return list(set(controls_list))


def concat(wb, dwg_dict, spec_list):
    """
    Creates the pdf container via pdf_cover_page(), then grabs all required drawing filepaths, then appends spec sheets
    from all the list_initial/list_controls spec sheet lists. Finally, writes the pdf file to the active folder.

    :param xw.Book wb: the project file generate_submittal is being called from
    :param dict dwg_dict: dictionary of dwgs as keys and values: [package key, dwg filepath]
    :param list spec_list: list of accessory filepaths in order of occurrence
    :return: None - file is written
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
    for pdf in spec_list:
        merger.addpages(PdfReader(f"{spec_loc}\\{pdf}").pages)

    merger.trailer.Root.AcroForm = acro_form

    # write pdf file to folder
    merger.write(target)


def pdf_cover_page(wb_sch):
    """
    Grabs and fills out the submittal cover page with information from the schedule.

    :param xw.Book wb_sch: the project file generate_submittal is being called from
    :return:
        - merger -
        - acro_form -
    :rtype: (PdfWriter, )
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
        'Buy Sell Rep': wb_sch.sheets['SCHEDULE'].range(co_name_cell).value,
        'Date': date.today().strftime('%m-%d-%Y'),
        'Quote Number': quote_number
    }

    # instantiate pdfrw object for field fill out
    template_pdf: PdfReader = PdfReader(cover_page_loc)

    # iterating in a loop of 1
    i: int = 0
    page:
    for page in template_pdf.pages:
        if i >= 1:
            break
        annotations = page[ANNOT_KEY]
        for annotation in annotations:
            if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
                if annotation[ANNOT_FIELD_KEY]:
                    key = annotation[ANNOT_FIELD_KEY][1:-1]
                    if key in data_dict.keys():
                        annotation.update(
                            PdfDict(V='{}'.format(data_dict[key])))
                        annotation.update(PdfDict(AP=''))
                        i += 1
    # makes the fields appear
    template_pdf.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject('true')))
    acro_form:  = template_pdf.Root.AcroForm

    # write pdf to file
    merger: PdfWriter = PdfWriter()
    merger.addpages(template_pdf.pages)

    return merger, acro_form
