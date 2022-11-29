# product_submittal.py
import xlwings as xw
import pandas as pd
from scripts.utils.rename import rename
from pdfrw import PdfReader, PdfWriter
import os
from determine_spec import lg_spec, typ_spec, sp_case_1, sp_case_2, controls
from DWG import DWG
"""
author: Sage Gendron
Master file to concatenate submittal drawings and spec sheets based on unique drawing codes. Takes user input to index 
available drawings, asks submittal refining questions and returns a full submittal file in a base directory.
"""
# IMMUTABLE GLOBAL VARIABLES USED FOR EASE IN UPDATING; THIS IS NOT BEST PRACTICE
spec_loc: str = r'C:\Estimating\Specification Pages'

# locate directory for drawings based on type to speed up walk
dir_all: str = r'C:\Estimating\CAD Drawings\Kits'
dir_1: str = r'C:\Estimating\CAD Drawings\Kits\TYPE 1'
dir_2: str = r'C:\Estimating\CAD Drawings\Kits\TYPE 2'
dir_3: str = r'C:\Estimating\CAD Drawings\Kits\TYPE 3'
dir_l: str = r'C:\Estimating\CAD Drawings\Kits\LARGE SIZE'

job_name_cell: str = 'C19'
co_name_cell: str = 'C23'


def build_dwgs(df):
    """
    Builds drawing objects based on information in provided DataFrame.
    Skips rows with no quantity, package key, or drawing provided. Only runs once per package key.

    :param pd.DataFrame df: data used to build drawing objects
    :return: dwg_sch - dictionary with package keys as keys and dwg objects as values
    :rtype: dict
    """
    dwg_sch: dict[str, DWG] = {}

    # loop through all rows on schedule
    i: int
    for index, row in df.iterrows():
        # skip if row quantity is 0
        if row['qty'] == 0:
            continue
        # skip if package key empty/NaN or if package key has already been included in submittal
        if type(pkg := row['pkg_key']) in (float, None) or pkg in dwg_sch:
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
            if size := row['size'] == 'SMALL' and type(rate := row['rate']) not in (None, str) and 0 < rate <= 5:
                dwg.set_sm()
        except TypeError:
            raise Exception('Please ensure rate fields are blank or filled with numeric values. Save. Try again.')

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
        dwg_sch[pkg] = dwg

    return dwg_sch


def find_dwgs(dwg_sch):
    """
    Generate filepaths for drawings in sch_dict

    :param dict dwg_sch: dict of deduplicated drawings from schedule as values, package keys as keys
    :return: sch_dict - dict of deduplicated drawings from schedule as values, package keys as keys
    :rtype: dict
    """
    # walk through folders to find drawing, append file path to list
    dwg: DWG
    for dwg in dwg_sch.values():
        name: str = dwg.name

        # check to ensure filetype is included from dwg column on schedule to counteract manually added drawing codes
        if name[-4:] != '.pdf':
            name = f"{name}.pdf"

        # loop through subfolders and files therein to find drawing location
        path: str
        subdir: list[str]
        files: list[str]
        # if no control kit
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
        # if large size drawing
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

    return dwg_sch


def concat(wb, dwg_dict, spec_list):
    """
    Creates the pdf container, then grabs all required drawing filepaths, then appends spec sheets from all the
    determine_spec functions. Finally, writes the pdf file to the active folder.

    :param xw.Book wb: the project file generate_submittal is being called from
    :param dict dwg_dict: dictionary of dwgs as keys and values: [package key, dwg filepath]
    :param list spec_list: list of accessory filepaths in order of occurrence
    :return: None - file is written
    """
    # instantiate pdf container
    merger: PdfWriter = PdfWriter()
    # instantiate list for ordered drawings from A to AN as a base
    dwg_list: list[None] = [None] * 40

    # name target file
    target: str = rename(wb, 'SUBMITTAL', 'pdf')

    # loop through dwg_dict to sort dwgs by package key
    pkg: str
    dwg: DWG
    for pkg, dwg in dwg_dict.items():
        # handle pkg_keys greater than Z (AA-AN) dynamically
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

    # write pdf file to folder
    merger.write(target)


def generate_submittal():
    """
    Master submittal generation function. Pulls information required from schedule, creates DWG objects, calls functions
    to location pdf drawings, calls functions to select spec sheets, then calls a function to concatenate/save them to
    the submittal file.

    :return: None - calls function concat to write pdf file
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
            spec_list_lg = lg_spec(parts, spec_list_lg)

        else:
            if dwg.sp_case_2:
                spec_list = sp_case_2(parts, spec_list)
            elif dwg.sp_case_1:
                spec_list = sp_case_1(parts, spec_list)
            else:
                spec_list = typ_spec(parts, dwg.sm, spec_list)

        # if dwg object flagged to include controls
        if type(dwg.ctrl_model) not in (None, float):
            controls_list = controls(parts[1], dwg.ctrl_model, dwg.ctrl_size, dwg.signal)

    # concatenate all lists into a single list for pdf merging
    final_spec_list: list[str] = spec_list + spec_list_lg + controls_list

    concat(wb, sch_dict, final_spec_list)
