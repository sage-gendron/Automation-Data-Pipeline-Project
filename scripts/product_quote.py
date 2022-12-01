# scripts/product_quote.py
"""
author: Sage Gendron
Primary logic for quote generation based on information contained in the engineered schedule.
"""
import pandas as pd
import xlwings as xw
import re
import json

from quote.connections import size_dict
from quote.controls import control_type_1, control_type_2
from quote.descriptions import generate_description, assign_descrip_to_quote
from quote.f_component import sm_f_component
from quote.io import assign_pn_to_quote, quoted_by, clear_quote
from quote.kit_sizes import quote_sm_kit, quote_large_kit
from utils.rename import rename

# Project file cells (strings) for header
sch_job_name: str = 'C19'
sch_customer: str = 'C23'
sch_attn: str = 'C24'
# Excel Quote cells (strings) for header
q_job_name: str = 'F7'
q_customer: str = 'F8'
q_attn: str = 'F9'
# Excel Quote cells for Shipping/ETA/Estimate Total
eta_cell: str = 'L7'
shipping_cell: str = 'L8'
quote_total: str = 'AA4'


def quote_pkg(sch_row):
    """
    Called for each package key that hasn't been quoted yet. \n
    Parses each part number to create the correct kit for each size, system type, connection type, control type/size.
    General structure is to instantiate two empty lists: part_list and part_quantity. These lists will grow and
    evolve to represent a quoted kit according to the drawing code and information from schedule.

    :param dict sch_row: sub-dictionary of scheduled information for a single row at a time
    :return:
        - new_part_list - list of new part numbers
        - new_part_quantity - list of part quantities
    :rtype: (list, list)
    """
    dwg: str = sch_row['dwg']
    size: str = size_dict[sch_row['size']]
    sys_type: str = sch_row['sys_type']

    large_size: bool = False
    is_sp_case_1: bool = False

    # check for system type blanks / NaNs
    if type(sys_type) is float:
        sys_type = 'MALE'

    # check if large size
    if size == 'LARGE':
        large_size = True

    # check for special_case_1
    if type(sch_row['sp_case_1']) is not float or sch_row['sp_case_1'] == 'YES':
        is_sp_case_1 = True

    # break apart kit number for parsing
    dwg_split: list[str] = re.findall(r"[\w'+=]+", dwg)

    # parse part numbers
    part_list: list[str] = []
    part_quantity: list[int] = []
    control_parts: list[str] = []
    control_qtys: list[int] = []

    # large size protocol
    if large_size:
        # quote large kit
        part_list, part_quantity = quote_large_kit(part_list, part_quantity, dwg_split, sch_row['rate'], size)

    else:
        # quote typical kit
        part_list, part_quantity, f_size, f_type = quote_sm_kit(part_list, part_quantity, dwg_split, sch_row, size,
                                                                is_sp_case_1)

        # quote f components
        part_list, part_quantity, f_check = sm_f_component(part_list, part_quantity, dwg_split, size, sys_type,
                                                           sch_row['conn_size'], sch_row['conn_type'], f_size,
                                                           f_type)

    # if control_type_1 found in dwg code previously, quote secondary control part with provided control part number
    if '+CTRL_1' in dwg_split[1]:
        control_parts, control_qtys = control_type_1(sch_row['control_pt'], sch_row['control_size_type'].split()[0],
                                                     sch_row['signal'])
    # if control_type_2 found in dwg code previously, quote secondary control part with provided control part number
    if '+CTRL_2' in dwg_split[1]:
        control_parts, control_qtys = control_type_2(sch_row['control_pt'], sch_row['signal'])

    # deduplicate
    new_part_list: list[str] = []
    new_part_quantity: list[int] = []
    p: int
    for p in range(len(part_list)):
        if part_list[p] in new_part_list:
            dex: int = new_part_list.index(part_list[p])
            new_part_quantity[dex] += part_quantity[p]
        else:
            new_part_list.append(part_list[p])
            new_part_quantity.append(part_quantity[p])

    # if control type 1 or 2 quoted, append those parts/quantities to the end
    if control_parts:
        new_part_list.extend(control_parts)
        new_part_quantity.extend(control_qtys)

    return new_part_list, new_part_quantity


def generate_quote():
    """
    Parent function that is called from the project file. Iterates through all schedule data, creates quote, quotes
    parts, enters all data into Excel quote, saves data package of all digested info from schedule as a JSON file.

    :return: data.json - saves sorted list of dictionaries in same folder to archive data for future mining
    :rtype: file
    """
    wb: xw.Book = xw.Book.caller()

    # if called from a revised project file, resets quote to blank from the template
    if 'R.' in wb.fullname.split('\\')[-1]:
        clear_quote(wb)

    # convert schedule to a DataFrame object
    df: pd.DataFrame = pd.read_excel(wb.fullname, sheet_name='SCHEDULE', header=0, index_col=0, usecols='A:AC',
                                     skiprows=31, nrows=1000)
    # remove rows with > 7 NaN values and re-index
    df.dropna(thresh=7, inplace=True)
    df.reset_index(drop=True)

    json_sch: list[dict[str, ...]] = []
    pkg_quantities: dict[str, int] = {}

    # iterate through all rows in DataFrame
    index: pd.Index
    row: pd.Series
    for index, row in df.iterrows():
        # skip row if package key is blank / NaN
        if type(row['pkg_key']) is not str or type(row['tag']) is not str:
            raise Exception('Please check to make sure you have filled in all tag and package key fields '
                            'and try again.')

        # if bypass or branch contained in tag, do not proceed on this loop (row)
        cur_tag: str = row['tag'].lower()
        if 'bypass' in cur_tag or 'branch' in cur_tag:
            continue
        # if row quantity is 0, do not proceed on this loop (row)
        if row['qty'] == 0:
            continue

        # summarize package quantities for quote
        pkg_quantities[row['pkg_key']] = pkg_quantities.get(row['pkg_key'], 0) + row['qty']

        # create new data structure to help facilitate coded quoting process
        json_sch.append({
            'qty': row['qty'],
            'eq_type': row['eq_type'],
            'tag': row['tag'],
            'rate': row['rate'],
            'pkg_key': row['pkg_key'],
            'size': row['size'],
            'sys_type': row['sys_type'],
            'conn_size': row['conn_size'],
            'conn_type': row['conn_type'],
            'engr_component': row['engr_component'],
            'control_method': row['control_method'],
            'control_size_type': row['control_size_type'],
            'control_type': row['control_type'],
            'control_pt': row['control_pt'],
            'signal': row['signal'],
            'dwg': row['dwg'],
            'sp_case_1': row['sp_case_1'],
            'quote_descrip': '',
            'part_numbers': [],
            'part_quantities': []
        })

    # create part number list for each kit type by pkg key
    sub_dwg_dict: dict[str, list[str, list[str], list[int]]] = {}
    kit: int
    for kit in range(len(json_sch)):
        if type(temp_dwg := json_sch[kit]['dwg']) is float or temp_dwg in ('SKIP', 'STACKED'):
            continue

        # if the given package key has already been quoted, increase package quantities and update json_sch
        if (temp_pkg_key := json_sch[kit]['pkg_key']) in sub_dwg_dict:
            json_sch[kit]['part_numbers'] = sub_dwg_dict[temp_pkg_key][1]
            json_sch[kit]['part_quantities'] = sub_dwg_dict[temp_pkg_key][2]
        # if the given package key has not been quoted, quote package based on info provided and update json_sch
        else:
            new_pns, new_qtys = quote_pkg(json_sch[kit])
            json_sch[kit]['part_numbers'] = new_pns
            json_sch[kit]['part_quantities'] = new_qtys
            sub_dwg_dict[temp_pkg_key] = [temp_dwg, new_pns, new_qtys]

    # generate quote descriptions
    json_sch, descrip_dict = generate_description(json_sch)

    # fill out quote sheet with descriptions
    assign_descrip_to_quote(wb, descrip_dict)
    # fill out quote sheet with part numbers
    assign_pn_to_quote(wb, sub_dwg_dict, pkg_quantities)
    # change quoted by cell to match initials in filename
    quoted_by(wb)

    # add per package key quantities as a final JSON entry in the data packet to aid sales order generation
    json_sch.append(pkg_quantities)

    # export JSON data packet with similar file naming scheme
    json_export_file: str = rename(wb, 'DATA', 'json')
    with open(json_export_file, 'w') as outfile:
        json.dump(json_sch, outfile, sort_keys=True, indent=4)
