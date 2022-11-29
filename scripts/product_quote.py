# product_quote.py
import pandas as pd
import xlwings as xw
import re
import json

from quote.controls import control_type_1, control_type_2
from quote.descriptions import generate_description, assign_descrip_to_quote
from quote.io import assign_pn_to_quote, quoted_by, clear_quote
from quote.kits import quote_sm_kit, quote_large_kit
from utils.rename import rename

"""
author: Sage Gendron
Logic for generating the quote from the schedule.
"""
# IMMUTABLE GLOBAL VARIABLES USED FOR EASE IN UPDATING; THIS IS NOT BEST PRACTICE
size_dict: dict[str, str] = {'1/2"': 'A', '5/8"': 'A', '3/4"': 'B', "3/4''": 'B', '7/8"': 'B', '1"': 'C', '1-1/4"': 'D',
                             '1 1/4"': 'D', '1-1/2"': 'E', '1 1/2"': 'E', '2"': 'F', '2-1/2"': 'G', '2 1/2"': 'G',
                             '3"': 'H', '4"': 'I', '6"': 'J', '8"': 'K', '10"': 'L', '12"': 'M', 'A': 'A', 'B': 'B',
                             'C': 'C', 'D': 'D', 'E': 'E', 'F': 'F'}
max_rate: dict[str, int] = {'A': 5, 'B': 9, 'C': 18, 'D': 27, 'E': 36, 'F': 64}
# Project file cells (strings) for header
sch_job_name: str = 'C19'
sch_customer: str = 'C23'
sch_attn: str = 'C24'
# Excel Quote cells (strings) for header
q_job_name: str = 'F7'
q_customer: str = 'F8'
q_attn: str = 'F9'
# Excel Quote cells for Freight/Lead Time
leadtime_cell: str = 'L7'
freight_cell: str = 'L8'
quote_total: str = 'AA4'


def quote_pkg(json_dict):
    """
    Called for each package key that hasn't been quoted yet. \n
    Parses each part number to create the correct coil kit for each size, pipe type, coil type, cv type/size, etc. \n
    General structure is to instantiate two empty lists: part_list and part_quantity. These lists will grow and
    evolve to represent a quoted kit according to the drawing code and information from schedule.

    :param dict json_dict: sub-dictionary of scheduled information for a single row (tag) at a time
    :return:
        - new_part_list - list of new part numbers
        - new_part_quantity - list of part quantities
    :rtype: (list, list)
    """
    dwg: str = json_dict['dwg']
    size: str = size_dict[json_dict['size']]
    sys_type: str = json_dict['sys_type']

    part_list: list[str] = []
    part_quantity: list[int] = []
    cv_parts: list[str] = []
    cv_qtys: list[int] = []

    large_size: bool = False
    is_sp_case_1: bool = False
    cv_check: bool = False
    picv_check: bool = False

    # check for pipe type blanks / NaNs
    if type(sys_type) is float:
        sys_type = 'MNPT'

    # check if large size
    if size == 'LARGE':
        large_size = True

    # check if SS column is not empty (NaN)
    if type(json_dict['sp_case_1']) is not float or json_dict['sp_case_1'] == 'YES':
        is_sp_case_1 = True

    # break apart kit number for parsing
    components: list[str] = re.findall(r"[\w'+=]+", dwg)

    # check for control valve markers to be caught in a boolean to call a separate quote function later
    cv_list: list[str] = ['+V2', '+V3', '+V4', '+BE', '+JCI']
    picv_list: list[str] = ['91', '92', '93', '85', '94']
    i: int = 0
    while i < len(cv_list) and cv_check is False and picv_check is False:
        if cv_list[i] in components[1]:
            cv_check = True
        if picv_list[i] in components[1]:
            picv_check = True
        i += 1

    # parse part numbers
    part_list: list[str]
    part_quantity: list[int]
    acc_list: list[str]
    if not large_size:
        # quote typical kit
        part_list, part_quantity, hose_size, hose_type, perm_is_tgv, is_tgv = quote_sm_kit(
            part_list, part_quantity, components, json_dict, json_dict['rate'], size, sys_type, is_sp_case_1)

        # quote hoses
        part_list, part_quantity, h_check = quote_f_component(part_list, part_quantity, components, size, sys_type,
                                                              json_dict['conn_size'], json_dict['conn_type'], hose_size,
                                                              hose_type, json_dict['hose_length'])

    # large size protocol
    elif large_size:
        # quote large kit
        part_list, part_quantity = quote_large_kit(part_list, part_quantity, components, json_dict['rate'], size)

    # if CV found in dwg code previously, quote actuator with provided CV pn
    if cv_check:
        cv_parts, cv_qtys = cv_quote_logic_OMNI.control_type_1(json_dict['control_pt'],
                                                               json_dict['control_size_type'].split()[0],
                                                               json_dict['signal'])
    # if PICV found in dwg code previously, quote actuator with provided PICV pn
    if picv_check:
        cv_parts, cv_qtys = scripts.quote.control_type_1.control_type_2(json_dict['control_pt'], json_dict['signal'])

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

    # if cv or PICV quoted, append those parts/quantities to the end
    if cv_check or picv_check:
        if type(cv_parts) is str:
            new_part_list.append(cv_parts)
            new_part_quantity.append(cv_qtys)
        else:
            new_part_list.extend(cv_parts)
            new_part_quantity.extend(cv_qtys)

    return new_part_list, new_part_quantity


def quote_f_component(part_list, part_quantity, components, size, sys_type, conn_size, conn_type, f_size, f_type,
                      f_length):
    """
    Quotes hoses based on information available specifically catches union size changes (generally) and tries to
    account for coil sizes where possible.

    :param list part_list: list of parts quoted thus far
    :param list part_quantity: list of quantities to pair with part_list entries by index
    :param list components: smart kit code (drawing name) split by hyphens
    :param str size: runout size value (in letter form) from cell in this particular row from the schedule
    :param str sys_type: runout type value from cell in this particular row from the schedule
    :param str conn_size: coil size value (in letter form) from cell in this particular row from the schedule
    :param str conn_type: coil type value from cell in this particular row from the schedule
    :param f_size: hose size (only if needed to be specified in special cases ie not matching runout size)
    :param f_type: hose type (only if needed to be specified so far)
    :param int f_length: hose length required as indicated on schedule (defaults to min length)
    :return:
        - part_list - list of parts quoted thus far
        - part_quantity - list of quantities to pair with part_list entries by index
        - h_check - are there hoses?
    :rtype: (list, list, bool)
    """
    h_check = None
    pn_hc_add = None

    try:
        if components[1][2] == 'H':
            # if hoses on runout side of kits, hose size/type to match pipe size/type
            if '=RHC' in components[1]:
                f_size = size
                f_type = sys_type

            # handle if hose size is not provided
            if type(f_size) is float or f_size is None:
                # if coil size not provided, hose size set to pipe size
                if type(conn_size) is float:
                    f_size = size
                # if coil size provided, set hose size equal to coil size
                else:
                    try:
                        f_size = size_dict[conn_size]
                    except KeyError:
                        f_size = conn_size

            # a bunch of catches to try to figure out what end the hose should have (this is a mess)
            # if no previously determined hose type and no coil type provided
            if f_type is None and type(conn_type) is float:
                f_type = 'M'
            # if no previously determined hose type and coil_type determined
            elif f_type is None and conn_type not in ('THD', 'TBD'):
                f_type = 'C' if conn_type == 'SWT' else 'M'
            # if no previously determined hose type or hose type threaded/TBD
            elif type(f_type) is float or f_type in ('THD', 'TBD'):
                f_type = 'M'
            # if hose type predetermined to be threaded
            # (treats male and female the same, does NOT auto build in couplings/labor)
            elif f_type in ('MNPT', 'FNPT'):
                f_type = 'M'
            # if hose type predetermined to be sweat
            elif f_type == 'SWT':
                f_type = 'C'
            # if all else fails and hose type is not predetermined, assume male
            elif f_type is None:
                f_type = 'M'

            # error handle mistyped numbers and NaN to default to minimum length
            if f_length % 3 != 0:
                # try to account for press coil types and hoses
                if f_type == 'PRESS':
                    if size in 'ABC':
                        pn_hc = f"HC-{f_size}MM-12"
                    elif size in 'DE':
                        pn_hc = f"HC-{f_size}MM-18"
                    elif size == 'F':
                        pn_hc = f"HC-{f_size}MM-24"
                    pn_hc_add = [f"PF-{f_size}F", f"TL-{f_size}"]  # are there PF-*Fs for these sizes?
                else:
                    if size in 'ABC':
                        pn_hc = f"HC-{f_size}M{f_type}-12"
                    elif size in 'DE':
                        pn_hc = f"HC-{f_size}M{f_type}-18"
                    elif size == 'F':
                        pn_hc = f"HC-{f_size}M{f_type}-24"
            else:
                if f_type == 'PRESS':
                    pn_hc = f"HC-{f_size}MM-{f_length}"
                    pn_hc_add = [f"PF-{f_size}F", f"TL-{f_size}"]  # are there PF-*Fs for these sizes?
                else:
                    pn_hc = f"HC-{f_size}M{f_type}-{int(f_length)}"
            # insert hose pn to part list and quantity list
            part_list.insert(1, pn_hc)
            part_quantity.insert(1, 2)
            # append pn_hc_add to part list and quantity list if exists
            if pn_hc_add:
                part_list.extend(pn_hc_add)
                for i in range(len(pn_hc_add)):
                    part_quantity.append(2)
            # add labor charges if L in smart kit number
            if components[1][3] == 'L':
                pn_tl = f"TL-{f_size}"
                part_list.append(pn_tl)
                # identify situations where only 1 labor for hose kits
                if components[0].startswith('3') and \
                        (components[0][1] == 'N' or 'BY' in components[1] or 'VY' in components[1]):
                    part_quantity.append(1)
                elif components[0][2] == 'O':
                    part_quantity.append(1)
                else:
                    part_quantity.append(2)
            h_check = True

        return part_list, part_quantity, h_check

    except Exception as e:
        raise Exception('Check dwg file codes to ensure there isn\'t a small kit with large dwg code.')


def generate_quote():
    """
    Parent function that is called from the project file. Iterates through all schedule data, creates quote, quotes
    parts, enters all data into Excel quote, saves data package of all digested info from schedule as a json file.

    :return: data.json - saves sorted json list of dictionaries in same folder to archive data for future handling
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
    package_quantities: dict[str, int] = {}

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
        package_quantities[row['pkg_key']] = package_quantities.get(row['pkg_key'], 0) + row['qty']

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
            'f_length': row['f'],
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
    assign_pn_to_quote(wb, sub_dwg_dict, package_quantities)
    # change quoted by cell to match initials in filename
    quoted_by(wb)

    # add per package key quantities to aid sales order generation
    json_sch.append(package_quantities)

    # export json data packet with similar file naming scheme
    json_export_file: str = rename(wb, 'DATA', 'json')
    with open(json_export_file, 'w') as outfile:
        json.dump(json_sch, outfile, sort_keys=True, indent=4)
