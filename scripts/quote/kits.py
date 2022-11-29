# quote_kits.py
from components import sp_case_1_component, typ_component, lg_component
from kit_types import type_two_kits, type_three_kits, no_control_kits

"""
author: Sage Gendron

"""

pass_chars: str = '234589LOPTX'


def quote_sm_kit(part_list, part_quantity, dwg_split, sch_row, rate, size, sys_type, is_sp_case_1):
    """
    Loops through characters in the first half of the smart kit code to identify which components are required. Sends
    the base part number for each component to different functions depending on its location in the kit.

    :param list part_list: already quoted parts for this package
    :param list part_quantity: already quoted quantities for this package
    :param list dwg_split: smart kit code (drawing name) split by hyphens
    :param dict sch_row: a row from the engineered schedule in json form
    :param float rate: the rate for the row from the schedule
    :param str size: the system size for the row from the schedule
    :param str sys_type: the connection type for the row from the schedule
    :param bool is_sp_case_1: indicates whether the package requires special case 1
    :return:
        - part_list - list of parts quoted thus far
        - part_quantity - list of quantities to pair with part_list entries by index
        - f_size - f component size (only if needed to be specified in special cases (ie not matching system size))
        - f_type - f component type (only if needed to be specified so far)
    :rtype: (list, list, , )
    """
    pn: str = ''
    f_size = None
    f_type = None
    i: int = 0

    char: str
    for char in dwg_split[0]:
        # handle pass chars '234589LOPTX': iterate and restart loop
        if char in pass_chars:
            i += 1
            continue
        # if sp_case_1 marked on schedule, send to sp_case_1_component() function
        elif is_sp_case_1:
            pn = sp_case_1_component(rate, size, char)
        # if typical component, send to typ_component() function
        else:
            pn = typ_component(rate, size, char)

        # parse through type 2 kits
        if dwg_split[0].startswith('2'):
            pn, f_size, f_type = type_two_kits(dwg_split, pn, size, sys_type, sch_row['conn_size'],
                                               sch_row['conn_type'], sch_row['control_size_type'], rate, i, f_size,
                                               f_type)
        # parse through type 3 kits
        elif dwg_split[0].startswith('3'):
            pn, f_size, f_type = type_three_kits(dwg_split, pn, size, sys_type, sch_row['conn_size'],
                                                 sch_row['conn_type'], sch_row['control_size_type'], rate,
                                                 i, f_size, f_type)
        # parse through kits with no controls
        else:
            pn, f_size, f_type = no_control_kits(dwg_split, pn, size, sys_type,
                                                 sch_row['conn_size'], sch_row['conn_type'], rate,
                                                 i, f_size, f_type)

        # add part number(s) to list with quantities
        if type(pn) == str:
            part_list.append(pn)
            part_quantity.append(1)
        else:
            part_list.extend(pn)
            for n in range(len(pn)):
                part_quantity.append(1)

        # increment iterative variable
        i += 1

    return part_list, part_quantity, f_size, f_type


def quote_large_kit(part_list, part_quantity, dwg_split, rate, size):
    """
    Loops through characters in the first half of the smart kit code to identify which components are required. Sends
    the character to lg_component() to identify the base part number. As all components designed to be flanged, no other
    connection type functions required.

    :param list part_list: list of parts quoted thus far
    :param list part_quantity: list of quantities to pair with part_list entries by index
    :param list dwg_split: smart kit code (drawing name) split by hyphens
    :param float rate: rate from this row from the schedule
    :param str size: system size value (in letter form) from cell in this particular row from the schedule
    :return:
        - part_list (:py:class:'list') - list of parts quoted thus far
        - part_quantity (:py:class:'list') - list of quantities to pair with part_list entries by index
    :rtype: (list, list)
    """
    char: str
    for char in dwg_split[0]:
        # handle pass chars '234589LOPTX': iterate and restart loop
        if char in pass_chars:
            continue
        else:
            pn = lg_component(rate, size, char)

        # add parts and part quantities to package list
        try:
            if type(pn) is str:
                part_list.append(pn)
                part_quantity.append(1)
            else:
                part_list.extend(pn)
                n: int
                for n in range(len(pn)):
                    part_quantity.append(1)
        except TypeError:
            raise Exception('Part list/quantity list not able to be generated. Please check dwg code for large '
                            'size kits and try again.')

    # handle flanged flex hoses
    if dwg_split[1][1] == 'F':
        part_list.append(f"F-{size}LG")
        part_quantity.append(2)

    return part_list, part_quantity
