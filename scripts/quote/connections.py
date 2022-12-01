# scripts/quote/connection.py
"""
author: Sage Gendron
Contains functions to assist in determining connection types and required connection sizes. Many of these intentionally
do not address all possible combinations so that they appear broken when placed into the quote Excel file. Future plan
was to implement a notes/errors logging system that prints out at bottom of each package for more obvious breaks like
these.
"""
# global reference variables that are hard equalities (for sizes) or hard limits (for rates)
size_dict: dict[str, str] = {'1': 'A', '2': 'B', '3': 'C', '4': 'D', '5': 'E', '6': 'F', '7': 'G', '8': 'H'}
max_rate: dict[str, int] = {'A': 10, 'B': 20, 'C': 30, 'D': 40, 'E': 50, 'F': 60}


def typ_system_type(pn, sys_type):
    """
    Identify correct system type for part number.

    :param str pn: part number to be modified
    :param str sys_type: system type value
    :return: pn - part number with altered connection type
    :rtype: str
    """
    # instantiate inverted type dictionary for system type lookup
    type_invert_ro: dict[str, str] = {'MALE': 'F', 'NON-THD': 'N', 'FEMALE': 'F', 'SP_CASE_2': 'SP2'}
    # alter part number to reflect correct system type
    if type(sys_type) is float or sys_type in ('TBD', 'THD'):
        pn = pn.replace('+', 'F', 1)
    else:
        pn = pn.replace('+', type_invert_ro[sys_type], 1)

    return pn


def supply_conn_type(pn, size, conn_size, conn_type, f_check=False):
    """
    Identify correct connection type for components with system type already evaluated.

    :param str pn: part number to be altered
    :param str size: system size (in letter form) used for reference if control size not provided
    :param str conn_size: connection size value
    :param str conn_type: connection type value
    :param bool f_check: are there f components?
    :return: pn - part number with altered connection size/type
    :rtype: str
    """
    # instantiate inverted type dictionaries for easy lookup
    type_invert_sc: dict[str, str] = {'MALE': 'F', 'NON-THD': 'N', 'FEMALE': 'M', 'SP_CASE_2': 'SP2'}
    type_invert_sc_e: dict[str, str] = {'MALE': 'EF', 'NON-THD': 'EN', 'FEMALE': 'EM', 'SP_CASE_2': 'ESP2'}
    type_invert_sc_2e: dict[str, str] = {'MALE': '2EF', 'NON-THD': '2EN', 'FEMALE': '2EM', 'SP_CASE_2': '2ESP2'}
    type_invert_sc_r: dict[str, str] = {'MALE': 'RF', 'NON-THD': 'RN', 'FEMALE': 'RM', 'SP_CASE_2': 'RSP2'}
    type_invert_sc_2r: dict[str, str] = {'MALE': '2RF', 'NON-THD': '2RN', 'FEMALE': '2RM', 'SP_CASE_2': '2RSP2'}

    # handle atypical/unexpected connection sizes
    if type(conn_size) is not float:
        if conn_size.upper() == 'TBD':
            conn_size = size
        else:
            try:
                conn_size = size_dict[conn_size]
            # if connection size not present in global variable size_dict, throw error
            except KeyError:
                raise Exception('Recheck connection sizes.')

    # change component type to accommodate f components if required or to reflect schedule info
    if f_check:
        conn_type = 'MALE'

    # handle a blank connection type
    if type(conn_type) is float:
        conn_type = 'TBD'

    # calculate numeric difference between system size and connection size
    ord_diff: int = ord(size) - ord(conn_size)

    # handles pns for no connection size or when connection size equals system size
    if type(conn_size) is float or ord_diff == 0:
        pn = pn.replace('+', type_invert_sc[conn_type], 1)
    # handles pns for connection sizes 1 to 2 sizes larger than system size
    elif ord_diff < 0:
        if ord_diff == -1:
            pn = pn.replace('+', type_invert_sc_e[conn_type], 1)
        elif ord_diff == -2:
            pn = pn.replace('+', type_invert_sc_2e[conn_type], 1)
    # handles pns for connection sizes 1 to 2 sizes smaller than system size
    elif ord_diff > 0:
        if ord_diff == 1:
            pn = pn.replace('+', type_invert_sc_r[conn_type], 1)
        elif ord_diff == 2:
            pn = pn.replace('+', type_invert_sc_2r[conn_type], 1)
    else:
        raise Exception(
            'Something went wrong trying to generate a part number. Please review coil types on schedule and try '
            'again.')

    return pn


def return_conn_type(pn, size, conn_size, conn_type, rate, f_size, f_type, f_check=False):
    """
    Evaluates connection size/type and alters part number to match. If reduced connection, component 2 will change size
    to match if rate does not exceed value in max_rate global variable.

    :param str pn: part number to be altered
    :param str size: system size (in letter form) used for reference if control size not provided
    :param str conn_size: connection size value
    :param str conn_type: connection type value
    :param float rate: engineer-provided rate
    :param str f_size: indicates f component size (since it can vary based max rates by size)
    :param str f_type: indicates f component type (since it can only vary between a few types)
    :param bool f_check: are there f components?
    :return:
        - pn - part number with altered body connection size/type
        - f_size - returns None if no f components required, else returns size of return connection
        - f_type - returns None if no f components required, else returns type of return connection
    :rtype: (str, str, str)
    """
    # instantiate inverted type dictionary for easy lookup
    type_invert_rc: dict[str, str] = {'MALE': 'F', 'NON-THD': 'N', 'FEMALE': 'F', 'SP_CASE_2': 'SP2'}

    # if connection size has a value, try to make that f component size, else error out
    if type(conn_size) is not float:
        if conn_size.upper() == 'TBD':
            conn_size = size
        else:
            try:
                conn_size = size_dict[conn_size]
            except KeyError:
                raise Exception('Recheck coil sizes.')

    # change component type to accommodate f components if required or to reflect schedule info
    if f_check:
        f_type = conn_type
        conn_type = 'MALE'

    # if no connection type provided or connection size = system size
    if type(conn_size) is float or conn_size == size:
        if type(conn_type) is float or conn_type in ('THD', 'TBD'):
            pn = pn.replace('+', 'F', 1)
        else:
            pn = pn.replace('+', type_invert_rc[conn_type], 1)
        # set f component size to system size
        f_size = size
    # if connection size is 1 less than system size and rate below max_rate global variable for its size
    elif ord(size) - ord(conn_size) == 1 and float(rate) <= float(max_rate[conn_size]):
        # take 1 off the value of the Unicode value of the size letter
        letnum: int = ord(pn[3]) - 1
        # rebuild existing part number with revised size
        pn = f"{pn[0:3]}{chr(letnum)}{pn[4:]}"
        # change connection side of component to correct type
        if type(conn_type) is float or conn_type in ('THD', 'TBD'):
            pn = pn.replace('+', 'F', 1)
        else:
            pn = pn.replace('+', type_invert_rc[conn_type], 1)
        # set f component size to reduced size
        f_size = chr(letnum)
    # if connection size is 2 less than system size and rate below max_rate global variable for its size
    elif ord(size) - ord(conn_size) == 2 and float(rate) <= float(max_rate[conn_size]):
        # take 2 off the value of the Unicode value of the size letter
        letnum: int = ord(pn[3]) - 2
        # rebuild existing part number with revised size
        pn = f"{pn[0:3]}{chr(letnum)}{pn[4:]}"
        # change connection side of component to correct type
        if type(conn_type) is float or conn_type in ('THD', 'TBD'):
            pn = pn.replace('+', 'F', 1)
        else:
            pn = pn.replace('+', type_invert_rc[conn_type], 1)
        # set f component size to reduced size
        f_size = chr(letnum)
    # if connection size is 1 or two less system size and rate above max_rate global variable for its size
    elif ord(size) - ord(conn_size) in [1, 2] and float(rate) > float(max_rate[conn_size]):
        # change connection side of component to correct type
        if type(conn_type) is float or conn_type in ('THD', 'TBD'):
            pn = pn.replace('+', 'F', 1)
        else:
            pn = pn.replace('+', type_invert_rc[conn_type], 1)
        # set f component size to system size
        f_size = size

    return pn, f_size, f_type


def ctrl_conn_type(pn, size, control_size_type, conn_size, f_check=False):
    """
    Assign connection type to part number to match control type/size if provided.

    :param str pn: part number to be altered
    :param str size: system size (in letter form) used for reference if control size not provided
    :param str control_size_type: control size/type provided
    :param str conn_size: connection size value
    :param bool f_check: are there f components?
    :return: pn - part number with altered connection size/type
    :rtype: str
    """
    # instantiate inverted type dictionaries for easy lookup
    type_invert_cv: dict[str, str] = {'MALE': 'F', 'NON-THD': 'N', 'FEMALE': 'M'}
    type_invert_cv_e: dict[str, str] = {'MALE': 'EF', 'NON-THD': 'EN', 'FEMALE': 'EM'}
    type_invert_cv_2e: dict[str, str] = {'MALE': '2EF', 'NON-THD': '2EN', 'FEMALE': '2EM'}
    type_invert_cv_r: dict[str, str] = {'MALE': 'RF', 'NON-THD': 'RN', 'FEMALE': 'RM'}
    type_invert_cv_2r: dict[str, str] = {'MALE': '2RF', 'NON-THD': '2RN', 'FEMALE': '2RM'}

    # check if control size/type was provided
    if type(control_size_type) is float or control_size_type == 'TBD':
        # check to see if conn_size was provided (for part between connection and control)
        control_size: str
        control_type: str
        if not f_check and type(conn_size) is not float:
            control_size = size_dict[conn_size]
        # if connection size not provided, assign control size to be system size
        else:
            control_size = size
        # control type assumed to be female if not provided
        control_type = 'FEMALE'
    # if control size/type provided, assign parameters accordingly (control size/type must be separated by a space)
    else:
        control_split: list[str] = control_size_type.split()
        control_size = size_dict[control_split[0]]
        control_type = control_split[1]

    # identify current part size (always first char after first hyphen)
    pn_size: str
    pn_size = pn[pn.index('-') + 1]

    # if control size matches component size, run through type-inversion dictionary
    size_diff: int = ord(pn_size) - ord(control_size)
    if size_diff == 0:
        pn = pn.replace('+', type_invert_cv[control_type], 1)
    elif size_diff > 0:
        # if control size 1 less than component size, run through type-inversion dictionary
        if size_diff == 1:
            pn = pn.replace('+', type_invert_cv_r[control_type], 1)
        # if control size 2 less than component size, run through type-inversion dictionary
        elif size_diff == 2:
            pn = pn.replace('+', type_invert_cv_2r[control_type], 1)
    elif size_diff < 0:
        # if cv size 1 more than component size, run through type-inversion dictionary
        if size_diff == -1:
            pn = pn.replace('+', type_invert_cv_e[control_type], 1)
        # if cv size 2 more than component size, run through type-inversion dictionary
        elif size_diff == -2:
            pn = pn.replace('+', type_invert_cv_2e[control_type], 1)

    return pn
