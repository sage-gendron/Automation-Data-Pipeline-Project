# connection_types.py
from scripts.product_quote import size_dict, max_rate

"""
author: Sage Gendron

"""


def typ_system_type(pn, sys_type):
    """
    Identify correct system type for part number and return.

    :param str pn: part number to be modified
    :param str sys_type: system type value from cell in this row from the schedule
    :return: pn - part number with altered connection type
    :rtype: str
    """
    # instantiate inverted type dictionary for system type lookup
    type_invert_ro: dict[str, str] = {'MNPT': 'F', 'SWT': 'C', 'FNPT': 'F', 'SP_CASE_2': 'SP2'}
    # alter part number to reflect correct system type
    if type(sys_type) is float or sys_type in ('TBD', 'THD'):
        pn = pn.replace('+', 'F', 1)
    else:
        pn = pn.replace('+', type_invert_ro[sys_type], 1)

    return pn


def supply_conn_type(pn, size, conn_size, conn_type, f_check=None):
    """
    Identify correct connection type for components with body type evaluated.

    :param str pn: part number to be altered
    :param str size: system size (in letter form) used for reference if control size not provided
    :param str conn_size: connection size value from cell in this particular row from the schedule
    :param str conn_type: connection type value from cell in this particular row from the schedule
    :param bool f_check: are there f components?
    :return: pn - part number with altered tailpiece connection size/type
    :rtype: str
    """
    # instantiate inverted type dictionaries for easy lookup
    type_invert_sc: dict[str, str] = {'MNPT': 'F', 'SWT': 'C', 'FNPT': 'M', 'PRESS': 'P', 'THD': 'F', 'TBD': 'F'}
    type_invert_sc_e: dict[str, str] = {'MNPT': 'EF', 'SWT': 'EC', 'FNPT': 'EM', 'PRESS': 'EP', 'THD': 'EF',
                                        'TBD': 'EF'}
    type_invert_sc_2e: dict[str, str] = {'MNPT': '2EF', 'SWT': '2EC', 'FNPT': '2EM', 'PRESS': '2EP', 'THD': '2EF',
                                         'TBD': '2EF'}
    type_invert_sc_r: dict[str, str] = {'MNPT': 'RF', 'SWT': 'RC', 'FNPT': 'RM', 'PRESS': 'RP', 'THD': 'RF',
                                        'TBD': 'RF'}
    type_invert_sc_2r: dict[str, str] = {'MNPT': '2RF', 'SWT': '2RC', 'FNPT': '2RM', 'PRESS': '2RP', 'THD': '2RF',
                                         'TBD': '2RF'}
    type_invert_sc_3r: dict[str, str] = {'MNPT': '3RF', 'SWT': '3RC', 'FNPT': '3RM', 'PRESS': '3RP', 'THD': '3RF',
                                         'TBD': '3RF'}
    type_invert_sc_4r: dict[str, str] = {'MNPT': '4RF', 'SWT': '4RC', 'FNPT': '4RM', 'PRESS': '4RP', 'THD': '4RF',
                                         'TBD': '4RF'}
    type_invert_sc_5r: dict[str, str] = {'MNPT': '5RF', 'SWT': '5RC', 'FNPT': '5RM', 'PRESS': '5RP', 'THD': '5RF',
                                         'TBD': '5RF'}

    # handle atypical/unexpected coil sizes
    if type(conn_size) is not float:
        if conn_size.upper() == 'TBD':
            conn_size = size
        else:
            try:
                conn_size = size_dict[conn_size]
            # if coil size not present in global variable size_dict, throw error
            except KeyError:
                raise Exception('Recheck coil sizes.')

    # change component type to accommodate flex hoses if required or to reflect schedule info
    if f_check is True:
        conn_type = 'MNPT'
    else:
        if conn_type == 'FLG':
            conn_type = 'MNPT'

    # handle a blank coil type
    if type(conn_type) is float:
        conn_type = 'TBD'

    # calculate numeric difference between pipe size and coil size
    ord_diff: int = ord(size) - ord(conn_size)

    # handles pns for no connection size or when connection size equals system size
    if type(conn_size) is float or ord_diff == 0:
        pn = pn.replace('+', type_invert_sc[conn_type], 1)

    # handles pns for connection sizes 1 through 5 sizes larger than system size
    elif ord_diff < 0:
        if ord_diff == -1:
            pn = pn.replace('+', type_invert_sc_e[conn_type], 1)
        elif ord_diff == -2:
            pn = pn.replace('+', type_invert_sc_2e[conn_type], 1)

    # handles pns for connection sizes 1 through 5 sizes smaller than system size
    elif ord_diff > 0:
        if ord_diff == 1:
            pn = pn.replace('+', type_invert_sc_r[conn_type], 1)
        elif ord_diff == 2:
            pn = pn.replace('+', type_invert_sc_2r[conn_type], 1)
        elif ord_diff == 3:
            pn = pn.replace('+', type_invert_sc_3r[conn_type], 1)
        elif ord_diff == 4:
            pn = pn.replace('+', type_invert_sc_4r[conn_type], 1)
        elif ord_diff == 5:
            pn = pn.replace('+', type_invert_sc_5r[conn_type], 1)
    else:
        raise Exception(
            'Something went wrong trying to generate a part number. Please review coil types on schedule and try '
            'again.')

    return pn


def return_conn_type(pn, size, conn_size, conn_type, rate, f_size, f_type, f_check=None):
    """
    Evaluates coil size/type and alters pn to match. If reduced coil, component 2 will change size to match if rate not
    too high.

    :param str pn: part number to be altered
    :param str size: system size (in letter form) used for reference if control size not provided
    :param str conn_size: connection size value from cell in this particular row from the schedule
    :param str conn_type: connection type value from cell in this particular row from the schedule
    :param float rate: rate value from cell in this particular row from the schedule
    :param str f_size: indicates f component size (since it can vary based max rates by size)
    :param str f_type: indicates f component type (since it can only vary between a few types)
    :param bool f_check: are there f components?
    :return:
        - pn - part number with altered body connection size/type
        - f_size - returns None if no f components required, else returns size of return connection component
        - f_type - returns None if no f components required, else returns type of return connection component
    :rtype: (str, str, str)
    """
    # instantiate inverted type dictionary for easy lookup
    type_invert_rc: dict[str, str] = {'MNPT': 'F', 'SWT': 'C', 'FNPT': 'F', 'PRESS': 'P', 'FLG': 'F'}

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
        conn_type = 'MNPT'

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
    # if connection size is 1 or two less system pipe size and rate above max_rate global variable for its size
    elif ord(size) - ord(conn_size) in [1, 2] and float(rate) > float(max_rate[conn_size]):
        # change connection side of component to correct type
        if type(conn_type) is float or conn_type in ('THD', 'TBD'):
            pn = pn.replace('+', 'F', 1)
        else:
            pn = pn.replace('+', type_invert_rc[conn_type], 1)
        # set f component size to system size
        f_size = size
    else:
        pass

    return pn, f_size, f_type


def ctrl_conn_type(pn, size, control_size_type, conn_size, f_check=None):
    """
    Assign tailpiece to part number to match control valve type/size if provided.

    :param str pn: part number to be altered
    :param str size: system size (in letter form) used for reference if control size not provided
    :param str control_size_type: control size/type provided in this row from schedule
    :param str conn_size: connection size value from cell in this row from the schedule
    :param bool f_check: are there f components?
    :return: pn - part number with altered connection size/type
    :rtype: str
    """
    # instantiate inverted type dictionaries for easy lookup
    type_invert_cv: dict[str, str] = {'MNPT': 'F', 'SWT': 'C', 'FNPT': 'M'}
    type_invert_cv_e: dict[str, str] = {'MNPT': 'EF', 'SWT': 'EC', 'FNPT': 'EM'}
    type_invert_cv_2e: dict[str, str] = {'MNPT': '2EF', 'SWT': '2EC', 'FNPT': '2EM'}
    type_invert_cv_r: dict[str, str] = {'MNPT': 'RF', 'SWT': 'RC', 'FNPT': 'RM'}
    type_invert_cv_2r: dict[str, str] = {'MNPT': '2RF', 'SWT': '2RC', 'FNPT': '2RM'}
    type_invert_cv_3r: dict[str, str] = {'MNPT': '3RF', 'SWT': '3RC', 'FNPT': '3RM'}
    type_invert_cv_4r: dict[str, str] = {'MNPT': '4RF', 'SWT': '4RC', 'FNPT': '4RM'}
    type_invert_cv_5r: dict[str, str] = {'MNPT': '5RF', 'SWT': '5RC', 'FNPT': '5RM'}

    # check if control size/type was provided
    if type(control_size_type) is float or control_size_type == 'TBD':
        # check to see if connection size was provided (for component between connection and control)
        cv_size: str
        cv_type: str
        if not f_check and type(conn_size) is not float:
            # check to see if connection size provided is in typical size list, if not, assign 1/2"
            try:
                cv_size = size_dict[conn_size]
            except KeyError:
                cv_size = 'A'
        # if connection size not provided, assign control size to be system size
        else:
            cv_size = size
        # control type assumed to be FNPT if not provided
        cv_type = 'FNPT'
    # if control size/type provided, assign parameters accordingly (control size/type must be separated by a space)
    else:
        cv_list: list[str] = control_size_type.split()
        cv_size = size_dict[cv_list[0]]
        cv_type = cv_list[1]

    pn_size: str
    if pn[3] == 'S':
        pn_size = pn[6] if pn[5] == '-' else pn[5]
    # part size handle clause for small kit size index situations
    elif pn[3] != '-':
        pn_size = pn[3]
    # assign component size as originally assumed (may differ if connection size varies)
    else:
        pn_size = pn[4]

    # if cv size matches pipe size, run through type-inversion dictionary
    if cv_size == pn_size:
        pn = pn.replace('+', type_invert_cv[cv_type], 1)
    # if cv size 1 less than pipe size, run through type-inversion dictionary
    elif ord(pn_size) - ord(cv_size) == 1:
        pn = pn.replace('+', type_invert_cv_r[cv_type], 1)
    # if cv size 2 less than component size, run through type-inversion dictionary
    elif ord(pn_size) - ord(cv_size) == 2:
        pn = pn.replace('+', type_invert_cv_2r[cv_type], 1)
    # if cv size 3 less than component size, run through type-inversion dictionary
    elif ord(pn_size) - ord(cv_size) == 3:
        pn = pn.replace('+', type_invert_cv_3r[cv_type], 1)
    # if cv size 4 less than component size, run through type-inversion dictionary
    elif ord(pn_size) - ord(cv_size) == 4:
        pn = pn.replace('+', type_invert_cv_4r[cv_type], 1)
    # if cv size 5 less than component size, run through type-inversion dictionary
    elif ord(pn_size) - ord(cv_size) == 5:
        pn = pn.replace('+', type_invert_cv_5r[cv_type], 1)
    # if cv size 1 more than component size, run through type-inversion dictionary
    elif ord(pn_size) - ord(cv_size) == -1:
        pn = pn.replace('+', type_invert_cv_e[cv_type], 1)
    # if cv size 2 more than component size, run through type-inversion dictionary
    elif ord(pn_size) - ord(cv_size) == -2:
        pn = pn.replace('+', type_invert_cv_2e[cv_type], 1)
    else:
        pass

    return pn
