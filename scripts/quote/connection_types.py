from scripts.quote.product_quote import size_dict, max_rate


def typ_runout_type(pn, sys_type):
    """
    Identify correct runout type for pn and return for further analysis.

    :param str pn: part number to be modified
    :param str sys_type: runout type value from cell in this particular row from the schedule
    :return: pn - part number with altered body connection type
    :rtype: str
    """
    # instantiate inverted type dictionary for pipe type lookup
    type_invert_ro: dict[str, str] = {'MNPT': 'F', 'SWT': 'C', 'FNPT': 'F', 'PRESS': 'P'}
    # alter part number to reflect correct runout type
    if type(sys_type) is float or sys_type in ('TBD', 'THD'):
        pn = pn.replace('+', 'F', 1)
    else:
        pn = pn.replace('+', type_invert_ro[sys_type], 1)

    return pn


def supply_coil_type(pn, size, conn_size, conn_type, h_check=None):
    """
    Identify correct tailpiece type for components with body type evaluated.

    :param str pn: part number to be altered
    :param str size: runout pipe size (in letter form) used for reference if cv size not provided
    :param str conn_size: coil size value from cell in this particular row from the schedule
    :param str conn_type: coil type value from cell in this particular row from the schedule
    :param bool h_check: are there hoses?
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
    if h_check is True:
        conn_type = 'MNPT'
    else:
        if conn_type == 'FLG':
            conn_type = 'MNPT'

    # handle a blank coil type
    if type(conn_type) is float:
        conn_type = 'TBD'

    # calculate numeric difference between pipe size and coil size
    ord_diff: int = ord(size) - ord(conn_size)

    # handles pns for no coil size or when coil size equals pipe size
    if type(conn_size) is float or ord_diff == 0:
        pn = pn.replace('+', type_invert_sc[conn_type], 1)

    # handles pns for coil sizes 1 through 5 sizes larger than pipe size
    elif ord_diff < 0:
        if ord_diff == -1:
            pn = pn.replace('+', type_invert_sc_e[conn_type], 1)
        elif ord_diff == -2:
            pn = pn.replace('+', type_invert_sc_2e[conn_type], 1)

    # handles pns for coil sizes 1 through 5 sizes smaller than pipe size
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


def return_coil_type(pn, size, conn_size, conn_type, rate, hose_size, hose_type, h_check=None):
    """
    Evaluates coil size/type and alters pn to match. If reduced coil, component 2 will change size to match if rate not
    too high.

    :param str pn: part number to be altered
    :param str size: runout pipe size (in letter form) used for reference if cv size not provided
    :param str conn_size: coil size value from cell in this particular row from the schedule
    :param str conn_type: coil type value from cell in this particular row from the schedule
    :param float rate: rate value from cell in this particular row from the schedule
    :param str hose_size: indicates hose size (since it can vary based on union size / max flow rates)
    :param str hose_type: indicates hose type (since it can only vary between sweat and threaded)
    :param h_check: are there hoses?
    :return:
        - pn - part number with altered body connection size/type
        - hose_size - returns None if no hoses required, else returns size of return coil component
        - hose_type - returns None if no hoses required, else returns type of return coil component
    :rtype: (str, , )
    """
    # instantiate inverted type dictionary for easy lookup
    type_invert_rc: dict[str, str] = {'MNPT': 'F', 'SWT': 'C', 'FNPT': 'F', 'PRESS': 'P', 'FLG': 'F'}

    # if coil size has a value, try to make that hose size, else default to 1/2" (to handle for 1/2" OD or smaller)
    if type(conn_size) is not float:
        if conn_size.upper() == 'TBD':
            conn_size = size
        else:
            try:
                conn_size = size_dict[conn_size]
            except KeyError:
                raise Exception('Recheck coil sizes.')

    # change component type to accommodate flex hoses if required or to reflect schedule info
    if h_check:
        hose_type = conn_type
        conn_type = 'MNPT'

    # if no coil type provided or coil size = pipe size
    if type(conn_size) is float or conn_size == size:
        if type(conn_type) is float or conn_type in ('THD', 'TBD'):
            pn = pn.replace('+', 'F', 1)
        else:
            pn = pn.replace('+', type_invert_rc[conn_type], 1)
        # set hose size to pipe size
        hose_size = size
    # if coil size is 1 less than pipe size and rate below max_rate global variable for its size
    elif ord(size) - ord(conn_size) == 1 and float(rate) <= float(max_rate[conn_size]):
        # take 1 off the value of the Unicode value of the size letter
        letnum: int = ord(pn[3]) - 1
        # rebuild existing part number with revised size
        pn = f"{pn[0:3]}{chr(letnum)}{pn[4:]}"
        # change coil side of component to correct type
        if type(conn_type) is float or conn_type in ('THD', 'TBD'):
            pn = pn.replace('+', 'F', 1)
        else:
            pn = pn.replace('+', type_invert_rc[conn_type], 1)
        # set hose size to reduced size
        hose_size = chr(letnum)
    # if coil size is 2 less than pipe size and rate below max_rate global variable for its size
    elif ord(size) - ord(conn_size) == 2 and float(rate) <= float(max_rate[conn_size]):
        # take 2 off the value of the Unicode value of the size letter
        letnum: int = ord(pn[3]) - 2
        # rebuild existing part number with revised size
        pn = f"{pn[0:3]}{chr(letnum)}{pn[4:]}"
        # change coil side of component to correct type
        if type(conn_type) is float or conn_type in ('THD', 'TBD'):
            pn = pn.replace('+', 'F', 1)
        else:
            pn = pn.replace('+', type_invert_rc[conn_type], 1)
        # set hose size to reduced size
        hose_size = chr(letnum)
    # if coil size is 1 or two less than pipe size and rate above max_rate global variable for its size
    elif ord(size) - ord(conn_size) in [1, 2] and float(rate) > float(max_rate[conn_size]):
        # change coil side of component to correct type
        if type(conn_type) is float or conn_type in ('THD', 'TBD'):
            pn = pn.replace('+', 'F', 1)
        else:
            pn = pn.replace('+', type_invert_rc[conn_type], 1)
        # set hose size to pipe size
        hose_size = size
    else:
        pass

    return pn, hose_size, hose_type


def cv_connection_type(pn, size, control_size_type, conn_size, h_check=None):
    """
    Assign tailpiece to part number to match control valve type/size if provided.

    :param str pn: part number to be altered
    :param str size: system size (in letter form) used for reference if control size not provided
    :param str control_size_type: control valve size/type provided in this particular row from schedule
    :param str conn_size: coil size value from cell in this particular row from the schedule
    :param bool h_check: are there hoses?
    :return: pn - part number with altered tailpiece connection size/type
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

    # check if cv size/type was provided
    if type(control_size_type) is float or control_size_type == 'TBD':
        # check to see if coil size was provided (for component between coil and cv)
        cv_size: str
        cv_type: str
        if not h_check and type(conn_size) is not float:
            # check to see if coil size provided is in typical size list, if not, assign 1/2"
            try:
                cv_size = size_dict[conn_size]
            except KeyError:
                cv_size = 'A'
        # if coil size not provided, assign cv size to be pipe size
        else:
            cv_size = size
        # cv type assumed to be FNPT if not provided
        cv_type = 'FNPT'
    # if cv size/type provided, assign parameters accordingly (cv size/type must be separated by a space)
    else:
        cv_list: list[str] = control_size_type.split()
        cv_size = size_dict[cv_list[0]]
        cv_type = cv_list[1]

    pn_size: str
    if pn[3] == 'S':
        pn_size = pn[6] if pn[5] == '-' else pn[5]
    # part size handle clause for TY1/TA1 index situations
    elif pn[3] != '-':
        pn_size = pn[3]
    # assign component size as originally assumed (may differ if coil size varies)
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
