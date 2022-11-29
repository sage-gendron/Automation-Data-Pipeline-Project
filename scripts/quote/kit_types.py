# scripts/quote/kit_types.py
"""
author: Sage Gendron
Directs incomplete part numbers to the correct functions based on their location in the package (and in the smart
drawing code schema).
"""
from connections import typ_system_type, ctrl_conn_type, supply_conn_type, return_conn_type


def type_two_kits(dwg_split, pn, size, sys_type, conn_size, conn_type, control_size_type, rate, i, f_size, f_type):
    """
    Generates a package quote for type two kits.
    Only works through one component at a time as indicated by the 'i' iterative variable.

    :param list dwg_split: smart kit code (drawing name) split by hyphens
    :param str pn: part number to be modified based on location in dwg code and sizes referenced
    :param str size: system size value (in letter form)
    :param str sys_type: system type value
    :param str conn_size: connection size value (in letter form)
    :param str conn_type: connection type value
    :param str control_size_type: control size/type value
    :param float rate: engineer-provided rate
    :param int i: iterated index of current letter in drawing code
    :param str f_size: f component size (only if needed to be specified in special cases (ie not matching system size))
    :param str f_type: f component type (only if needed to be specified)
    :return:
        - pn - completed part number
        - f_size - f component size (only if needed to be specified in special cases (ie not matching system size))
        - f_type - f component type (only if needed to be specified)
    :rtype: (str, str, str)
    """
    if i == 1:
        # check for system-side f components in case certain types force components to be female
        if '=RF' in dwg_split[1]:
            pn = pn.replace('+', 'F', 1)
        else:
            pn = typ_system_type(pn, sys_type)
        # check if f component in package to force a female connection
        pn = supply_conn_type(pn, size, conn_size, conn_type, 'F' in dwg_split[1][:4])

    elif i == 2:
        # check if f component in package to force a female connection
        pn, f_size, f_type = return_conn_type(pn, size, conn_size, conn_type, rate, f_size, f_type,
                                              'F' in dwg_split[1][:4])
        pn = ctrl_conn_type(pn, size, control_size_type, conn_size, 'F' in dwg_split[1][:4])

    elif i == 3:
        pn = typ_system_type(pn, sys_type)
        # check for supply control to alter connection type function called
        if '=SC' in dwg_split[1]:
            pn = supply_conn_type(pn, size, conn_size, conn_type)
        else:
            # check if f component in package to force a female connection
            pn = ctrl_conn_type(pn, size, control_size_type, conn_size, 'F' in dwg_split[1][:4])

    return pn, f_size, f_type


def type_three_kits(dwg_split, pn, size, sys_type, conn_size, conn_type, control_size_type, rate, i, f_size, f_type):
    """
    Generates a package quote for type three kits.
    Only works through one component at a time as indicated by the 'i' iterative variable.

    :param list dwg_split: smart kit code (drawing name) split by hyphens
    :param str pn: current base part number to be modified based on location in dwg code and sizes referenced
    :param str size: system size value (in letter form)
    :param str sys_type: system type value
    :param str conn_size: connection size value (in letter form)
    :param str conn_type: connection type value
    :param str control_size_type: control size/type value
    :param float rate: engineer-provided rate
    :param int i: iterative variable to identify index of part/location of component being quoted
    :param str f_size: f component size (only if needed to be specified in special cases ie not matching system size)
    :param str f_type: f component type (only if needed to be specified so far)
    :return:
        - pn - completed part number
        - f_size - f component size (only if needed to be specified in special cases (ie not matching system size))
        - f_type - f component type (only if needed to be specified so far)
    :rtype: (str, str, str)
    """
    if i == 1:
        # check for system-side f components in case certain types force components to be female
        if '=RF' in dwg_split[1]:
            pn = pn.replace('+', 'F', 1)
        else:
            pn = typ_system_type(pn, sys_type)
        # check for supply control to alter connection type function called
        if '=SC' in dwg_split[1]:
            pn = ctrl_conn_type(pn, size, control_size_type, conn_size)
        else:
            pn = supply_conn_type(pn, size, conn_size, conn_type, 'F' in dwg_split[1][:4])

    elif i == 2:
        # check for f components to force female
        pn, f_size, f_type = return_conn_type(pn, size, conn_size, conn_type, rate, f_size, f_type,
                                              'F' in dwg_split[1][:4])
        pn = ctrl_conn_type(pn, size, control_size_type, conn_size, 'F' in dwg_split[1][:4])

    elif i == 3:
        # check for system-side f components in case certain types force components to be female
        if '=RF' in dwg_split[1]:
            pn = pn.replace('+', 'F', 1)
        else:
            pn = typ_system_type(pn, sys_type)
        # check for f components to force female
        pn = ctrl_conn_type(pn, size, control_size_type, conn_size, 'F' in dwg_split[1][:4])

    elif i == 4:
        pn = typ_system_type(pn, sys_type)
        # check for engineered component type on the type 3 kits, so it can force it to match system type if required
        if dwg_split[0][4] == 'C':
            pn = typ_system_type(pn, sys_type)
        else:
            # change to bypass component matching control size when available
            pn = ctrl_conn_type(pn, size, control_size_type, conn_size, 'F' in dwg_split[1][:4])

    return pn, f_size, f_type


def no_control_kits(dwg_split, pn, size, sys_type, conn_size, conn_type, rate, i, f_size, f_type):
    """
    Generates a package quote for kits without controls.
    Only works through one component at a time as indicated by the 'i' iterative variable.

    :param list dwg_split: smart kit code (drawing name) split by hyphens
    :param str pn: previously generated base part number to be edited with correct connections
    :param str size: system size value (in letter form)
    :param str sys_type: system type value
    :param str conn_size: connection size value (in letter form)
    :param str conn_type: connection type value
    :param float rate: engineer-provided rate
    :param int i: iterative variable indicating location of currently quoted component in drawing name
    :param str f_size: indicates f component size (since it can vary based on max rates)
    :param str f_type: indicates f component type (since it can only vary between threaded and not-threaded)
    :return:
        - pn - completed part number
        - f_size - f component size (only if needed to be specified in special cases (ie not matching system size))
        - f_type - f component type (only if needed to be specified so far)
    :rtype: (str, str, str)
    """
    if i == 0:
        if len(dwg_split) > 1:
            # check for system-side f components in case certain types force components to be female
            if '=RF' in dwg_split[1]:
                pn = pn.replace('+', 'F', 1)
            else:
                pn = typ_system_type(pn, sys_type)
            # check if f components in package to force female to accommodate f components
            pn = supply_conn_type(pn, size, conn_size, conn_type, 'F' in dwg_split[1][:4])
        elif len(dwg_split) == 1:
            pn = typ_system_type(pn, sys_type)
            pn = supply_conn_type(pn, size, conn_size, conn_type)

    elif i == 1:
        pn, f_size, f_type = return_conn_type(pn, size, conn_size, conn_type, rate, f_size, f_type,
                                              'F' in dwg_split[1][:4])
        pn = supply_conn_type(pn, size, conn_size, conn_type)

    elif i == 2:
        # check for system-side f components in case certain types force components to be female
        if '=RF' in dwg_split[1]:
            pn = pn.replace('+', 'F', 1)
        else:
            pn = typ_system_type(pn, sys_type)
        # check if f component in package to force female to accommodate f components
        pn = supply_conn_type(pn, size, conn_size, conn_type, 'F' in dwg_split[1][:4])

    return pn, f_size, f_type
