# kit_types.py
from connections import typ_system_type, ctrl_conn_type, supply_conn_type, return_conn_type

"""
author: Sage Gendron

"""


def type_two_kits(dwg_split, pn, size, sys_type, conn_size, conn_type, control_size_type, rate, i, f_size, f_type):
    """
    Builds type two kit component part numbers with correct end connections and sizes matching given information.

    :param list dwg_split: smart kit code (drawing name) split by hyphens
    :param str pn: current base part number to be modified based on location in dwg code and sizes referenced
    :param str size: system size value (in letter form) from cell in this particular row from the schedule
    :param str sys_type: system type value from cell in this particular row from the schedule
    :param str conn_size: connection size value (in letter form) from cell in this particular row from the schedule
    :param str conn_type: connection type value from cell in this particular row from the schedule
    :param str control_size_type: control size/type value from cell in this particular row from the schedule
    :param float rate: rate value from cell in this particular row from the schedule
    :param int i: iterated index of current letter in drawing code
    :param f_size: f component size (only if needed to be specified in special cases ie not matching system size)
    :param f_type: f component type (only if needed to be specified so far)
    :return:
        - pn - current base part number to be modified based on location in dwg code and sizes referenced
        - f_size - f component size (only if needed to be specified in special cases ie not matching system size)
        - f_type - f component type (only if needed to be specified so far)
    :rtype: (str, , )
    """
    if i == 1:
        # check runout hoses in case SWT pipe type forces components to be THD to accommodate hose
        if '=RF' in dwg_split[1]:
            pn = pn.replace('+', 'F', 1)
        else:
            pn = typ_system_type(pn, sys_type)

        # check if f component in package to force FNPT
        pn = supply_conn_type(pn, size, conn_size, conn_type, 'F' in dwg_split[1][:4])
    elif i == 2:
        # check if f component in package to force FNPT
        pn, f_size, f_type = return_conn_type(pn, size, conn_size, conn_type, rate, f_size, f_type,
                                              'F' in dwg_split[1][:4])
        pn = ctrl_conn_type(pn, size, control_size_type, conn_size, 'F' in dwg_split[1][:4])
    elif i == 3:
        pn = typ_system_type(pn, sys_type)

        # check for supply control to alter connection type
        if '=SCV' in dwg_split[1]:
            pn = supply_conn_type(pn, size, conn_size, conn_type)
        else:
            # check if f component in package to force FNPT
            pn = ctrl_conn_type(pn, size, control_size_type, conn_size, 'F' in dwg_split[1][:4])

    return pn, f_size, f_type


def type_three_kits(dwg_split, pn, size, sys_type, conn_size, conn_type, control_size_type, rate, i, f_size, f_type):
    """
    Builds type three kit component part number with correct end connections and sizes matching given information.

    :param list dwg_split: smart kit code (drawing name) split by hyphens
    :param str pn: current base part number to be modified based on location in dwg code and sizes referenced
    :param str size: system size value (in letter form) from cell in this particular row from the schedule
    :param str sys_type: system type value from cell in this particular row from the schedule
    :param str conn_size: connection size value (in letter form) from cell in this particular row from the schedule
    :param str conn_type: connection type value from cell in this particular row from the schedule
    :param str control_size_type: control size/type value from cell in this particular row from the schedule
    :param float rate: rate from rate cell in this particular row from the schedule
    :param int i: iterative variable to identify index of part/location of component being quoted
    :param f_size: f component size (only if needed to be specified in special cases ie not matching runout size)
    :param f_type: f component type (only if needed to be specified so far)
    :return:
        - pn - current base part number to be modified based on location in dwg code and sizes referenced
        - hose_size - hose size (only if needed to be specified in special cases ie not matching runout size)
        - hose_type - hose type (only if needed to be specified so far)
    :rtype: (str, , )
    """
    if i == 1:
        # check runout hoses in case SWT pipe type forces components to be THD to accommodate hose
        if '=RF' in dwg_split[1]:
            pn = pn.replace('+', 'F', 1)
        else:
            pn = typ_system_type(pn, sys_type)
        # check for supply control valve to alter tailpiece type
        if '=SCV' in dwg_split[1]:
            pn = ctrl_conn_type(pn, size, control_size_type, conn_size)
        else:
            pn = supply_conn_type(pn, size, conn_size, conn_type, 'F' in dwg_split[1][:4])
    elif i == 2:
        # check for hoses to force FNPT
        pn, f_size, f_type = return_conn_type(pn, size, conn_size, conn_type, rate, f_size, f_type,
                                              'F' in dwg_split[1][:4])
        pn = ctrl_conn_type(pn, size, control_size_type, conn_size, 'F' in dwg_split[1][:4])
    elif i == 3:
        # check system f components in case SWT pipe type forces components to be THD to accommodate f component
        if '=RF' in dwg_split[1]:
            pn = pn.replace('+', 'F', 1)
        else:
            pn = typ_system_type(pn, sys_type)
        # check for hoses to force FNPT
        pn = ctrl_conn_type(pn, size, control_size_type, conn_size, 'F' in dwg_split[1][:4])
    elif i == 4:
        pn = typ_system_type(pn, sys_type)
        # check for engineered component type on the type 3 kits, so it can force it to match system type
        if dwg_split[0][4] == 'C':
            pn = typ_system_type(pn, sys_type)
        else:
            # change to bypass component matching control size when available
            pn = ctrl_conn_type(pn, size, control_size_type, conn_size, 'F' in dwg_split[1][:4])

    return pn, f_size, f_type


def no_control_kits(dwg_split, pn, size, sys_type, conn_size, conn_type, rate, i,
                    f_size, f_type):
    """
    Generates a package quote for kits without control valves. Only works through one component at a time as indicated
    by the 'i' iterative variable taken as a parameter.

    :param list dwg_split: smart kit code (drawing name) split by hyphens
    :param str pn: previously generated base part number to be edited with correct connection types
    :param str size: system size value (in letter form) from cell in this particular row from the schedule
    :param str sys_type: system type value from cell in this particular row from the schedule
    :param str conn_size: connection size value (in letter form) from cell in this particular row from the schedule
    :param str conn_type: connection type value from cell in this particular row from the schedule
    :param float rate: rate from rate cell in this particular row from the schedule
    :param int i: iterative variable indicating location of currently quoted component in drawing name
    :param f_size: indicates f component size (since it can vary based on max rates)
    :param f_type: indicates f component type (since it can only vary between sweat and threaded)
    :return:
        - pn - current base part number to be modified based on location in dwg code and sizes referenced
        - f_size - f component size (only if needed to be specified in special cases ie not matching runout size)
        - f_type - f component type (only if needed to be specified so far)
    :rtype: (str, , )
    """
    # if first component in kit
    if i == 0:
        if len(dwg_split) > 1:
            # check system f components in case SWT pipe type forces components to be THD to accommodate f components
            if '=RF' in dwg_split[1]:
                pn = pn.replace('+', 'F', 1)
            else:
                pn = typ_system_type(pn, sys_type)
            # check if f components in package to force FNPT to accommodate f components
            pn = supply_conn_type(pn, size, conn_size, conn_type, 'F' in dwg_split[1][:4])
        elif len(dwg_split) == 1:
            pn = typ_system_type(pn, sys_type)
            pn = supply_conn_type(pn, size, conn_size, conn_type)
    # if second component in kit
    elif i == 1:
        pn, f_size, f_type = return_conn_type(pn, size, conn_size, conn_type, rate, f_size, f_type,
                                              'F' in dwg_split[1][:4])
        pn = supply_conn_type(pn, size, conn_size, conn_type)
    # if third component in kit
    elif i == 2:
        # check system f components in case SWT pipe type forces components to be FNPT to accommodate f components
        if '=RF' in dwg_split[1]:
            pn = pn.replace('+', 'F', 1)
        else:
            pn = typ_system_type(pn, sys_type)
        # check if f component in package to force FNPT to accommodate f components
        pn = supply_conn_type(pn, size, conn_size, conn_type, 'F' in dwg_split[1][:4])

    return pn, f_size, f_type
