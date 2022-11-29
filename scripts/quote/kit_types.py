# kit_types.py
from connection_types import typ_system_type, ctrl_conn_type, supply_conn_type, return_conn_type

"""
author: Sage Gendron

"""


def type_two_kits(components, pn, size, sys_type, conn_size, conn_type, cv_size_type, rate, ta_size, i, f_size,
                  f_type):
    """
    Builds two-way kit component part number with correct end connections and sizes matching given information.

    :param list components: smart kit code (drawing name) split by hyphens
    :param str pn: current base part number to be modified based on location in dwg code and sizes referenced
    :param str size: system size value (in letter form) from cell in this particular row from the schedule
    :param str sys_type: system type value from cell in this particular row from the schedule
    :param str conn_size: connection size value (in letter form) from cell in this particular row from the schedule
    :param str conn_type: connection type value from cell in this particular row from the schedule
    :param str cv_size_type: cv size/type value from cell in this particular row from the schedule
    :param float rate: rate value from cell in this particular row from the schedule
    :param int i: iterated index of current letter in drawing code
    :param f_size: f component size (only if needed to be specified in special cases ie not matching runout size)
    :param f_type: f component type (only if needed to be specified so far)
    :return:
        - pn - current base part number to be modified based on location in dwg code and sizes referenced
        - f_size - f component size (only if needed to be specified in special cases ie not matching runout size)
        - f_type - f component type (only if needed to be specified so far)
    :rtype: (str, , )
    """
    if i == 1:
        # check runout hoses in case SWT pipe type forces components to be THD to accommodate hose
        if '=RHC' in components[1]:
            pn = pn.replace('+', 'F', 1)
        else:
            pn = typ_system_type(pn, sys_type)

        # check if f component in package to force FNPT
        if 'F' in components[1][:4]:
            pn = supply_conn_type(pn, size, conn_size, conn_type, True)
        else:
            pn = supply_conn_type(pn, size, conn_size, conn_type)
    elif i == 2:
        # check if f component in package to force FNPT
        if 'F' in components[1][:4]:
            pn, f_size, f_type = return_conn_type(pn, size, conn_size, conn_type, rate, f_size, f_type,
                                                  True)
            pn = ctrl_conn_type(pn, size, cv_size_type, conn_size, True)
        else:
            pn, f_size, f_type = return_conn_type(pn, size, conn_size, conn_type, rate, f_size, f_type)
            pn = ctrl_conn_type(pn, size, cv_size_type, conn_size)
    elif i == 3:
        pn = typ_system_type(pn, sys_type)

        # check for supply control to alter connection type
        if '=SCV' in components[1]:
            pn = supply_conn_type(pn, size, conn_size, conn_type)
        else:
            # check if f component in package to force FNPT
            if 'F' in components[1][:4]:
                pn = ctrl_conn_type(pn, size, cv_size_type, conn_size, True)
            else:
                pn = ctrl_conn_type(pn, size, cv_size_type, conn_size)

    return pn, f_size, f_type


def type_three_kits(components, pn, pipe_size, pipe_type, coil_size, coil_type, cv_size_type, rate, ta_size, i, hose_size,
                    hose_type):
    """
    Builds three-way kit component part number with correct end connections and sizes matching given information.

    :param list components: smart kit code (drawing name) split by hyphens
    :param str pn: current base part number to be modified based on location in dwg code and sizes referenced
    :param str pipe_size: runout size value (in letter form) from cell in this particular row from the schedule
    :param str pipe_type: runout type value from cell in this particular row from the schedule
    :param str coil_size: coil size value (in letter form) from cell in this particular row from the schedule
    :param str coil_type: coil type value from cell in this particular row from the schedule
    :param str cv_size_type: cv size/type value from cell in this particular row from the schedule
    :param float rate: flow rate from rate cell in this particular row from the schedule
    :param ta_size: list of additional components required if TA different size than kit
    :param int i: iterative variable to identify index of part/location of component being quoted
    :param hose_size: hose size (only if needed to be specified in special cases ie not matching runout size)
    :param hose_type: hose type (only if needed to be specified so far)
    :return:
        - pn - current base part number to be modified based on location in dwg code and sizes referenced
        - hose_size - hose size (only if needed to be specified in special cases ie not matching runout size)
        - hose_type - hose type (only if needed to be specified so far)
    :rtype: (str, , )
    """
    if i == 1:
        # check runout hoses in case SWT pipe type forces components to be THD to accommodate hose
        if '=RHC' in components[1]:
            pn = pn.replace('+', 'F', 1)
        else:
            pn = typ_system_type(pn, pipe_type)

        # check for supply control valve to alter tailpiece type
        if '=SCV' in components[1]:
            pn = ctrl_conn_type(pn, pipe_size, cv_size_type, coil_size)
        # check for press coils and if hoses are required (as this is handled differently than hoses or press alone)
        elif coil_type == 'PRESS' and 'H' in components[1][:4]:
            pn = supply_conn_type(pn, pipe_size, coil_size, coil_type, True)
        # check for hoses without press coil type
        elif 'H' in components[1][:4]:
            pn = supply_conn_type(pn, pipe_size, coil_size, coil_type, True)
        else:
            pn = supply_conn_type(pn, pipe_size, coil_size, coil_type)
    elif i == 2:
        # check for hoses to force FNPT
        if 'H' in components[1][:4]:
            pn, hose_size, hose_type = return_conn_type(pn, pipe_size, coil_size, coil_type, rate, hose_size, hose_type,
                                                        True)
            pn = ctrl_conn_type(pn, pipe_size, cv_size_type, coil_size, True)
        else:
            pn, hose_size, hose_type = return_conn_type(pn, pipe_size, coil_size, coil_type, rate, hose_size, hose_type)
            pn = ctrl_conn_type(pn, pipe_size, cv_size_type, coil_size)
    elif i == 3:
        # check to ensure TA size did not have to get altered in typ_component() or ss_component()
        if not ta_size:
            # check runout hoses in case SWT pipe type forces components to be THD to accommodate hose
            if '=RHC' in components[1]:
                pn = pn.replace('+', 'F', 1)
            else:
                pn = typ_system_type(pn, pipe_type)
        # check for hoses to force FNPT
        if 'H' in components[1][:4]:
            pn = ctrl_conn_type(pn, pipe_size, cv_size_type, coil_size, True)
        else:
            pn = ctrl_conn_type(pn, pipe_size, cv_size_type, coil_size)
    elif i == 4:
        pn = typ_system_type(pn, pipe_type)
        # check for manual balance valve on the three-way bypass, so it can force it to match runout type
        if components[0][4] == 'B':
            pn = typ_system_type(pn, pipe_type)
        else:
            # change to bypass component matching CV size when available
            if 'H' in components[1][:4]:
                pn = ctrl_conn_type(pn, pipe_size, cv_size_type, coil_size, True)
            else:
                pn = ctrl_conn_type(pn, pipe_size, cv_size_type, coil_size)

    return pn, hose_size, hose_type


def no_control_kits(component_length, components, pn, pipe_size, pipe_type, coil_size, coil_type, rate, ta_size, i,
                    f_size, f_type):
    """
    Generates a package quote for kits without control valves. Only works through one component at a time as indicated
    by the 'i' iterative variable taken as a parameter.

    :param int component_length: length of the smart kit code split by hyphens. If only len of 1 (single component), treat differently
    :param list components: smart kit code (drawing name) split by hyphens
    :param str pn: previously generated base part number to be edited with correct connection types
    :param str pipe_size: runout size value (in letter form) from cell in this particular row from the schedule
    :param str pipe_type: runout type value from cell in this particular row from the schedule
    :param str coil_size: coil size value (in letter form) from cell in this particular row from the schedule
    :param str coil_type: coil type value from cell in this particular row from the schedule
    :param float rate: flow rate from rate cell in this particular row from the schedule
    :param int i: iterative variable indicating location of currently quoted component in drawing name
    :param f_size: indicates f component size (since it can vary based on union size / max flow rates)
    :param f_type: indicates f component type (since it can only vary between sweat and threaded)
    :return:
        - pn - current base part number to be modified based on location in dwg code and sizes referenced
        - f_size - hose size (only if needed to be specified in special cases ie not matching runout size)
        - f_type - hose type (only if needed to be specified so far)
    :rtype: (str, , )
    """
    # if first component in kit
    if i == 0:
        if component_length > 1:
            # check runout hoses in case SWT pipe type forces components to be THD to accommodate hose
            if '=RHC' in components[1]:
                pn = pn.replace('+', 'F', 1)
            else:
                pn = typ_system_type(pn, pipe_type)

            # check if f components in package to force FNPT to accommodate hoses
            if 'F' in components[1][:4]:
                pn = supply_conn_type(pn, pipe_size, coil_size, coil_type, True)
            else:
                pn = supply_conn_type(pn, pipe_size, coil_size, coil_type)
        elif component_length == 1:
            pn = typ_system_type(pn, pipe_type)
            pn = supply_conn_type(pn, pipe_size, coil_size, coil_type)
    # if second component in kit
    elif i == 1:
        # check if f components in package to force FNPT to accommodate hoses
        if 'F' in components[1][:4]:
            pn, f_size, f_type = return_conn_type(pn, pipe_size, coil_size, coil_type, rate, f_size, f_type,
                                                  True)
        else:
            pn, f_size, f_type = return_conn_type(pn, pipe_size, coil_size, coil_type, rate, f_size, f_type)
        pn = supply_conn_type(pn, pipe_size, coil_size, coil_type)
    # if third component in kit
    elif i == 2:
        # check runout hoses in case SWT pipe type forces components to be FNPT to accommodate hoses
        if '=RHC' in components[1]:
            pn = pn.replace('+', 'F', 1)
        else:
            pn = typ_system_type(pn, pipe_type)
        # check if f component in package to force FNPT to accommodate hoses
        if 'F' in components[1][:4]:
            pn = supply_conn_type(pn, pipe_size, coil_size, coil_type, True)
        else:
            pn = supply_conn_type(pn, pipe_size, coil_size, coil_type)

    return pn, f_size, f_type
