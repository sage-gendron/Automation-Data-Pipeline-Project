# controls.py
"""
author:Sage Gendron

"""


def control_type_1(pn, cv_size, act_signal):
    """
    Uses pn, cv_size, act_signal parameters from schedule to use control valve part number to either concatenate with,
    or create list with, actuator. Actuators only selected based on data from act_signal and size of control valve.

    :param str pn: control valve part number taken from schedule selection
    :param str cv_size: control valve size taken from schedule selection
    :param str act_signal: actuator signal type taken from schedule selection
    :returns:
        - cv_parts (:py:class:'list') - control valve part number in index 0 and actuator in index 1 if separate;
                                        else control valve/actuator in index 0
        - cv_qtys (:py:class:'list') - contains an int(1) for length of cv_parts
    :rtype: (list, list)
    """
    # instantiate variables to be returned
    cv_parts: list = []
    cv_qtys: list = []

    # use cv_pn function to get actuator required
    pn = control_1_parts(pn, cv_size, act_signal)

    # add parts and part quantities to package list
    if type(pn) == str:
        cv_parts.append(pn)
        cv_qtys.append(1)
    else:
        cv_parts.extend(pn)
        for n in range(len(pn)):
            cv_qtys.append(1)

    return cv_parts, cv_qtys


def control_1_parts(pn, cv_size, act_signal):
    """
    Takes given control information and selects an actuator. \n
    Correctly creates a second part number for separated actuator bodies or adds the actuator to the control valve
    part number for VE411s/VE431s.

    :param str pn: control valve part number taken from schedule selection
    :param str cv_size: control valve size taken from schedule selection
    :param str act_signal: actuator signal type taken from schedule selection
    :returns: cv_pn_list (:py:class:'list') - control valve part number in index 0 and actuator in index 1 if separate;
                                              else control valve/actuator in index 0
    :rtype: list
    """
    # instantiate actuator dictionaries based on CV model
    v243_act = {'ON/OFF, FC': 'ME4530W', 'ON/OFF, FO': 'ME4430W', 'MOD, FLP': 'ME4340', 'MOD, FC': 'ME4940',
                'MOD, FO': 'ME4840'}
    v321_act = {'ON/OFF, FC': 'ME4430W', 'ON/OFF, FO': 'ME4530W', 'MOD, FLP': 'ME4340', 'MOD, FC': 'ME4840',
                'MOD, FO': 'ME4940'}
    v411_act = {'ON/OFF, FC': {'1"': '5430', '1-1/2"': '5430', '2"': '5440', '3"': '5850-ON', '4"': '5850-ON'},
                'ON/OFF, FO': {'1"': '5430', '1-1/2"': '5430', '2"': '5440', '3"': '5850-ON', '4"': '5850-ON'},
                'MOD, FLP': {'1"': '5330', '1-1/2"': '5330', '2"': '5340', '3"': '5350', '4"': '5350'},
                'MOD, FC': {'1"': '5830', '1-1/2"': '5830', '2"': '5840', '3"': '5850-ON', '4"': '5850-ON'},
                'MOD, FO': {'1"': '5830', '1-1/2"': '5830', '2"': '5840', '3"': '5850-ON', '4"': '5850-ON'}
                }

    cv_pn_list: list = []

    # if cv pn starts with V243 select actuator from v243_act
    if pn.startswith('V243'):
        cv_pn_list.extend([pn, v243_act[act_signal]])
    # if cv pn starts with V321 select actuator from v321_act
    elif pn.startswith('V321'):
        cv_pn_list.extend([pn, v321_act[act_signal]])
    # if cv pn starts with VE select actuator from v411_act (also includes V431 actuators)
    elif pn.startswith('VE'):
        cv_pn_list.append(f"{pn}-{v411_act[act_signal][cv_size]}")
    # if cv pn starts with Z (is a Belimo zone CV), return just the cv pn
    elif pn.startswith('Z'):
        return cv_pn_list.append(pn)
    # else, raise an error if control valve pn is not included in the above
    else:
        raise Exception('CV part number not found. Please review selected part numbers.')

    return cv_pn_list


def control_type_2(pn, act_signal):
    """
    Uses cv_pn and act_signal columns from schedule to use control valve part number to create a list with actuator.
    Actuators only selected based on data from act_signal and PICV part number.

    :param str pn: PICV part number taken from schedule selection
    :param str act_signal: actuator signal type taken from schedule selection
    :returns:
        - cv_parts (:py:class:'list') - PICV part number in index 0 and actuator in index 1
        - cv_qtys (:py:class:'list') - contains an int(1) for length of cv_parts
    :rtype: (list, list)
    """
    # instantiate variables to be returned
    cv_parts: list[str] = []
    cv_qtys: list[int] = []

    # use picv_pn function to get actuator required
    pn = control_2_parts(pn, act_signal)

    # add parts and part quantities to package list
    if type(pn) == str:
        cv_parts.append(pn)
        cv_qtys.append(1)
    else:
        cv_parts.extend(pn)
        for _ in range(len(pn)):
            cv_qtys.append(1)

    return cv_parts, cv_qtys


def control_2_parts(pn, act_signal):
    """
    Takes given PICV information and selects an actuator.

    :param str pn: PICV part number taken from schedule selection
    :param str act_signal: actuator signal type taken from schedule selection
    :returns: picv_pn_list (:py:class:'list') - PICV part number in index 0 and actuator in index 1
    :rtype: list
    """
    # instantiate actuator dictionaries based on PICV model
    ninetwo_actuator_dict: dict[str, str] = {'ON/OFF, FC': 'ME4530W', 'ON/OFF, FO': 'ME4430W',
                                             'MOD, FLP': 'VA-7482-8002-RA', 'MOD, FC': 'ME4940', 'MOD, FO': 'ME4840'}
    eightfive_actuator_dict: dict[str, str] = {'MOD, FLP': 'VA9310-HGA-2', 'MOD, FC': 'VA9208-GGA-2',
                                               'MOD, FO': 'VA9208-GGA-2'}

    picv_pn_list: list[str] = [pn]

    # if picv pn starts with T92 select actuator from ninetwo_actuator_dict
    if pn.startswith('T92'):
        picv_pn_list.append(ninetwo_actuator_dict[act_signal])
    # if PICV pn starts with T85 select actuator from eightfive_actuator_dict
    elif pn.startswith('T85'):
        picv_pn_list.append(eightfive_actuator_dict[act_signal])
    # if 94 PICV, return just PICV pn
    elif pn.startswith('T94'):
        return picv_pn_list
    # else, raise an error if PICV pn is not included in the above
    else:
        raise Exception('PICV part number not found. Please review selected part numbers.')

    return picv_pn_list
