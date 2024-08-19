
def get_prot_elements(device):
    """ Get all of the time overcurrent and instantneous overcurrent
    protection elements that are active and return useful attributes
    """
    prot_elements = ['RelToc', 'RelIoc']

    elements = [element for element in device.GetContents('*.Rel?oc', 1)
                if element.GetAttribute('outserv') == 0
                if element.GetParent().GetAttribute('outserv') == 0
                if element.GetClassName() in prot_elements]

    return elements


def get_active_elements(elements, fault_type: str):
    """
    From a list of protection elements get those elements that are capable of detecting the fault type
    elements str: 'Phase-Ground', '2-Phase', '3-Phase'
    """

    earth_fault_type = ['1ph', '3I0', 'S3I0', 'd1m', 'I0']
    negative_sequence_type = ['I2', '3I2']
    phase_type = ['3ph', 'phA', 'phB', 'phC', 'd3m']

    active_elements = []
    for element in elements:
        element_type = element.typ_id
        if fault_type == 'Phase-Ground':
            # all elements are active
            active_elements.append(element)
        elif fault_type == '2-Phase':
            if element_type in negative_sequence_type or phase_type:
                # Only negative sequence and phase elements are active
                active_elements.append(element)
        elif fault_type == '3-Phase':
            if element_type in phase_type:
                # Only phase elements are active
                active_elements.append(element)

    return active_elements


def get_fuse_current(fuse):
    """

    """
    return fuse.GetAttribute("c:labs")


def get_measured_current(element, fault_level, fault_type):
    """ For a given element, look at the meaurement type according to the
    elements type. Then for the element get the appropriate current that
    the relay is using to make a decision.
    """
    MeasurementType = element.typ_id.atype

    if MeasurementType in ['3ph', 'd3m']:  # 3 phase current
        return fault_level
    elif MeasurementType in ['3I0', 'S3I0']:  # Earth current & sensitive earth current
        return convert_to_i0(fault_level, threei0=True)
    elif MeasurementType in ['I0', '1ph']:  # Zero seq current & 1 phase current
        return fault_level
    elif MeasurementType in ['d1m']:  # 1 phase current
        return fault_level
    elif MeasurementType in ['I2']:  # Neg seq. current
        return convert_to_i2(fault_level, fault_type, threei2=False)
    elif MeasurementType in ['3I2']:  # 3 x neg seq current
        return convert_to_i2(fault_level, fault_type, threei2=True)
    else:
        # If an unhandelled measurement type ie encountered return the 3 phase current
        return fault_level


def convert_to_i2(fault_current, fault_type: str, threei2=False):
    """
    Convert 2 phase fault or earth fault phase current to negative sequence current.
    This function assumes 2 phase fault angles tend towards 180 degrees apart.
    fault_type str: '3-phase', '2-Phase','Phase-Ground'.
    """

    if fault_type == '2-Phase':
        ia = complex(0, fault_current)
        ib = complex(0, -fault_current)
    elif fault_type == 'Phase-Ground':
        ia = complex(0, fault_current)
        ib = complex(0, 0)
    else:
        return 0
    ic = complex(0, 0)

    a = complex(-0.5, 0.866)
    a2 = complex(-0.5, -0.866)

    aib = ib * a
    a2ib = ib * a2
    aic = ic * a
    a2ic = ic * a2

    ia2 = (ia + a2ib + aic) / 3
    ib2 = ia2 * a
    ic2 = ia2 * a2

    ia2 = abs(ia2)
    ib2 = abs(ib2)
    ic2 = abs(ic2)

    if not threei2:
        result = (ia2 + ib2 + ic2) / 3
    else:
        result = (ia2 + ib2 + ic2)
    return result

def convert_to_i0(fault_current, threei0=False):
    """
    """
    if threei0:
        return 3 * fault_current
    else:
        return fault_current
