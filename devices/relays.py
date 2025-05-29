import math
from fault_study import fault_impedance

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
        element_type = element.typ_id.atype
        if fault_type == 'Phase-Ground':
            # all elements are active
            active_elements.append(element)
        elif fault_type == '2-Phase':
            if element_type in negative_sequence_type or element_type in phase_type:
                # Only negative sequence and phase elements are active
                active_elements.append(element)
        elif fault_type == '3-Phase':
            if element_type in phase_type:
                # Only phase elements are active
                active_elements.append(element)

    return active_elements


# def device_pickup(device, fault_type: str):
#     """
#     Obtain the device pickup for the given fault type
#     fault_type: 'Phase-Ground', '2-Phase', '3-Phase'
#     """
#
#     if device.GetClassName() == 'ElmRelay':
#         elements = get_prot_elements(device)
#         active_elements = get_active_elements(elements, fault_type)
#
#         minimum_pickup = 9999
#         for element in active_elements:
#             if element.GetClassName() == 'RelToc':
#                 pickup = element.GetAttribute("e:cpIpset")
#             elif element.GetClassName() == 'RelIoc':
#                 pickup = element.GetAttribute("e:cpIpset")
#             else:
#                 pickup = 9999
#             if pickup < minimum_pickup:
#                 minimum_pickup = pickup
#         return minimum_pickup
#     else:
#         # Device is a fuse
#         pickup = device.irat * 2
#     return minimum_pickup


def determine_pickup_values(device_pf):
    """The values returned from this function will be used to calculate the
    reach factor."""
    # If the devices is a fuse then it will have a known size. A fuse factor
    # of two will be applied
    if device_pf.GetClassName() == "RelFuse":
        fuse_size = int(device_pf.GetAttribute("r:typ_id:e:irat"))
        setting_values = [fuse_size * 2, fuse_size * 2, 0]
        return setting_values
    # It has been assumed that only IDMT elements will be used to reach for
    # Phase and earth faults.
    idmt_elements = [
        idmt_element
        for idmt_element in device_pf.GetContents("*.RelToc", True)
        if not idmt_element.GetAttribute("e:outserv")
    ]
    # idmt_elements =  device_pf.GetContents('*.RelToc', True)
    # app.PrintInfo(f"relay = {device_pf}, Element = {idmt_elements}")
    # Determine the OC pickup
    oc_idmt_elements = [
        oc_idmt_element
        for oc_idmt_element in idmt_elements
        if oc_idmt_element.GetAttribute("r:typ_id:e:sfiec") == "I>t"
        if "definite" not in oc_idmt_element.pcharac.loc_name.lower()
    ]
    # Not all devices use IDMT elements. If a relay does not have a configured
    # IDMT then look to include the INST element to perform reach.
    if not oc_idmt_elements:
        oc_idmt_elements = [
            oc_inst_element
            for oc_inst_element in device_pf.GetContents("*.RelIoc", True)
            if oc_inst_element.GetAttribute("r:typ_id:e:sfiec") == "I>>"
            if not oc_inst_element.IsOutOfService()
            if oc_inst_element.GetAttribute("r:typ_id:e:irecltarget")
        ]
    # Select the largest pickup. This is assuming you can have multiple pickups
    # and one trip is dependent on this particular setting
    highest_oc_pickup = 0
    for oc_idmt_element in oc_idmt_elements:
        pickup = oc_idmt_element.GetAttribute("e:cpIpset")
        if pickup > highest_oc_pickup:
            highest_oc_pickup = pickup
    # Determine the EF pickup
    ef_idmt_elements = [
        ef_idmt_element
        for ef_idmt_element in idmt_elements
        if ef_idmt_element.GetAttribute("r:typ_id:e:sfiec") == "IE>t"
        if "definite" not in ef_idmt_element.pcharac.loc_name.lower()
    ]
    if not ef_idmt_elements:
        ef_idmt_elements = [
            ef_inst_element
            for ef_inst_element in device_pf.GetContents("*.RelIoc", True)
            if ef_inst_element.GetAttribute("r:typ_id:e:sfiec") == "IE>>"
            if not ef_inst_element.IsOutOfService()
            if ef_inst_element.GetAttribute("r:typ_id:e:irecltarget")
        ]
    # Select the largest pickup. This is assuming you can have multiple pickups
    # and one trip is dependent on this particular setting
    highest_ef_pickup = 0
    for ef_idmt_element in ef_idmt_elements:
        pickup = ef_idmt_element.GetAttribute("e:cpIpset")
        if pickup > highest_ef_pickup:
            highest_ef_pickup = pickup
    # Determine the NPS pickup
    nps_idmt_elements = [
        nps_idmt_element
        for nps_idmt_element in idmt_elements
        if nps_idmt_element.GetAttribute("r:typ_id:e:sfiec") == "I2>t"
    ]
    nps_inst_elements = [
        nps_inst_element
        for nps_inst_element in device_pf.GetContents("*.RelIoc", True)
        if nps_inst_element.GetAttribute("r:typ_id:e:sfiec") == "I2>>"
        if nps_inst_element.GetAttribute("r:typ_id:e:irecltarget")
        if not nps_inst_element.IsOutOfService()
    ]
    nps_elements = nps_idmt_elements + nps_inst_elements
    # Select the largest pickup. This is assuming you can have multiple pickups
    # and one trip is dependent on this particular setting
    highest_nps_pickup = 0
    for nps_idmt_element in nps_elements:
        pickup = nps_idmt_element.GetAttribute("e:cpIpset")
        if pickup > highest_nps_pickup:
            highest_nps_pickup = pickup
    if highest_nps_pickup > 0.1:
        highest_nps_pickup = 0
    setting_values = [round(highest_oc_pickup), round(highest_ef_pickup), round(highest_nps_pickup)]
    return setting_values


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
        return convert_to_i0(fault_level, threei0=False)
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


def measure_type(device):
    pass


def get_device_trips(device):
    """Get number of trips for each device."""

    if device.GetClassName() == 'ElmRelay':
        try:
            reclosing_element = device.GetContents("*.RelRecl", True)[0]
            trips = reclosing_element.GetAttribute("oplockout")
        except IndexError:
            trips = 1  # It's a feeder relay
    else:
        trips = 1  # device  is a fuse
    return trips


def device_reach_factors(region, device):
    """

    :param device:
    :return:
    """

    # PRIMARY PICKUPS
    ef_pickup = determine_pickup_values(device.object)[1]
    ph_pickup = determine_pickup_values(device.object)[0]
    nps_pickup = determine_pickup_values(device.object)[2]

    # Phase elements can see earth faults
    if ef_pickup > 0 and ph_pickup > 0:
        effective_ef_pickup = min(ef_pickup, ph_pickup)
    elif ph_pickup > 0:
        effective_ef_pickup = ph_pickup
    elif ef_pickup > 0:
        effective_ef_pickup = ef_pickup
    else:
        effective_ef_pickup = 0

    # PRIMARY REACH FACTORS
    if effective_ef_pickup > 0:
        ef_rf = []
        for term in device.sect_terms:
            term_fl_pg = fault_impedance.term_pg_fl(region, term)
            device_fl = swer_transform(device, term, term_fl_pg)
            if device_fl != term_fl_pg:
                # device is seeing 2P fault current from a 1P SWER term
                ef_rf.append(round(device_fl / ph_pickup, 2))
            else:
                ef_rf.append(round(device_fl / effective_ef_pickup, 2))
    else:
        ef_rf = ['NA'] * len(device.sect_terms)
    if ph_pickup > 0:
        ph_rf = [round(term.min_fl_ph / ph_pickup, 2) for term in device.sect_terms]
    else:
        ph_rf = ['NA'] * len(device.sect_terms)
    if nps_pickup > 0:
        nps_ef_rf = []
        for term in device.sect_terms:
            term_fl_pg = fault_impedance.term_pg_fl(region, term)
            device_fl = swer_transform(device, term, term_fl_pg)
            if  device_fl == term_fl_pg:
                # There is no SWER, the device sees earth fault
                nps_ef_rf.append(round(device_fl / 3 / nps_pickup, 2))
            else:
                # There is SWER, the device sees 2 phase fault current
                nps_ef_rf.append(round(device_fl / math.sqrt(3) / nps_pickup, 2))
        nps_ph_rf = [round(term.min_fl_ph/math.sqrt(3) / nps_pickup, 2) for term in device.sect_terms]
    else:
        nps_ef_rf = ['NA'] * len(device.sect_terms)
        nps_ph_rf = ['NA'] * len(device.sect_terms)

    # BACK-UP PICKUPS
    if device.us_devices:
        # Obtain the lowest pickup setting of all bu devices
        bu_devices = device.us_devices
        bu_ef_pickup = None
        bu_ph_pickup = None
        bu_nps_pickup = None
        for bu_device in bu_devices:
            bu_ef_pickup_bu_device = determine_pickup_values(bu_device.object)[1]
            bu_ph_pickup_bu_device = determine_pickup_values(bu_device.object)[0]
            bu_nps_pickup_bu_device = determine_pickup_values(bu_device.object)[2]
            if not bu_ef_pickup or (bu_ef_pickup and bu_ef_pickup_bu_device < bu_ef_pickup):
                bu_ef_pickup = bu_ef_pickup_bu_device
            if not bu_ph_pickup or (bu_ph_pickup and bu_ph_pickup_bu_device < bu_ph_pickup):
                bu_ph_pickup = bu_ph_pickup_bu_device
            if not bu_nps_pickup or (bu_nps_pickup and bu_nps_pickup_bu_device < bu_nps_pickup):
                bu_nps_pickup = bu_nps_pickup_bu_device

        # Phase elements can see earth faults
        if bu_ef_pickup > 0 and bu_ph_pickup > 0:
            effective_bu_ef_pickup = min(bu_ef_pickup, bu_ph_pickup)
        elif bu_ph_pickup > 0:
            effective_bu_ef_pickup = bu_ph_pickup
        elif bu_ef_pickup > 0:
            effective_bu_ef_pickup = bu_ef_pickup
        else:
            effective_bu_ef_pickup = 0

        # BACK-UP REACH FACTORS
        if effective_bu_ef_pickup > 0:
            bu_ef_rf = []
            for term in device.sect_terms:
                term_fl_pg = fault_impedance.term_pg_fl(region, term)
                bu_device_fl = swer_transform(bu_devices[0], term, term_fl_pg)
                if bu_device_fl != term_fl_pg:
                    # device is seeing 2P fault current from a 1P SWER term
                    bu_ef_rf.append(round(bu_device_fl / bu_ph_pickup, 2))
                else:
                    bu_ef_rf.append(round(bu_device_fl / effective_bu_ef_pickup, 2))
        else:
            bu_ef_rf = ['NA'] * len(device.sect_terms)
        if bu_ph_pickup > 0:
            bu_ph_rf = [round(term.min_fl_ph / bu_ph_pickup, 2) for term in device.sect_terms]
        else:
            bu_ph_rf = ['NA'] * len(device.sect_terms)
        if bu_nps_pickup > 0:
            bu_nps_ef_rf = []
            for term in device.sect_terms:
                term_fl_pg = fault_impedance.term_pg_fl(region, term)
                bu_device_fl = swer_transform(bu_devices[0], term, term_fl_pg)
                if bu_device_fl == term_fl_pg:
                    # There is no SWER, the device sees earth fault
                    bu_nps_ef_rf.append(round(bu_device_fl / 3 / bu_nps_pickup, 2))
                else:
                    # There is SWER, the device sees 2 phase fault current
                    bu_nps_ef_rf.append(round(bu_device_fl / math.sqrt(3) / bu_nps_pickup, 2))
            bu_nps_ph_rf = [round(term.min_fl_ph / math.sqrt(3) / bu_nps_pickup, 2) for term in device.sect_terms]
        else:
            bu_nps_ef_rf = ['NA'] * len(device.sect_terms)
            bu_nps_ph_rf = ['NA'] * len(device.sect_terms)
    else:
        # NO BACK-UP
        bu_device = None
        bu_ef_pickup = 'NA'
        bu_ph_pickup = 'NA'
        bu_nps_pickup = 'NA'
        bu_ef_rf = ['NA'] * len(device.sect_terms)
        bu_ph_rf = ['NA'] * len(device.sect_terms)
        bu_nps_ef_rf = ['NA'] * len(device.sect_terms)
        bu_nps_ph_rf = ['NA'] * len(device.sect_terms)

    device_reach_factors = {
        'ef_pickup' : [ef_pickup] * len(device.sect_terms),
        'ph_pickup' : [ph_pickup] * len(device.sect_terms),
        'nps_pickup' : [nps_pickup] * len(device.sect_terms),
        'ef_rf' : ef_rf,
        'ph_rf' : ph_rf,
        'nps_ef_rf' : nps_ef_rf,
        'nps_ph_rf' : nps_ph_rf,
        'bu_ef_pickup' : [bu_ef_pickup] * len(device.sect_terms),
        'bu_ph_pickup' : [bu_ph_pickup] * len(device.sect_terms),
        'bu_nps_pickup' : [bu_nps_pickup] * len(device.sect_terms),
        'bu_ef_rf' : bu_ef_rf,
        'bu_ph_rf' : bu_ph_rf,
        'bu_nps_ef_rf' : bu_nps_ef_rf,
        'bu_nps_ph_rf' : bu_nps_ph_rf
    }
    return device_reach_factors


def swer_transform(device, term, term_fl_pg):
    """
    The fault level seen by the device depends on any voltage and phase transformations that occur between the device
    and the fault.
    Currently, this function handles single phase SWER only.
    :param device:
    :param term:
    :return:
    """

    if term.l_l_volts != device.l_l_volts and term.phases == 1 and device.phases > 1:
        device_fl = (term.l_l_volts * term_fl_pg / device.l_l_volts) / math.sqrt(3)
    else:
        device_fl = term_fl_pg
    return device_fl