import math
from devices import relays
from cond_damage import conditional_formatting as cf
from cond_damage import apply_results as ar
import logging

from importlib import reload
reload(relays)
reload(cf)
reload(ar)


def cond_damage(app, devices):

    cf.set_up(app)
    fl_step = 10

    for device in devices:
        dev_obj = device.object
        app.PrintPlain(f"Performing conductor damage assessment for {dev_obj.loc_name}...")
        lines = device.sect_lines
        total_trips = relays.get_device_trips(dev_obj)
        fault_type = '2-Phase'
        for line in lines:
            relays.reset_reclosing(dev_obj)
            trip_count = 1
            total_energy = 0
            while trip_count <= total_trips:
                block_service_status = relays.set_enabled_elements(app, dev_obj)
                min_fl_clear_times, _ = fault_clear_times(app, device, line, fl_step, fault_type)
                max_energy, max_fl, max_clear_time = worst_case_energy(line, min_fl_clear_times, fault_type, device,
                                                                       False)
                relays.reset_block_service_status(block_service_status)
                trip_count  = relays.trip_count(dev_obj, increment=True)
                total_energy += max_energy
                if max_clear_time is not None:
                    line.ph_clear_time = max_clear_time
                    line.ph_fl = max_fl
                else:
                    logging.info(f"{dev_obj.loc_name} {fault_type} trip {trip_count} "
                                 f"fault clearing time calculation error.")
            line.ph_energy = total_energy
        ar.rewrite_results(app, lines, fault_type)

        line_fault_type = 'Phase-Ground'
        for line in lines:
            relays.reset_reclosing(dev_obj)
            trip_count = 1
            total_energy = 0
            while trip_count <= total_trips:
                block_service_status = relays.set_enabled_elements(app, dev_obj)
                min_fl_clear_times, device_fault_type = fault_clear_times(app, device, line, fl_step, line_fault_type)
                transposition = False
                if line_fault_type != device_fault_type:
                    transposition = True
                max_energy, max_fl, max_clear_time = worst_case_energy(line, min_fl_clear_times, fault_type, device,
                                                                       transposition)
                relays.reset_block_service_status(block_service_status)
                trip_count  = relays.trip_count(dev_obj, increment=True)
                total_energy += max_energy
                if max_clear_time is not None:
                    line.pg_clear_time = max_clear_time
                    line.pg_fl = max_fl
                else:
                    logging.info(f"{dev_obj.loc_name} {fault_type} trip {trip_count} "
                                 f"fault clearing time calculation error.")
            line.pg_energy = total_energy
        ar.rewrite_results(app, lines, line_fault_type)


def fault_clear_times(app, device, line, fl_step, fault_type):
    """

    :param app:
    :param device:
    :param line:
    :param fl_step:
    :param fault_type: 'Phase-Ground', '2-Phase', '3-Phase'
    :return: dictionary {
    fl_1: trip_time_1, fl_2: trip_time_2...}
    """

    if fault_type in ['2-Phase', '3-Phase']:
        min_fl = line.min_fl_ph
        max_fl = line.max_fl_ph
    else:
        # Check if this is a SWER line, and does the device see the same current?
        min_fl, max_fl, fault_type = swer_transform(app, device, line, fault_type)

    device_obj = device.object
    # Create a list of fault levels in the interval of min and max fault
    # currents. Two intervals may be assessed:
    # fl_interval_1 is composed of equidistant step sizes between
    # min fault level and max fault level
    # fl_interval_2 is composed of only the element hisets between
    # min fault level and max fault level

    fl_interval_1 = range(min_fl, max_fl + 1, fl_step)
    # Initialise fl_interval_2
    # fl_interval_2 = [min_fl, max_fl]

    # Select only the elements capable of detecting the fault type
    # and enabled for the current auto-reclose iteration
    if device_obj.GetClassName() == 'RelFuse':
        active_elements = [device_obj]
    else:
        elements = relays.get_prot_elements(device_obj)
        active_elements = relays.get_active_elements(elements, fault_type)
        # hisets = [
        #     element.GetAttribute("e:cpIpset") - 1 for element in active_elements
        #           if element.GetClassName() == 'RelIoc']
        # fl_interval_2 = fl_interval_2 + hisets

    # Initialise fault level:min operating time dictionary
    min_fl_clear_times = {fault_level: None for fault_level in fl_interval_1}
    for element in active_elements:
        for fault_level in fl_interval_1:
            # Calculate protection operate time for element and fl
            if element.GetClassName() == 'RelFuse':
                operate_time = fuse_clear_time(element, fault_level)
                switch_operate_time = 0
            else:
                element_current = relays.get_measured_current(
                    element, fault_level, fault_type)
                operate_time = element_trip_time(element, element_current)
                switch_operate_time = 0.05
            if not operate_time or operate_time <= 0:
                continue
            clear_time = operate_time + switch_operate_time
            # If this is the minimum fault clear time for that fault level,
            # update the dictionary accordingly
            if (min_fl_clear_times[fault_level] is None
                    or clear_time < min_fl_clear_times[fault_level]):
                min_fl_clear_times[fault_level] = round(clear_time, 3)

    return min_fl_clear_times, fault_type


def swer_transform(app, device, line, fault_type):
    """

    :param device:
    :param line:
    :param fault_type:
    :return:
    """

    min_fl = line.min_fl_pg
    max_fl = line.max_fl_pg

    line_type = line.object.typ_id
    if (('SWER' in line_type.loc_name)
            and line.phases == 1
            and device.phases > 1):
            min_fl = round((line.l_l_volts * min_fl / device.l_l_volts) / math.sqrt(3))
            max_fl = round((line.l_l_volts * max_fl / device.l_l_volts) / math.sqrt(3))
            fault_type = '2-Phase'
    return min_fl, max_fl, fault_type


def worst_case_energy(line, min_fl_clear_times, fault_type, device, transpose):
    """
    line: [min_fl, max_fl]
    Input: min_fl_op_times = {fl_1: trip_time_1, fl_2: trip_time_2...}
    Output: [fault level, operating time] that corresponded
    to worst case energy for the line
    """

    max_energy = 0
    max_fl = None
    max_clear_time = None
    max_pair = [None, None]

    for fl, clear_time in min_fl_clear_times.items():
        if clear_time is None:
            continue
        energy = fl ** 2 * clear_time
        if energy > max_energy:
            max_energy = energy
            max_fl = fl
            max_clear_time = clear_time

    # If transposed due to SWER, need to transpose back to what is seen by the line
    if fault_type == 'Phase-Ground' and transpose and max_fl is not None:
        # reverse the fl transposition
        max_fl = round((math.sqrt(3) * max_fl * device.l_l_volts) / line.l_l_volts)
    return max_energy, max_fl, max_clear_time



def fuse_clear_time(fuse, flt_cur):
    """
    For a given fault current, use fuse element setting attributes
    to calculate the fault clear time.
    Interpolates linearly between a list of [fault level-total clear time]
    values representing a hermite polynomial
    :param fuse: RelFuse element
    :param flt_cur: float, integer: fault current (A)
    :return: float :fuse total clear time (sec)
    """

    op_time = None

    type_fuse = fuse.GetAttribute("e:typ_id")
    # melt curve
    typechatoc = type_fuse.GetAttribute("e:pmelt")
    # curve type
    curve_type = typechatoc.GetAttribute("e:i_type")
    # curve equation variables
    curve_var = typechatoc.GetAttribute("e:vmat")

    number_of_rows = len(curve_var)

    def total_clear_time(k, p):
        interpolate_1 = ((flt_cur - curve_var[k][p])
                         / (curve_var[k + 1][p] - curve_var[k][p]))
        interpolate_2 = ((curve_var[k][p+1] - curve_var[k + 1][p+1])
                         * interpolate_1)
        total_clear = curve_var[k][p+1] - interpolate_2
        return total_clear

    # Hermite Polynomial
    if curve_type == 6:
        curve_count = typechatoc.GetAttribute("e:i_curves")
        if curve_count == 1:
            p = curve_count - 1
        elif curve_count == 2:
            p = curve_count
        else:
            # Unhandled curve count
            return op_time
        if flt_cur < curve_var[0][p]:
            return op_time
        if flt_cur > curve_var[(number_of_rows - 1)][p]:
            return curve_var[(number_of_rows - 1)][p+1]
        k = 0
        while k < (number_of_rows - 1):
            if curve_var[k][p] <= flt_cur <= curve_var[k + 1][p]:
                op_time = total_clear_time(k, p)
                break
            k += 1
    else:
        # Unhandled curve type
        return op_time
    return op_time


def element_trip_time(element, flt_cur):
    """
    For a given fault current,
    use the relay element setting attributes to calculate the operate time.
    :param element: RelToc or RelIoc element
    :param flt_cur: float, integer: fault current (A)
    :return min_fl_op_times: float: relay operate time (sec)
    """

    op_time = None

    if element.GetClassName() == 'RelToc':
        pickup = element.GetAttribute("e:cpIpset")
        if flt_cur <= pickup:
            return op_time
        time_dial = element.GetAttribute("e:Tpset")
        # curve characteristic
        curve_char = element.GetAttribute("e:pcharac")
        # curve type
        curve_type = curve_char.GetAttribute("e:i_type")
        # curve equation variables
        curve_var = curve_char.GetAttribute("e:vmat")

        # Definite time
        if curve_type == 0:
            a1 = curve_var[0][0]
            op_time = time_dial * a1
        # IEC 255-3
        elif curve_type == 1:
            a1 = curve_var[0][0]
            a2 = curve_var[1][0]
            a3 = curve_var[2][0]
            op_time = time_dial * a1 / ((flt_cur / pickup) ** a2 - a3)
        # ANSI/IEEE
        elif curve_type == 2:
            a1 = curve_var[0][0]
            a2 = curve_var[1][0]
            a3 = curve_var[2][0]
            a4 = curve_var[3][0]
            op_time = time_dial * (a1 / ((flt_cur / pickup) ^ a2 - a3) + a4)
        # ANSI/IEEE squared
        elif curve_type == 3:
            a1 = curve_var[0][0]
            a2 = curve_var[1][0]
            op_time = (time_dial * a1 + a2)/((flt_cur / pickup) ^ 2)
        # ABB/Westinghouse
        elif curve_type == 4:
            a1 = curve_var[0][0]
            a2 = curve_var[1][0]
            a3 = curve_var[2][0]
            a4 = curve_var[3][0]
            a5 = curve_var[4][0]
            if (flt_cur / pickup) >= 1.5:
                op_time = ((a1 + a2)/(((flt_cur / pickup) - a3) ^ a4)
                           * time_dial / 24000)
            else:
                op_time = a5 / (flt_cur / pickup - 1) * time_dial / 24000
        # Linear approximation
        elif curve_type == 5:
            # Unhandled curve type
            pass
        # Hermite Polynomial
        elif curve_type == 6:
            number_of_rows = len(curve_var)
            i_ip = flt_cur / pickup
            curve_count = curve_char.GetAttribute("e:i_curves")  # number of curves

            def clear_time(m):
                interpolate_1 = (i_ip - curve_var[m][0]) / (curve_var[m + 1][0] - curve_var[m][0])
                interpolate_2 = ((curve_var[m][1] - curve_var[m + 1][1])
                                 * interpolate_1)
                t1 = curve_var[m][1] - interpolate_2
                return t1

            if curve_count > 1:
                # Unhandled curve count
                return op_time
            if i_ip < curve_var[0][0]:
                return op_time
            if i_ip > curve_var[(number_of_rows - 1)][0]:
                return curve_var[(number_of_rows - 1)][1]
            k = 0
            while k < (number_of_rows - 1):
                if curve_var[k][0] <= i_ip <= curve_var[k + 1][0]:
                    op_time = clear_time(k) * time_dial
                    break
                k += 1
        # DSL - Equation
        elif curve_type == 7:
            # Unhandled curve type
            pass
        # Special Equation
        elif curve_type == 8:
            a1 = curve_var[0][0]
            a2 = curve_var[1][0]
            a3 = curve_var[2][0]
            b1 = curve_var[3][0]
            b2 = curve_var[4][0]
            b3 = curve_var[5][0]
            op_time = ((time_dial * a1) / (((flt_cur / pickup) + b1)
                                           ^ b2 + b3)
                       + time_dial * a2 + a3)
        # I sqr T (based on Ir)
        elif curve_type == 9:
            # Unhandled curve type
            pass
        # I sqr T (based on In)
        elif curve_type == 10:
            # Unhandled curve type
            pass
        # I sqr T (based on Ip)
        elif curve_type == 11:
            # Unhandled curve type
            pass
    elif element.GetClassName() == 'RelIoc':
        min_time = element.GetAttribute("e:cptotime")
        pickup = element.GetAttribute("e:cpIpset")
        if flt_cur >= pickup:
            op_time = min_time

    return op_time