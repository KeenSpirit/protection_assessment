import math
import logging
import pandas as pd
from devices import relays

def rewrite_results(app, lines, fault_type):
    """

    :param app:
    :param lines:
    :param fault_type:
    :return:
    """
    # Based on the selected folder create a list of available matrices
    app.SetGraphicUpdate(0)
    for line in lines:
        if fault_type == '2-Phase':
            line_energy = line.ph_energy
            dpl_num = "dpl1"
        else:
            line_energy = line.pg_energy
            dpl_num = "dpl2"
        try:
            line_type = line.obj.typ_id
            if fault_type == '2-Phase' and 'SWER' in line_type.loc_name:
                line.obj.SetAttribute(f"e:{dpl_num}", 3)
                continue
        except AttributeError:
            pass
        try:
            allowable_energy = line.thermal_rating ** 2 * 1
            if line_energy < allowable_energy:
                # No conductor damage
                line.obj.SetAttribute(f"e:{dpl_num}", 2)
            elif line_energy > allowable_energy:
                # Conductor damage
                line.obj.SetAttribute(f"e:{dpl_num}", 1)
        except Exception:
            # No data
            line.obj.SetAttribute(f"e:{dpl_num}", 0)
            logging.exception(f"{line.obj.loc_name} No data")
            logging.info(line.__dict__)


def cond_damage_results(devices):
    """

    :param devices:
    :return:
    """

    def _allowable_fl(thermal_rating, clear_time, trips):
        try:
            allowable_fl = round(thermal_rating / (math.sqrt(clear_time * trips)))
        except:
            allowable_fl = None
        return allowable_fl

    def _damage_test(line, _allowable_fl, no_trips, fault_type):
        try:
            line_type = line.obj.typ_id
            if fault_type == 'Phase' and 'SWER' in line_type.loc_name:
                return "SWER"
        except AttributeError:
            pass
        thermal_rating = line.thermal_rating
        if fault_type == 'Phase':
            fl = line.ph_fl
            clear_time = line.ph_clear_time
        else:
            fl = line.pg_fl
            clear_time = line.pg_clear_time
        acceptable_fl = _allowable_fl(thermal_rating, clear_time, no_trips)
        if not acceptable_fl:
            return "NO DATA"
        elif fl > acceptable_fl:
            return "FAIL"
        else :
            return "PASS"

    line_list = []
    for device in devices:
        trips = relays.get_device_trips(device.obj)
        list_length = len(device.sect_lines)
        line_df = pd.DataFrame({
            "Device":
                [device.obj.loc_name]*list_length,
            "Trips":
                trips,
            "Line":
                [line.obj.loc_name for line in device.sect_lines],
            "Line Type":
                [line.line_type for line in device.sect_lines],
            "Worst case energy ph flt lvl":
                [line.ph_fl for line in device.sect_lines],
            "Worst case energy ph flt clear time":
                [line.ph_clear_time for line in device.sect_lines],
            "Allowable phase fault level":
                [_allowable_fl(line.thermal_rating, line.ph_clear_time, trips) for line in device.sect_lines],
            "Phase fault conductor damage":
                [_damage_test(line, _allowable_fl, trips, fault_type='Phase') for line in device.sect_lines],
            "Worst case energy gnd flt lvl":
                [line.pg_fl for line in device.sect_lines],
            "Worst case energy gnd flt clear time":
                [line.pg_clear_time for line in device.sect_lines],
            "Allowable ground fault level":
                [_allowable_fl(line.thermal_rating, line.pg_clear_time, trips) for line in device.sect_lines],
            "Ground fault conductor damage":
                [_damage_test(line, _allowable_fl, trips, fault_type='Ground') for line in device.sect_lines],
        })
        line_list.append(line_df)
    cond_damage_df = pd.concat(line_list)

    return cond_damage_df
