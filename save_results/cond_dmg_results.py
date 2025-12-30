import math
import pandas as pd
from relays import reclose
from importlib import reload
reload(relays)

def cond_damage_results(devices):
    """
    For saving to Excel file
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
        trips = reclose.get_device_trips(device.obj)
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