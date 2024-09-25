import math
import pandas as pd
from devices import relays

def rewrite_results(app, lines, fault_type, trips):
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
            fault_level = line.ph_fl
            clear_time = line.ph_clear_time
            dpl_num = "dpl1"
        else:
            fault_level = line.pg_fl
            clear_time = line.pg_clear_time
            dpl_num = "dpl2"
        try:
            allowable_fl = line.thermal_rating / (math.sqrt(clear_time) * (trips + 1))
            if fault_level < allowable_fl:
                # No conductor damage
                line.object.SetAttribute(f"e:{dpl_num}", 2)
                # app.PrintPlain(f"line: {line.object.loc_name} {fault_type}: GREEN")
            elif fault_level > allowable_fl:
                # Conductor damage
                line.object.SetAttribute(f"e:{dpl_num}", 1)
                # app.PrintPlain(f"line: {line.object.loc_name} {fault_type}: RED")
                # app.PrintPlain(f"rating: {line.thermal_rating}, clear time: {clear_time}")
                # app.PrintPlain(f"fault_level: {fault_level}, allowable_fl: {allowable_fl}")
        except Exception:
            # No data
            line.object.SetAttribute(f"e:{dpl_num}", 0)
            app.PrintPlain(f"line: {line.object.loc_name} {fault_type}: GREY")
            app.PrintPlain(f"rating: {line.thermal_rating}, clear time: {clear_time}")


def cond_damage_results(devices):
    """

    :param devices:
    :return:
    """

    def _allowable_fl(thermal_rating, clear_time, trips):
        try:
            allowable_fl = round(thermal_rating / (math.sqrt(clear_time) * (trips + 1)))
        except:
            allowable_fl = None
        return allowable_fl

    def _damage_test(fl, _allowable_fl, thermal_rating, clear_time, trips):
        acceptable_fl = _allowable_fl(thermal_rating, clear_time, trips)
        if not acceptable_fl:
            return "NO DATA"
        elif fl > acceptable_fl:
            return "FAIL"
        else :
            return "PASS"

    line_list = []
    for device in devices:
        trips = relays.get_device_trips(device.object)
        list_length = len(device.sect_lines)
        line_df = pd.DataFrame({
            "Device":
                [device.object.loc_name]*list_length,
            "Trips":
                trips,
            "Line":
                [line.object.loc_name for line in device.sect_lines],
            "Line Type":
                [line.line_type for line in device.sect_lines],
            "Phase fault level":
                [line.ph_fl for line in device.sect_lines],
            "Phase fault clear time":
                [line.ph_clear_time for line in device.sect_lines],
            "Allowable phase fault level":
                [_allowable_fl(line.thermal_rating, line.ph_clear_time, trips) for line in device.sect_lines],
            "Phase fault conductor damage":
                [_damage_test(line.ph_fl, _allowable_fl, line.thermal_rating, line.ph_clear_time, trips) for line in device.sect_lines],
            "Ground fault level":
                [line.pg_fl for line in device.sect_lines],
            "Ground fault clear time":
                [line.pg_clear_time for line in device.sect_lines],
            "Allowable ground fault level":
                [_allowable_fl(line.thermal_rating, line.pg_clear_time, trips) for line in device.sect_lines],
            "Ground fault conductor damage":
                [_damage_test(line.pg_fl, _allowable_fl, line.thermal_rating, line.pg_clear_time, trips) for line in device.sect_lines],
        })
        line_list.append(line_df)
    cond_damage_df = pd.concat(line_list)

    return cond_damage_df
