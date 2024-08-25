import math


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
            fault_level = line.ph_fl
            clear_time = line.ph_clear_time
            dpl_num = "dpl1"
        else:
            fault_level = line.pg_fl
            clear_time = line.pg_clear_time
            dpl_num = "dpl2"
        try:
            allowable_fl = line.thermal_rating / math.sqrt(clear_time)
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
