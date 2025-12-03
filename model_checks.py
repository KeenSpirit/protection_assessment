import sys


def chk_empty_fdrs(app, feeders_devices):
    """
    Check that the selected feeders have protection devices created.
    :param app:
    :param feeders_devices:
    :return:
    """

    empty_feeders = [feeder for feeder, devices in feeders_devices.items() if devices == []]

    if len(empty_feeders) == len(feeders_devices):
        app.PrintError("No protection devices were detected in the model for the selected feeders. \n"
                       "Please add and configure the required protection devices and re-run the script.")
        sys.exit(0)
    for empty_feeder in empty_feeders:
            app.PrintWarn(f"No protection devices were detected in the model for feeder {empty_feeder}. \n"
                          "This feeder will be excluded from the study.")
            del feeders_devices[empty_feeder]
    return feeders_devices


def relay_checks(app, relays) -> dict:
    """ Relays must be configured with the correct information"""

    relay_issues_detected = {}
    for device in relays:
        if device.GetClassName() != 'ElmRelay':
            continue
        relay_issues_detected = relay_type_check(device, relay_issues_detected)
        if not relay_issues_detected:
            relay_issues_detected = ct_phase_check(device, relay_issues_detected)
    if relay_issues_detected:
        app.PrintWarn(f"Warnings: {relay_issues_detected}")


def relay_type_check(elmrelay, relay_issues_detected):
    """
    The relay must have relay type assigned
    :param elmrelay:
    :param relay_issues_detected:
    :return:
    """

    if elmrelay.GetAttribute("e:typ_id") is None:
        _add_issue(
            relay_issues_detected, elmrelay.loc_name, f"{elmrelay} has no typ_id assigned"
        )
    return relay_issues_detected


def ct_phase_check(elmrelay, relay_issues_detected) -> dict:
    """
    Each relay must have a CT wired to it, and the relay measured phase count must equal the CT phase count.
    """
    # elmrelay CT must be present.
    if elmrelay.GetAttribute("e:cpCt") is None:
        _add_issue(
            relay_issues_detected, elmrelay.loc_name, f"{elmrelay} has no CT assigned"
        )
    else:
        # elmrelay CT phase count must match relay measured phases
        ct = elmrelay.GetAttribute("e:cpCt")
        ct_phases = ct.GetAttribute("e:iphase")
        typmeas = elmrelay.GetContents("*.RelMeasure")[0].GetAttribute("typ_id")
        measure_type = typmeas.GetAttribute("atype")
        three_phase = ['3rms', '3drm', '3pui', '3dui', 'ulil']
        one_phase = ['1rms', '1pui', '1ph', 'selr']
        if ct_phases == 3 and measure_type in one_phase:
            _add_issue(
                relay_issues_detected, ct.loc_name, f"{elmrelay} measured phase count != CT phase count"
            )
        if ct_phases == 1 and measure_type in three_phase:
            _add_issue(
                relay_issues_detected, ct.loc_name, f"{elmrelay} measured phase count != CT phase count"
            )
    return relay_issues_detected


def _add_issue(dictionary, key, value):
    """Takes a dictionary, a key, and a value as input and appends the
    key and value or initializes the dictionary if unable to append"""
    try:
        dictionary[key].append(value)
    except KeyError:
        dictionary[key] = [value]
    return dictionary