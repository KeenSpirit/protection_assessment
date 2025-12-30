import sys
from pf_config import pft
from typing import List, Dict, Union


def relay_checks(app: pft.Application, relays: List):
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


def relay_type_check(elmrelay: pft.ElmRelay, relay_issues_detected: Dict):
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


def ct_phase_check(elmrelay: pft.ElmRelay, relay_issues_detected: Dict) -> Dict:
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