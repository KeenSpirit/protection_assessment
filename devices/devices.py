from dataclasses import dataclass
from save_results import map_functs as mf


def create_fuse(app, relay):


    tr_size = relay.ds_tr_size
    typfuse = get_fuse_element(app, tr_size)
    if not typfuse:
        return []
    fuse_name = relay.ds_tr_site
    equip = app.GetProjectFolder("equip")
    fuse_folder = equip.GetContents("Fuses.IntFolder", True)[0]
    lib_contents = fuse_folder.GetContents("*.RelFuse", 0)
    if lib_contents:
        for fuse in lib_contents:
            if fuse.loc_name == fuse_name:
                return [fuse]
    rel_fuse = fuse_folder.CreateObject("RelFuse", f'{fuse_name}')
    rel_fuse.loc_name = fuse_name
    rel_fuse.SetAttribute("typ_id", typfuse)
    return [rel_fuse]


def get_fuse_element(app, tr_size: int):
    fuse_object = None
    fuse_types = f_types(app)
    try:
        fuse_string = pole_fuse_sizes[tr_size]
    except KeyError:
        return None
    for fuse in fuse_types:
        if fuse.loc_name == fuse_string:
            fuse_object = fuse
            break
    return fuse_object


def f_types(app):
    # Create a list of all the fuse types
    ergon_lib = app.GetGlobalLibrary()
    fuse_folder = ergon_lib.SearchObject(r"\ErgonLibrary\Protection\Fuses.IntFolder")
    fuse_types = fuse_folder.GetContents("*.TypFuse", 0)
    return fuse_types


def format_devices(study_results, user_selection, site_name_map):

    """

    :param study_results:
    :param site_name_map:
    :return:
    """

    (sect_phase_max, sect_pg_max, sect_phase_min, sect_pg_min, sect_tr_phase_max, sect_tr_pg_max, device_max_load,
     us_devices, ds_devices, ds_capacity, device_lines) = study_results
    selected_devices = list(user_selection.values())[0]

    def str_check(dict):
        if dict == 'no terminations':
            return 0
        (value,) = dict.values()
        return value

    device_list = []
    # Store results in dictionary:
    for device, value in site_name_map.items():
        if [device][0] in selected_devices:
            (elmterm,) = value.values()
            max_ph_fl = str_check(sect_phase_max[elmterm])
            (max_pg_fl,) = sect_pg_max[elmterm].values()
            min_ph_fl = str_check(sect_phase_min[elmterm])
            (min_pg_fl,) = sect_pg_min[elmterm].values()
            tr_max_ph = str_check(sect_tr_phase_max[elmterm])
            if sect_tr_pg_max[elmterm] == 'no terminations':
                tr_max_pg = 0
                tr_max_name = ""
            else:
                tr_max_pg = next(iter(sect_tr_pg_max[elmterm].values()))
                tr_max_name = next(iter(sect_tr_pg_max[elmterm]))
                tr_max_name = tr_max_name.cpSubstat.loc_name
            max_tr_size = device_max_load[elmterm]
            ds_device_names = [mf.term_element(device, site_name_map, element=True) for device in ds_devices[elmterm]]
            us_device_names = [mf.term_element(device, site_name_map, element=True) for device in us_devices[elmterm]]
            section_lines = device_lines[elmterm]

            device = Device(
                device,
                round(max_ph_fl),
                round(max_pg_fl),
                round(min_ph_fl),
                round(min_pg_fl),
                tr_max_name,
                round(max_tr_size),
                round(tr_max_pg),
                round(tr_max_ph),
                ds_device_names,
                us_device_names,
                section_lines
                )
            device_list.append(device)

    return device_list


@dataclass
class Device:
    object: object
    max_fl_ph: float
    max_fl_pg: float
    min_fl_ph: float
    min_fl_pg: float
    ds_tr_site: str
    ds_tr_size: int
    ds_tr_ph: float
    ds_tr_pg: float
    ds_devices: list
    us_device: list
    sect_device: list


SWER_f_sizes = {
    10: 'FUSE LINK 11/22/33kV 3A CLASS K',
    25: 'FUSE LINK 11/22/33kV 3A CLASS K'
}

pole_fuse_sizes = {
    0: None,
    10: 'FUSE LINK 11/22/33kV 8A CLASS T',
    15: 'FUSE LINK 11/22/33kV 8A CLASS T',
    25: 'FUSE LINK 11/22/33kV 8A CLASS T',
    50: 'FUSE LINK 11/22/33kV 8A CLASS T',
    63: 'FUSE LINK 11/22/33kV 8A CLASS T',
    100: 'FUSE LINK 11/22/33kV 3A CLASS K',
    200: 'FUSE LINK 11/22/33kV 20A CLASS K',
    300: 'FUSE LINK 11/22/33kV 25A CLASS K',
    315: 'FUSE LINK 11/22/33kV 25A CLASS K',
    500: 'FUSE LINK 11/22/33kV 40A CLASS K',
    750: 'FUSE LINK 11/22/33kV 50A CLASS K',
    1000: 'FUSE LINK 11/22/33kV 65A CLASS K',
    1500: 'FUSE LINK 11/22/33kV 80A CLASS K',
}

rmu_fuse_sizes = {
    0: None,
    300: '300_35.5A_Air',
    315: '300_35.5A_Air',
    500: '500_40A_Air',
    750: '750_63A_Air',
    1000: '1000_80A_Air',
    1500: '1500_100A_Air'
}





