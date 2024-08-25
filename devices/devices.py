from dataclasses import dataclass


def create_fuse(app, relay):


    tr_size = relay.max_tr_size
    typfuse = get_fuse_element(app, tr_size)
    if not typfuse:
        return []
    fuse_name = relay.max_ds_tr
    equip = app.GetProjectFolder("equip")
    protection = equip.GetContents("Protection.IntFolder", True)
    if not protection:
        protection = equip.CreateObject("IntFolder", "Protection")
    else:
        protection = protection[0]
    fuse_folder = protection.GetContents("Fuses.IntFolder", True)
    if not fuse_folder:
        fuse_folder = protection.CreateObject("IntFolder", "Fuses")
    else:
        fuse_folder = fuse_folder[0]
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


@dataclass
class Device:
    object: object
    cubicle: object
    term: object
    ds_capacity: float
    max_fl_ph: float
    max_fl_pg: float
    min_fl_ph: float
    min_fl_pg: float
    max_ds_tr: str
    max_tr_size: int
    tr_max_ph: float
    tr_max_pg: float
    sect_terms: list
    sect_loads: list
    sect_lines: list
    us_devices: list
    ds_devices: list


@dataclass
class Line:
    object: object
    max_fl_ph: float
    max_fl_pg: float
    min_fl_ph: float
    min_fl_pg: float
    line_type: str
    thermal_rating: float
    ph_clear_time: float
    ph_fl: float
    pg_clear_time: float
    pg_fl: float


@dataclass
class Termination:
    object: object
    max_fl_ph: float
    max_fl_pg: float
    min_fl_ph: float
    min_fl_pg: float


@dataclass
class Load:
    object: object
    term: object
    load_kva: float


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





