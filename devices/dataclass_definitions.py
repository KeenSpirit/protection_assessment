from dataclasses import dataclass


@dataclass
class Device:
    object: object
    cubicle: object
    term: object
    phases: int
    l_l_volts: float
    ds_capacity: (object, None)
    max_fl_ph: (object, None)
    max_fl_pg: (object, None)
    min_fl_ph: (object, None)
    min_fl_pg: (object, None)
    max_ds_tr: (object, None)
    sect_terms: list
    sect_loads: list
    sect_lines: list
    us_devices: list
    ds_devices: list


@dataclass
class Line:
    object: object
    phases: int
    l_l_volts: float
    max_fl_ph: float
    max_fl_pg: float
    min_fl_ph: float
    min_fl_pg: float
    line_type: str
    thermal_rating: float
    ph_energy: float
    ph_clear_time: float
    ph_fl: float
    pg_energy: float
    pg_clear_time: float
    pg_fl: float


@dataclass
class Termination:
    object: object
    constr: (None, str)
    phases: int
    l_l_volts: float
    max_fl_ph: float
    max_fl_pg: float
    min_fl_ph: float
    min_fl_pg: float
    min_fl_pg10: float
    min_fl_pg50: float


@dataclass
class Tfmr:
    obj: object
    term: object
    load_kva: (float, None)
    max_ph: (str, None)
    max_pg: (str, None)
    fuse: (str, None)
    insulation: (str, None)
    impedance: (str, None)