"""
All data as per Technical Instructoin TSD0019k
"""

########################################################################################################################
# Ergon Energy fuses
########################################################################################################################

# SWER Isolating TR FL < 8kA

ee_swer_isol_tr_11_11 = {
    100: 'FUSE LINK 11/22/33kV 25A CLASS K',
    200: 'FUSE LINK 11/22/33kV 40A CLASS K'
}

ee_swer_isol_tr_11_127 = {
    100: 'FUSE LINK 11/22/33kV 25A CLASS K',
    200: 'FUSE LINK 11/22/33kV 40A CLASS K'
}

ee_swer_isol_tr_11_191 = {
    100: 'FUSE LINK 11/22/33kV 31A CLASS K',
    200: 'FUSE LINK 11/22/33kV 50A CLASS K'
}

ee_swer_isol_tr_22_127 = {
    100: 'FUSE LINK 11/22/33kV 20A CLASS K',
    200: 'FUSE LINK 11/22/33kV 31A CLASS K'
}

ee_swer_isol_tr_22_191 = {
    100: 'FUSE LINK 11/22/33kV 20A CLASS K',
    200: 'FUSE LINK 11/22/33kV 31A CLASS K'
}

ee_swer_isol_tr_33_127 = {
    100: 'FUSE LINK 11/22/33kV 16A CLASS K',
    200: 'FUSE LINK 11/22/33kV 25A CLASS K'
}

ee_swer_isol_tr_33_191 = {
    100: 'FUSE LINK 11/22/33kV 16A CLASS K',
    200: 'FUSE LINK 11/22/33kV 25A CLASS K'
}

# SWER Distribution TR

ee_swer_dist_tr_11 = {
    10: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    25: 'FUSE LINK 11/22/33kV  6/20A CLASS K',
    50: 'FUSE LINK 11/22/33kV 10A CLASS K',
    75: 'FUSE LINK 11/22/33kV  10/30A CLASS K'
}

ee_swer_dist_tr_127 = {
    10: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    25: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    50: 'FUSE LINK 11/22/33kV 10A CLASS K',
    75: '1FUSE LINK 11/22/33kV 0/30A CLASS K'
}

ee_swer_dist_tr_191 = {
    10: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    25: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    50: 'FUSE LINK 11/22/33kV  6/20A CLASS K',
    75: 'FUSE LINK 11/22/33kV  6/20A CLASS K'
}

# TR 11kV EDO FL < 8kA
# Assuming TR LV fuse
ee_tr_11_1p = {
    10: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    25: 'FUSE LINK 11/22/33kV  6/20A CLASS K',
    50: 'FUSE LINK 11/22/33kV 16A CLASS K'
}
ee_tr_11_3p = {
    25: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    50: 'FUSE LINK 11/22/33kV  6/20A CLASS K',
    63: 'FUSE LINK 11/22/33kV  6/20A CLASS K',
    100: 'FUSE LINK 11/22/33kV 16A CLASS K',
    200: 'FUSE LINK 11/22/33kV 25A CLASS K',
    300: 'FUSE LINK 11/22/33kV 31A CLASS K',
    315: 'FUSE LINK 11/22/33kV 31A CLASS K',
    500: 'FUSE LINK 11/22/33kV 50A CLASS K',
    750: 'FUSE LINK 11/22/33kV 50A CLASS K'
}

# TR 22kV EDO FL < 6kA
# Assuming TR LV fuse
ee_tr_22_1p = {
    10: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    25: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    50: 'FUSE LINK 11/22/33kV 16A CLASS K'
}
ee_tr_22_3p = {
    25: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    50: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    63: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    100: 'FUSE LINK 11/22/33kV  6/20A CLASS K',
    200: 'FUSE LINK 11/22/33kV 16A CLASS K',
    300: 'FUSE LINK 11/22/33kV 20A CLASS K',
    315: 'FUSE LINK 11/22/33kV 20A CLASS K',
    500: 'FUSE LINK 11/22/33kV 31A CLASS K',
    750: 'FUSE LINK 11/22/33kV 50A CLASS K'
}

# TR 33kV EDO FL < 4kA
# Assuming TR LV fuse
ee_tr_33_1p = {
    10: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    25: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    50: 'FUSE LINK 11/22/33kV  3/10A CLASS K'
}

ee_tr_33_3p = {
    25: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    50: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    63: 'FUSE LINK 11/22/33kV 3/10A  CLASS K',
    100: 'FUSE LINK 11/22/33kV  3/10A CLASS K',
    200: 'FUSE LINK 11/22/33kV 12A CLASS K',
    300: 'FUSE LINK 11/22/33kV 16A CLASS K',
    315: 'FUSE LINK 11/22/33kV 16A CLASS K',
    500: 'FUSE LINK 11/22/33kV 20A CLASS K',
    750: 'FUSE LINK 11/22/33kV 50A CLASS K'
}

########################################################################################################################
# Energex fuses
########################################################################################################################

ex_SWER_f_sizes = {
    10: 'FUSE LINK 11/22/33kV 3A CLASS K',
    25: 'FUSE LINK 11/22/33kV 3A CLASS K'
}

ex_pole_1p_fuses = {
    0: None,
    10: 'FUSE LINK 11/22/33kV 8A CLASS T',
    15: 'FUSE LINK 11/22/33kV 8A CLASS T',
    25: 'FUSE LINK 11/22/33kV 8A CLASS T',
    50: 'FUSE LINK 11/22/33kV 16A CLASS K',
    63: 'FUSE LINK 11/22/33kV 16A CLASS K',
}

ex_pole_3p_fuses = {
    0: None,
    10: 'FUSE LINK 11/22/33kV 8A CLASS T',
    15: 'FUSE LINK 11/22/33kV 8A CLASS T',
    25: 'FUSE LINK 11/22/33kV 8A CLASS T',
    50: 'FUSE LINK 11/22/33kV 8A CLASS T',
    63: 'FUSE LINK 11/22/33kV 8A CLASS T',
    100: 'FUSE LINK 11/22/33kV 16A CLASS K',
    200: 'FUSE LINK 11/22/33kV 20A CLASS K',
    300: 'FUSE LINK 11/22/33kV 25A CLASS K',
    315: 'FUSE LINK 11/22/33kV 25A CLASS K',
    500: 'FUSE LINK 11/22/33kV 40A CLASS K',
    750: 'FUSE LINK 11/22/33kV 50A CLASS K',
    1000: 'FUSE LINK 11/22/33kV 65A CLASS K',
    1500: 'FUSE LINK 11/22/33kV 80A CLASS K',
}

ex_rmu_air_fuses = {
    0: None,
    50: 'FUSE LINK LV/11kV SIBA 30 004 03  6.3',
    100: 'FUSE LINK LV/11kV SIBA 30 004 03  10',
    200: 'FUSE LINK LV/11kV SIBA 30 004 03  16',
    300: 'FUSE LINK LV/11kV SIBA 30 004 03  25',
    315: 'FUSE LINK LV/11kV SIBA 30 004 03  25',
    500: 'FUSE LINK LV/11kV SIBA 30 004 03  40',
    '750 LowZ': 'FUSE LINK LV/11kV SIBA 30 012 03  63',
    '750 HighZ': 'FUSE LINK LV/11kV SIBA 30 012 03  63',
    '1000 LowZ': 'FUSE LINK LV/11kV Bussmann 12FXLSJ80',
    '1000 HighZ': 'FUSE LINK LV/11kV SIBA 30 012 03  63',
    '1500 LowZ': 'FUSE LINK LV/11kV SIBA 30 521 03  100',
    '1500 HighZ': 'FUSE LINK LV/11kV Bussmann 12FXLSJ80'
}

ex_rmu_oil_fuses = {
    0: None,
    50: 'FUSE LINK LV/11kV Bussmann 12OEFMA10',
    100: 'FUSE LINK LV/11kV Bussmann 12OEFMA10',
    200: 'FUSE LINK LV/11kV SIBA 30 144 36  20',
    300: 'FUSE LINK LV/11kV SIBA 30 144 36  31.5',
    315: 'FUSE LINK LV/11kV SIBA 30 144 36  31.5',
    500: 'FUSE LINK LV/11kV SIBA 30 144 36  40',
    750: 'FUSE LINK LV/11kV Bussmann 12OEFMA63',
    1000: 'FUSE LINK LV/11kV SIBA 30 144 36  80',
    1500: 'FUSE LINK LV/11kV SIBA 30 237 36  125',
}