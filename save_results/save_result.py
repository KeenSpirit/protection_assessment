from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import sys
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft
from cond_damage import apply_results as ar
from fault_study import fault_impedance
from devices import relays
from importlib import reload
reload(fault_impedance)
reload(relays)

def save_dataframe(app, region, study_selections, grid_data, all_fault_studies):
    """ saves the dataframe in the user directory.
    If the user is connected through citrix, the file should
    be saved local users PowerFactoryResults folder
    """

    import os
    import time
    project = app.GetActiveProject()
    derived_proj = project.der_baseproject
    try:
        der_proj_name = derived_proj.GetFullName()
    except AttributeError:
        der_proj_name = project.loc_name
    try:
        project_version = project.der_baseversion
    except AttributeError:
        project_version = 'NA'

    app.PrintPlain("Saving Fault Level Study...")
    date_string = time.strftime("%Y%m%d-%H%M%S")
    filename = 'Fault Study Results ' + app.GetActiveStudyCase().loc_name + ' ' + date_string + '.xlsx'
    filename = fix_string(filename)

    user = Path.home().name
    basepath = Path('//client/c$/LocalData') / user

    if basepath.exists():
        clientpath = basepath
    else:
        clientpath = Path('c:/LocalData') / user
    filepath = os.path.join(clientpath, filename)

    grid_data_df = pd.DataFrame(grid_data)
    fault_studies_pd = format_fl_results(region, all_fault_studies)

    if region == 'SEQ':
        oh_z = '0'
        ug_z = '0'
    else:
        oh_z = '50'
        ug_z = '10'

    variations = app.GetActiveNetworkVariations()

    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        # General Information sheet
        workbook = writer.book
        grid_data_df.to_excel(writer, sheet_name='General Information', startrow=21, index=False)
        worksheet = workbook['General Information']
        worksheet['A1'] = app.GetActiveStudyCase().loc_name
        worksheet['A2'] = "Fault Level Study"
        worksheet['A4'] = 'Script Run Date'
        worksheet['A5'] = date_string
        worksheet['A6'] = 'Base Project:'
        worksheet['A7'] = der_proj_name
        worksheet['A8'] = 'Used Version:'
        worksheet['A9'] = project_version.loc_name
        worksheet['A10'] = 'Fault Calculation Method: complete'
        worksheet['A12'] = 'Maximum Fault Level Study Parameters:'
        worksheet['A13'] = 'Conductor Temperature: 20 deg C'
        worksheet['A14'] = 'Voltage c factor 1.1'
        worksheet['A15'] = 'Minimum Fault Levels Study Parameters:'
        worksheet['A16'] = 'Conductor Temperature: Individual annealing temperature'
        worksheet['A17'] = 'Voltage c factor 1.0'
        worksheet['A18'] = f'Minimum phase-ground fault resistance for overhead lines set to {oh_z} ohms'
        worksheet['A19'] = f'Minimum phase-ground fault resistance for underground lines set to {ug_z} ohms'
        worksheet['A21'] = 'External Grid Data:'

        i = 0
        for feeder, key in fault_studies_pd.items():
            study_results = key[0]
            dfls_list = key[1]
            study_results.to_excel(writer, sheet_name='Study Results', startrow=i+1, index=False)
            sheet = workbook['Study Results']
            sheet[f'A{i+1}'].font = Font(size=11, bold=True)
            sheet[f'A{i+1}'] = fix_string(feeder)
            i += 15

            # Detailed Fault Levels sheet
            for j, device in enumerate(dfls_list):
                count = (j + 1) * 20 - 20
                if j == 0:
                    device_name = device.columns[1]
                else:
                    device_name = device.columns[1]
                device.to_excel(writer, sheet_name=feeder, startrow=1, startcol=count, index=False)
                sheet = workbook[feeder]
                start_col_letter = get_column_letter(count+1)
                sheet[f"{start_col_letter}1"].font = Font(size=11, bold=True)
                sheet[f"{start_col_letter}1"] = f'Primary protection: {device_name}'


            # Conductor Damage Results
            if 'Conductor Damage Assessment' in study_selections:
                devices = all_fault_studies[feeder]
                cond_damage_df = ar.cond_damage_results(devices)
                sheet_name = feeder + " Cond Dmg Res"
                cond_damage_df.to_excel(writer, sheet_name=sheet_name, startrow=0, index=False)

    app.PrintPlain("Adjusting column widths...")
    wb = load_workbook(filepath)
    ws = wb['General Information']
    adjust_col_width(ws)
    ws = wb['Study Results']
    adjust_col_width(ws)
    for feeder in fault_studies_pd:
        ws = wb[feeder]
        adjust_col_width(ws)
        if 'Conductor Damage Assessment' in study_selections:
            sheet_name = feeder + " Cond Dmg Res"
            ws = wb[sheet_name]
            adjust_col_width(ws)

    # Save the adjusted workbook
    wb.save(filepath)
    app.PrintPlain("Output file saved to " + filepath)


def format_fl_results(region, all_fault_studies):
    """

    :param grid_data:
    :param all_fault_studies:
    :return:
    """

    fault_studies_pd = {}
    for feeder, devices in all_fault_studies.items():
        study_results_df = format_study_results(devices)
        dfls_list = format_detailed_results(region, devices)

        fault_studies_pd[feeder] = [study_results_df, dfls_list]

    return fault_studies_pd


def format_study_results(devices):
    """

    :param devices:
    :return:
    """

    device_list = format_devices()
    # Store results in dictionary:
    for device in devices:
        ds_device_names = [device.object.loc_name for device in device.ds_devices]
        us_device_names = [device.object.loc_name for device in device.us_devices]
        device_list[device.object.loc_name] = [
            round(device.ds_capacity),
            round(device.max_fl_ph),
            round(device.max_fl_pg),
            round(device.min_fl_ph),
            round(device.min_fl_pg),
            device.max_ds_tr,
            round(device.max_tr_size),
            round(device.tr_max_pg),
            round(device.tr_max_ph),
            ', '.join(ds_device_names),
            ', '.join(us_device_names),

        ]

    # Format 'Study Results' data
    formatted_dev_pd = pd.DataFrame.from_dict(device_list)
    study_results_df = formatted_dev_pd.fillna("")

    return study_results_df


def format_detailed_results(region, devices):
    """

    :param devices:
    :return:
    """
    dfls_list = []
    for device in devices:
        device_reach_factors = relays.device_reach_factors(region, device)
        df = pd.DataFrame({
                device.object.loc_name: [term.object.loc_name for term in device.sect_terms],
            # FAULT LEVELS
                'Max PG fault': [term.max_fl_pg for term in device.sect_terms],
                'Max PH fault': [term.max_fl_ph for term in device.sect_terms],
                'Min PG fault': [fault_impedance.term_pg_fl(region, term) for term in device.sect_terms],
                'Min 2P fault': [term.min_fl_ph for term in device.sect_terms],
            # PICKUPS
                'EF PRI PU': device_reach_factors['ef_pickup'],
                'EF BU PU': device_reach_factors['bu_ef_pickup'],
                'PH PRI PU': device_reach_factors['ph_pickup'],
                'PH BU PU': device_reach_factors['bu_ph_pickup'],
                'NPS PRI PU': device_reach_factors['nps_pickup'],
                'NPS BU PU': device_reach_factors['bu_nps_pickup'],
            # REACH FACTORS
                'EF PRI RF': device_reach_factors['ef_rf'],
                'EF BU RF': device_reach_factors['bu_ef_rf'],
                'PH PRI RF': device_reach_factors['ph_rf'],
                'PH BU RF': device_reach_factors['bu_ph_rf'],
                'NPS EF PRI RF': device_reach_factors['nps_ef_rf'],
                'NPS EF BU RF': device_reach_factors['bu_nps_ef_rf'],
                'NPS PH PRI RF': device_reach_factors['nps_ph_rf'],
                'NPS PH BU RF': device_reach_factors['bu_nps_ph_rf']
            })
        # Sort fault levels by Max PG fault
        df_sorted = df.sort_values(by=df.columns[1], ascending=False)
        df_sorted.insert(0, 'Tfmr Size (kVA)', '')

        tr_dict = {load.term.loc_name:round(load.load_kva) for load in device.sect_loads}
        df_sorted['Tfmr Size (kVA)'] = df_sorted[device.object.loc_name].map(tr_dict).fillna('')

        dfls_list.append(df_sorted)

    return dfls_list


def format_devices() -> dict[str:list]:
    """

    """

    device_list = {
        'Site Name':
            [
                'DS capacity  (kVA)',
                'Max Ph FL',
                'Max PG FL',
                'Min Ph FL',
                'Min PG FL',
                'Max DS TR (Site name)',
                'Max TR size (kVA)',
                'TR Max Ph ',
                'TR max PG',
                'DS devices',
                'BU device'
             ]
    }

    return device_list


def fix_string(file_name):
    """
    Excel does not allow special characters in file names. Remove any such cases and replace with a '_'
    :param file_name: string
    :return:
    """
    forbidden_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    file_name_list = list(file_name)
    for i in range (len(file_name_list)):
        if file_name_list[i] in forbidden_chars:
            file_name_list[i] = '_'
    return ''.join(file_name_list)


def adjust_col_width(ws):
    """
    Adjust column width of given Excel sheet
    :param ws:
    :return:
    """

    # Adjust the column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name

        # Iterate over all cells in the column to find the maximum length
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass

        # Set the column width to the maximum length found
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

