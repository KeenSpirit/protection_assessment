from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
import sys
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft
from cond_damage import apply_results as ar


def save_dataframe(app, grid_data, all_fault_studies):
    """ saves the dataframe in the user directory.
    If the user is connected through citrix, the file should
    be saved local users PowerFactoryResults folder
    """
    import os
    import time

    app.PrintPlain("Saving Fault Level Study...")
    date_string = time.strftime("%Y%m%d-%H%M%S")
    filename = 'Fault Study Results ' + app.GetActiveStudyCase().loc_name + ' ' + date_string + '.xlsx'

    user = Path.home().name
    basepath = Path('//client/c$/LocalData') / user

    if basepath.exists():
        clientpath = basepath
    else:
        clientpath = Path('c:/LocalData') / user
    filepath = os.path.join(clientpath, filename)

    grid_data_df = pd.DataFrame(grid_data)
    fault_studies_pd = format_fl_results(all_fault_studies)

    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        # General Information sheet
        grid_data_df.to_excel(writer, sheet_name='General Information', startrow=9, index=False)
        workbook = writer.book
        worksheet = workbook['General Information']
        worksheet['A1'] = app.GetActiveStudyCase().loc_name
        worksheet['A2'] = "Fault Level Study"
        worksheet['A4'] = 'Script Run Date'
        worksheet['A5'] = date_string
        worksheet['A8'] = 'External Grid Data:'

        i = 0
        for feeder, key in fault_studies_pd.items():
            study_results = key[0]
            dfls_list = key[1]
            study_results.to_excel(writer, sheet_name='Study Results', startrow=i+1, index=False)
            sheet = workbook['Study Results']
            sheet[f'A{i+1}'].font = Font(size=11, bold=True)
            sheet[f'A{i+1}'] = feeder
            i += 15

            # Detailed Fault Levels sheet
            for j, device in enumerate(dfls_list):
                count = (j + 1) * 6 - 6
                device.to_excel(writer, sheet_name=feeder, startrow=0, startcol=count, index=False)

            # Conductor Damage Results
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
        sheet_name = feeder + " Cond Dmg Res"
        ws = wb[sheet_name]
        adjust_col_width(ws)

    # Save the adjusted workbook
    wb.save(filepath)
    app.PrintPlain("Output file saved to " + filepath)


def format_fl_results(all_fault_studies):
    """

    :param grid_data:
    :param all_fault_studies:
    :return:
    """

    fault_studies_pd = {}
    for feeder, devices in all_fault_studies.items():
        study_results_df = format_study_results(devices)
        dfls_list = format_detailed_results(devices)

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


def format_detailed_results(devices):
    """

    :param devices:
    :return:
    """

    dfls_list = []
    for device in devices:
        df = pd.DataFrame({
                device.object.loc_name: [term.object.loc_name for term in device.sect_terms],
                'Max PG fault': [term.max_fl_pg for term in device.sect_terms],
                'Max 3P fault': [term.max_fl_ph for term in device.sect_terms],
                'Min PG fault': [term.min_fl_pg for term in device.sect_terms],
                'Min 2P fault': [term.min_fl_ph for term in device.sect_terms]
            })
        # Sort fault levels by Max PG fault
        df_sorted = df.sort_values(by=df.columns[1], ascending=False)
        df_sorted.insert(0, 'Tfmr Size (kVA)', '')

        tr_dict = {load.term.loc_name:load.load_kva for load in device.sect_loads}
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