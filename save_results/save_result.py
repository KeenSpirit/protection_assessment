from pathlib import Path
import os
import time
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
import sys
from pf_config import pft
from fault_study import fault_impedance
from relays.reach_factors import device_reach_factors
from save_results import cond_dmg_results as cd
from importlib import reload
import re
import numpy as np

reload(fault_impedance)
reload(relays)
reload(cd)

def save_dataframe(app, region, study_selections, external_grid, feeders):
    """ saves the dataframe in the user directory.
    If the user is connected through citrix, the file should
    be saved local users PowerFactoryResults folder
    """

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

    # PowerFactory model external grid data read to put in a pd.DataFrame
    formatted_grid_data = format_grid_data(external_grid)
    grid_data_df = pd.DataFrame(formatted_grid_data)
    grid_data_df = clean_dataframe(grid_data_df)

    fault_studies_pd = format_fl_results(app, region, feeders)

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
        grid_data_df.to_excel(writer, sheet_name='General Information', startrow=26, index=False)
        worksheet = workbook['General Information']

        # Use safe_set_cell for all cell assignments
        safe_set_cell(worksheet, 'A1', str(app.GetActiveStudyCase().loc_name))
        safe_set_cell(worksheet, 'A2', "Fault Level Study")
        safe_set_cell(worksheet, 'A4', 'Script Run Date-Time')
        safe_set_cell(worksheet, 'A5', str(date_string))
        safe_set_cell(worksheet, 'A6', 'Base Project:')
        safe_set_cell(worksheet, 'A7', str(der_proj_name))
        safe_set_cell(worksheet, 'A8', 'Used Version:')
        safe_set_cell(worksheet, 'A9',
                      str(project_version.loc_name if hasattr(project_version, 'loc_name') else project_version))
        safe_set_cell(worksheet, 'A10', 'Fault Calculation Method: complete')
        safe_set_cell(worksheet, 'A12', 'Maximum Fault Level Study Parameters:')
        safe_set_cell(worksheet, 'A13', 'Conductor Temperature: 20 deg C')
        safe_set_cell(worksheet, 'A14', 'Voltage c factor 1.1')
        safe_set_cell(worksheet, 'A15', 'Minimum Fault Levels Study Parameters:')
        safe_set_cell(worksheet, 'A16', 'Conductor Temperature: Individual annealing temperature')
        safe_set_cell(worksheet, 'A17', 'Voltage c factor 1.0')
        safe_set_cell(worksheet, 'A18',
                      f'Minimum phase-ground fault resistance for overhead lines set to {oh_z} ohms')
        safe_set_cell(worksheet, 'A19',
                      f'Minimum phase-ground fault resistance for underground lines set to {ug_z} ohms')
        safe_set_cell(worksheet, 'A21', f'Conductor Damage Study Parameters:')
        safe_set_cell(worksheet, 'A22',
                      f'Total energy let-through is summed across autoreclose trip sequence.')
        safe_set_cell(worksheet, 'A23',
                      f'Tabulated fault clearing time shown is for final trip of the trip sequence.')
        safe_set_cell(worksheet, 'A25', 'External Grid Data:')

        i = 1
        for feeder, key in fault_studies_pd.items():
            study_results = key[0]
            fdr_open_points = key[1]
            dfls_list = key[2]

            # Clean the study results DataFrame
            study_results = clean_dataframe(study_results)
            fdr_open_points = clean_dataframe(fdr_open_points)

            col_letter = get_column_letter(i)
            study_results.to_excel(writer, sheet_name='Summary Results', startrow=2, startcol=i-1, index=False)
            fdr_open_points.to_excel(writer, sheet_name='Summary Results', startrow=21, startcol=i-1, index=False)
            sheet = workbook['Summary Results']
            sheet[f'{col_letter}1'].font = Font(size=12, bold=True)
            safe_set_cell(sheet, f'{col_letter}1', fix_string(str(feeder)))
            study_results_len =len(study_results.columns)
            i = i + study_results_len + 1

            # Create safe sheet name for detailed fault levels
            safe_feeder_name = create_safe_sheet_name(f"{str(feeder)} Detailed Results")

            # Detailed Fault Levels sheet
            for j, device in enumerate(dfls_list):
                count = (j + 1) * 27 - 27
                device_name = str(device.columns[1]) if len(device.columns) > 1 else "Unknown Device"

                # Clean the device DataFrame and ensure proper numeric types
                device = clean_dataframe(device)
                device = ensure_numeric_types(device)

                device.to_excel(writer, sheet_name=safe_feeder_name, startrow=1, startcol=count, index=False)
                sheet = workbook[safe_feeder_name]
                start_col_letter = get_column_letter(count + 1)
                sheet[f"{start_col_letter}1"].font = Font(size=11, bold=True)
                safe_set_cell(sheet, f"{start_col_letter}1", f'Primary protection: {device_name}')

            # Conductor Damage Results
            if 'Conductor Damage Assessment' in study_selections:
                devices = [fdr.devices for fdr in feeders if fdr.obj.loc_name == feeder][0]
                cond_damage_df = cd.cond_damage_results(devices)
                cond_damage_df = clean_dataframe(cond_damage_df)
                cond_damage_df = ensure_numeric_types(cond_damage_df)

                sheet_name = create_safe_sheet_name(f"{feeder} Cond Dmg Res")
                cond_damage_df.to_excel(writer, sheet_name=sheet_name, startrow=0, index=False)

    wb = load_workbook(filepath)
    ws = wb['General Information']
    adjust_gen_info_col_size(ws)
    ws = wb['Summary Results']
    adjust_summ_col_size(ws)
    for feeder in fault_studies_pd:
        safe_feeder_name = create_safe_sheet_name(f"{str(feeder)} Detailed Results")
        if safe_feeder_name in wb.sheetnames:
            ws = wb[safe_feeder_name]
            adjust_detailed_col_size(ws)
            if 'Conductor Damage Assessment' in study_selections:
                sheet_name = create_safe_sheet_name(f"{feeder} Cond Dmg Res")
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    adjust_cond_damage_col_width(ws)

    # Save the adjusted workbook
    wb.save(filepath)
    app.PrintPlain("Output file saved to " + filepath)


def format_grid_data(ext_grid):
    """
    :param external_grid:
    :return:
    """

    formatted_grid_data = {}
    for grid, attributes in ext_grid.items():
        formatted_grid_data['Parameter'] = ['3-P fault level (A):', 'R/X:', 'Z2/Z1:', 'X0/X1:', 'R0/X0:']
        formatted_grid_data[f'{grid.loc_name} Maximum'] = attributes[:5]
        formatted_grid_data[f'{grid.loc_name} Minimum'] = attributes[5:10]
        formatted_grid_data[f'{grid.loc_name} Sys Norm Minimum'] = attributes[-5:]
    return formatted_grid_data


def ensure_numeric_types(df):
    """
    Ensure that numeric values are stored as proper numeric types, not strings
    """
    if df is None or df.empty:
        return df

    df_numeric = df.copy()

    for col in df_numeric.columns:
        # Skip columns that should remain as strings
        col_str = str(col).lower()
        if any(keyword in col_str for keyword in ['name', 'construction', 'site', 'device']):
            continue

        # Try to convert to numeric, keeping strings where conversion fails
        df_numeric[col] = pd.to_numeric(df_numeric[col], errors='ignore')

        # For columns that are now numeric, ensure integers are displayed as integers
        if df_numeric[col].dtype in ['float64', 'float32']:
            # Check if all non-null values are whole numbers
            non_null_values = df_numeric[col].dropna()
            if len(non_null_values) > 0:
                # If all values are whole numbers, convert to int64 (but keep NaN as NaN)
                if all(pd.isna(val) or (isinstance(val, (int, float)) and val == int(val)) for val in non_null_values):
                    # Use nullable integer type to preserve NaN values
                    df_numeric[col] = df_numeric[col].astype('Int64')

    return df_numeric


def clean_dataframe(df):
    """
    Clean DataFrame to ensure Excel compatibility
    """
    if df is None or df.empty:
        return df

    # Create a copy to avoid modifying the original
    df_clean = df.copy()

    # Replace problematic values
    df_clean = df_clean.replace([float('inf'), float('-inf')], None)

    # Clean string columns (but don't convert everything to string)
    for col in df_clean.columns:
        if df_clean[col].dtype == 'object':
            # Only clean string values, don't convert numbers to strings
            df_clean[col] = df_clean[col].apply(lambda x: clean_string_value(x) if isinstance(x, str) else x)

    # Clean column names
    df_clean.columns = [clean_string_value(str(col)) for col in df_clean.columns]

    return df_clean


def clean_string_value(value):
    """
    Clean string values to remove problematic characters
    """
    if pd.isna(value) or value is None:
        return ''

    value_str = str(value)

    # Remove or replace problematic characters
    # Remove control characters (ASCII 0-31 except tab, newline, carriage return)
    value_str = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', value_str)

    # Replace problematic characters with safe alternatives
    replacements = {
        '\x00': '',  # null character
        '\x01': '',  # start of heading
        '\x02': '',  # start of text
        '\x03': '',  # end of text
        '\x04': '',  # end of transmission
        '\x05': '',  # enquiry
        '\x06': '',  # acknowledge
        '\x07': '',  # bell
        '\x08': '',  # backspace
        '\x0B': ' ',  # vertical tab -> space
        '\x0C': ' ',  # form feed -> space
        '\x0E': '',  # shift out
        '\x0F': '',  # shift in
    }

    for old, new in replacements.items():
        value_str = value_str.replace(old, new)

    # Limit length to prevent Excel issues
    if len(value_str) > 32767:  # Excel cell character limit
        value_str = value_str[:32767]

    return value_str


def safe_set_cell(worksheet, cell_reference, value):
    """
    Safely set cell value with proper cleaning
    """
    try:
        cleaned_value = clean_string_value(value)
        worksheet[cell_reference] = cleaned_value
    except Exception as e:
        # If setting fails, set as string
        worksheet[cell_reference] = str(value)[:32767] if value else ''


def create_safe_sheet_name(name):
    """
    Create a safe Excel sheet name
    """
    if not name:
        name = "Sheet"

    # Convert to string and clean
    name_str = clean_string_value(str(name))

    # Excel sheet name restrictions:
    # - Max 31 characters
    # - Cannot contain: \ / ? * [ ] :
    # - Cannot be empty
    # - Cannot start or end with single quote

    forbidden_chars = ['\\', '/', '?', '*', '[', ']', ':']
    for char in forbidden_chars:
        name_str = name_str.replace(char, '_')

    # Remove leading/trailing single quotes
    name_str = name_str.strip("'")

    # Ensure it's not empty
    if not name_str or name_str.isspace():
        name_str = "Sheet"

    # Truncate to 31 characters
    if len(name_str) > 31:
        name_str = name_str[:31]

    # Remove trailing spaces
    name_str = name_str.rstrip()

    return name_str


def format_fl_results(app, region, feeders):
    """

    :param grid_data:
    :param all_fault_studies:
    :return:
    """

    fault_studies_pd = {}
    for feeder in feeders:
        study_results_df = format_study_results(app, feeder.devices)
        fdr_open_points = format_fdr_open_points(feeder)
        dfls_list = format_detailed_results(app, region, feeder.devices)

        fault_studies_pd[feeder.obj.loc_name] = [study_results_df, fdr_open_points, dfls_list]

    return fault_studies_pd


def format_study_results(app, devices):
    """

    :param devices:
    :return:
    """

    device_list = format_devices()
    # Store results in dictionary:
    for device in devices:
        ds_device_names = [str(device.obj.loc_name) for device in device.ds_devices]
        us_device_names = [str(device.obj.loc_name) for device in device.us_devices]

        max_ds_tr = device.max_ds_tr

        # Keep numeric values as numbers, not strings
        device_values = [
            safe_numeric(device.l_l_volts),
            safe_numeric(device.phases),
            safe_numeric(device.ds_capacity),
            safe_numeric(device.max_fl_3ph),
            safe_numeric(device.max_fl_2ph),
            safe_numeric(device.max_fl_pg),
            safe_numeric(device.min_fl_3ph),
            safe_numeric(device.min_fl_2ph),
            safe_numeric(device.min_fl_pg),
            safe_numeric(device.min_sn_fl_2ph),
            safe_numeric(device.min_sn_fl_pg),
            str(max_ds_tr.term.cpSubstat.loc_name) if max_ds_tr.term is not None else '',
            safe_numeric(max_ds_tr.load_kva),
            safe_numeric(max_ds_tr.max_ph),
            safe_numeric(max_ds_tr.max_pg),
            ', '.join(ds_device_names),
            ', '.join(us_device_names),
        ]

        device_list[str(device.obj.loc_name)] = device_values

    # Format 'Summary Results' data
    formatted_dev_pd = pd.DataFrame.from_dict(device_list)
    study_results_df = formatted_dev_pd.fillna("")

    return study_results_df


def format_fdr_open_points(feeder):

    open_points = feeder.open_points
    safe_open_points = [str(open_point.loc_name) for open_point in open_points]

    fdr_open_points = {
        'Feeder Open Points': safe_open_points
    }
    # Format open point data
    formatted_fdr_open_points = pd.DataFrame.from_dict(fdr_open_points)
    formatted_fdr_open_points = formatted_fdr_open_points.fillna("")

    return formatted_fdr_open_points

def safe_numeric(value):
    """
    Safely convert a value to numeric, returning the original numeric value
    """
    try:
        if value is None or pd.isna(value):
            return np.nan
        if value == float('inf') or value == float('-inf'):
            return np.nan

        # If it's already a number, return as-is
        if isinstance(value, (int, float)):
            return value

        # Try to convert string to number
        return float(value)
    except (ValueError, TypeError, OverflowError):
        return np.nan


def format_detailed_results(app, region, devices):
    """

    :param devices:
    :return:
    """
    dfls_list = []
    for device in devices:
        device_reach_factors = device_reach_factors(region, device, device.sect_terms)

        # Safely extract device name and terms
        device_name = str(device.obj.loc_name) if device.obj is not None else 'Unknown Device'
        sect_terms = getattr(device, 'sect_terms', [])

        df = pd.DataFrame({
            device_name: [
                str(term.obj.loc_name) if term.obj is not None else str(term)
                for term in sect_terms],
            # Data
            'Construction': [str(getattr(term, 'constr', '')) for term in sect_terms],
            'L-L Voltage': [safe_numeric(getattr(term, 'l_l_volts', 0)) for term in sect_terms],
            'No. Phases': [safe_numeric(getattr(term, 'phases', 0)) for term in sect_terms],
            # FAULT LEVELS
            'Max 3PH fault': [safe_numeric(getattr(term, 'max_fl_3ph', 0)) for term in sect_terms],
            'Max 2PH fault': [safe_numeric(getattr(term, 'max_fl_2ph', 0)) for term in sect_terms],
            'Max PG fault': [safe_numeric(getattr(term, 'max_fl_pg', 0)) for term in sect_terms],
            'Min 3P fault': [safe_numeric(getattr(term, 'min_fl_3ph', 0)) for term in sect_terms],
            'Min 2P fault': [safe_numeric(getattr(term, 'min_fl_2ph', 0)) for term in sect_terms],
            'Min PG fault': [safe_numeric(
                fault_impedance.get_terminal_pg_fault(region, term)) for term in sect_terms],
            'Min Sys Norm 2P fault':
                [safe_numeric(getattr(term, 'min_fl_3ph', 0)) for term in sect_terms],
            'Min Sys NormPG fault': [safe_numeric(
                fault_impedance.get_terminal_pg_fault(
                    region, term, system_normal=True)) for term in sect_terms],
            # PICKUPS
            'EF PRI PU': device_reach_factors.get('ef_pickup', []),
            'EF BU PU': device_reach_factors.get('bu_ef_pickup', []),
            'PH PRI PU': device_reach_factors.get('ph_pickup', []),
            'PH BU PU': device_reach_factors.get('bu_ph_pickup', []),
            'NPS PRI PU': device_reach_factors.get('nps_pickup', []),
            'NPS BU PU': device_reach_factors.get('bu_nps_pickup', []),
            # REACH FACTORS
            'EF PRI RF': device_reach_factors.get('ef_rf', []),
            'EF BU RF': device_reach_factors.get('bu_ef_rf', []),
            'PH PRI RF': device_reach_factors.get('ph_rf', []),
            'PH BU RF': device_reach_factors.get('bu_ph_rf', []),
            'NPS EF PRI RF': device_reach_factors.get('nps_ef_rf', []),
            'NPS EF BU RF': device_reach_factors.get('bu_nps_ef_rf', []),
            'NPS PH PRI RF': device_reach_factors.get('nps_ph_rf', []),
            'NPS PH BU RF': device_reach_factors.get('bu_nps_ph_rf', [])
        })

        # Sort fault levels by Max PG fault if column exists and has data
        if 'Max PG fault' in df.columns and not df.empty:
            try:
                df_sorted = df.sort_values(by='Max PG fault', ascending=False)
            except AttributeError:
                df_sorted = df
        else:
            df_sorted = df

        df_sorted.insert(0, 'Tfmr Size (kVA)', np.nan)

        # Safely handle transformer data - keep as numeric
        try:
            tr_dict = {
                str(load.term.loc_name): safe_numeric(load.load_kva) for load in device.sect_loads
                    if hasattr(load, 'term') and hasattr(load, 'load_kva')
            }
            df_sorted['Tfmr Size (kVA)'] = df_sorted[device_name].map(tr_dict)
        except AttributeError:
            # If mapping fails, leave the column with NaN values
            pass

        dfls_list.append(df_sorted)

    return dfls_list


def format_devices() -> dict[str:list]:
    """

    """

    device_list = {
        'Site Name':
            [
                'L-L Voltage (kV)',
                'No. Phases',
                'DS Capacity  (kVA)',
                'Max 3Ph FL',
                'Max 2Ph FL',
                'Max PG FL',
                'Min 3Ph FL',
                'Min 2Ph FL',
                'Min PG FL',
                'Min SN 2P FL',
                'Min SN PG FL',
                'Max DS TR (Site name)',
                'Max TR size (kVA)',
                'TR Max Ph ',
                'TR Max PG',
                'Downstream Devices',
                'Back-up Device'
            ]
    }

    return device_list


def fix_string(file_name):
    """
    Excel does not allow special characters in file names. Remove any such cases and replace with a '_'
    :param file_name: string
    :return:
    """
    if not file_name:
        return "default_filename"

    file_name_str = str(file_name)
    forbidden_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']

    for char in forbidden_chars:
        file_name_str = file_name_str.replace(char, '_')

    # Remove control characters
    file_name_str = re.sub(r'[\x00-\x1F\x7F]', '_', file_name_str)

    return file_name_str


def adjust_gen_info_col_size(ws):

    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"].font = Font(bold=True, size=12)

    for col in ws.columns:
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = 18.0

    for cell in ws[27]:
        cell.alignment = Alignment(wrap_text=True)


def adjust_summ_col_size(ws):
    """
    Adjust column width of given Excel sheet
    :param ws:
    :return:
    """

    # Adjust the column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name

        second_row_cell = col[1] if len(col) > 1 else None
        target_column = (second_row_cell and second_row_cell.value and str(second_row_cell.value) == "Tfmr Size (kVA)")
        if target_column:
            ws.column_dimensions[column].width = 13.71
        else:
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


def adjust_detailed_col_size(ws):
    """
    Adjust column width of given Excel sheet
    :param ws:
    :return:
    """

    # Set second row height to 30.00
    ws.row_dimensions[2].height = 30.00

    # Apply "Wrap Text" formatting to cells in the second row
    for cell in ws[2]:  # Get all cells in row 2
        cell.alignment = Alignment(wrap_text=True)

    # Adjust the column widths
    device_col_letter = None
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name

        second_row_cell = col[1] if len(col) > 1 else None
        tfmr_column = (second_row_cell and second_row_cell.value and str(second_row_cell.value) == "Tfmr Size (kVA)")
        if tfmr_column:
            ws.column_dimensions[column].width = 13.71
            second_row_cell = col[1]
            tr_col_num = second_row_cell.column
            tr_col_letter = second_row_cell.column_letter
            device_col_num = tr_col_num + 1
            device_col_letter = get_column_letter(device_col_num)
        elif column == device_col_letter:
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
        else:
            ws.column_dimensions[column].width = 15


def adjust_cond_damage_col_width(ws):
    """
    Adjust column width of given Excel sheet
    :param ws: Excel worksheet object
    :return:
    """

    # Set top row height to 30.00
    ws.row_dimensions[1].height = 30.00

    # Apply "Wrap Text" formatting to cells in the top row
    wrap_alignment = Alignment(wrap_text=True)

    for cell in ws[1]:  # Get all cells in row 1
        cell.alignment = wrap_alignment

    # Define fixed widths for columns E onwards
    fixed_widths = {
        'E': 13.86,
        'F': 17.14,
        'G': 14.71,
        'H': 16.43,
        'I': 13.57,
        'J': 17.14,
        'K': 15.71,
        'L': 16.29
    }

    # Adjust column widths
    for col in ws.columns:
        column_letter = col[0].column_letter  # Get the column name

        # For columns A to D, adjust width based on maximum cell length
        if column_letter in ['A', 'B', 'C', 'D']:
            max_length = 0

            # Iterate over all cells in the column to find the maximum length
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except (AttributeError, ValueError):
                    pass

            # Set the column width to the maximum length found
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width

        # For columns E onwards, use fixed widths
        elif column_letter in fixed_widths:
            ws.column_dimensions[column_letter].width = fixed_widths[column_letter]


