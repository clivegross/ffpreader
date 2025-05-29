import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo

# function to load text file
def load_text(filename):
    file = open(filename, 'r')
    text = file.read()
    file.close()
    return text

# function to read text and separate out sections in square brackets [ ] return as list
def separate_sections(text):
    sections = []
    start = 0
    while True:
        start = text.find('[', start)
        if start == -1:
            break
        end = text.find(']', start)
        if end == -1:
            break
        # trim leading and trailing white space
        trimmed = text[start+1:end].strip()
        sections.append(trimmed)
        start = end + 1
    return sections

# function to read list of strings and only return those that contain a keyword in the first line
def filter_sections(sections, keyword):
    filtered = []
    for section in sections:
        if keyword in section.split('\n')[0]:
            filtered.append(section)
    return filtered

# function to read list of strings and only return those that contain a keyword at start of string
def filter_sections_start(sections, keyword):
    filtered = []
    for section in sections:
        if section.startswith(keyword):
            filtered.append(section)
    return filtered

# function to read list of strings and only return those that contain a keyword at end of first line
def filter_sections_end(sections, keyword):
    filtered = []
    for section in sections:
        if section.split('\n')[0].endswith(keyword):
            filtered.append(section)
    return filtered

# function to read list of strings from ffp file and only return loop info sections (first line starts with M and ends with 'X 1')
def filter_loop_info_sections(sections):
    filtered = filter_sections_start(sections, "M")
    filtered = filter_sections_end(filtered, "X 1")
    return filtered

# function to read list of strings from ffp file and only return loop device sections (first line starts with M and ends with 'X 2')
def filter_loop_device_sections(sections):
    filtered = filter_sections_start(sections, "M")
    filtered = filter_sections_end(filtered, "X 2")
    return filtered

# function to read list of strings from ffp file and only return node sections (first line starts with P)
def filter_node_sections(sections):
    '''
    Example node:
    [ P 110000 P 1
    xxxx

    ]
    '''
    filtered = filter_sections_start(sections, "P")
    return filtered

# function to read list of strings from ffp file and only return zone sections (first line starts with Z)
def filter_zone_sections(sections):
    '''
    Example zone:
    [ Z 1 Z 1
    Y	TOWER 2 BASEMENT 5	N	N	0	0	N	N	N	N	0	0	N	N	N	
    Y	TOWER 3 BASEMENT 5 	N	N	0	0	N	N	N	N	0	0	N	N	N
    ...	
    N	Unassigned Text	N	N	0	0	N	N	N	N	0	0	N	N	N	
    N	Unassigned Text	N	N	0	0	N	N	N	N	0	0	N	N	N	
    ]
    '''
    filtered = filter_sections_start(sections, "Z")
    return filtered

# function to read a string, skip first line and parse the rest as a tab separated table
def parse_tsv(text):
    '''
    Example:
    [ M 140102 X 2
    22	IRD-ICG-L00-CSV-01 BOH CORRIDOR	x02	OPT	0	0	N	Y	Y	Y	Y	Y	N	Y	6	0	0	0	0	180	N	100	80	0	0	0	0	0	NA	
    22	IRD-IT1-L00-SEL-02 ELEC CUPBOARD	x02	OPT	0	0	N	Y	Y	Y	Y	Y	N	Y	6	0	0	0	0	180	N	100	80	0	0	0	0	0	NA...
    ]
    Returns:
    skips the first header line and treats the remaining string as a tab separated values (tsv) format data and returns as a list of lists
    [
        ['22', 'IRD-ICG-L00-CSV-01 BOH CORRIDOR', 'x02', 'OPT', '0', '0', 'N', 'Y', 'Y', 'Y', 'Y', 'Y', 'N', 'Y', '6', '0', '0', '0', '0', '180', 'N', '100', '80', '0', '0', '0', '0', '0', 'NA'],
        ['22', 'IRD-IT1-L00-SEL-02 ELEC CUPBOARD', 'x02', 'OPT', '0', '0', 'N', 'Y', 'Y', 'Y', 'Y', 'Y', 'N', 'Y', '6', '0', '0', '0', '0', '180', 'N', '100', '80', '0', '0', '0', '0', '0', 'NA...']
    ]
    '''
    lines = text.split('\n')[1:]
    table = [line.split('\t') for line in lines]
    return table

# function to parse_tsv a list of strings
def parse_tsvs(sections):
    tables = [parse_tsv(section) for section in sections]
    return tables

def parse_section_header_info(section):
    '''
    parses standard info from first row of section
    Example:
    [ M 110101 X 1
    23	Apollo Loop No: 23	0	0	0	0	0	0	550	2500	1	R
    ]
    Returns:
    {'node': 11, 'id': '110101', 'raw': 'M 110101 X 1'}
    '''
    section_header_info = {}
    lines = section.split('\n')
    # split first line by spaces and take the second element, eg '110101' from 'M 110101 X 1'
    id = parse_node_loop_id(section)
    # remove the last 4 digits, eg '11' from '110101'
    section_header_info['node'] = parse_node_id(id)
    # save the id
    section_header_info['id'] = id
    # save the raw first line
    section_header_info['raw'] = lines[0]
    return section_header_info

def parse_sections_header_info(sections):
    section_header_infos = [parse_section_header_info(section) for section in sections]
    return section_header_infos

# function to read a loop section string and return the 'id'  eg '110101' from string 'M 110101 X 1/n...'
def parse_node_loop_id(section):
    lines = section.split('\n')
    '''
    split first line by spaces and take the second element, eg '110101' from 'M 110101 X 1'
    '''
    loop_id = lines[0].split(' ')[1]
    return loop_id

# function to read a node loop id and return the node , ie remove the last 4 digits, eg '11' from '110101'
def parse_node_id(id):
    return int(id[:-4])

# function to read a node section and return a dict containing node info
def parse_node_section(section):
    '''
    Example:
    [ P 110000 P 1
    IT4 DATA GATHERING POINT 1	11
                
    ]
    Returns:
    {'node': 11, 'name': 'IT4 DATA GATHERING POINT 1', 'id': '110101', 'raw': 'P 110000 P 1'}
    '''
    node_info = {}
    node_info.update(parse_section_header_info(section))
    # parse_tsv returns [['IT4 DATA GATHERING POINT 1', '11']]
    # take the first element of the first list, eg 'IT4 DATA GATHERING POINT 1'
    node_description = parse_tsv(section)[0][0]
    node_info['description'] = node_description
    # reorder the node_info dict id then decription then the rest
    node_info = {k: node_info[k] for k in ['node', 'description', 'id', 'raw']}
    return node_info

# function to read a list of node sections and return a list of node dicts
def parse_node_sections(sections):
    node_list = [parse_node_section(section) for section in sections]
    return node_list    

# function to read a section and return a dict containing loop info
def parse_loop_info_section(section):
    '''
    Example:
    [ M 110101 X 1
    23	Apollo Loop No: 23	0	0	0	0	0	0	550	2500	1	R
    ]
    Returns:
    {'loop': 23, 'node': 11, 'id': '110101', 'raw': 'M 110101 X 1'}
    '''
    table = parse_tsv(section)
    loop_info = {}
    loop_info['loop'] = int(table[0][0])
    loop_info.update(parse_section_header_info(section))
    return loop_info

# function to read a list of loop info sections and return a list of loop info dicts
def parse_loop_info_sections(sections):
    loop_info_list = [parse_loop_info_section(section) for section in sections]
    return loop_info_list

# function to read a section and return a dict containing loop device info
def parse_loop_device_section(section):
    '''
    Example:
    [ M 140102 X 2
    22	IRD-ICG-L00-CSV-01 BOH CORRIDOR	x02	OPT	0	0	N	Y	Y	Y	Y	Y	N	Y	6	0	0	0	0	180	N	100	80	0	0	0	0	0	NA	
    22	IRD-IT1-L00-SEL-02 ELEC CUPBOARD	x02	OPT	0	0	N	Y	Y	Y	Y	Y	N	Y	6	0	0	0	0	180	N	100	80	0	0	0	0	0	NA...
    ]
    Returns:
    {
        'loop': 22, 'node': 14, 'id': '140102', 'raw': 'M 140102 X 2',
        'devices': ['22', 'IRD-ICG-L00-CSV-01 BOH CORRIDOR', 'x02', 'OPT', '0', '0', 'N', ...],
    }
    '''
    loop_devices = parse_section_header_info(section)
    loop_devices['devices'] = create_loop_device_df(section)
    return loop_devices

# function to read a list of loop device sections and return a list of loop device dicts
def parse_loop_device_sections(sections):
    loop_device_list = [parse_loop_device_section(section) for section in sections]
    return loop_device_list

# function to read a section and return a dict containing zone section info and zone list df
def parse_zone_section(section):
    '''
    Example:
    [ Z 1 Z 1
    Y	TOWER 2 BASEMENT 5	N	N	0	0	N	N	N	N	0	0	N	N	N	
    Y	TOWER 3 BASEMENT 5 	N	N	0	0	N	N	N	N	0	0	N	N	N
    ...	
    N	Unassigned Text	N	N	0	0	N	N	N	N	0	0	N	N	N	
    N	Unassigned Text	N	N	0	0	N	N	N	N	0	0	N	N	N	
    ]
    Returns:
    {
        'raw': 'Z 1 Z 1',
        'zones': pd.DataFrame(
            Columns: ['zone', 'description']
            Index: [0, 1, 2, ...]
            Data: [
                [1, 'TOWER 2 BASEMENT 5'],
                [2, 'TOWER 3 BASEMENT 5'],
                ...
                [4999, 'Unassigned Text'],
                [5000, 'Unassigned Text']
            ]
                )
    }
    '''
    zone_info = {}
    zone_info['raw'] = section.split('\n')[0]
    zone_info['zones'] = create_zone_df(section)
    return zone_info

# function to read a list of zone sections and return a list of zone dicts
def parse_zone_sections(sections):
    zone_list = [parse_zone_section(section) for section in sections]
    return zone_list

# create a pandas dataframe from a zone section
def create_zone_df(section):
    raw_list_of_lists = parse_tsv(section)
    # set the column names ['zone', 'description', ...rest is not required]
    df = list_to_df(raw_list_of_lists)
    # delete the first column
    df = df.drop(df.columns[0], axis=1)
    # set the new first column name to 'description'
    df = df.rename(columns={df.columns[0]: 'description'})
    # remove the remaining columns
    df = df.drop(df.columns[1:], axis=1)
    # add a 1-indexed column for the zone ID
    df['zone'] = df.index + 1
    # move the zone_id to the front
    cols = df.columns.tolist()
    cols = cols[-1:] + cols[:-1]
    df = df[cols]
    # add a column containing the raw section header
    df['raw'] = section.split('\n')[0]
    return df
    
# create a pandas dataframe from a loop devices section
def create_loop_device_df(section):
    raw_list_of_lists = parse_tsv(section)
    # set the column names ['zone', 'description', 'subtype', 'type', ...rest is not required]
    df = list_to_df(raw_list_of_lists)
    # Assuming df is your DataFrame
    col_names = ['zone', 'description', 'subtype', 'type']
    # new_names = ['new_name1', 'new_name2', 'new_name3', 'new_name4']
    old_names = df.columns[:4]
    df.rename(columns=dict(zip(old_names, col_names)), inplace=True)
    # add a 1-indexed column for the device ID
    df['device'] = df.index + 1
    # move the device_id to the front
    cols = df.columns.tolist()
    cols = cols[-1:] + cols[:-1]
    df = df[cols]
    return df

# add the 'loop' ID to the loop device df in the merged loop device dict and return the dict
def add_loop_id_to_loop_device_df(loop_device_dict):
    loop_id = loop_device_dict['loop']
    loop_device_df = loop_device_dict['devices']
    loop_device_df['loop'] = loop_id
    # move the loop_id to the front
    cols = loop_device_df.columns.tolist()
    cols = cols[-1:] + cols[:-1]
    loop_device_df = loop_device_df[cols]
    loop_device_dict['devices'] = loop_device_df
    return loop_device_dict

# function to merge a list of loop info dicts with a list of loop device dicts on the 'id' key
def merge_loop_info_and_devices(loop_info_list, loop_device_list):
    '''
    Example:
    loop_info_list = [{'id': '110101', 'loop': 23, 'node': 11, 'raw': 'M 110101 X 1'}, ...]
    loop_device_list = [
        {
            'id': '140102', 'loop': 22, 'node': 14, 'raw': 'M 140102 X 2',
            'devices': ['22', 'IRD-ICG-L00-CSV-01 BOH CORRIDOR', 'x02', 'OPT', '0', '0', 'N', ...]
        },
        ...
    ]
    Returns:
    [{'id': '110101', 'loop': 23, 'node': 11, 'raw': 'M 110101 X 1', 'devices': []}, ...]
    '''
    merged = []
    for loop_info in loop_info_list:
        for loop_device in loop_device_list:
            if loop_info['id'] == loop_device['id']:
                loop_info['devices'] = loop_device['devices']
                loop_info = add_loop_id_to_loop_device_df(loop_info)
                merged.append(loop_info)
                break
    return merged

# function to read a list of sections and return a list of dicts containing loop info and devices
def filter_and_parse_loop_devices_sections(sections):
    loop_device_sections = filter_loop_device_sections(sections)
    loop_device_list = parse_loop_device_sections(loop_device_sections)
    loop_info_sections = filter_loop_info_sections(sections)
    loop_info_list = parse_loop_info_sections(loop_info_sections)
    merged = merge_loop_info_and_devices(loop_info_list, loop_device_list)
    return merged

# function to combine all dfs in the 'devices' column of a list of dicts into a single df
def combine_dfs(merged, data_key='devices'):
    '''
    takes a list of dicts containing a 'data' key with a pd.DataFrame value and combines them into a single pd.DataFrameq
    Example:
    merged = [
        {'devices': pd.DataFrame(...), 'id': '110101', 'loop': 23, 'node': 11, 'raw': 'M 110101 X 1'},
        ...
    ]
    Returns:
    pd.DataFrame(...)
    '''
    devices_dfs = [loop[data_key] for loop in merged]
    combined = pd.concat(devices_dfs)
    return combined

# function to convert list of lists to pandas dataframe
def list_to_df(list_of_lists, col_names=None):
    if col_names is None:
        df = pd.DataFrame(list_of_lists)
    else:
        df = pd.DataFrame(list_of_lists, columns=col_names)
    return df

# function to write a list of dicts to csv
def write_csv_from_list_of_dicts(filename, list_of_dicts):
    import csv
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=list_of_dicts[0].keys())
        writer.writeheader()
        for row in list_of_dicts:
            writer.writerow(row)

# function to write a dataframe to excel and format as a table
def write_df_to_excel_and_format(data, excel_filepath, as_table=False):
    '''
    expects dict in format:
    {
        "sheet_name": "loops",
        "data": pandas.df(xxx)
    }
    or list of dicts
    '''
    if type(data) is not list:
        data = [data]
        
    writer = pd.ExcelWriter(excel_filepath, engine='openpyxl')

    for sheet_data in data:
        print(f"writing DataFrame to excel sheet '{sheet_data['sheet_name']}' in '{excel_filepath}'...")
        sheet_data['data'].to_excel(writer, index=False, sheet_name=sheet_data['sheet_name'])

    writer.close()
    
    if as_table:
        for sheet_data in data:
            print(f"reload the workbook to apply Excel Table object table name '{sheet_data['sheet_name']}' to data...")
            book = load_workbook(excel_filepath)
            # Ensure the writer will write to the correct sheet
            sheets = {ws.title: ws for ws in book.worksheets}
            print(sheets)
            sheet = book[sheet_data['sheet_name']]

            print("calculating dimensions...")
            dimensions = sheet.calculate_dimension()
            print("dimensions: ", dimensions)

            print("add table to sheet...")
            tab = Table(displayName=sheet_data['sheet_name'], ref=sheet.calculate_dimension())
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                showLastColumn=False, showRowStripes=True)
            tab.tableStyleInfo = style
            sheet.add_table(tab)

            print("Saving workbook...")
        book.save(excel_filepath)

    print("done")

def clean_df(df):

    if 'zone' in df.columns:
        # Remove rows where 'zone' is 0
        df = df[df['zone'] != 0]

    if 'description' in df.columns:
        # Remove rows where 'description' is empty or NaN
        df = df[df['description'].ne('') & df['description'].notna()]

        # Remove rows where 'description' is 'Unassigned Text'
        df = df[df['description'] != 'Unassigned Text']

        # Remove leading and trailing whitespace from 'description' and 'type'
        df['description'] = df['description'].str.strip()

        # Replace illegal characters in 'description' and 'type'
        df['description'] = df['description'].str.replace('/', '-')
        df['description'] = df['description'].str.replace('&', '+')

    # Replace illegal characters in 'type' only if 'type' column exists
    if 'type' in df.columns:
        df['type'] = df['type'].str.strip()
        df['type'] = df['type'].str.replace('/', '-')
    
    return df


def clean_devices_df(df):
    # initial clean
    df = clean_df(df)

    # Select all columns up to and including 'type'
    df = df.loc[:, :'type']

     # Add 'name' column
    df['name'] = 'L' + df['loop'].astype(str) + ' - ' + 'D' + df['device'].astype(str) + ' - ' + 'Z' + df['zone'].astype(str) + ' - ' + df['description'] + ' - ' + df['type']

    # Add 'locationId' column
    df['locationId'] = df['description'].apply(lambda x: x.split(' ')[0] if x.startswith('IRD-') else '')

    # Sort by 'loop'
    df = df.sort_values(['loop', 'device'])

    return df

def split_df_for_modbus_mapping(df, equipment_type='devices'):
    '''
    Split the DataFrame into separate DataFrames for each modbus register mapping
    type = 'devices' or 'loops' or 'nodes' or 'zones'
    '''
    # if df is a dict first convert to a DataFrame
    if type(df) is dict:
        df = pd.DataFrame.from_dict(df)
    # if df is list of lists first convert to a DataFrame
    if type(df) is list:
        df = list_to_df(df)
    if equipment_type == 'devices' or equipment_type == 'loops':
        # Split the DataFrame
        df1 = df[df['loop'] <= 90]
        df2 = df[(df['loop'] > 90) & (df['loop'] <= 180)]
        df3 = df[df['loop'] > 180]

        # Create a list of dictionaries
        data = [
            {
                'sheet_name': equipment_type+'_L1_to_L90',
                'data': df1,
            },
            {
                'sheet_name': equipment_type+'_L91_to_L180',
                'data': df2,
            },
            {
                'sheet_name': equipment_type+'_L181_to_L250',
                'data': df3
            }
        ]
    elif equipment_type == 'nodes':
        return [
            {
                'sheet_name': equipment_type,
                'data': df
            }
        ]
    elif equipment_type == 'zones':
        # Split the DataFrame
        df1 = df[df['zone'] <= 1000]
        df2 = df[(df['zone'] > 1001) & (df['zone'] <= 2000)]
        df3 = df[df['zone'] > 2000]
        return [
            {
                'sheet_name': equipment_type+'_Z1_to_Z1000',
                'data': df1
            },
            {
                'sheet_name': equipment_type+'_Z1001_to_Z2000',
                'data': df2
            },
            {
                'sheet_name': equipment_type+'_Z2001_to_Z2500',
                'data': df3
            }
        ]
    else:
        return []

    return data


if __name__ == "__main__":
    # input ffp database file
    input_dir = "./data/input"
    ffp_database_filename = 'QWP 16.02.24.ffp'
    ffp_database_filename_2 = 'QWP 17.02.25 ASE change rev 2 .ffp'
    ffp_database_filepath = os.path.join(input_dir, ffp_database_filename_2)

    # output csv and excel files
    output_dir = "./data/output"
    excel_file = ffp_database_filename_2+'.xlsx'
    excel_filepath = os.path.join(output_dir, excel_file)
    device_sheet = 'devices'
    device_csv = os.path.join(output_dir, device_sheet+'.csv')
    loop_sheet = 'loops'
    loop_csv = os.path.join(output_dir, loop_sheet+'.csv')
    node_sheet = 'nodes'
    node_csv = os.path.join(output_dir, node_sheet+'.csv')
    zone_sheet = 'zones'
    zone_csv = os.path.join(output_dir, zone_sheet+'.csv')

    # load text file
    text = load_text(ffp_database_filepath)
    # separate out into list of sections seperated by []
    sections = separate_sections(text)

    # ==================
    # make list of zones and write to file
    zone_sections = filter_zone_sections(sections)
    zone_sections_list = parse_zone_sections(zone_sections)
    print("number of zone sections: ", len(zone_sections))
    # ==================
    # combine all zones dfs
    # only appears to be one zone section in the example file but just in case
    print("combined loop devices")
    combined_zones_df = combine_dfs(zone_sections_list, 'zones')
    print(combined_zones_df.head())
    # write the entire loop devices table to csv, including unused addresses
    combined_zones_df.to_csv(zone_csv, index=False)
    cleaned_zones_df = clean_df(combined_zones_df)

    # ==================
    # make list of nodes and write to file
    node_sections = filter_node_sections(sections)
    node_list = parse_node_sections(node_sections)
    print("number of node sections: ", len(node_sections))
    print(node_list[:10])
    write_csv_from_list_of_dicts(node_csv, node_list)

    # ==================
    # make list of loops and write to file
    loop_info_sections = filter_loop_info_sections(sections)
    loop_info_list = parse_loop_info_sections(loop_info_sections)
    write_csv_from_list_of_dicts(loop_csv, loop_info_list)
    print("number of loop info sections: ", len(loop_info_sections))
    # print(loop_info_sections[:5])
    print(loop_info_list[:5])

    # ==================
    # make list of loop devices and write to file
    # loop_device_sections = filter_loop_device_sections(sections)
    # loop_device_sections_header_info = parse_sections_header_info(loop_device_sections)
    # print("loop device sample")
    # print("loop devices")
    # loop_devices = parse_loop_device_sections(loop_device_sections)
    # print(loop_devices[:5])
    # print(loop_device_sections_header_info[:5])
    # print("number of loop device sections: ", len(loop_device_sections))
    # ==================
    # merge loop info and devices
    print("merged loop info and devices")
    loop_devices_by_loop = filter_and_parse_loop_devices_sections(sections)
    print(loop_devices_by_loop[:5])
    # ==================
    # combine all devices dfs
    print("combined loop devices")
    combined_loop_devices_df = combine_dfs(loop_devices_by_loop)
    print(combined_loop_devices_df.head())
    # write to csv but drop the index column

    # write the entire loop devices table to csv, including unused addresses
    combined_loop_devices_df.to_csv(device_csv, index=False)

    # clean the loop devices table before writing to Excel
    cleaned_loop_devices_df = clean_devices_df(combined_loop_devices_df)

    # split the data into separate sheets for each modbus register mapping
    print("splitting data into separate sheets for each modbus register mapping")
    zones_data = split_df_for_modbus_mapping(cleaned_zones_df, 'zones')
    nodes_data = split_df_for_modbus_mapping(node_list, 'nodes')
    devices_data = split_df_for_modbus_mapping(cleaned_loop_devices_df, 'devices')
    loops_data = split_df_for_modbus_mapping(loop_info_list, 'loops')
    # combine the data
    data = devices_data + loops_data + nodes_data + zones_data
    print(data)

    write_df_to_excel_and_format(data, excel_filepath)


