import pandas as pd


def load_text(filename):
    """load text file, return as string"""
    with open(filename, "r") as file:
        text = file.read()
    return text


def filter_sections_start(sections, keyword):
    """read list of strings and only return those that contain a keyword at start of string"""
    filtered = []
    for section in sections:
        if section.startswith(keyword):
            filtered.append(section)
    return filtered


def filter_sections_end(sections, keyword):
    """read list of strings and only return those that contain a keyword at end of first line"""
    filtered = []
    for section in sections:
        if section.split("\n")[0].endswith(keyword):
            filtered.append(section)
    return filtered


def parse_tsv(text):
    """
    Read a string, skip first line and parse the rest as a tab separated value (tsv) list
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
    """
    lines = text.split("\n")[1:]
    table = [line.split("\t") for line in lines]
    return table


def parse_tsvs(sections):
    """function to parse_tsv a list of strings"""
    tables = [parse_tsv(section) for section in sections]
    return tables


def list_to_df(list_of_lists, col_names=None):
    """convert list of lists to pandas dataframe"""
    if col_names is None:
        df = pd.DataFrame(list_of_lists)
    else:
        df = pd.DataFrame(list_of_lists, columns=col_names)
    return df


def combine_named_dfs_from_list_of_dicts(merged, data_key="devices"):
    """
    Combine all dfs in the 'data_key' column of a list of dicts into a single df
    takes a list of dicts containing a 'data_key' key with a pd.DataFrame value and combines them into a single pd.DataFrameq
    Example:
    merged = [
        {'devices': pd.DataFrame(...), 'id': '110101', 'loop': 23, 'node': 11, 'raw': 'M 110101 X 1'},
        {'devices': pd.DataFrame(...), 'id': 'xxxx', 'loop': 24, 'node': 11, 'raw': 'M xxxx X 1'},
        ...
    ]
    Returns:
    pd.DataFrame(...)
    """
    devices_dfs = [loop[data_key] for loop in merged]
    combined = pd.concat(devices_dfs)
    return combined
