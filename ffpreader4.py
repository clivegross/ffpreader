import os
import pandas as pd
from utils import (
    load_text,
    filter_sections_start,
    parse_tsv,
    list_to_df,
    combine_named_dfs_from_list_of_dicts,
)


class FFPReader:
    # Internal constants
    _NODE_SECTION_FLAG = "P"
    _ZONE_SECTION_FLAG = "Z"
    _ZONE_SECTION_NAME = "zones"
    _LOOP_SECTION_FLAG = "M"
    _LOOP_INFO_SUFFIX = "X 1"
    _LOOP_DEVICE_SUFFIX = "X 2"

    def __init__(self, ffp_filepath):
        self.ffp_filepath = ffp_filepath
        # self.raw_text = None
        self.sections = self._load_and_separate_sections()
        self.zones = self._filter_parse_combine_zone_sections()
        self.nodes = self._filter_parse_load_node_sections()
        self.loop_info_list = None  # list of dicts
        self.loop_device_list = None  # list of dicts
        self.combined_devices_df = None
        self.cleaned_devices_df = None

    @property
    def cleaned_zones(self):
        """Return a cleaned version of the zones DataFrame."""
        if self.zones is not None:
            return self._clean_df(self.zones)
        return None

    def _load_and_separate_sections(self):
        """Read ffp file as text and separate out sections in square brackets [ ] return as list"""
        text = load_text(self.ffp_filepath)
        sections = []
        start = 0
        while True:
            start = text.find("[", start)
            if start == -1:
                break
            end = text.find("]", start)
            if end == -1:
                break
            # trim leading and trailing white space
            trimmed = text[start + 1 : end].strip()
            sections.append(trimmed)
            start = end + 1
        return sections

    def _parse_section_header_info(self, section):
        """
        parses standard info from first row of section
        Example:
        [ M 110101 X 1
        23	Apollo Loop No: 23	0	0	0	0	0	0	550	2500	1	R
        ]
        Returns:
        {'node': 11, 'id': '110101', 'raw': 'M 110101 X 1'}
        """
        section_header_info = {}
        lines = section.split("\n")
        # split first line by spaces and take the second element, eg '110101' from 'M 110101 X 1'
        id = self._parse_node_loop_id(section)
        # remove the last 4 digits, eg '11' from '110101'
        section_header_info["node"] = self._parse_node_id(id)
        # save the id
        section_header_info["id"] = id
        # save the raw first line
        section_header_info["raw"] = lines[0]
        return section_header_info

    def _parse_node_id(self, id):
        """read a node loop id and return the node id , ie remove the last 4 digits, eg '11' from '110101'"""
        return int(id[:-4])

    # function to read a loop section string and return the 'id'  eg '110101' from string 'M 110101 X 1/n...'
    def _parse_node_loop_id(self, section):
        lines = section.split("\n")
        """
        split first line by spaces and take the second element, eg '110101' from 'M 110101 X 1'
        """
        loop_id = lines[0].split(" ")[1]
        return loop_id

    def _filter_parse_combine_zone_sections(self):
        """
        Read list of strings from ffp file and only return zone sections (first line starts with Z)
        In testing, only ever one 'zones' section in file but will treat as though more than one section is possible and combine them if so.
        Example zone:
        [ Z 1 Z 1
        Y	TOWER 2 BASEMENT 5	N	N	0	0	N	N	N	N	0	0	N	N	N
        Y	TOWER 3 BASEMENT 5 	N	N	0	0	N	N	N	N	0	0	N	N	N
        ...
        N	Unassigned Text	N	N	0	0	N	N	N	N	0	0	N	N	N
        N	Unassigned Text	N	N	0	0	N	N	N	N	0	0	N	N	N
        ]
        """
        filtered = filter_sections_start(self.sections, self._ZONE_SECTION_FLAG)
        zone_list = [self._parse_zone_section(section) for section in filtered]
        return combine_named_dfs_from_list_of_dicts(zone_list, "zones")

    # function to read a section and return a dict containing zone section info and zone list df
    def _parse_zone_section(self, section):
        """
        takes a zones section text string and parses into a dict containing:
            - 'raw': the raw string
            - 'zones': a dataframe of zones
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
        """
        zone_info = {}
        zone_info["raw"] = section.split("\n")[0]
        zone_info["zones"] = self._to_zone_df(section)
        return zone_info

    # create a pandas dataframe from a zone section
    def _to_zone_df(self, section):
        raw_list_of_lists = parse_tsv(section)
        # set the column names ['zone', 'description', ...rest is not required]
        df = list_to_df(raw_list_of_lists)
        # delete the first column
        df = df.drop(df.columns[0], axis=1)
        # set the new first column name to 'description'
        df = df.rename(columns={df.columns[0]: "description"})
        # remove the remaining columns
        df = df.drop(df.columns[1:], axis=1)
        # add a 1-indexed column for the zone ID
        df["zone"] = df.index + 1
        # move the zone_id to the front
        cols = df.columns.tolist()
        cols = cols[-1:] + cols[:-1]
        df = df[cols]
        # add a column containing the raw section header
        df["raw"] = section.split("\n")[0]
        return df

    def _filter_parse_load_node_sections(self):
        """
        Example node:
        [ P 110000 P 1
        xxxx

        ]
        """
        filtered = filter_sections_start(self.sections, self._NODE_SECTION_FLAG)
        nodes_list_of_dicts = [
            self._parse_node_section(section) for section in filtered
        ]
        return pd.DataFrame(nodes_list_of_dicts)

    # function to read a node section and return a dict containing node info
    def _parse_node_section(self, section):
        """
        Example:
        [ P 110000 P 1
        IT4 DATA GATHERING POINT 1	11

        ]
        Returns:
        {'node': 11, 'name': 'IT4 DATA GATHERING POINT 1', 'id': '110101', 'raw': 'P 110000 P 1'}
        """
        node_info = {}
        node_info.update(self._parse_section_header_info(section))
        # parse_tsv returns [['IT4 DATA GATHERING POINT 1', '11']]
        # take the first element of the first list, eg 'IT4 DATA GATHERING POINT 1'
        node_description = parse_tsv(section)[0][0]
        node_info["description"] = node_description
        # reorder the node_info dict id then decription then the rest
        node_info = {k: node_info[k] for k in ["node", "description", "id", "raw"]}
        return node_info

    def _clean_df(self, df):
        """
        Takes a zones, loops, devices or nodes df and returns a "cleaned" df with:
            - rows representing empty addresses removed
            - string fields trimmed
            - illegal charceters removed from string fields
        """

        if "zone" in df.columns:
            # Remove rows where 'zone' is 0
            df = df[df["zone"] != 0]

        if "description" in df.columns:
            # Remove rows where 'description' is empty or NaN
            df = df[df["description"].ne("") & df["description"].notna()]

            # Remove rows where 'description' is 'Unassigned Text'
            df = df[df["description"] != "Unassigned Text"]

            # Remove leading and trailing whitespace from 'description' and 'type'
            df["description"] = df["description"].str.strip()

            # Replace illegal characters in 'description' and 'type'
            df["description"] = df["description"].str.replace("/", "-")
            df["description"] = df["description"].str.replace("&", "+")

        # Replace illegal characters in 'type' only if 'type' column exists
        if "type" in df.columns:
            df["type"] = df["type"].str.strip()
            df["type"] = df["type"].str.replace("/", "-")

        return df


if __name__ == "__main__":
    input_dir = "./data/input"
    output_dir = "./data/output"
    ffp_database_filename = (
        "QWP 17.02.25 ASE change rev 2 .ffp"  # Example filename from your code
    )
    ffp_database_filepath = os.path.join(input_dir, ffp_database_filename)

    # Instantiate the reader
    reader = FFPReader(ffp_database_filepath)
    # print(reader.sections)
    print("\n####\nZones")
    print(reader.zones)

    print("\n####\nCleaned Zones")
    print(reader.cleaned_zones)

    print("\n####\nNodes")
    print(reader.nodes)
