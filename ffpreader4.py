import os
import pandas as pd
from utils import (
    load_text,
    filter_sections_start,
    filter_sections_end,
    parse_tsv,
    list_to_df,
)


class FFPReader:
    # Internal constants
    _NODE_SECTION_FLAG = "P"
    _ZONE_SECTION_FLAG = "Z"
    _ZONE_SECTION_NAME = "zones"
    _LOOP_OR_LOOP_DEVICE_SECTION_FLAG = "M"
    _LOOP_INFO_SECTION_SUFFIX = "X 1"
    _DEVICE_SECTION_SUFFIX = "X 2"

    def __init__(self, ffp_filepath):
        self.ffp_filepath = ffp_filepath
        # self.raw_text = None
        self.sections = self._load_and_separate_sections()
        self.zones = self._filter_parse_load_zone_section_to_df()
        self.nodes = self._filter_parse_load_node_sections_to_df()
        self.loops = self._filter_parse_load_loop_info_sections_to_df()
        self.devices = self._filter_parse_and_load_loop_devices_sections_to_df()
        self.combined_devices_df = None

    @property
    def cleaned_zones(self):
        """Return a cleaned version of the zones DataFrame."""
        if self.zones is not None:
            return self._clean_df(self.zones)
        return None

    @property
    def cleaned_devices(self):
        # initial clean
        df = self._clean_df(self.devices)

        # Select all columns up to and including 'type'
        df = df.loc[:, :"type"]

        # Add 'name' column
        df["name"] = (
            "L"
            + df["loop"].astype(str)
            + " - "
            + "D"
            + df["device"].astype(str)
            + " - "
            + "Z"
            + df["zone"].astype(str)
            + " - "
            + df["description"]
            + " - "
            + df["type"]
        )

        # Try and determine location assuming programmer has set device description in format '{location} {device details}'
        df["locationId"] = df["description"].apply(
            lambda x: x.split(" ")[0] if x.startswith("IRD-") else ""
        )

        # Sort by 'loop'
        df = df.sort_values(["loop", "device"])

        return df

    def _load_and_separate_sections(self):
        """
        The FFP system configuration is described in "sections" seperated by square brackets [ ]
        The first line of each section identifies the type of configuration provided in the section.
        eg:
        - Zones: "Z 1 Z 1"
        - Loop info: "M 90102 X 1"
        - Loop devices: "M 90102 X 2"
        - Node info: "P 10000 P 1"
        Read the ffp file and return as list of sections as strings, with section square brackets removed and whitespace trimmed.

        Example:

        text = "...
        [ Z 1 Z 1
        Y	TOWER 2 BASEMENT 5	N	N	0	0	N	N	N	N	0	0	N	N	N
        Y	TOWER 3 BASEMENT 5 	N	N	0	0	N	N	N	N	0	0	N	N	N
        ...
        ]
        [ M 90101 X 3
        2	O1	IRD-ICG-L05M-FCG-01 MGF OVERFLOW	Y
        2	O2	IRD-ICG-L05M-FCG-01 MGF OVERFLOW	Y
        ...
        ]
        [ M 90102 X 1
        12	Apollo Loop No: 12	0	0	0	0	0	0	550	2500	1	R
        ]
        [ M 90102 X 2
        90	IRD-ICG-L05M-FCG-01 MGF OVERFLOW	x02	OPT	0	0	N	Y	Y	N	N	N	N	Y	0	0	0	0	0	179	N	80	80	0	0	0	0	0	NA
        91	IRD-IT3-L05M-SFS-01 FIRE CUPBOARD	x02	OPT	0	0	N	Y	Y	N	N	N	N	Y	0	0	0	0	0	179	N	80	80	0	0	0	0	0	NA
        ]
        [ P 10000 P 1
        MASD-FIP-ICG-L02-01 T4 L02 MFIP	1

        ]"

        returns [
            "Z 1 Z 1
            Y	TOWER 2 BASEMENT 5 ...",
            "M 90101 X 3
            2	O1	IRD-ICG-L05M-FCG-01 MGF OVERFLOW	Y ...",
            "M 90102 X 1
                12	Apollo Loop No: 12	0	0	0	0	0	0	550	2500	1	R",
            "M 90102 X 2
            90	IRD-ICG-L05M-FCG-01 MGF OVERFLOW ...",
            "P 10000 P 1
            MASD-FIP-ICG-L02-01 T4 L02 MFIP	1"
        ]
        """
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

    def _filter_parse_load_zone_section_to_df(self):
        """
        Read list of strings from ffp file and only return zone sections (first line starts with Z)
        In testing, only ever one 'zones' section in file. If there was ever a file with more than one zones section, we wouldnt know how to re-index the zoone ID/address so should raise an error.
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
        if len(filtered) != 1:
            raise ValueError(
                f"Expected exactly one zone section, found {len(filtered)}. Cannot re-index zones, check FFP file."
            )
        else:
            print(f"Found {len(filtered)} zone sections, proceeding with parsing.")
        # parse first section in list (only one item)
        filtered = filtered[0]
        return self._parse_zone_section_to_df(filtered)

    def _parse_zone_section_to_df(self, section):
        """
        takes a zones section text string and parses into a df containing columns:
            - 'zone': the zone address/ID
            - 'description': the name/description assigned to the zone
        Returns:
        pd.DataFrame(
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
        """
        raw_list_of_lists = parse_tsv(section)
        # set the column names ['zone', 'description', ...rest is not required]
        df = list_to_df(raw_list_of_lists)
        # delete the first column
        df = df.drop(df.columns[0], axis=1)
        # set the new first column name to 'description'
        df = df.rename(columns={df.columns[0]: "description"})
        # remove the remaining columns
        df = df.drop(df.columns[1:], axis=1)
        # The zone number/zone address isnt explicitly stated, its implied in the row number it occupies within the zones section
        # add a 1-indexed column for the zone ID
        df["zone"] = df.index + 1
        # move the zone_id to the front
        cols = df.columns.tolist()
        cols = cols[-1:] + cols[:-1]
        df = df[cols]
        # add a column containing the raw section header
        df["raw"] = section.split("\n")[0]
        return df

    def _filter_parse_load_node_sections_to_df(self):
        """
        Example node:
        [ P 110000 P 1
        xxxx

        ]
        """
        filtered = filter_sections_start(self.sections, self._NODE_SECTION_FLAG)
        nodes_list_of_dicts = [
            self._parse_node_section_to_dict(section) for section in filtered
        ]
        return pd.DataFrame(nodes_list_of_dicts)

    # function to read a node section and return a dict containing node info
    def _parse_node_section_to_dict(self, section):
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

    def _filter_parse_load_loop_info_sections_to_df(self):
        """read list of strings from ffp file and only return loop info sections (first line starts with M and ends with 'X 1')"""
        # filter sections that start with M and end with 'X 1'
        filtered = filter_sections_start(
            self.sections, self._LOOP_OR_LOOP_DEVICE_SECTION_FLAG
        )
        filtered = filter_sections_end(filtered, self._LOOP_INFO_SECTION_SUFFIX)
        # parse each section and return a list of dicts containing loop info
        loop_info_list_of_dicts = [
            self._parse_loop_info_section_to_dict(section) for section in filtered
        ]
        return pd.DataFrame(loop_info_list_of_dicts)

    def _parse_loop_info_section_to_dict(self, section):
        """
        Read a loop info section and return a dict containing loop info
        Example:
        [ M 110101 X 1
        23	Apollo Loop No: 23	0	0	0	0	0	0	550	2500	1	R
        ]
        Returns:
        {'loop': 23, 'node': 11, 'id': '110101', 'raw': 'M 110101 X 1'}
        """
        table = parse_tsv(section)
        loop_info = {}
        loop_info["loop"] = int(table[0][0])
        loop_info.update(self._parse_section_header_info(section))
        print(f"Loop info: {loop_info}")
        return loop_info

    def _filter_parse_and_load_loop_devices_sections_to_df(self):
        """read a list of sections and return a list of dicts containing loop info and devices"""
        # filter sections where the first line starts with M and end with 'X 2'
        filtered = filter_sections_start(
            self.sections, self._LOOP_OR_LOOP_DEVICE_SECTION_FLAG
        )
        filtered = filter_sections_end(filtered, self._DEVICE_SECTION_SUFFIX)
        loop_devices_list_of_dfs = [
            self._parse_loop_device_section_to_df(section) for section in filtered
        ]
        # Combine the DataFrames into a single DataFrame
        devices = pd.concat(loop_devices_list_of_dfs, ignore_index=True)
        return devices

    def _parse_loop_device_section_to_df(self, section):
        """
        read a section and return a dict containing loop device info
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
        """
        # Get the loop ID for the section
        loop = self._get_loop_id_for_loop_device_section(section)
        # Parse the section into a list of lists
        raw_list_of_lists = parse_tsv(section)
        # Convert the list of lists into a DataFrame
        df = list_to_df(raw_list_of_lists)

        # set the column names ['zone', 'description', 'subtype', 'type', ...rest is not required]
        col_names = ["zone", "description", "subtype", "type"]
        old_names = df.columns[:4]
        df.rename(columns=dict(zip(old_names, col_names)), inplace=True)
        # The loop device number/device address isnt explicitly stated, its impplied in the row number it occupies within the loop device section
        # add a 1-indexed column for the device ID
        df["device"] = df.index + 1
        # Add the loop number as a new column
        df["loop"] = loop
        # Reorder columns to move 'device' and 'loop' to the front
        cols = ["device", "loop"] + [
            col for col in df.columns if col not in ["device", "loop"]
        ]
        df = df[cols]

        return df

    def _get_loop_id_for_loop_device_section(self, section):
        """
        Retrieve the loop number for a given loop device section.

        Here's a typical loop info section (loop 12):
        [ M 90102 X 1
        12	Apollo Loop No: 12	0	0	0	0	0	0	550	2500	1	R
        ]
        *Info for all loops have been saved to the self.loops df.

        Here's a typical loop device section, for the above loop (loop 12):
        [ M 90102 X 2
        90	IRD-ICG-L05M-FCG-01 MGF OVERFLOW	x02	OPT	0	0	N	Y	Y	N	N	N	N	Y	0	0	0	0	0	179	N	80	80	0	0	0	0	0	NA
        91	IRD-IT3-L05M-SFS-01 FIRE CUPBOARD	x02	OPT	0	0	N	Y	Y	N	N	N	N	Y	0	0	0	0	0	179	N	80	80	0	0	0	0	0	NA
        ...
        ]

        The loop device section can be mapped to the corresponding loop by the "90102" identifier in the first row.
        This is stored for each loop in the self.loops.id column.

        This method returns 12 for the above example when the loop device section is passed in as an argument.

        Args:
            section (str): The loop device section to process.

        Returns:
            str or None: The loop ID if a match is found in `self.loops`, otherwise None.
        """
        # Parse the header information from the section
        loop_info = self._parse_section_header_info(section)

        # Ensure 'id' exists in the parsed loop_info
        if "id" not in loop_info:
            return None

        # Find the row in self.loops where 'id' matches loop_info['id']
        matching_row = self.loops[self.loops["id"] == loop_info["id"]]

        # If a match is found, return the 'loop' value; otherwise, return None
        if not matching_row.empty:
            return matching_row.iloc[0]["loop"]
        return None

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
    ffp_database_filename = "QWP 17.02.25 ASE change rev 2 .ffp"
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

    print("\n####\nLoops")
    print(reader.loops)

    print("\n####\nDevices")
    print(reader.devices)

    print("\n####\nCleaned Devices")
    print(reader.cleaned_devices)

    output_file = os.path.join(output_dir, ffp_database_filename + ".zones" + ".csv")
    reader.zones.to_csv(output_file, index=False)
