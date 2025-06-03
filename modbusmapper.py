import pandas as pd
import numpy as np


class ModbusMapper:
    _ZONE_COLNAME = "zone"
    _LOOP_COLNAME = "loop"
    _DEVICE_COLNAME = "device"
    _NODE_COLNAME = "node"
    _GATEWAY_COLNAME = "gateway"
    _HOLDING_REGISTER_COLNAME = "holding_register"
    _ALARM_BIT_OFFSET_COLNAME = "alarm_bit_offset"
    _PREALARM_BIT_OFFSET_COLNAME = "prealarm_bit_offset"
    _FAULT_BIT_OFFSET_COLNAME = "fault_bit_offset"
    _ISOLATE_BIT_OFFSET_COLNAME = "isolate_bit_offset"
    _ALARM_DECIMAL_COLNAME = "alarm_decimal"
    _PREALARM_DECIMAL_COLNAME = "prealarm_decimal"
    _FAULT_DECIMAL_COLNAME = "fault_decimal"
    _ISOLATE_DECIMAL_COLNAME = "isolate_decimal"
    _OPEN_CIRCUIT_BIT_OFFSET_COLNAME = "open_circuit_bit_offset"
    _SHORT_CIRCUIT_A_BIT_OFFSET_COLNAME = "short_circuit_a_bit_offset"
    _SHORT_CIRCUIT_B_BIT_OFFSET_COLNAME = "short_circuit_b_bit_offset"
    _LOOP_DOWN_BIT_OFFSET_COLNAME = "loop_down_bit_offset"
    _OVER_CURRENT_BIT_OFFSET_COLNAME = "over_current_bit_offset"
    _NON_CONFIGURED_BIT_OFFSET_COLNAME = "non_configured_bit_offset"
    _LOOP_MODULE_FAULT_BIT_OFFSET_COLNAME = "loop_module_fault_bit_offset"
    _OPEN_CIRCUIT_DECIMAL_COLNAME = "open_circuit_decimal"
    _SHORT_CIRCUIT_A_DECIMAL_COLNAME = "short_circuit_a_decimal"
    _SHORT_CIRCUIT_B_DECIMAL_COLNAME = "short_circuit_b_decimal"
    _LOOP_DOWN_DECIMAL_COLNAME = "loop_down_decimal"
    _OVER_CURRENT_DECIMAL_COLNAME = "over_current_decimal"
    _NON_CONFIGURED_DECIMAL_COLNAME = "non_configured_decimal"
    _LOOP_MODULE_FAULT_DECIMAL_COLNAME = "loop_module_fault_decimal"
    _GROUPS = {
        "ALARM": [_ALARM_BIT_OFFSET_COLNAME, _ALARM_DECIMAL_COLNAME],
        "PREALARM": [_PREALARM_BIT_OFFSET_COLNAME, _PREALARM_DECIMAL_COLNAME],
        "FAULT": [_FAULT_BIT_OFFSET_COLNAME, _FAULT_DECIMAL_COLNAME],
        "ISOLATE": [_ISOLATE_BIT_OFFSET_COLNAME, _ISOLATE_DECIMAL_COLNAME],
        "LOOP_OPEN_CIRCUIT": [
            _OPEN_CIRCUIT_BIT_OFFSET_COLNAME,
            _OPEN_CIRCUIT_DECIMAL_COLNAME,
        ],
        "LOOP_SHORT_CIRCUIT_A": [
            _SHORT_CIRCUIT_A_BIT_OFFSET_COLNAME,
            _SHORT_CIRCUIT_A_DECIMAL_COLNAME,
        ],
        "LOOP_SHORT_CIRCUIT_B": [
            _SHORT_CIRCUIT_B_BIT_OFFSET_COLNAME,
            _SHORT_CIRCUIT_B_DECIMAL_COLNAME,
        ],
        "LOOP_DOWN": [_LOOP_DOWN_BIT_OFFSET_COLNAME, _LOOP_DOWN_DECIMAL_COLNAME],
        "LOOP_OVER_CURRENT": [
            _OVER_CURRENT_BIT_OFFSET_COLNAME,
            _OVER_CURRENT_DECIMAL_COLNAME,
        ],
        "LOOP_NON_CONFIGURED": [
            _NON_CONFIGURED_BIT_OFFSET_COLNAME,
            _NON_CONFIGURED_DECIMAL_COLNAME,
        ],
        "LOOP_MODULE_FAULT": [
            _LOOP_MODULE_FAULT_BIT_OFFSET_COLNAME,
            _LOOP_MODULE_FAULT_DECIMAL_COLNAME,
        ],
    }

    def __init__(
        self, configuration=None, nodes=None, zones=None, loops=None, devices=None
    ):
        """
        Initialize ModbusMapper with optional configuration from FFPReader or individual DataFrames.

        Args:
            configuration (dict, optional): Should contain keys 'nodes', 'zones', 'loops', 'devices',
                                            each mapping to a DataFrame.
            nodes, zones, loops, devices (pd.DataFrame, optional): Individual DataFrames can be provided directly.
        """
        # Start with empty DataFrames
        self._nodes = pd.DataFrame()
        self._zones = pd.DataFrame()
        self._loops = pd.DataFrame()
        self._devices = pd.DataFrame()

        # If configuration dict is provided, use its values
        if configuration is not None:
            self.nodes = configuration.get("nodes", self._nodes)
            self.zones = configuration.get("zones", self._zones)
            self.loops = configuration.get("loops", self._loops)
            self.devices = configuration.get("devices", self._devices)

        # If individual DataFrames are provided, override the corresponding attributes
        if nodes is not None:
            self.nodes = nodes
        if zones is not None:
            self.zones = zones
        if loops is not None:
            self.loops = loops
        if devices is not None:
            self.devices = devices

    @property
    def zones(self):
        return self._zones

    @zones.setter
    def zones(self, value):
        self._zones = value if value is not None else pd.DataFrame()
        if not self._zones.empty:
            self.add_zone_modbus_mapping()

    @property
    def nodes(self):
        return self._nodes

    @nodes.setter
    def nodes(self, value):
        self._nodes = value if value is not None else pd.DataFrame()
        if not self._nodes.empty:
            self.add_node_modbus_mapping()

    @property
    def loops(self):
        return self._loops

    @loops.setter
    def loops(self, value):
        self._loops = value if value is not None else pd.DataFrame()
        if not self._loops.empty:
            self.add_loop_modbus_mapping()

    @property
    def devices(self):
        return self._devices

    @devices.setter
    def devices(self, value):
        self._devices = value if value is not None else pd.DataFrame()
        if not self._devices.empty:
            self.add_device_modbus_mapping()

    @property
    def modbus_configuration(self):
        return {
            "nodes": self.split_by_modbus_gateway(equipment_type="nodes"),
            "zones": self.split_by_modbus_gateway(equipment_type="zones"),
            "loops": self.split_by_modbus_gateway(equipment_type="loops"),
            "devices": self.split_by_modbus_gateway(equipment_type="devices"),
        }

    def _add_modbus_mapping(
        self, df, id_col, gateway_reg_func, bit_offset_func, bit_offset_cols
    ):
        """
        Add Modbus mapping columns to a DataFrame for a given equipment type.

        Args:
            df (pd.DataFrame): The DataFrame to modify.
            id_col (str): The column name containing the numeric equipment ID (e.g., 'zone', 'node').
            gateway_reg_func (callable): Function that takes an ID and returns (gateway, register).
            bit_offset_func (callable): Function that takes an ID and returns a tuple of bit offsets.
            bit_offset_cols (list[str]): List of column names for the bit offsets.

        Returns:
            pd.DataFrame: A copy of the input DataFrame with added Modbus mapping columns.
        """
        if df is None or df.empty:
            return df
        df = df.copy()
        if id_col is not None:
            df[self._GATEWAY_COLNAME] = df[id_col].apply(
                lambda x: gateway_reg_func(x)[0]
            )
            df[self._HOLDING_REGISTER_COLNAME] = df[id_col].apply(
                lambda x: gateway_reg_func(x)[1]
            )
            df[bit_offset_cols] = df[id_col].apply(
                lambda x: pd.Series(bit_offset_func(x))
            )
        else:
            df[self._GATEWAY_COLNAME] = df.apply(
                lambda row: gateway_reg_func(row)[0], axis=1
            )
            df[self._HOLDING_REGISTER_COLNAME] = df.apply(
                lambda row: gateway_reg_func(row)[1], axis=1
            )
            df[bit_offset_cols] = df.apply(
                lambda row: pd.Series(bit_offset_func(row)), axis=1
            )

        # Add decimal columns for each bit offset column using _GROUPS
        for group, (bit_offset_col, decimal_col) in self._GROUPS.items():
            if bit_offset_col in df.columns:
                df[decimal_col] = df[bit_offset_col].apply(
                    lambda x: (
                        self.calculate_register_decimal(x) if self._is_notnull(x) else 0
                    )
                )

        return df

    def _is_notnull(self, x):
        if isinstance(x, (list, tuple, np.ndarray)):
            # True if at least one element is not null
            return any(pd.notnull(e) for e in x)
        else:
            return pd.notnull(x)

    def add_zone_modbus_mapping(self):
        """
        Adds Modbus mapping columns to the zones DataFrame:
        - gateway
        - holding_register
        - alarm_bit_offset
        - prealarm_bit_offset
        - fault_bit_offset
        - isolate_bit_offset

        Each zone uses 4 bits of a 16-bit register, then next zone uses the next 4 bits:
        Bit 0 = alarm
        Bit 1 = pre-alarm/inv-alarm
        Bit 2 = fault
        Bit 3 = isolate

        Gateway 1 holding register start 6002, end 6251, Zone status bits for 1000 zones
        (1 … 1000).

        Gateway 2 holding register start 12352, end 12601, Zone status bits for 1000 zones
        (1001 … 2000).

        Gateway 3 start 18702 end 18826, Zone status bits for 500 zones
        (2001 … 2500).
        """
        if self._zones is None or self._zones.empty:
            return

        def get_gateway_and_register(zone_num):
            # Gateway 1: zones 1-1000, registers 6002-6251
            if 1 <= zone_num <= 1000:
                gateway = 1
                reg_base = 6002
                reg = reg_base + ((zone_num - 1) // 4)
            # Gateway 2: zones 1001-2000, registers 12352-12601
            elif 1001 <= zone_num <= 2000:
                gateway = 2
                reg_base = 12352
                reg = reg_base + ((zone_num - 1001) // 4)
            # Gateway 3: zones 2001-2500, registers 18702-18826
            elif 2001 <= zone_num <= 2500:
                gateway = 3
                reg_base = 18702
                reg = reg_base + ((zone_num - 2001) // 4)
            else:
                gateway = None
                reg = None
            return gateway, reg

        def get_bit_offsets(zone_num):
            # Each register holds 4 zones, each zone uses 4 bits
            offset = ((zone_num - 1) % 4) * 4
            return offset, offset + 1, offset + 2, offset + 3

        self._zones = self._add_modbus_mapping(
            self._zones,
            self._ZONE_COLNAME,
            get_gateway_and_register,
            get_bit_offsets,
            [
                self._ALARM_BIT_OFFSET_COLNAME,
                self._PREALARM_BIT_OFFSET_COLNAME,
                self._FAULT_BIT_OFFSET_COLNAME,
                self._ISOLATE_BIT_OFFSET_COLNAME,
            ],
        )

    def add_loop_modbus_mapping(self):
        """
        Each loop uses 8 bits:
            Bit 0 = open circuit
            Bit 1 = short circuit on side A
            Bit 2 = short circuit on side B
            Bit 2 & 3 = loop is down
            Bit 3 = over current,
            Bit 4 = non configured
            Bit 5 = loop module fault

        Adds Modbus mapping columns to the loops DataFrame:
        - gateway
        - holding_register
        - open_circuit_bit_offset
        - short_circuit_a_bit_offset
        - short_circuit_b_bit_offset
        - loop_down_bit_offset
        - over_current_bit_offset
        - non_configured_bit_offset
        - loop_module_fault_bit_offset

        Gateway 1: registers 152-196, loops 1-90 (8 bits per loop)
        Gateway 2: registers 6502-6546, loops 91-180 (8 bits per loop)
        Gateway 3: registers 12852-12886, loops 181-250 (8 bits per loop)
        """
        if self._loops is None or self._loops.empty:
            return

        def get_gateway_and_register(loop_num):
            # Gateway 1: loops 1-90, registers 152-196
            if 1 <= loop_num <= 90:
                gateway = 1
                reg_base = 152
                reg = reg_base + ((loop_num - 1) // 2)
            # Gateway 2: loops 91-180, registers 6502-6546
            elif 91 <= loop_num <= 180:
                gateway = 2
                reg_base = 6502
                reg = reg_base + ((loop_num - 91) // 2)
            # Gateway 3: loops 181-250, registers 12852-12886
            elif 181 <= loop_num <= 250:
                gateway = 3
                reg_base = 12852
                reg = reg_base + ((loop_num - 181) // 2)
            else:
                gateway = None
                reg = None
            return gateway, reg

        def get_bit_offsets(loop_num):
            # Each register holds 2 loops, each loop uses 8 bits
            offset = ((loop_num - 1) % 2) * 8
            # Bit mapping for each loop
            return (
                offset,  # open circuit
                offset + 1,  # short circuit A
                offset + 2,  # short circuit B
                [offset + 2, offset + 3],  # loop is down (uses bits 2 & 3)
                offset + 3,  # over current (also part of loop down)
                offset + 4,  # non configured
                offset + 5,  # loop module fault
            )

        self._loops = self._add_modbus_mapping(
            self._loops,
            self._LOOP_COLNAME,
            get_gateway_and_register,
            get_bit_offsets,
            [
                "open_circuit_bit_offset",
                "short_circuit_a_bit_offset",
                "short_circuit_b_bit_offset",
                "loop_down_bit_offset",
                "over_current_bit_offset",
                "non_configured_bit_offset",
                "loop_module_fault_bit_offset",
            ],
        )

    def add_node_modbus_mapping(self):
        """
        Adds Modbus mapping columns to the nodes DataFrame:
        - gateway
        - holding_register
        - alarm_bit_offset
        - fault_bit_offset
        - isolate_bit_offset

        Gateway 1, registers 102 to 126, Front Panel MCP status bits for 100 nodes.
        Each MCP uses 4 bits:
        Bit 0 = alarm
        Bit 1 = not used
        Bit 2 = fault
        Bit 3 = isolate

        Gateway 2 and 3: no node data.
        """
        if self._nodes is None or self._nodes.empty:
            return

        def get_gateway_and_register(node_num):
            # Gateway 1: nodes 1-100, registers 102-126
            if 1 <= node_num <= 100:
                gateway = 1
                reg_base = 102
                reg = reg_base + ((node_num - 1) // 4)
            else:
                gateway = None
                reg = None
            return gateway, reg

        def get_bit_offsets(node_num):
            # Each register holds 4 nodes, each node uses 4 bits
            offset = ((node_num - 1) % 4) * 4
            # Bit 1 is not used
            return offset, offset + 2, offset + 3

        self._nodes = self._add_modbus_mapping(
            self._nodes,
            self._NODE_COLNAME,
            get_gateway_and_register,
            get_bit_offsets,
            [
                self._ALARM_BIT_OFFSET_COLNAME,
                self._FAULT_BIT_OFFSET_COLNAME,
                self._ISOLATE_BIT_OFFSET_COLNAME,
            ],
        )

    def add_device_modbus_mapping(self):
        """
        Adds Modbus mapping columns to the devices DataFrame:
        - gateway
        - holding_register
        - alarm_bit_offset
        - prealarm_bit_offset
        - fault_bit_offset
        - isolate_bit_offset

        Devices are addressed by (loop, device) pair.
        Each loop has up to 128 devices, each device uses 4 bits.
        Each loop uses 32 registers.
        Each register holds 4 devices.
        """

        if self._devices is None or self._devices.empty:
            return

        def get_gateway_and_register(row):
            loop = row[self._LOOP_COLNAME]
            device = row[self._DEVICE_COLNAME]

            # Which device block within the loop
            register_offset_in_loop = (device - 1) // 4
            loop_offset = None
            reg_base = None

            if 1 <= loop <= 90:
                gateway = 1
                loop_offset = loop - 1
                reg_base = 242
            elif 91 <= loop <= 180:
                gateway = 2
                loop_offset = loop - 91
                reg_base = 6592
            elif 181 <= loop <= 250:
                gateway = 3
                loop_offset = loop - 181
                reg_base = 12942
            else:
                return None, None

            # Each loop uses 32 registers
            reg = reg_base + (loop_offset * 32) + register_offset_in_loop
            return gateway, reg

        def get_bit_offsets(row):
            device = row[self._DEVICE_COLNAME]
            offset = ((device - 1) % 4) * 4  # 4 bits per device
            return offset, offset + 1, offset + 2, offset + 3

        self._devices = self._add_modbus_mapping(
            self._devices,
            id_col=None,  # not using a single ID column; pass row in lambda
            gateway_reg_func=lambda row: get_gateway_and_register(row),
            bit_offset_func=lambda row: get_bit_offsets(row),
            bit_offset_cols=[
                self._ALARM_BIT_OFFSET_COLNAME,
                self._PREALARM_BIT_OFFSET_COLNAME,
                self._FAULT_BIT_OFFSET_COLNAME,
                self._ISOLATE_BIT_OFFSET_COLNAME,
            ],
        )

    def split_by_modbus_gateway(self, equipment_type="devices"):
        """
        Split the DataFrame into separate objects for Modbus gateway mapping.

        This method splits the data for a specified equipment type into multiple
        DataFrames based on predefined ranges for Modbus gateways. The split is
        performed differently depending on the equipment type.

        Args:
            equipment_type (str, optional): The type of equipment to process.
                Must be one of the following:
                - "devices": Split by loop ranges (L1 to L90, L91 to L180, L181 to L250).
                - "loops": Split by loop ranges (L1 to L90, L91 to L180, L181 to L250).
                - "nodes": No split; returns a single gateway with all nodes.
                - "zones": Split by zone ranges (Z1 to Z1000, Z1001 to Z2000, Z2001 to Z2500).
                Defaults to "devices".

        Returns:
            list[dict]: A list of dictionaries, where each dictionary represents a Modbus
            gateway and contains the following keys:
                - "gateway" (int): The gateway number.
                - "description" (str): A description of the data range for the gateway.
                - "data" (pd.DataFrame): The DataFrame containing the data for the gateway.

        Notes:
            - For "devices" and "loops", the split is based on the "loop" column:
                - Gateway 1: Loops 1 to 90
                - Gateway 2: Loops 91 to 180
                - Gateway 3: Loops 181 to 250
            - For "zones", the split is based on the _ZONE_COLNAME column:
                - Gateway 1: Zones 1 to 1000
                - Gateway 2: Zones 1001 to 2000
                - Gateway 3: Zones 2001 to 2500
            - For "nodes", no splitting is performed; all nodes are assigned to Gateway 1.
            - If an invalid `equipment_type` is provided, an empty list is returned.
        """
        if equipment_type == "devices" or equipment_type == "loops":
            if equipment_type == "loops":
                df = self.loops.copy()
            else:
                df = self.devices.copy()
            # Split the DataFrame
            df1 = df[df["loop"] <= 90]
            df2 = df[(df["loop"] > 90) & (df["loop"] <= 180)]
            df3 = df[df["loop"] > 180]

            # Create a list of dictionaries
            data = [
                {
                    "gateway": 1,
                    "description": equipment_type + "_L1_to_L90",
                    "data": df1,
                },
                {
                    "gateway": 2,
                    "description": equipment_type + "_L91_to_L180",
                    "data": df2,
                },
                {
                    "gateway": 3,
                    "description": equipment_type + "_L181_to_L250",
                    "data": df3,
                },
            ]
        elif equipment_type == "nodes":
            return [{"gateway": 1, "description": equipment_type, "data": self.nodes}]
        elif equipment_type == "zones":
            df = self.zones.copy()
            # Split the DataFrame
            df1 = df[df[self._ZONE_COLNAME] <= 1000]
            df2 = df[(df[self._ZONE_COLNAME] > 1001) & (df[self._ZONE_COLNAME] <= 2000)]
            df3 = df[df[self._ZONE_COLNAME] > 2000]
            return [
                {
                    "gateway": 1,
                    "description": equipment_type + "_Z1_to_Z1000",
                    "data": df1,
                },
                {
                    "gateway": 2,
                    "description": equipment_type + "_Z1001_to_Z2000",
                    "data": df2,
                },
                {
                    "gateway": 3,
                    "description": equipment_type + "_Z2001_to_Z2500",
                    "data": df3,
                },
            ]
        else:
            return []

        return data

    def calculate_register_decimal(self, *bit_offsets):
        """
        CCalculate the decimal representation of a 16-bit register based on given bit offsets.
        Accepts either multiple integer arguments or a single list-like argument.
        Returns the sum of 2^offset for each offset.

        Args:
            *bit_offsets (int or list-like): One or more bit positions (0–15) to set to 1,
                                            or a single list-like of bit positions to AND.

        Returns:
            int: Decimal value of the 16-bit register.
        """
        import collections.abc

        # If a single argument and it's list-like (but not a string), treat as a list of offsets
        if (
            len(bit_offsets) == 1
            and isinstance(bit_offsets[0], collections.abc.Iterable)
            and not isinstance(bit_offsets[0], str)
        ):
            offsets = list(bit_offsets[0])
        else:
            offsets = list(bit_offsets)

        value = 0
        for offset in offsets:
            if 0 <= offset < 16:
                value |= 1 << offset
            else:
                raise ValueError(
                    f"Invalid bit offset: {offset}. Must be between 0 and 15."
                )
        return value

    def extend_with_all_bit_decimals(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Extend the DataFrame by computing and adding decimal values
        for all *_BIT_OFFSET columns in the register layout.
        """

        def normalize_to_bit_list(offsets):
            if pd.isna(offsets):
                return []
            if isinstance(offsets, int):
                return [offsets]
            if isinstance(offsets, str):
                return [
                    int(o.strip()) for o in offsets.split(",") if o.strip().isdigit()
                ]
            if isinstance(offsets, list):
                return offsets
            raise TypeError(f"Unsupported type for bit offsets: {type(offsets)}")

        # Handle all known bit offset column names
        bit_offset_cols = [
            self._ALARM_BIT_OFFSET_COLNAME,
            self._PREALARM_BIT_OFFSET_COLNAME,
            self._FAULT_BIT_OFFSET_COLNAME,
            self._ISOLATE_BIT_OFFSET_COLNAME,
        ]

        for offset_col in bit_offset_cols:
            if offset_col in df.columns:
                decimal_col = offset_col.replace("_OFFSET", "_DECIMAL")
                df[decimal_col] = df[offset_col].apply(
                    lambda x: self.calculate_register_decimal(*normalize_to_bit_list(x))
                )

        return df


if __name__ == "__main__":
    from ffpreader import FFPReader
    from utils import write_dfs_to_excel_and_format
    import os

    input_dir = "./data/input"
    output_dir = "./data/output"
    ffp_database_filename = "QWP 17.02.25 ASE change rev 2 .ffp"
    ffp_database_filepath = os.path.join(input_dir, ffp_database_filename)
    reader = FFPReader(ffp_database_filepath)

    # Example usage with configuration dict
    mapper = ModbusMapper(configuration=reader.configuration)
    print("\n####\nNodes")
    print(mapper.nodes)
    print("\n####\nZones")
    print(mapper.zones)
    print("\n####\nLoops")
    print(mapper.loops)
    print("\n####\nDevices")
    print(mapper.devices)

    # Example usage with individual DataFrames
    mapper2 = ModbusMapper()
    mapper2.zones = reader.zones
    print(mapper2.zones)

    print("\n####\nModbus Configuration")
    # print(mapper.modbus_configuration)
    excel_file = os.path.join(
        output_dir, ffp_database_filename + ".modbus config" + ".xlsx"
    )
    write_dfs_to_excel_and_format(mapper.modbus_configuration, excel_file)
