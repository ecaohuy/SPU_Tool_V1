"""Dynamic mapping engine for SPU Processing Tool.

Parses the Mapping sheet and applies mappings to transform input data to output format.
"""

import pandas as pd
import re


def safe_int(value, default=None):
    """Safely convert value to int, returning default if conversion fails."""
    if pd.isna(value):
        return default
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return default


class MappingEngine:
    """Engine to apply dynamic mappings from Mapping sheet."""

    def __init__(self, mapping_df, config, version="V1.70.26"):
        """Initialize the mapping engine.

        Args:
            mapping_df: DataFrame from the Mapping sheet
            config: Loaded config.json dictionary
            version: Version string to filter mappings (e.g., "V1.70.26")
        """
        self.config = config
        self.version = version
        self.mappings = self._parse_mappings(mapping_df)
        self.spu_config = self._get_spu_config()

    def _get_spu_config(self):
        """Get SPU configuration for the current version."""
        spu = self.config.get("SPU", {})
        return spu.get(self.version, {})

    def _parse_mappings(self, mapping_df):
        """Parse mapping rules from the Mapping sheet.

        Args:
            mapping_df: DataFrame with columns: Version, Sheet, Column,
                       SourceSheet, SourceColumn, FixedValue, Note, Meaning

        Returns:
            dict: Mappings organized by target sheet
        """
        mappings = {}

        if mapping_df.empty:
            return mappings

        # Filter by version if Version column exists
        if "Version" in mapping_df.columns:
            # Include rows where version matches or version is empty/NaN
            version_mask = (
                (mapping_df["Version"] == self.version) |
                (mapping_df["Version"].isna()) |
                (mapping_df["Version"] == "")
            )
            filtered_df = mapping_df[version_mask]
        else:
            filtered_df = mapping_df

        # Group by target sheet
        for _, row in filtered_df.iterrows():
            target_sheet = str(row.get("Sheet", "")).strip()
            if not target_sheet:
                continue

            if target_sheet not in mappings:
                mappings[target_sheet] = []

            mapping_rule = {
                "target_column": str(row.get("Column", "")).strip(),
                "source_sheet": str(row.get("SourceSheet", "")).strip(),
                "source_column": str(row.get("SourceColumn", "")).strip(),
                "fixed_value": row.get("FixedValue"),
                "note": str(row.get("Note", "")).strip(),
                "meaning": str(row.get("Meaning", "")).strip()
            }

            # Only add if there's a valid target column
            if mapping_rule["target_column"]:
                mappings[target_sheet].append(mapping_rule)

        return mappings

    def get_target_sheets(self):
        """Get list of target sheets that have mappings defined.

        Returns:
            list: List of target sheet names
        """
        return list(self.mappings.keys())

    def apply_mapping(self, source_data, target_sheet):
        """Apply mappings to generate data for a target sheet.

        Args:
            source_data: Dictionary of DataFrames {sheet_name: DataFrame}
            target_sheet: Name of the target sheet to generate

        Returns:
            pd.DataFrame: Mapped data for the target sheet
        """
        if target_sheet not in self.mappings:
            return pd.DataFrame()

        rules = self.mappings[target_sheet]
        if not rules:
            return pd.DataFrame()

        # Determine the number of rows from source sheets
        max_rows = 0
        for rule in rules:
            source_sheet = rule["source_sheet"]
            if source_sheet in source_data and not source_data[source_sheet].empty:
                max_rows = max(max_rows, len(source_data[source_sheet]))

        if max_rows == 0:
            return pd.DataFrame()

        # Build the output DataFrame
        result = {}
        for rule in rules:
            target_col = rule["target_column"]
            values = self._apply_single_rule(rule, source_data, max_rows)
            if values is not None:
                result[target_col] = values

        return pd.DataFrame(result)

    def _apply_single_rule(self, rule, source_data, num_rows):
        """Apply a single mapping rule.

        Args:
            rule: Mapping rule dictionary
            source_data: Dictionary of DataFrames
            num_rows: Number of rows to generate

        Returns:
            list or pd.Series: Values for the target column
        """
        fixed_value = rule["fixed_value"]
        source_sheet = rule["source_sheet"]
        source_column = rule["source_column"]
        note = rule["note"]

        # Case 1: Fixed value
        if pd.notna(fixed_value) and str(fixed_value).strip():
            fixed_val = str(fixed_value).strip()
            # Check if it's a config lookup (e.g., "config:mcc")
            if fixed_val.startswith("config:"):
                config_key = fixed_val[7:]  # Remove "config:" prefix
                config_value = self._get_config_value(config_key)
                return [config_value] * num_rows
            return [fixed_val] * num_rows

        # Case 2: Source column mapping
        if source_sheet and source_column:
            if source_sheet in source_data:
                df = source_data[source_sheet]
                if source_column in df.columns:
                    values = df[source_column].tolist()
                    # Pad with None if needed
                    while len(values) < num_rows:
                        values.append(None)
                    return values[:num_rows]

        # Case 3: Config mapping based on note
        if note:
            config_values = self._apply_config_mapping(note, source_data, num_rows)
            if config_values is not None:
                return config_values

        return [None] * num_rows

    def _get_config_value(self, key_path):
        """Get value from config using dot notation path.

        Args:
            key_path: Path like "mcc" or "SPU.V1.70.26.bandwidth_mapping"

        Returns:
            Value from config or empty string if not found
        """
        keys = key_path.split(".")
        value = self.config

        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return ""

        return str(value) if value is not None else ""

    def _apply_config_mapping(self, note, source_data, num_rows):
        """Apply configuration-based mapping.

        Args:
            note: Note field that may contain mapping instructions
            source_data: Dictionary of DataFrames
            num_rows: Number of rows

        Returns:
            list: Mapped values or None if no mapping applied
        """
        note_lower = note.lower()

        # Bandwidth mapping
        if "bandwidth" in note_lower:
            return self._apply_bandwidth_mapping(source_data, num_rows)

        # EARFCN mapping
        if "earfcn" in note_lower or "frequency" in note_lower:
            return self._apply_earfcn_mapping(source_data, num_rows)

        # MME mapping
        if "mme" in note_lower:
            return self._apply_mme_mapping(source_data, num_rows)

        # AMF mapping
        if "amf" in note_lower:
            return self._apply_amf_mapping(source_data, num_rows)

        return None

    def _apply_bandwidth_mapping(self, source_data, num_rows):
        """Apply bandwidth mapping from config."""
        bw_mapping = self.spu_config.get("bandwidth_mapping", {})

        # Get bandwidth values from Radio 4G or Radio 5G
        result = []
        for i in range(num_rows):
            # Try Radio 4G first
            if "Radio 4G" in source_data and not source_data["Radio 4G"].empty:
                df = source_data["Radio 4G"]
                if i < len(df) and "dlChannelBandwidth" in df.columns:
                    bw = df.iloc[i]["dlChannelBandwidth"]
                    mapped = bw_mapping.get(str(bw), bw)
                    result.append(mapped)
                    continue
            result.append(None)

        return result if any(v is not None for v in result) else None

    def _apply_earfcn_mapping(self, source_data, num_rows):
        """Apply EARFCN to frequency mapping from config."""
        earfcn_mapping = self.spu_config.get("earfcn_mapping", {})

        result = []
        for i in range(num_rows):
            if "Radio 4G" in source_data and not source_data["Radio 4G"].empty:
                df = source_data["Radio 4G"]
                if i < len(df) and "arfcndl" in df.columns:
                    earfcn = df.iloc[i]["arfcndl"]
                    earfcn_int = safe_int(earfcn)
                    if earfcn_int is not None:
                        mapped = earfcn_mapping.get(str(earfcn_int), earfcn)
                        result.append(mapped)
                        continue
            result.append(None)

        return result if any(v is not None for v in result) else None

    def _apply_mme_mapping(self, source_data, num_rows):
        """Apply MME IP mapping from config."""
        mme_config = self.config.get("mme", {})

        result = []
        for i in range(num_rows):
            if "IP" in source_data and not source_data["IP"].empty:
                df = source_data["IP"]
                if i < len(df) and "MME" in df.columns:
                    mme_names = str(df.iloc[i]["MME"]).split()
                    ips = []
                    for mme_name in mme_names:
                        if mme_name in mme_config:
                            ips.extend(mme_config[mme_name])
                    result.append(";".join(ips) if ips else None)
                    continue
            result.append(None)

        return result if any(v is not None for v in result) else None

    def _apply_amf_mapping(self, source_data, num_rows):
        """Apply AMF IP mapping from config."""
        amf_config = self.config.get("amf", {})

        result = []
        for i in range(num_rows):
            if "IP" in source_data and not source_data["IP"].empty:
                df = source_data["IP"]
                if i < len(df) and "AMF" in df.columns:
                    amf_names = str(df.iloc[i]["AMF"]).split()
                    ips = []
                    for amf_name in amf_names:
                        if amf_name in amf_config:
                            ips.extend(amf_config[amf_name])
                    result.append(";".join(ips) if ips else None)
                    continue
            result.append(None)

        return result if any(v is not None for v in result) else None

    def get_groups(self, source_data):
        """Get unique groups from the IP sheet.

        Args:
            source_data: Dictionary of DataFrames

        Returns:
            list: List of unique group names
        """
        if "IP" in source_data and not source_data["IP"].empty:
            df = source_data["IP"]
            if "Group" in df.columns:
                groups = df["Group"].dropna().unique().tolist()
                return [str(g) for g in groups]
        return ["default"]

    def filter_by_group(self, source_data, group):
        """Filter source data by group.

        Args:
            source_data: Dictionary of DataFrames
            group: Group name to filter by

        Returns:
            dict: Filtered DataFrames
        """
        filtered = {}

        for sheet_name, df in source_data.items():
            if df.empty:
                filtered[sheet_name] = df
                continue

            # Get NE_Names for this group from IP sheet
            if "IP" in source_data and "Group" in source_data["IP"].columns:
                ip_df = source_data["IP"]
                group_ne_names = ip_df[ip_df["Group"] == group]["NE_Name"].tolist()

                if "NE_Name" in df.columns:
                    filtered[sheet_name] = df[df["NE_Name"].isin(group_ne_names)]
                else:
                    filtered[sheet_name] = df
            else:
                filtered[sheet_name] = df

        return filtered
