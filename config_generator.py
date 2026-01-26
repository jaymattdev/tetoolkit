"""
Configuration Generator Module
Converts Excel specification to TOML configuration files for each document type.
"""

import pandas as pd
import toml
import os
import logging
from pathlib import Path
import ast

# Configure module logger
logger = logging.getLogger(__name__)


class ConfigGenerator:
    """Generate TOML configuration files from Excel specifications."""

    def __init__(self, excel_path, output_dir="configs"):
        """
        Initialize the configuration generator.

        Args:
            excel_path: Path to Excel file with extraction specifications
            output_dir: Directory to save TOML configuration files
        """
        self.excel_path = excel_path
        self.output_dir = output_dir
        self.df = None

    def load_excel(self):
        """Load and validate Excel specification file."""
        try:
            self.df = pd.read_excel(self.excel_path)
            required_cols = ['Source Name', 'Elements to Extract',
                           'Name Start Anchor', 'Name Stop Anchor']

            missing_cols = [col for col in required_cols if col not in self.df.columns]
            if missing_cols:
                raise ValueError(f"Missing required columns: {missing_cols}")

            logger.info(f"Successfully loaded {len(self.df)} document type specifications")
            print(f"Successfully loaded {len(self.df)} document type specifications")
            return True

        except Exception as e:
            logger.error(f"Error loading Excel file: {e}")
            print(f"Error loading Excel file: {e}")
            return False

    def parse_list_column(self, value):
        """
        Parse string representation of list into actual list.

        Args:
            value: String like '["DOB", "SSN"]' or NaN

        Returns:
            List or empty list
        """
        if pd.isna(value) or value == '':
            return []

        try:
            # Handle string representation of list
            if isinstance(value, str):
                return ast.literal_eval(value)
            return value
        except:
            # If parsing fails, return empty list
            return []

    def create_toml_config(self, row):
        """
        Create TOML configuration dictionary from Excel row.

        Args:
            row: Pandas Series representing one document type

        Returns:
            Dictionary formatted for TOML output
        """
        source_name = row['Source Name']
        elements = self.parse_list_column(row['Elements to Extract'])
        name_start = self.parse_list_column(row['Name Start Anchor'])
        name_stop = self.parse_list_column(row['Name Stop Anchor'])

        # Determine if name extraction is needed
        has_name_extraction = len(name_start) > 0 and len(name_stop) > 0

        # Check if ID is in filename (optional column)
        id_in_file = row.get('ID in File', False)
        if pd.isna(id_in_file):
            id_in_file = False

        # Get additional cleaning keywords (optional column)
        additional_keywords = self.parse_list_column(row.get('Additional Keywords', []))

        # Check if names are in reverse order (optional column)
        reverse_name_order = row.get('Reverse Name Order', False)
        if pd.isna(reverse_name_order):
            reverse_name_order = False

        config = {
            'Document': {
                'document_source': source_name,
                'number_of_elements': len(elements),
                'name_extraction': has_name_extraction,
                'id_in_file': bool(id_in_file)
            },
            'Extraction': {
                'active_patterns': elements,
                'name_start_anchors': name_start,
                'name_stop_anchors': name_stop
            },
            'Cleaning': {
                'additional_keywords': additional_keywords,
                'spell_check_threshold': 0.85  # Jaro-Winkler similarity threshold
            },
            'Parsing': {
                'reverse_name_order': bool(reverse_name_order)  # For "Last, First" format
            },
            'Validation': {
                'enable_date_logic': True,
                'enable_positional_outliers': True,
                'enable_within_document_gaps': True,
                'enable_value_reasonableness': True,
                'positional_outlier_threshold': 3.0,
                'within_document_gap_threshold': 2000,
                'date_future_tolerance_days': 0,
                'critical_elements': []  # Can be customized per document type
            },
            'Output': {
                'output_file_name': f"{source_name}_extractions.xlsx"
            }
        }

        logger.debug(f"Created config for {source_name} with {len(elements)} elements")
        return config

    def generate_all_configs(self):
        """Generate TOML config files for all document types in Excel."""
        if self.df is None:
            if not self.load_excel():
                return False

        # Create output directory if it doesn't exist
        os.makedirs(self.output_dir, exist_ok=True)

        generated_configs = []

        for idx, row in self.df.iterrows():
            try:
                config = self.create_toml_config(row)
                source_name = config['Document']['document_source']

                # Save TOML file
                config_path = os.path.join(self.output_dir, f"{source_name}.toml")
                with open(config_path, 'w') as f:
                    toml.dump(config, f)

                generated_configs.append(config_path)
                logger.info(f"Generated config: {config_path}")
                print(f"Generated config: {config_path}")

            except Exception as e:
                logger.error(f"Error generating config for row {idx}: {e}")
                print(f"Error generating config for row {idx}: {e}")
                continue

        logger.info(f"Successfully generated {len(generated_configs)} configuration files")
        print(f"\nSuccessfully generated {len(generated_configs)} configuration files")
        return generated_configs


def main():
    """Example usage of ConfigGenerator."""
    # Example: generator = ConfigGenerator("extraction_specs.xlsx")
    # generator.generate_all_configs()
    pass


if __name__ == "__main__":
    main()
