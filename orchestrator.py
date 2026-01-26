"""
Extraction Orchestrator Module
Coordinates extraction across multiple document types and configurations.
"""

import os
from pathlib import Path
from typing import Dict, List
import pandas as pd
from extractor import TextExtractor


class ExtractionOrchestrator:
    """Orchestrate extraction across multiple document types."""

    def __init__(self, plan_folder, configs_folder, patterns_dict=None):
        """
        Initialize the orchestrator.

        Args:
            plan_folder: Path to plan folder containing document type subfolders
            configs_folder: Path to folder containing TOML configuration files
            patterns_dict: Dictionary mapping element names to regex patterns
        """
        self.plan_folder = Path(plan_folder)
        self.configs_folder = Path(configs_folder)
        self.patterns_dict = patterns_dict or {}
        self.all_extractions = []
        self.statistics = {}

    def get_config_for_source(self, source_name):
        """
        Find configuration file for a given source name.

        Args:
            source_name: Name of the document source

        Returns:
            Path to config file or None
        """
        config_path = self.configs_folder / f"{source_name}.toml"

        if config_path.exists():
            return str(config_path)
        else:
            print(f"Warning: No config found for source '{source_name}'")
            return None

    def process_single_source(self, source_folder):
        """
        Process all documents in a single source folder.

        Args:
            source_folder: Path to folder containing documents of one type

        Returns:
            List of extraction dictionaries
        """
        source_name = source_folder.name
        print(f"\n{'='*60}")
        print(f"Processing source: {source_name}")
        print(f"{'='*60}")

        # Get configuration for this source
        config_path = self.get_config_for_source(source_name)

        if not config_path:
            print(f"Skipping {source_name} - no configuration found")
            return []

        # Initialize extractor with config and patterns
        extractor = TextExtractor(config_path, self.patterns_dict)

        # Extract from all files in the folder
        # Only support .txt files
        file_extensions = ['.txt']
        extractions = extractor.extract_from_folder(str(source_folder), file_extensions)

        print(f"Extracted {len(extractions)} total items from {source_name}")

        return extractions

    def process_all_sources(self):
        """
        Process all source folders in the plan folder.

        Returns:
            List of all extraction dictionaries
        """
        if not self.plan_folder.exists():
            print(f"Error: Plan folder not found: {self.plan_folder}")
            return []

        # Get all subdirectories in the plan folder
        source_folders = [f for f in self.plan_folder.iterdir() if f.is_dir()]

        if not source_folders:
            print(f"Warning: No source folders found in {self.plan_folder}")
            return []

        print(f"Found {len(source_folders)} source folders to process")

        self.all_extractions = []

        # Process each source folder
        for source_folder in source_folders:
            extractions = self.process_single_source(source_folder)
            self.all_extractions.extend(extractions)

        print(f"\n{'='*60}")
        print(f"TOTAL EXTRACTIONS: {len(self.all_extractions)}")
        print(f"{'='*60}\n")

        return self.all_extractions

    def get_extractions_dataframe(self):
        """
        Convert extractions to a pandas DataFrame.

        Returns:
            DataFrame with all extraction results
        """
        if not self.all_extractions:
            return pd.DataFrame()

        df = pd.DataFrame(self.all_extractions)

        # Reorder columns for better readability
        column_order = [
            'source', 'filename', 'element', 'value',
            'extraction_order', 'extraction_position'
        ]

        # Add any additional columns that might exist
        for col in df.columns:
            if col not in column_order:
                column_order.append(col)

        # Only use columns that actually exist in the dataframe
        column_order = [col for col in column_order if col in df.columns]

        return df[column_order]

    def calculate_statistics(self):
        """
        Calculate extraction statistics.

        Returns:
            Dictionary with statistics by source
        """
        if not self.all_extractions:
            return {}

        df = pd.DataFrame(self.all_extractions)

        stats = {}

        # Group by source
        for source in df['source'].unique():
            source_df = df[df['source'] == source]

            # Count unique documents
            num_documents = source_df['filename'].nunique()

            # Count elements found vs not found
            total_elements = len(source_df)
            found_elements = len(source_df[source_df['value'].notna()])
            not_found_elements = total_elements - found_elements

            # Calculate percentage
            found_percentage = (found_elements / total_elements * 100) if total_elements > 0 else 0

            # Get element-level statistics
            element_stats = []
            for element in source_df['element'].unique():
                element_df = source_df[source_df['element'] == element]
                element_found = len(element_df[element_df['value'].notna()])
                element_total = len(element_df)

                element_stats.append({
                    'Element': element,
                    'Found': element_found,
                    'Not Found': element_total - element_found,
                    'Total': element_total,
                    'Found %': round(element_found / element_total * 100, 2) if element_total > 0 else 0
                })

            stats[source] = {
                'documents_processed': num_documents,
                'total_elements': total_elements,
                'found_elements': found_elements,
                'not_found_elements': not_found_elements,
                'found_percentage': round(found_percentage, 2),
                'element_breakdown': element_stats
            }

        self.statistics = stats
        return stats

    def get_statistics_dataframe(self):
        """
        Convert statistics to a pandas DataFrame for output.

        Returns:
            DataFrame with statistics
        """
        if not self.statistics:
            self.calculate_statistics()

        # Create summary statistics
        summary_data = []

        for source, stats in self.statistics.items():
            summary_data.append({
                'Source': source,
                'Documents Processed': stats['documents_processed'],
                'Total Elements': stats['total_elements'],
                'Found': stats['found_elements'],
                'Not Found': stats['not_found_elements'],
                'Found %': stats['found_percentage']
            })

        summary_df = pd.DataFrame(summary_data)

        # Create detailed element-level statistics
        detailed_data = []

        for source, stats in self.statistics.items():
            for elem_stat in stats['element_breakdown']:
                detailed_data.append({
                    'Source': source,
                    **elem_stat
                })

        detailed_df = pd.DataFrame(detailed_data)

        return summary_df, detailed_df


def main():
    """Example usage of ExtractionOrchestrator."""
    # Example:
    # patterns = {
    #     'DOB': r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b',
    #     'SSN': r'\b\d{3}-\d{2}-\d{4}\b'
    # }
    #
    # orchestrator = ExtractionOrchestrator(
    #     plan_folder="plan_ABC",
    #     configs_folder="configs",
    #     patterns_dict=patterns
    # )
    #
    # orchestrator.process_all_sources()
    # stats = orchestrator.calculate_statistics()
    pass


if __name__ == "__main__":
    main()
