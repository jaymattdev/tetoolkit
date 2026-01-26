"""
Output Manager Module
Handles saving extraction results to Excel and pickle formats.
"""

import pandas as pd
import pickle
import os
from pathlib import Path
from datetime import datetime


class OutputManager:
    """Manage output of extraction results to various formats."""

    def __init__(self, output_folder="output"):
        """
        Initialize the output manager.

        Args:
            output_folder: Directory to save output files
        """
        self.output_folder = Path(output_folder)
        self.output_folder.mkdir(parents=True, exist_ok=True)

    def generate_filename(self, plan_name, file_type="xlsx", suffix=""):
        """
        Generate timestamped filename for output.

        Args:
            plan_name: Name of the plan being processed
            file_type: File extension (xlsx or pkl)
            suffix: Optional suffix to add (e.g., "HIGH_CONFIDENCE", "LOW_CONFIDENCE")

        Returns:
            String filename
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if suffix:
            return f"{plan_name}_extractions_{suffix}_{timestamp}.{file_type}"
        return f"{plan_name}_extractions_{timestamp}.{file_type}"

    def save_to_excel(self, extractions_df, summary_stats_df, detailed_stats_df,
                     plan_name, custom_filename=None, suffix="", additional_stats=None):
        """
        Save extraction results and statistics to Excel file with multiple sheets.

        Args:
            extractions_df: DataFrame containing all extraction results
            summary_stats_df: DataFrame with summary statistics
            detailed_stats_df: DataFrame with detailed element-level statistics
            plan_name: Name of the plan being processed
            custom_filename: Optional custom filename (otherwise auto-generated)
            suffix: Optional suffix for filename (e.g., "HIGH_CONFIDENCE")
            additional_stats: Optional dict of additional stat DataFrames

        Returns:
            Path to saved Excel file
        """
        if custom_filename:
            filename = custom_filename
        else:
            filename = self.generate_filename(plan_name, "xlsx", suffix)

        output_path = self.output_folder / filename

        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Write extractions to first sheet
                extractions_df.to_excel(writer, sheet_name='Raw Extractions', index=False)

                # Write legacy statistics
                summary_stats_df.to_excel(writer, sheet_name='Summary Statistics', index=False)
                detailed_stats_df.to_excel(writer, sheet_name='Detailed Statistics', index=False)

                # Write additional statistics if provided
                if additional_stats:
                    if 'element_statistics' in additional_stats and not additional_stats['element_statistics'].empty:
                        additional_stats['element_statistics'].to_excel(
                            writer, sheet_name='Element Stats', index=False)

                    if 'parsing_statistics' in additional_stats and not additional_stats['parsing_statistics'].empty:
                        additional_stats['parsing_statistics'].to_excel(
                            writer, sheet_name='Parsing Stats', index=False)

                    if 'confidence_statistics' in additional_stats and not additional_stats['confidence_statistics'].empty:
                        additional_stats['confidence_statistics'].to_excel(
                            writer, sheet_name='Confidence Stats', index=False)

                    if 'flag_statistics' in additional_stats and not additional_stats['flag_statistics'].empty:
                        additional_stats['flag_statistics'].to_excel(
                            writer, sheet_name='Flag Stats', index=False)

                    if 'participant_statistics' in additional_stats and not additional_stats['participant_statistics'].empty:
                        additional_stats['participant_statistics'].to_excel(
                            writer, sheet_name='Participant Stats', index=False)

                    if 'timing_statistics' in additional_stats and not additional_stats['timing_statistics'].empty:
                        additional_stats['timing_statistics'].to_excel(
                            writer, sheet_name='Timing Stats', index=False)

                # Auto-adjust column widths for better readability
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter

                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass

                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width

            print(f"Excel file saved: {output_path}")
            return str(output_path)

        except Exception as e:
            print(f"Error saving Excel file: {e}")
            return None

    def save_to_pickle(self, extractions_list, plan_name, custom_filename=None, suffix=""):
        """
        Save raw extraction list to pickle file for downstream processing.

        Args:
            extractions_list: List of extraction dictionaries
            plan_name: Name of the plan being processed
            custom_filename: Optional custom filename (otherwise auto-generated)
            suffix: Optional suffix for filename (e.g., "HIGH_CONFIDENCE")

        Returns:
            Path to saved pickle file
        """
        if custom_filename:
            filename = custom_filename
        else:
            filename = self.generate_filename(plan_name, "pkl", suffix)

        output_path = self.output_folder / filename

        try:
            with open(output_path, 'wb') as f:
                pickle.dump(extractions_list, f, protocol=pickle.HIGHEST_PROTOCOL)

            print(f"Pickle file saved: {output_path}")
            return str(output_path)

        except Exception as e:
            print(f"Error saving pickle file: {e}")
            return None

    def load_from_pickle(self, pickle_path):
        """
        Load extraction results from pickle file.

        Args:
            pickle_path: Path to pickle file

        Returns:
            List of extraction dictionaries
        """
        try:
            with open(pickle_path, 'rb') as f:
                data = pickle.load(f)

            print(f"Loaded {len(data)} extractions from {pickle_path}")
            return data

        except Exception as e:
            print(f"Error loading pickle file: {e}")
            return None

    def save_all_outputs(self, extractions_list, summary_stats_df, detailed_stats_df, plan_name, suffix="", additional_stats=None):
        """
        Save all output formats (Excel and pickle).

        Args:
            extractions_list: List of extraction dictionaries
            summary_stats_df: DataFrame with summary statistics
            detailed_stats_df: DataFrame with detailed element-level statistics
            plan_name: Name of the plan being processed
            suffix: Optional suffix for filename (e.g., "HIGH_CONFIDENCE", "LOW_CONFIDENCE")
            additional_stats: Optional dict of additional stat DataFrames

        Returns:
            Tuple of (excel_path, pickle_path)
        """
        if not suffix:
            print(f"\nSaving outputs for plan: {plan_name}")
            print(f"Output directory: {self.output_folder}")

        # Convert extractions list to DataFrame
        extractions_df = pd.DataFrame(extractions_list)

        # Save Excel
        excel_path = self.save_to_excel(
            extractions_df,
            summary_stats_df,
            detailed_stats_df,
            plan_name,
            suffix=suffix,
            additional_stats=additional_stats
        )

        # Save pickle
        pickle_path = self.save_to_pickle(extractions_list, plan_name, suffix=suffix)

        if not suffix:
            print(f"\nOutput files created:")
            print(f"  Excel: {excel_path}")
            print(f"  Pickle: {pickle_path}")

        return excel_path, pickle_path

    def create_extraction_summary_report(self, summary_stats_df):
        """
        Print a formatted summary report to console.

        Args:
            summary_stats_df: DataFrame with summary statistics
        """
        print("\n" + "="*80)
        print("EXTRACTION SUMMARY REPORT")
        print("="*80)

        for _, row in summary_stats_df.iterrows():
            print(f"\n{row['Source']}:")
            print(f"  Documents Processed: {row['Documents Processed']}")
            print(f"  Total Elements: {row['Total Elements']}")
            print(f"  Found: {row['Found']} ({row['Found %']}%)")
            print(f"  Not Found: {row['Not Found']}")

        print("\n" + "="*80 + "\n")


def main():
    """Example usage of OutputManager."""
    # Example:
    # output_mgr = OutputManager(output_folder="output")
    # output_mgr.save_all_outputs(
    #     extractions_list=all_extractions,
    #     summary_stats_df=summary_df,
    #     detailed_stats_df=detailed_df,
    #     plan_name="plan_ABC"
    # )
    pass


if __name__ == "__main__":
    main()
