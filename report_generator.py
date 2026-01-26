"""
Interactive Excel Report Generator
Creates user-friendly Excel reports for database building and data review.

Features:
- Filterable extraction tables with SharePoint links
- Participant-level summary views
- Separate tabs for high/low confidence data
- Direct links to documents and sources
- Conditional formatting for missing values
"""

import pandas as pd
import logging
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo

# Configure module logger
logger = logging.getLogger(__name__)


class ReportGenerator:
    """Generate interactive Excel reports for database building."""

    def __init__(self, sharepoint_base_url: str = "https://yourcompany.sharepoint.com/sites/yoursite/"):
        """
        Initialize the report generator.

        Args:
            sharepoint_base_url: Base URL for SharePoint document links
        """
        self.sharepoint_base_url = sharepoint_base_url
        logger.info("Initialized ReportGenerator")

    def create_sharepoint_link(self, source: str, filename: str) -> str:
        """
        Create SharePoint link for a document.

        Args:
            source: Source folder name
            filename: Document filename

        Returns:
            SharePoint URL (placeholder format)
        """
        # Placeholder format - user will replace with actual SharePoint structure
        return f"{self.sharepoint_base_url}{source}/{filename}"

    def prepare_extraction_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Prepare extraction data for report with links and formatting.

        Args:
            df: DataFrame with validated extractions

        Returns:
            Prepared DataFrame
        """
        if df.empty:
            logger.warning("Empty dataframe provided")
            return df

        logger.info(f"Preparing {len(df)} extractions for report")

        # Create a copy
        report_df = df.copy()

        # Create SharePoint links
        report_df['Document Link'] = report_df.apply(
            lambda row: self.create_sharepoint_link(row['source'], row['filename']),
            axis=1
        )

        # Create source link
        report_df['Source Link'] = report_df['source'].apply(
            lambda source: f"{self.sharepoint_base_url}{source}/"
        )

        # For missing cleaned values, use document link
        if 'cleaned_value' in report_df.columns:
            mask = report_df['cleaned_value'].isna()
            report_df.loc[mask, 'cleaned_value'] = report_df.loc[mask, 'Document Link']

        # Select and reorder columns for report
        columns = []

        # Core columns
        if 'participant_id' in report_df.columns:
            columns.append('participant_id')
        columns.extend(['source', 'filename', 'element'])

        # Value columns
        if 'value' in report_df.columns:
            columns.append('value')
        if 'cleaned_value' in report_df.columns:
            columns.append('cleaned_value')

        # Name columns if present
        name_cols = ['FNAME', 'LNAME', 'BFNAME', 'BLNAME', 'SFNAME', 'SLNAME']
        for col in name_cols:
            if col in report_df.columns:
                columns.append(col)

        # Confidence and flags
        if 'confidence' in report_df.columns:
            columns.append('confidence')
        if 'flags' in report_df.columns:
            columns.append('flags')
        if 'flag_reasons' in report_df.columns:
            columns.append('flag_reasons')

        # Link columns at end
        columns.extend(['Document Link', 'Source Link'])

        # Only include columns that exist
        columns = [col for col in columns if col in report_df.columns]

        report_df = report_df[columns]

        # Rename columns for readability
        rename_map = {
            'participant_id': 'Participant ID',
            'source': 'Source',
            'filename': 'Filename',
            'element': 'Element',
            'value': 'Raw Value',
            'cleaned_value': 'Cleaned Value',
            'confidence': 'Confidence',
            'flags': 'Flags',
            'flag_reasons': 'Flag Reasons'
        }
        report_df = report_df.rename(columns=rename_map)

        logger.info(f"Prepared report with {len(report_df)} rows and {len(report_df.columns)} columns")
        return report_df

    def create_participant_summary(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Create participant-level summary across all sources.

        Args:
            df: DataFrame with extractions

        Returns:
            Participant summary DataFrame
        """
        if df.empty or 'participant_id' not in df.columns:
            logger.info("No participant ID column - skipping participant summary")
            return pd.DataFrame()

        logger.info("Creating participant summary")

        # Remove rows without participant ID
        df_with_id = df[df['participant_id'].notna()].copy()

        if df_with_id.empty:
            logger.warning("No extractions with participant IDs")
            return pd.DataFrame()

        # Pivot data by participant and element
        summary_rows = []

        for participant_id, group in df_with_id.groupby('participant_id'):
            participant_data = {'Participant ID': participant_id}

            # Get document link for this participant (use first one)
            first_row = group.iloc[0]
            participant_data['Document Link'] = self.create_sharepoint_link(
                first_row['source'], first_row['filename']
            )

            # Add each element as a column
            for _, row in group.iterrows():
                element = row['element']
                source = row['source']

                # Use cleaned value if available, otherwise raw value
                value = row.get('cleaned_value')
                if pd.isna(value):
                    value = row.get('value')

                # Create column name with source prefix for clarity
                col_name = f"{source}_{element}"

                participant_data[col_name] = value

            summary_rows.append(participant_data)

        summary_df = pd.DataFrame(summary_rows)

        # Sort columns: Participant ID first, then alphabetically
        cols = ['Participant ID']
        if 'Document Link' in summary_df.columns:
            cols.append('Document Link')

        other_cols = sorted([col for col in summary_df.columns if col not in cols])
        cols.extend(other_cols)

        summary_df = summary_df[cols]

        logger.info(f"Created participant summary with {len(summary_df)} participants")
        return summary_df

    def apply_excel_formatting(self, filepath: str):
        """
        Apply Excel formatting to the report.

        Args:
            filepath: Path to Excel file
        """
        try:
            logger.info(f"Applying formatting to {filepath}")
            wb = load_workbook(filepath)

            # Define styles
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            warning_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            link_font = Font(color="0563C1", underline="single")
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]

                # Skip empty sheets
                if ws.max_row == 1:
                    continue

                # Format header row
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')

                # Auto-adjust column widths and apply borders
                for column in ws.columns:
                    max_length = 0
                    column_letter = column[0].column_letter

                    for cell in column:
                        # Apply borders
                        cell.border = border

                        # Calculate max length
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass

                        # Format links
                        if isinstance(cell.value, str):
                            if cell.value.startswith('http://') or cell.value.startswith('https://'):
                                cell.font = link_font
                                cell.hyperlink = cell.value

                        # Highlight missing values (if cleaned value is a link, it was missing)
                        if cell.column == ws.max_column - 1:  # Cleaned Value column (usually)
                            if isinstance(cell.value, str) and cell.value.startswith('http'):
                                cell.fill = warning_fill

                    # Set column width
                    adjusted_width = min(max_length + 2, 50)
                    ws.column_dimensions[column_letter].width = adjusted_width

                # Freeze top row
                ws.freeze_panes = 'A2'

                # Enable autofilter on header row
                if ws.max_row > 1:
                    ws.auto_filter.ref = ws.dimensions

            wb.save(filepath)
            logger.info("Formatting applied successfully")

        except Exception as e:
            logger.error(f"Error applying formatting: {e}", exc_info=True)

    def generate_interactive_report(
        self,
        validated_df: pd.DataFrame,
        output_path: str,
        plan_name: str
    ) -> str:
        """
        Generate comprehensive interactive Excel report.

        Args:
            validated_df: DataFrame with validated extractions
            output_path: Path to save report
            plan_name: Name of the plan

        Returns:
            Path to generated report
        """
        logger.info(f"Generating interactive report for {plan_name}")

        # Separate by confidence
        high_conf_df = validated_df[validated_df.get('confidence', 'HIGH') == 'HIGH'].copy()
        low_conf_df = validated_df[validated_df.get('confidence', 'LOW') == 'LOW'].copy()

        # Prepare data
        high_conf_report = self.prepare_extraction_data(high_conf_df)
        low_conf_report = self.prepare_extraction_data(low_conf_df)

        # Create participant summaries
        high_conf_participant = self.create_participant_summary(high_conf_df)
        low_conf_participant = self.create_participant_summary(low_conf_df)

        # Create Excel writer
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # High confidence extractions (main tab)
            if not high_conf_report.empty:
                high_conf_report.to_excel(writer, sheet_name='High Confidence Data', index=False)
                logger.info(f"Added High Confidence Data sheet with {len(high_conf_report)} rows")

            # High confidence participant summary
            if not high_conf_participant.empty:
                high_conf_participant.to_excel(writer, sheet_name='Participant Summary', index=False)
                logger.info(f"Added Participant Summary sheet with {len(high_conf_participant)} participants")

            # Low confidence extractions (review tab)
            if not low_conf_report.empty:
                low_conf_report.to_excel(writer, sheet_name='Low Confidence - Review', index=False)
                logger.info(f"Added Low Confidence Review sheet with {len(low_conf_report)} rows")

            # Low confidence participant summary
            if not low_conf_participant.empty:
                low_conf_participant.to_excel(writer, sheet_name='Participant Review', index=False)
                logger.info(f"Added Participant Review sheet with {len(low_conf_participant)} participants")

            # Add instructions sheet
            instructions = self._create_instructions(plan_name)
            instructions.to_excel(writer, sheet_name='Instructions', index=False)

        # Apply formatting
        self.apply_excel_formatting(output_path)

        logger.info(f"Interactive report generated: {output_path}")
        return output_path

    def _create_instructions(self, plan_name: str) -> pd.DataFrame:
        """
        Create instructions sheet for report users.

        Args:
            plan_name: Name of the plan

        Returns:
            DataFrame with instructions
        """
        instructions = [
            {'Section': 'OVERVIEW', 'Details': f'Interactive Extraction Report for {plan_name}'},
            {'Section': '', 'Details': ''},
            {'Section': 'SHEET GUIDE', 'Details': ''},
            {'Section': '1. High Confidence Data', 'Details': 'All extractions that passed validation - ready for database import'},
            {'Section': '2. Participant Summary', 'Details': 'Per-participant view of all extractions across sources'},
            {'Section': '3. Low Confidence - Review', 'Details': 'Flagged extractions requiring manual review'},
            {'Section': '4. Participant Review', 'Details': 'Per-participant view of flagged extractions'},
            {'Section': '', 'Details': ''},
            {'Section': 'HOW TO USE', 'Details': ''},
            {'Section': 'Filtering', 'Details': 'Use filter dropdowns in header row to filter by Source, Participant ID, or Element'},
            {'Section': 'Document Links', 'Details': 'Click "Document Link" to open source document in SharePoint (requires link update)'},
            {'Section': 'Source Links', 'Details': 'Click "Source Link" to open source folder in SharePoint'},
            {'Section': 'Missing Values', 'Details': 'Cells highlighted in red have missing cleaned values - click link to review document'},
            {'Section': '', 'Details': ''},
            {'Section': 'UPDATING SHAREPOINT LINKS', 'Details': ''},
            {'Section': 'Step 1', 'Details': 'Find & Replace: https://yourcompany.sharepoint.com/sites/yoursite/'},
            {'Section': 'Step 2', 'Details': 'Replace with your actual SharePoint base URL'},
            {'Section': 'Step 3', 'Details': 'Verify folder structure matches: {base_url}/{source_name}/{filename}'},
            {'Section': '', 'Details': ''},
            {'Section': 'DATABASE BUILDING WORKFLOW', 'Details': ''},
            {'Section': '1. Review High Confidence Data', 'Details': 'Import high-confidence extractions into database'},
            {'Section': '2. Check Participant Summary', 'Details': 'Verify data completeness per participant'},
            {'Section': '3. Review Flagged Items', 'Details': 'Manually verify low-confidence extractions'},
            {'Section': '4. Fill Missing Data', 'Details': 'Use document links to retrieve missing values from source documents'},
            {'Section': '5. Verify Cross-Source', 'Details': 'Compare same elements across different sources for consistency'},
            {'Section': '', 'Details': ''},
            {'Section': 'TIPS', 'Details': ''},
            {'Section': 'Sort by Participant', 'Details': 'Easy to review all data for one person'},
            {'Section': 'Sort by Element', 'Details': 'Easy to review all values of same type (DOB, SSN, etc.)'},
            {'Section': 'Sort by Source', 'Details': 'Easy to review all extractions from one document type'},
            {'Section': 'Filter by Flags', 'Details': 'Focus on specific validation issues'},
            {'Section': '', 'Details': ''},
            {'Section': 'COLUMNS EXPLAINED', 'Details': ''},
            {'Section': 'Participant ID', 'Details': 'Unique identifier (SSN, Employee ID, etc.)'},
            {'Section': 'Source', 'Details': 'Document type (Birth_Certificate, Employment_Application, etc.)'},
            {'Section': 'Filename', 'Details': 'Name of source document'},
            {'Section': 'Element', 'Details': 'Data element extracted (DOB, SSN, NAME, etc.)'},
            {'Section': 'Raw Value', 'Details': 'Original extracted text from document'},
            {'Section': 'Cleaned Value', 'Details': 'Standardized/parsed value (dates formatted, names split, etc.)'},
            {'Section': 'FNAME/LNAME', 'Details': 'Parsed name components (if applicable)'},
            {'Section': 'Confidence', 'Details': 'HIGH = passed validation, LOW = flagged for review'},
            {'Section': 'Flags', 'Details': 'Validation issue types (positional_outlier, date_logic_violation, etc.)'},
            {'Section': 'Flag Reasons', 'Details': 'Detailed explanation of validation flags'},
            {'Section': 'Document Link', 'Details': 'Direct link to source document (SharePoint)'},
            {'Section': 'Source Link', 'Details': 'Direct link to source folder (SharePoint)'},
        ]

        return pd.DataFrame(instructions)


def main():
    """Example usage of ReportGenerator."""
    # Example:
    # report_gen = ReportGenerator(sharepoint_base_url="https://mycompany.sharepoint.com/sites/mysite/")
    #
    # df = pd.DataFrame(validated_extractions)
    # report_path = report_gen.generate_interactive_report(
    #     validated_df=df,
    #     output_path="plans/plan_ABC/output/plan_ABC_INTERACTIVE_REPORT.xlsx",
    #     plan_name="plan_ABC"
    # )
    pass


if __name__ == "__main__":
    main()
