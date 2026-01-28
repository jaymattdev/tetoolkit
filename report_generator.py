"""
Interactive Excel Report Generator - Redesigned
Creates user-friendly Excel reports optimized for data comparison and review.

New Structure:
1. All Extracted Data - Unified view with rich filtering
2. Value Comparison - Cross-source and within-source conflict detection
3. Review Needed - Low confidence + conflicts requiring attention
4. Source Statistics - Performance metrics by source
5. Participant Statistics - Completeness metrics by participant
"""

import pandas as pd
import logging
import numpy as np
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, Reference
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
import matplotlib.pyplot as plt
import io
from openpyxl.drawing.image import Image as OpenpyxlImage

# Configure module logger
logger = logging.getLogger(__name__)


class ReportGenerator:
    """Generate comprehensive interactive Excel reports for data review."""

    def __init__(self, sharepoint_base_url: str = "https://yourcompany.sharepoint.com/sites/yoursite/"):
        """
        Initialize the report generator.

        Args:
            sharepoint_base_url: Base URL for SharePoint document links
        """
        self.sharepoint_base_url = sharepoint_base_url
        logger.info("Initialized ReportGenerator (Redesigned)")

    def create_sharepoint_link(self, source: str, filename: str) -> str:
        """
        Create SharePoint link for a document.

        Args:
            source: Source folder name
            filename: Document filename

        Returns:
            SharePoint URL
        """
        return f"{self.sharepoint_base_url}{source}/{filename}"

    # ====================
    # TAB 1: ALL EXTRACTED DATA
    # ====================

    def create_all_data_tab(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Create unified view of all extracted data with rich filtering.

        Args:
            df: DataFrame with validated extractions

        Returns:
            Formatted DataFrame for All Extracted Data tab
        """
        if df.empty:
            logger.warning("Empty dataframe provided for all data tab")
            return pd.DataFrame()

        logger.info(f"Creating All Extracted Data tab with {len(df)} rows")

        # Create a copy
        report_df = df.copy()

        # Add SharePoint links
        report_df['Document Link'] = report_df.apply(
            lambda row: self.create_sharepoint_link(row['source'], row['filename']),
            axis=1
        )

        # Select columns for report
        columns = []

        if 'participant_id' in report_df.columns:
            columns.append('participant_id')

        columns.extend(['source', 'filename', 'element'])

        if 'value' in report_df.columns:
            columns.append('value')
        if 'cleaned_value' in report_df.columns:
            columns.append('cleaned_value')

        if 'confidence' in report_df.columns:
            columns.append('confidence')
        if 'flags' in report_df.columns:
            columns.append('flags')
        if 'flag_reasons' in report_df.columns:
            columns.append('flag_reasons')

        columns.append('Document Link')

        # Only include columns that exist
        columns = [col for col in columns if col in report_df.columns]
        report_df = report_df[columns]

        # Rename for readability
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

        logger.info(f"All Extracted Data tab ready with {len(report_df)} rows")
        return report_df

    # ====================
    # TAB 2: VALUE COMPARISON
    # ====================

    def create_value_comparison_tab(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Create value comparison view showing cross-source and within-source conflicts.

        Args:
            df: DataFrame with validated extractions

        Returns:
            Comparison DataFrame
        """
        if df.empty or 'participant_id' not in df.columns:
            logger.warning("No participant IDs for value comparison")
            return pd.DataFrame()

        logger.info("Creating Value Comparison tab")

        # Remove rows without participant ID
        df_with_id = df[df['participant_id'].notna()].copy()

        if df_with_id.empty:
            logger.warning("No extractions with participant IDs")
            return pd.DataFrame()

        comparison_rows = []

        # Group by participant and element
        for (participant_id, element), group in df_with_id.groupby(['participant_id', 'element']):

            # Get all unique values for this element
            values_by_source = {}

            for _, row in group.iterrows():
                source = row['source']
                value = row.get('cleaned_value')
                if pd.isna(value):
                    value = row.get('value')

                # Track multiple values from same source
                if source not in values_by_source:
                    values_by_source[source] = []

                if value not in values_by_source[source]:
                    values_by_source[source].append(value)

            # Determine status
            all_values = []
            for source_values in values_by_source.values():
                all_values.extend(source_values)

            # Remove None/NaN values for comparison
            non_null_values = [v for v in all_values if pd.notna(v)]
            unique_values = list(set(non_null_values))

            if len(unique_values) == 0:
                status = "Missing"
            elif len(unique_values) == 1:
                status = "Match"
            else:
                # Check if conflict is within-source or cross-source
                within_source_conflict = any(len(vals) > 1 for vals in values_by_source.values())
                cross_source_conflict = len(unique_values) > 1 and len(values_by_source) > 1

                if within_source_conflict and cross_source_conflict:
                    status = "Conflict (Within + Cross)"
                elif within_source_conflict:
                    status = "Conflict (Within Source)"
                elif cross_source_conflict:
                    status = "Conflict (Cross Source)"
                else:
                    status = "Single Source"

            # Build row
            row_data = {
                'Participant ID': participant_id,
                'Element': element,
                'Status': status,
                'Unique Values': len(unique_values)
            }

            # Add source columns
            for source, values in sorted(values_by_source.items()):
                # Join multiple values with " | "
                value_str = " | ".join([str(v) if pd.notna(v) else "-" for v in values])
                row_data[f"{source}"] = value_str if value_str else "-"

            # Add action column
            if status.startswith("Conflict"):
                row_data['Action'] = "REVIEW"
            elif status == "Missing":
                row_data['Action'] = "MISSING"
            else:
                row_data['Action'] = "OK"

            comparison_rows.append(row_data)

        comparison_df = pd.DataFrame(comparison_rows)

        # Sort by action priority (REVIEW first, then MISSING, then OK)
        action_order = {'REVIEW': 0, 'MISSING': 1, 'OK': 2}
        comparison_df['_sort'] = comparison_df['Action'].map(action_order)
        comparison_df = comparison_df.sort_values(['_sort', 'Participant ID', 'Element'])
        comparison_df = comparison_df.drop('_sort', axis=1)

        logger.info(f"Value Comparison tab ready with {len(comparison_df)} comparisons")
        return comparison_df

    # ====================
    # TAB 3: REVIEW NEEDED
    # ====================

    def create_review_needed_tab(self, df: pd.DataFrame, comparison_df: pd.DataFrame) -> pd.DataFrame:
        """
        Create review needed tab combining low confidence and conflicts.

        Args:
            df: DataFrame with validated extractions
            comparison_df: DataFrame from value comparison

        Returns:
            Review DataFrame
        """
        logger.info("Creating Review Needed tab")

        review_rows = []

        # Section 1: Low Confidence Extractions
        if 'confidence' in df.columns:
            low_conf = df[df['confidence'] == 'LOW'].copy()

            for _, row in low_conf.iterrows():
                review_rows.append({
                    'Participant ID': row.get('participant_id', 'N/A'),
                    'Source': row['source'],
                    'Element': row['element'],
                    'Value': row.get('cleaned_value') if pd.notna(row.get('cleaned_value')) else row.get('value'),
                    'Issue Type': 'Low Confidence',
                    'Details': row.get('flag_reasons', ''),
                    'Priority': 'Medium',
                    'Document Link': self.create_sharepoint_link(row['source'], row['filename'])
                })

        # Section 2: Cross-Source and Within-Source Conflicts
        if not comparison_df.empty:
            conflicts = comparison_df[comparison_df['Status'].str.contains('Conflict', na=False)].copy()

            for _, row in conflicts.iterrows():
                # Get all source values
                source_cols = [col for col in row.index if col not in ['Participant ID', 'Element', 'Status', 'Unique Values', 'Action']]
                values = [f"{col}: {row[col]}" for col in source_cols if pd.notna(row[col]) and row[col] != '-']

                review_rows.append({
                    'Participant ID': row['Participant ID'],
                    'Source': 'Multiple',
                    'Element': row['Element'],
                    'Value': ' | '.join(values),
                    'Issue Type': row['Status'],
                    'Details': f"{row['Unique Values']} different values found",
                    'Priority': 'High',
                    'Document Link': ''
                })

        # Section 3: Missing Critical Elements
        # (Will need critical_elements list passed in - for now skip)

        if not review_rows:
            logger.info("No items requiring review")
            return pd.DataFrame(columns=['Participant ID', 'Source', 'Element', 'Value', 'Issue Type', 'Details', 'Priority', 'Document Link'])

        review_df = pd.DataFrame(review_rows)

        # Sort by priority
        priority_order = {'High': 0, 'Medium': 1, 'Low': 2}
        review_df['_sort'] = review_df['Priority'].map(priority_order)
        review_df = review_df.sort_values(['_sort', 'Participant ID'])
        review_df = review_df.drop('_sort', axis=1)

        logger.info(f"Review Needed tab ready with {len(review_df)} items")
        return review_df

    # ====================
    # TAB 4: SOURCE STATISTICS
    # ====================

    def create_source_statistics_tab(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Create source-level statistics showing extraction performance.

        Args:
            df: DataFrame with validated extractions

        Returns:
            Source statistics DataFrame
        """
        logger.info("Creating Source Statistics tab")

        if df.empty:
            return pd.DataFrame()

        stats_rows = []

        for source, group in df.groupby('source'):
            total_extractions = len(group)

            # Confidence breakdown
            if 'confidence' in group.columns:
                high_conf = len(group[group['confidence'] == 'HIGH'])
                low_conf = len(group[group['confidence'] == 'LOW'])
                high_pct = (high_conf / total_extractions * 100) if total_extractions > 0 else 0
                low_pct = (low_conf / total_extractions * 100) if total_extractions > 0 else 0
            else:
                high_conf = total_extractions
                low_conf = 0
                high_pct = 100
                low_pct = 0

            # Missing data
            missing = len(group[group['cleaned_value'].isna() & group['value'].isna()]) if 'cleaned_value' in group.columns else 0
            missing_pct = (missing / total_extractions * 100) if total_extractions > 0 else 0

            # Most problematic elements (most low confidence)
            if 'confidence' in group.columns:
                problem_elements = group[group['confidence'] == 'LOW']['element'].value_counts().head(3)
                problem_str = ", ".join([f"{elem} ({count})" for elem, count in problem_elements.items()])
            else:
                problem_str = "None"

            # Needs review flag
            needs_review = "YES" if high_pct < 70 else "NO"

            stats_rows.append({
                'Source': source,
                'Total Extractions': total_extractions,
                'High Confidence': high_conf,
                'High Conf %': round(high_pct, 1),
                'Low Confidence': low_conf,
                'Low Conf %': round(low_pct, 1),
                'Missing Data': missing,
                'Missing %': round(missing_pct, 1),
                'Top Problem Elements': problem_str,
                'Needs Review': needs_review
            })

        stats_df = pd.DataFrame(stats_rows)

        # Sort by high confidence percentage (ascending to show worst first)
        stats_df = stats_df.sort_values('High Conf %', ascending=True)

        logger.info(f"Source Statistics tab ready with {len(stats_df)} sources")
        return stats_df

    # ====================
    # TAB 5: PARTICIPANT STATISTICS
    # ====================

    def create_participant_statistics_tab(self, df: pd.DataFrame, comparison_df: pd.DataFrame) -> pd.DataFrame:
        """
        Create participant-level statistics showing completeness.

        Args:
            df: DataFrame with validated extractions
            comparison_df: DataFrame from value comparison

        Returns:
            Participant statistics DataFrame
        """
        logger.info("Creating Participant Statistics tab")

        if df.empty or 'participant_id' not in df.columns:
            return pd.DataFrame()

        df_with_id = df[df['participant_id'].notna()].copy()

        if df_with_id.empty:
            return pd.DataFrame()

        stats_rows = []

        for participant_id, group in df_with_id.groupby('participant_id'):
            total_elements = len(group)

            # High confidence percentage
            if 'confidence' in group.columns:
                high_conf = len(group[group['confidence'] == 'HIGH'])
                high_pct = (high_conf / total_elements * 100) if total_elements > 0 else 0
            else:
                high_conf = total_elements
                high_pct = 100

            # Completeness (non-null values)
            if 'cleaned_value' in group.columns:
                completed = len(group[group['cleaned_value'].notna()])
            else:
                completed = len(group[group['value'].notna()])
            completeness_pct = (completed / total_elements * 100) if total_elements > 0 else 0

            # Cross-source conflicts
            if not comparison_df.empty:
                participant_conflicts = comparison_df[
                    (comparison_df['Participant ID'] == participant_id) &
                    (comparison_df['Status'].str.contains('Conflict', na=False))
                ]
                conflicts = len(participant_conflicts)
            else:
                conflicts = 0

            # Missing critical elements (placeholder - need critical elements list)
            missing_critical = 0

            # Review priority
            if conflicts > 2 or completeness_pct < 50:
                priority = "High"
            elif conflicts > 0 or completeness_pct < 80:
                priority = "Medium"
            else:
                priority = "Low"

            # Needs review flag
            needs_review = "YES" if priority in ["High", "Medium"] else "NO"

            stats_rows.append({
                'Participant ID': participant_id,
                'Total Elements': total_elements,
                'Completed': completed,
                'Completeness %': round(completeness_pct, 1),
                'High Confidence %': round(high_pct, 1),
                'Conflicts': conflicts,
                'Missing Critical': missing_critical,
                'Review Priority': priority,
                'Needs Review': needs_review
            })

        stats_df = pd.DataFrame(stats_rows)

        # Sort by priority (High first), then completeness (ascending)
        priority_order = {'High': 0, 'Medium': 1, 'Low': 2}
        stats_df['_sort'] = stats_df['Review Priority'].map(priority_order)
        stats_df = stats_df.sort_values(['_sort', 'Completeness %'])
        stats_df = stats_df.drop('_sort', axis=1)

        logger.info(f"Participant Statistics tab ready with {len(stats_df)} participants")
        return stats_df

    # ====================
    # EXCEL FORMATTING
    # ====================

    def apply_excel_formatting(self, filepath: str):
        """
        Apply comprehensive Excel formatting to all tabs.

        Args:
            filepath: Path to Excel file
        """
        try:
            logger.info(f"Applying formatting to {filepath}")
            wb = load_workbook(filepath)

            # Define styles
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")

            # Confidence colors
            high_conf_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Light green
            low_conf_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Light red

            # Status colors
            match_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Green
            conflict_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")  # Yellow
            missing_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Red

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
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

                # Get column indices for specific formatting
                header_row = [cell.value for cell in ws[1]]

                conf_col = header_row.index('Confidence') + 1 if 'Confidence' in header_row else None
                status_col = header_row.index('Status') + 1 if 'Status' in header_row else None
                action_col = header_row.index('Action') + 1 if 'Action' in header_row else None
                priority_col = header_row.index('Priority') + 1 if 'Priority' in header_row else None
                needs_review_col = header_row.index('Needs Review') + 1 if 'Needs Review' in header_row else None

                # Format data rows
                for row in range(2, ws.max_row + 1):
                    for col in range(1, ws.max_column + 1):
                        cell = ws.cell(row, col)
                        cell.border = border

                        # Confidence coloring (Tab 1)
                        if conf_col and col == conf_col:
                            if cell.value == 'HIGH':
                                cell.fill = high_conf_fill
                            elif cell.value == 'LOW':
                                cell.fill = low_conf_fill

                        # Status coloring (Tab 2)
                        if status_col and col == status_col:
                            if cell.value == 'Match':
                                cell.fill = match_fill
                            elif 'Conflict' in str(cell.value):
                                cell.fill = conflict_fill
                            elif cell.value == 'Missing':
                                cell.fill = missing_fill

                        # Action coloring (Tab 2 & 3)
                        if action_col and col == action_col:
                            if cell.value == 'REVIEW':
                                cell.fill = conflict_fill
                                cell.font = Font(bold=True)
                            elif cell.value == 'MISSING':
                                cell.fill = missing_fill

                        # Priority coloring (Tab 3 & 5)
                        if priority_col and col == priority_col:
                            if cell.value == 'High':
                                cell.fill = missing_fill
                                cell.font = Font(bold=True, color="9C0006")
                            elif cell.value == 'Medium':
                                cell.fill = conflict_fill

                        # Needs Review coloring (Tab 4 & 5)
                        if needs_review_col and col == needs_review_col:
                            if cell.value == 'YES':
                                cell.fill = missing_fill
                                cell.font = Font(bold=True)

                        # Format links
                        if isinstance(cell.value, str):
                            if cell.value.startswith('http://') or cell.value.startswith('https://'):
                                cell.font = link_font
                                cell.hyperlink = cell.value

                # Auto-adjust column widths
                for column in ws.columns:
                    max_length = 0
                    column_letter = column[0].column_letter

                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass

                    adjusted_width = min(max_length + 2, 50)
                    ws.column_dimensions[column_letter].width = adjusted_width

                # Freeze top row
                ws.freeze_panes = 'A2'

                # Enable autofilter
                if ws.max_row > 1:
                    ws.auto_filter.ref = ws.dimensions

            wb.save(filepath)
            logger.info("Formatting applied successfully")

        except Exception as e:
            logger.error(f"Error applying formatting: {e}", exc_info=True)

    # ====================
    # INSTRUCTIONS TAB
    # ====================

    def _create_instructions(self, plan_name: str) -> pd.DataFrame:
        """Create instructions sheet for the report."""
        instructions = [
            ['Interactive Extraction Report', f'Plan: {plan_name}'],
            ['', ''],
            ['TAB GUIDE:', ''],
            ['1. All Extracted Data', 'Complete dataset with all extractions. Filter by confidence, source, participant, or element.'],
            ['', 'GREEN = High confidence | RED = Low confidence'],
            ['', ''],
            ['2. Value Comparison', 'Side-by-side comparison of values across sources and within sources.'],
            ['', 'Identifies conflicts that need review. Use filters to find specific participants or elements.'],
            ['', 'GREEN = Match | YELLOW = Conflict | RED = Missing'],
            ['', ''],
            ['3. Review Needed', 'All items requiring attention: low confidence extractions and conflicts.'],
            ['', 'Prioritized by urgency (High/Medium/Low). Start here for data cleaning.'],
            ['', ''],
            ['4. Source Statistics', 'Performance metrics by document source.'],
            ['', 'Shows which sources had best/worst extraction rates. Identifies problem areas.'],
            ['', ''],
            ['5. Participant Statistics', 'Completeness metrics by participant.'],
            ['', 'Shows data completeness and conflicts per participant. Identifies which need review.'],
            ['', ''],
            ['USAGE TIPS:', ''],
            ['• Use Excel filters', 'Click dropdown arrows in headers to filter data'],
            ['• Sort columns', 'Click column headers to sort'],
            ['• Document links', 'Blue underlined links open source documents'],
            ['• Color coding', 'Green=good, Yellow=review, Red=problem'],
            ['• Review workflow', 'Start with Tab 3 (Review Needed), use Tab 2 (Comparison) for context'],
            ['', ''],
            ['SHAREPOINT LINKS:', ''],
            ['Base URL', self.sharepoint_base_url],
            ['Note', 'Update base URL in code if different'],
        ]

        return pd.DataFrame(instructions, columns=['Topic', 'Description'])

    # ====================
    # MAIN REPORT GENERATION
    # ====================

    def generate_interactive_report(
        self,
        validated_df: pd.DataFrame,
        output_path: str,
        plan_name: str
    ) -> str:
        """
        Generate comprehensive interactive Excel report with new structure.

        Args:
            validated_df: DataFrame with validated extractions
            output_path: Path to save report
            plan_name: Name of the plan

        Returns:
            Path to generated report
        """
        logger.info(f"Generating redesigned interactive report for {plan_name}")

        # Generate all tabs
        all_data = self.create_all_data_tab(validated_df)
        comparison = self.create_value_comparison_tab(validated_df)
        review_needed = self.create_review_needed_tab(validated_df, comparison)
        source_stats = self.create_source_statistics_tab(validated_df)
        participant_stats = self.create_participant_statistics_tab(validated_df, comparison)
        instructions = self._create_instructions(plan_name)

        # Create Excel writer
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write all tabs
            if not all_data.empty:
                all_data.to_excel(writer, sheet_name='All Extracted Data', index=False)
                logger.info(f"Added All Extracted Data: {len(all_data)} rows")

            if not comparison.empty:
                comparison.to_excel(writer, sheet_name='Value Comparison', index=False)
                logger.info(f"Added Value Comparison: {len(comparison)} comparisons")

            if not review_needed.empty:
                review_needed.to_excel(writer, sheet_name='Review Needed', index=False)
                logger.info(f"Added Review Needed: {len(review_needed)} items")

            if not source_stats.empty:
                source_stats.to_excel(writer, sheet_name='Source Statistics', index=False)
                logger.info(f"Added Source Statistics: {len(source_stats)} sources")

            if not participant_stats.empty:
                participant_stats.to_excel(writer, sheet_name='Participant Statistics', index=False)
                logger.info(f"Added Participant Statistics: {len(participant_stats)} participants")

            # Instructions always added
            instructions.to_excel(writer, sheet_name='Instructions', index=False)
            logger.info("Added Instructions tab")

        # Apply formatting
        self.apply_excel_formatting(output_path)

        logger.info(f"Redesigned interactive report generated: {output_path}")
        return output_path


def main():
    """Example usage of ReportGenerator."""
    pass


if __name__ == "__main__":
    main()
