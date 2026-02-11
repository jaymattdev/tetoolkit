"""
Interactive Excel Report Generator - Refactored
Creates user-friendly Excel reports optimized for data comparison and review.
"""

import pandas as pd
import logging
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

# Constants
QUALITY_ORDER = {'Conflict': 0, 'Review': 1, 'Missing': 2, 'Good': 3}

QUALITY_FILLS = {
    'good': PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
    'yellow': PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
    'red': PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
    'light_red': PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
}


class ReportGenerator:
    """Generate comprehensive interactive Excel reports for data review."""

    def __init__(self, sharepoint_base_url: str = "https://yourcompany.sharepoint.com/sites/yoursite/"):
        self.sharepoint_base_url = sharepoint_base_url
        logger.info("Initialized ReportGenerator")

    # ====================
    # HELPER METHODS
    # ====================

    def _create_sharepoint_link(self, source: str, filename: str) -> str:
        """Create SharePoint link for a document."""
        # Ensure filename has .pdf extension
        if pd.notna(filename):
            if not filename.endswith('.pdf'):
                # Strip existing extension if any and add .pdf
                filename = filename.rsplit('.', 1)[0] + '.pdf'
        return f"{self.sharepoint_base_url}{source}/{filename}"

    def _validate_participant_df(self, df: pd.DataFrame, context: str) -> pd.DataFrame:
        """Validate and filter DataFrame for participant-based operations."""
        if df.empty or 'participant_id' not in df.columns:
            logger.warning(f"No participant IDs for {context}")
            return pd.DataFrame()

        df_filtered = df[df['participant_id'].notna()].copy()
        if df_filtered.empty:
            logger.warning(f"No extractions with participant IDs for {context}")

        return df_filtered

    def _get_element_quality(self, element_data: pd.DataFrame) -> tuple[str, str, list]:
        """
        Determine quality and best value for an element.

        Returns:
            tuple: (quality, best_value, unique_values_list)
        """
        values = element_data['cleaned_value'].dropna()
        if values.empty:
            values = element_data['value'].dropna()

        unique_values = list(values.unique())

        if len(unique_values) == 0:
            return 'Missing', '', []
        elif len(unique_values) == 1:
            has_low_conf = (element_data['confidence'] == 'LOW').any()
            has_flags = element_data['flags'].notna().any() and (element_data['flags'] != '').any()
            quality = 'Review' if (has_low_conf or has_flags) else 'Good'
            return quality, unique_values[0], unique_values
        else:
            return 'Conflict', ' | '.join([str(v) for v in unique_values]), unique_values

    def _sort_by_quality(self, df: pd.DataFrame, *other_sort_cols) -> pd.DataFrame:
        """Sort DataFrame by quality priority and other columns."""
        if df.empty or 'Quality' not in df.columns:
            return df

        df['_sort'] = df['Quality'].map(QUALITY_ORDER)
        sort_cols = ['_sort'] + list(other_sort_cols)
        df = df.sort_values(sort_cols).drop('_sort', axis=1)
        return df

    def _prepare_filename_and_link(self, df: pd.DataFrame) -> pd.DataFrame:
        """Strip file extensions and add SharePoint links."""
        df['_original_filename'] = df['filename']
        df['filename'] = df['filename'].apply(
            lambda x: x.rsplit('.', 1)[0] if pd.notna(x) and '.' in x else x
        )
        df['Document Link'] = df.apply(
            lambda row: self._create_sharepoint_link(row['source'], row['_original_filename']),
            axis=1
        )
        return df

    # ====================
    # TAB GENERATION
    # ====================

    def create_all_data_tab(self, df: pd.DataFrame) -> pd.DataFrame:
        """Create unified view of all extracted data."""
        if df.empty:
            return pd.DataFrame()

        logger.info(f"Creating All Extracted Data tab with {len(df)} rows")
        report_df = self._prepare_filename_and_link(df.copy())

        # Select columns dynamically
        base_cols = ['participant_id', 'source', 'filename', 'element', 'value', 'cleaned_value']
        optional_cols = ['extraction_order', 'extraction_position', 'flags', 'flag_reasons']
        columns = [c for c in base_cols + optional_cols if c in report_df.columns] + ['Document Link']

        report_df = report_df[columns]

        # Sort by participant_id and source for easier reading
        sort_cols = []
        if 'participant_id' in report_df.columns:
            sort_cols.append('participant_id')
        if 'source' in report_df.columns:
            sort_cols.append('source')
        if sort_cols:
            report_df = report_df.sort_values(sort_cols)

        return report_df

    def create_participant_summary(self, df: pd.DataFrame) -> pd.DataFrame:
        """Create compact participant summary for large datasets."""
        df_with_id = self._validate_participant_df(df, "participant summary")
        if df_with_id.empty:
            return pd.DataFrame()

        logger.info("Creating Participant Summary")
        all_elements = sorted(df_with_id['element'].unique())
        total_elements = len(all_elements)
        summary_rows = []

        for participant_id, participant_group in df_with_id.groupby('participant_id'):
            quality_counts = {'Good': 0, 'Review': 0, 'Conflict': 0, 'Missing': 0}
            issue_elements = {'Good': [], 'Conflict': [], 'Review': [], 'Missing': []}

            for element in all_elements:
                element_data = participant_group[participant_group['element'] == element]
                quality, _, _ = self._get_element_quality(element_data)
                quality_counts[quality] += 1
                issue_elements[quality].append(element)

            # Calculate metrics
            completeness_pct = (quality_counts['Good'] / total_elements * 100) if total_elements > 0 else 0
            non_missing = quality_counts['Good'] + quality_counts['Review'] + quality_counts['Conflict']
            quality_pct = (quality_counts['Good'] / non_missing * 100) if non_missing > 0 else 0

            # Format element lists
            complete_str = ', '.join(issue_elements['Good']) if issue_elements['Good'] else ''
            missing_str = ', '.join(issue_elements['Missing']) if issue_elements['Missing'] else ''
            conflict_str = ', '.join(issue_elements['Conflict']) if issue_elements['Conflict'] else ''

            summary_rows.append({
                'Participant ID': participant_id,
                'Completeness %': round(completeness_pct, 1),
                'Quality %': round(quality_pct, 1),
                'Good': quality_counts['Good'],
                'Review': quality_counts['Review'],
                'Conflicts': quality_counts['Conflict'],
                'Missing': quality_counts['Missing'],
                'Total': total_elements,
                'Sources': len(participant_group['source'].unique()),
                'Complete Elements': complete_str,
                'Missing Elements': missing_str,
                'Conflicting Elements': conflict_str
            })

        summary_df = pd.DataFrame(summary_rows)

        # Sort by conflicts (high to low), then completeness (low to high)
        summary_df = summary_df.sort_values(['Conflicts', 'Completeness %'], ascending=[False, True])

        logger.info(f"Participant Summary ready with {len(summary_df)} participants")
        return summary_df
    def create_source_statistics_tab(self, df: pd.DataFrame) -> pd.DataFrame:
        """Create source-level statistics."""
        if df.empty:
            return pd.DataFrame()

        logger.info("Creating Source Statistics tab")
        stats_rows = []

        for source, group in df.groupby('source'):
            # Count missing elements (no value extracted)
            missing = len(group[group['cleaned_value'].isna() & group['value'].isna()])

            # Count actual extractions (exclude missing - only count rows with values)
            actual_extractions = len(group) - missing

            # High/Low confidence only applies to actual extractions
            high_conf = len(group[group.get('confidence', 'HIGH') == 'HIGH']) if 'confidence' in group.columns else actual_extractions
            low_conf = len(group[group.get('confidence', 'LOW') == 'LOW']) if 'confidence' in group.columns else 0

            # Calculate percentages based on total elements (actual + missing)
            total_elements = len(group)
            high_pct = (high_conf / total_elements * 100) if total_elements > 0 else 0
            low_pct = (low_conf / total_elements * 100) if total_elements > 0 else 0
            missing_pct = (missing / total_elements * 100) if total_elements > 0 else 0

            problem_elements = "None"
            if 'confidence' in group.columns:
                problems = group[group['confidence'] == 'LOW']['element'].value_counts().head(3)
                problem_elements = ", ".join([f"{elem} ({count})" for elem, count in problems.items()])

            stats_rows.append({
                'Source': source,
                'Total Extractions': actual_extractions,
                'High Confidence': high_conf,
                'High Conf %': round(high_pct, 1),
                'Low Confidence': low_conf,
                'Low Conf %': round(low_pct, 1),
                'Missing Data': missing,
                'Missing %': round(missing_pct, 1),
                'Top Problem Elements': problem_elements
            })

        stats_df = pd.DataFrame(stats_rows).sort_values('High Conf %', ascending=True)
        logger.info(f"Source Statistics tab ready with {len(stats_df)} sources")
        return stats_df

    def create_participant_statistics_tab(self, df: pd.DataFrame) -> pd.DataFrame:
        """Create participant-level statistics."""
        df_with_id = self._validate_participant_df(df, "participant statistics")
        if df_with_id.empty:
            return pd.DataFrame()

        logger.info("Creating Participant Statistics tab")

        # Count total elements
        all_elements = df_with_id['element'].unique()
        total_elements = len(all_elements)
        stats_rows = []

        for participant_id, group in df_with_id.groupby('participant_id'):
            elements_extracted = group['element'].nunique()
            completeness = (elements_extracted / total_elements * 100) if total_elements > 0 else 0

            # Count conflicts by checking each element for multiple unique values
            conflicts = 0
            for element in all_elements:
                element_data = group[group['element'] == element]
                quality, _, _ = self._get_element_quality(element_data)
                if quality == 'Conflict':
                    conflicts += 1

            stats_rows.append({
                'Participant ID': participant_id,
                'Elements Extracted': elements_extracted,
                'Total Elements': total_elements,
                'Completeness %': round(completeness, 1),
                'Conflicts': conflicts
            })

        stats_df = pd.DataFrame(stats_rows).sort_values('Completeness %', ascending=False)
        logger.info(f"Participant Statistics tab ready with {len(stats_df)} participants")
        return stats_df

    # ====================
    # EXCEL FORMATTING
    # ====================

    def _apply_conditional_formatting_rule(self, ws, col_letter, value, fill, font=None):
        """Helper to apply conditional formatting rule."""
        try:
            rule = CellIsRule(operator='equal', formula=[f'"{value}"'], fill=fill)
            if font:
                rule.font = font
            ws.conditional_formatting.add(f'{col_letter}2:{col_letter}{ws.max_row}', rule)
        except Exception as e:
            logger.warning(f"Could not apply conditional formatting for {value} in {col_letter}: {e}")

    def _apply_participant_color_banding(self, ws, headers):
        """Apply alternating grey/white color banding by participant ID."""
        # Find participant_id column
        pid_col = None
        for header, idx in headers.items():
            if header == 'participant_id':
                pid_col = idx
                break

        if not pid_col:
            logger.warning("No participant_id column found for color banding")
            return

        # Define alternating fills
        grey_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        # Track participant changes and apply banding
        current_participant = None
        use_grey = False

        for row_idx in range(2, ws.max_row + 1):
            participant = ws.cell(row_idx, pid_col).value

            # Check if participant changed
            if participant != current_participant:
                current_participant = participant
                use_grey = not use_grey  # Toggle color

            # Apply fill to entire row
            fill = grey_fill if use_grey else white_fill
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row_idx, col_idx)
                # Only apply if cell doesn't already have special formatting
                if cell.fill.start_color.rgb in ('00000000', None, 'FFFFFF', 'F2F2F2'):
                    cell.fill = fill

        logger.info(f"Applied participant color banding to {ws.max_row - 1} rows")

    def apply_excel_formatting(self, filepath: str):
        """Apply comprehensive Excel formatting to all tabs."""
        try:
            logger.info(f"Applying formatting to {filepath}")
            wb = load_workbook(filepath)

            # Define styles
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            link_font = Font(color="0563C1", underline="single")

            for ws in wb.worksheets:
                if ws.max_row == 1:
                    continue

                # Format headers
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font

                # Build column map
                headers = {cell.value: idx for idx, cell in enumerate(ws[1], 1)}
                col_map = {}

                for header, idx in headers.items():
                    col_letter = ws.cell(1, idx).column_letter

                    # Map columns for conditional formatting
                    if header == 'Quality':
                        col_map['quality'] = col_letter

                # Apply conditional formatting
                self._apply_formatting_rules(ws, col_map)

                # Apply color banding by participant for All Extracted Data tab
                if ws.title == 'All Extracted Data':
                    self._apply_participant_color_banding(ws, headers)

                # Process hyperlinks
                self._process_hyperlinks(ws, link_font)

                # Auto-adjust widths
                self._auto_adjust_columns(ws)

                # Freeze and filter
                ws.freeze_panes = 'A2'
                if ws.max_row > 1:
                    ws.auto_filter.ref = ws.dimensions

            wb.save(filepath)
            logger.info("Formatting applied successfully")

        except Exception as e:
            logger.error(f"Error applying formatting: {e}", exc_info=True)

    def _apply_formatting_rules(self, ws, col_map):
        """Apply conditional formatting rules for Quality column."""
        fills = QUALITY_FILLS

        if 'quality' in col_map:
            self._apply_conditional_formatting_rule(ws, col_map['quality'], 'Good', fills['good'])
            self._apply_conditional_formatting_rule(ws, col_map['quality'], 'Review', fills['yellow'])
            self._apply_conditional_formatting_rule(ws, col_map['quality'], 'Conflict', fills['red'])
            self._apply_conditional_formatting_rule(ws, col_map['quality'], 'Missing', fills['light_red'])

    def _process_hyperlinks(self, ws, link_font):
        """Process and convert hyperlinks, creating hidden columns for additional links."""
        # Process regular link columns
        header_row = [cell.value for cell in ws[1]]
        link_columns = [idx for idx, h in enumerate(header_row, 1) if 'Link' in str(h) or 'Documents' in str(h)]

        for col_idx in link_columns:
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row_idx, col_idx)
                cell_value = str(cell.value or '')

                if cell_value and cell_value.startswith('http'):
                    if ' | ' in cell_value:
                        urls = cell_value.split(' | ')
                        cell.hyperlink = urls[0]
                        cell.value = f"Open File ({len(urls)} docs)"
                    else:
                        cell.hyperlink = cell_value
                        cell.value = "Open File"
                    cell.font = link_font

        # Track columns that need helper columns
        columns_needing_helpers = {}  # {col_idx: max_files_needed}

        # First pass: identify columns with multiple files
        for col_idx in range(1, ws.max_column + 1):
            max_files = 0
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row_idx, col_idx)
                cell_value = str(cell.value or '')

                if '|||' in cell_value and ' | ' in cell_value:
                    file_count = cell_value.count(' | ') + 1
                    max_files = max(max_files, file_count)

            if max_files > 1:
                columns_needing_helpers[col_idx] = max_files

        # Insert helper columns (right to left to maintain indices)
        for col_idx in sorted(columns_needing_helpers.keys(), reverse=True):
            max_files = columns_needing_helpers[col_idx]
            header_name = ws.cell(1, col_idx).value

            # Insert helper columns after this column
            for helper_idx in range(max_files - 1):
                ws.insert_cols(col_idx + 1)
                helper_col = col_idx + 1 + helper_idx
                ws.cell(1, helper_col).value = f"{header_name}_Link{helper_idx + 2}"
                ws.column_dimensions[get_column_letter(helper_col)].hidden = True

        # Second pass: process hyperlinks and populate helper columns
        for col_idx in range(1, ws.max_column + 1):
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row_idx, col_idx)
                cell_value = str(cell.value or '')

                if '|||' in cell_value:
                    # Check if this has multiple pipe-separated entries
                    if ' | ' in cell_value:
                        # Multiple files: "file1.pdf|||link1 | file2.pdf|||link2"
                        parts = cell_value.split(' | ')
                        display_parts = []
                        links = []

                        for part in parts:
                            if '|||' in part:
                                value, link = part.split('|||', 1)
                                display_parts.append(value)
                                if link and link.startswith('http'):
                                    links.append(link)

                        # Main cell: show all filenames, link to first
                        cell.value = ' | '.join(display_parts)
                        if links:
                            cell.hyperlink = links[0]
                            cell.font = link_font

                        # Populate helper columns with additional links
                        for helper_idx, link in enumerate(links[1:], start=1):
                            helper_cell = ws.cell(row_idx, col_idx + helper_idx)
                            helper_cell.value = display_parts[helper_idx] if helper_idx < len(display_parts) else ''
                            helper_cell.hyperlink = link
                            helper_cell.font = link_font
                    else:
                        # Single file: "file.pdf|||link"
                        value, link = cell_value.split('|||', 1)
                        if link and link.startswith('http'):
                            cell.value = value
                            cell.hyperlink = link
                            cell.font = link_font

    def _auto_adjust_columns(self, ws, sample_size=100):
        """Auto-adjust column widths based on content."""
        for col_idx in range(1, ws.max_column + 1):
            col_letter = ws.cell(1, col_idx).column_letter
            max_length = len(str(ws.cell(1, col_idx).value or ''))

            for row_idx in range(2, min(sample_size + 2, ws.max_row + 1)):
                try:
                    cell_value = str(ws.cell(row_idx, col_idx).value or '')
                    max_length = max(max_length, len(cell_value))
                except:
                    pass

            ws.column_dimensions[col_letter].width = min(max_length + 2, 50)

    # ====================
    # MAIN REPORT GENERATION
    # ====================

    def generate_interactive_report(self, validated_df: pd.DataFrame, output_path: str, plan_name: str) -> str:
        """Generate comprehensive interactive Excel report."""
        logger.info(f"Generating interactive report for {plan_name}")

        # Generate tabs
        participant_summary = self.create_participant_summary(validated_df)
        all_data = self.create_all_data_tab(validated_df)
        source_stats = self.create_source_statistics_tab(validated_df)
        participant_stats = self.create_participant_statistics_tab(validated_df)

        # Write to Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            tabs = [
                (participant_summary, 'Participant Summary'),
                (all_data, 'All Extracted Data'),
                (source_stats, 'Source Statistics'),
                (participant_stats, 'Participant Statistics')
            ]

            for df, sheet_name in tabs:
                if not df.empty:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    logger.info(f"Added {sheet_name}: {len(df)} rows")

        # Apply formatting (including color banding)
        self.apply_excel_formatting(output_path)

        logger.info(f"Interactive report generated: {output_path}")
        return output_path


def main():
    """Example usage."""
    pass


if __name__ == "__main__":
    main()
