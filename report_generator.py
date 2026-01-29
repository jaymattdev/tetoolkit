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

logger = logging.getLogger(__name__)

# Constants
QUALITY_ORDER = {'Conflict': 0, 'Review': 1, 'Missing': 2, 'Good': 3}
PRIORITY_ORDER = {'High': 0, 'Medium': 1, 'Low': 2, 'None': 3}

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
        optional_cols = ['confidence', 'extraction_order', 'extraction_position', 'flags', 'flag_reasons']
        columns = [c for c in base_cols + optional_cols if c in report_df.columns] + ['Document Link']

        return report_df[columns]

    def create_value_comparison_tab(self, df: pd.DataFrame) -> pd.DataFrame:
        """Create element-by-element comparison across sources."""
        if df.empty:
            return pd.DataFrame()

        logger.info("Creating Value Comparison tab")
        df_with_id = self._validate_participant_df(df, "value comparison")
        if df_with_id.empty:
            return pd.DataFrame()

        comparison_rows = []
        for (participant_id, element), group in df_with_id.groupby(['participant_id', 'element']):
            values_by_source = {}
            for _, row in group.iterrows():
                source = row['source']
                value = row.get('cleaned_value') if pd.notna(row.get('cleaned_value')) else row.get('value')
                if source not in values_by_source:
                    values_by_source[source] = []
                if pd.notna(value):
                    values_by_source[source].append(str(value))

            # Determine status
            all_values = [v for vals in values_by_source.values() for v in vals]
            unique_values = list(set(all_values))

            if len(unique_values) == 0:
                status = 'Missing'
            elif len(unique_values) == 1:
                status = 'Match'
            else:
                sources_with_conflicts = sum(1 for vals in values_by_source.values() if len(set(vals)) > 1)
                if sources_with_conflicts > 0:
                    status = 'Conflict (Within + Cross)'
                else:
                    status = 'Conflict (Cross Source)'

            # Build compact values string
            sources_with_values = [
                f"{source}: {' | '.join(vals)}"
                for source, vals in sorted(values_by_source.items())
                if vals
            ]
            compact_values = " || ".join(sources_with_values) if sources_with_values else "No values"

            # Get document links
            doc_links = [
                self._create_sharepoint_link(row['source'], row.get('_original_filename', row['filename']))
                for _, row in group.iterrows()
            ]
            doc_links_str = ' | '.join(set(doc_links[:5]))

            comparison_rows.append({
                'Participant ID': participant_id,
                'Element': element,
                'Status': status,
                'Values': compact_values,
                'Sources Checked': len(values_by_source),
                'Unique Values': len(unique_values),
                'Documents': doc_links_str,
                'Action': 'REVIEW' if status.startswith("Conflict") else ('MISSING' if status == 'Missing' else 'OK')
            })

        comparison_df = pd.DataFrame(comparison_rows)

        # Sort by action priority
        action_order = {'REVIEW': 0, 'MISSING': 1, 'OK': 2}
        comparison_df['_sort'] = comparison_df['Action'].map(action_order)
        comparison_df = comparison_df.sort_values(['_sort', 'Participant ID', 'Element']).drop('_sort', axis=1)

        logger.info(f"Value Comparison tab ready with {len(comparison_df)} comparisons")
        return comparison_df

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
            issue_elements = {'Conflict': [], 'Review': [], 'Missing': []}

            for element in all_elements:
                element_data = participant_group[participant_group['element'] == element]
                quality, _, _ = self._get_element_quality(element_data)
                quality_counts[quality] += 1
                if quality != 'Good':
                    issue_elements[quality].append(element)

            # Calculate metrics
            completeness_pct = (quality_counts['Good'] / total_elements * 100) if total_elements > 0 else 0
            non_missing = quality_counts['Good'] + quality_counts['Review'] + quality_counts['Conflict']
            quality_pct = (quality_counts['Good'] / non_missing * 100) if non_missing > 0 else 0

            # Determine status and action
            if quality_counts['Conflict'] > 0:
                status, priority = 'Conflicts', 'High'
                action = f"Resolve {quality_counts['Conflict']} conflicts: {', '.join(issue_elements['Conflict'][:3])}"
                action += "..." if len(issue_elements['Conflict']) > 3 else ""
            elif quality_counts['Review'] > 0:
                status = 'Review'
                priority = 'Medium' if quality_counts['Review'] > 5 else 'Low'
                action = f"Review {quality_counts['Review']} items: {', '.join(issue_elements['Review'][:3])}"
                action += "..." if len(issue_elements['Review']) > 3 else ""
            elif quality_counts['Missing'] > 0:
                status = 'Incomplete'
                priority = 'Medium' if quality_counts['Missing'] > 10 else 'Low'
                action = f"Fill {quality_counts['Missing']} missing: {', '.join(issue_elements['Missing'][:3])}"
                action += "..." if len(issue_elements['Missing']) > 3 else ""
            else:
                status, priority, action = 'Complete', 'None', 'Database Ready ✓'

            summary_rows.append({
                'Participant ID': participant_id,
                'View Details': '#Master_View!A1',  # Updated with VBA later
                'Status': status,
                'Priority': priority,
                'Completeness %': round(completeness_pct, 1),
                'Quality %': round(quality_pct, 1),
                'Good': quality_counts['Good'],
                'Review': quality_counts['Review'],
                'Conflicts': quality_counts['Conflict'],
                'Missing': quality_counts['Missing'],
                'Total': total_elements,
                'Sources': len(participant_group['source'].unique()),
                'Action Required': action
            })

        summary_df = pd.DataFrame(summary_rows)

        # Sort by priority
        summary_df['_sort'] = summary_df['Priority'].map(PRIORITY_ORDER)
        summary_df = summary_df.sort_values(['_sort', 'Completeness %']).drop('_sort', axis=1)

        logger.info(f"Participant Summary ready with {len(summary_df)} participants")
        return summary_df

    def create_participant_master_view(self, df: pd.DataFrame) -> pd.DataFrame:
        """Create participant master view with one row per element per participant."""
        df_with_id = self._validate_participant_df(df, "participant master view")
        if df_with_id.empty:
            return pd.DataFrame()

        logger.info("Creating Participant Master View")

        all_elements = sorted(df_with_id['element'].unique())
        all_participants = sorted(df_with_id['participant_id'].unique())
        all_sources = sorted(df_with_id['source'].unique())

        master_rows = []

        for participant_id in all_participants:
            participant_data = df_with_id[df_with_id['participant_id'] == participant_id]

            for element in all_elements:
                element_data = participant_data[participant_data['element'] == element]

                # Build source-specific values and links
                source_data = {}
                sources_checked = set()

                for _, row in element_data.iterrows():
                    source = row['source']
                    sources_checked.add(source)

                    if source not in source_data:
                        value = row.get('cleaned_value') if pd.notna(row.get('cleaned_value')) else row.get('value')
                        if pd.notna(value):
                            filename = row.get('_original_filename', row['filename'])
                            link = self._create_sharepoint_link(source, filename)
                            source_data[source] = f"{value}|||{link}"

                # Get quality and best value
                quality, best_value, unique_values = self._get_element_quality(element_data)

                # Build row
                row_data = {
                    'Participant ID': participant_id,
                    'Element': element,
                    'Best Value': best_value,
                    'Quality': quality,
                    'Unique Values': len(unique_values)
                }

                # Add source columns
                for source in all_sources:
                    if source in source_data:
                        row_data[source] = source_data[source]
                    elif source in sources_checked:
                        row_data[source] = '-'
                    else:
                        row_data[source] = ''

                master_rows.append(row_data)

        master_df = pd.DataFrame(master_rows)
        master_df = self._sort_by_quality(master_df, 'Participant ID', 'Element')

        logger.info(f"Participant Master View ready with {len(master_df)} rows")
        return master_df

    def create_review_needed_tab(self, df: pd.DataFrame, comparison_df: pd.DataFrame) -> pd.DataFrame:
        """Create review needed tab combining low confidence and conflicts."""
        logger.info("Creating Review Needed tab")
        review_rows = []

        # Low Confidence Extractions
        if 'confidence' in df.columns:
            low_conf = df[df['confidence'] == 'LOW']
            for _, row in low_conf.iterrows():
                review_rows.append({
                    'Participant ID': row.get('participant_id', 'N/A'),
                    'Source': row['source'],
                    'Element': row['element'],
                    'Value': row.get('cleaned_value') if pd.notna(row.get('cleaned_value')) else row.get('value'),
                    'Issue Type': 'Low Confidence',
                    'Details': row.get('flag_reasons', ''),
                    'Priority': 'Medium',
                    'Document Link': self._create_sharepoint_link(row['source'], row['filename'])
                })

        # Conflicts
        if not comparison_df.empty:
            conflicts = comparison_df[comparison_df['Status'].str.contains('Conflict', na=False)]
            for _, row in conflicts.iterrows():
                review_rows.append({
                    'Participant ID': row['Participant ID'],
                    'Source': 'Multiple',
                    'Element': row['Element'],
                    'Value': row['Values'],
                    'Issue Type': row['Status'],
                    'Details': f"{row['Unique Values']} different values found",
                    'Priority': 'High',
                    'Document Link': ''
                })

        if not review_rows:
            return pd.DataFrame(columns=['Participant ID', 'Source', 'Element', 'Value', 'Issue Type', 'Details', 'Priority', 'Document Link'])

        review_df = pd.DataFrame(review_rows)
        priority_order_map = {'High': 0, 'Medium': 1, 'Low': 2}
        review_df['_sort'] = review_df['Priority'].map(priority_order_map)
        review_df = review_df.sort_values(['_sort', 'Participant ID']).drop('_sort', axis=1)

        logger.info(f"Review Needed tab ready with {len(review_df)} items")
        return review_df

    def create_source_statistics_tab(self, df: pd.DataFrame) -> pd.DataFrame:
        """Create source-level statistics."""
        if df.empty:
            return pd.DataFrame()

        logger.info("Creating Source Statistics tab")
        stats_rows = []

        for source, group in df.groupby('source'):
            total = len(group)
            high_conf = len(group[group.get('confidence', 'HIGH') == 'HIGH'])
            low_conf = len(group[group.get('confidence', 'LOW') == 'LOW'])
            missing = len(group[group['cleaned_value'].isna() & group['value'].isna()])

            high_pct = (high_conf / total * 100) if total > 0 else 0
            low_pct = (low_conf / total * 100) if total > 0 else 0
            missing_pct = (missing / total * 100) if total > 0 else 0

            problem_elements = "None"
            if 'confidence' in group.columns:
                problems = group[group['confidence'] == 'LOW']['element'].value_counts().head(3)
                problem_elements = ", ".join([f"{elem} ({count})" for elem, count in problems.items()])

            stats_rows.append({
                'Source': source,
                'Total Extractions': total,
                'High Confidence': high_conf,
                'High Conf %': round(high_pct, 1),
                'Low Confidence': low_conf,
                'Low Conf %': round(low_pct, 1),
                'Missing Data': missing,
                'Missing %': round(missing_pct, 1),
                'Top Problem Elements': problem_elements,
                'Needs Review': "YES" if high_pct < 70 else "NO"
            })

        stats_df = pd.DataFrame(stats_rows).sort_values('High Conf %', ascending=True)
        logger.info(f"Source Statistics tab ready with {len(stats_df)} sources")
        return stats_df

    def create_participant_statistics_tab(self, df: pd.DataFrame, comparison_df: pd.DataFrame) -> pd.DataFrame:
        """Create participant-level statistics."""
        df_with_id = self._validate_participant_df(df, "participant statistics")
        if df_with_id.empty:
            return pd.DataFrame()

        logger.info("Creating Participant Statistics tab")

        # Count total elements
        total_elements = len(df_with_id['element'].unique())
        stats_rows = []

        for participant_id, group in df_with_id.groupby('participant_id'):
            elements_extracted = group['element'].nunique()
            completeness = (elements_extracted / total_elements * 100) if total_elements > 0 else 0

            # Count conflicts from comparison
            conflicts = 0
            if not comparison_df.empty:
                participant_conflicts = comparison_df[
                    (comparison_df['Participant ID'] == participant_id) &
                    (comparison_df['Status'].str.contains('Conflict', na=False))
                ]
                conflicts = len(participant_conflicts)

            stats_rows.append({
                'Participant ID': participant_id,
                'Elements Extracted': elements_extracted,
                'Total Elements': total_elements,
                'Completeness %': round(completeness, 1),
                'Conflicts': conflicts,
                'Needs Review': "YES" if conflicts > 0 or completeness < 80 else "NO"
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
                    header_str = str(header) if header else ''

                    # Map columns
                    if header == 'Confidence':
                        col_map['confidence'] = col_letter
                    elif header == 'Status':
                        col_map['status'] = col_letter
                    elif header == 'Quality':
                        col_map['quality'] = col_letter
                    elif header in ('Action', 'Action Required'):
                        col_map['action'] = col_letter
                    elif header == 'Priority':
                        col_map['priority'] = col_letter
                    elif header == 'Needs Review':
                        col_map['needs_review'] = col_letter

                # Apply conditional formatting
                self._apply_formatting_rules(ws, col_map)

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
        """Apply all conditional formatting rules."""
        fills = QUALITY_FILLS

        if 'confidence' in col_map:
            self._apply_conditional_formatting_rule(ws, col_map['confidence'], 'HIGH', fills['good'])
            self._apply_conditional_formatting_rule(ws, col_map['confidence'], 'LOW', fills['red'])

        if 'status' in col_map:
            for val in ['Match', 'Complete']:
                self._apply_conditional_formatting_rule(ws, col_map['status'], val, fills['good'])
            for val in ['Missing']:
                self._apply_conditional_formatting_rule(ws, col_map['status'], val, fills['red'])
            for val in ['Incomplete', 'Conflicts', 'Review', 'Conflict (Within Source)', 'Conflict (Cross Source)', 'Conflict (Within + Cross)']:
                self._apply_conditional_formatting_rule(ws, col_map['status'], val, fills['yellow'])

        if 'quality' in col_map:
            self._apply_conditional_formatting_rule(ws, col_map['quality'], 'Good', fills['good'])
            self._apply_conditional_formatting_rule(ws, col_map['quality'], 'Review', fills['yellow'])
            self._apply_conditional_formatting_rule(ws, col_map['quality'], 'Conflict', fills['red'])
            self._apply_conditional_formatting_rule(ws, col_map['quality'], 'Missing', fills['light_red'])

        if 'action' in col_map:
            self._apply_conditional_formatting_rule(ws, col_map['action'], 'REVIEW', fills['yellow'], Font(bold=True))
            self._apply_conditional_formatting_rule(ws, col_map['action'], 'MISSING', fills['red'])

        if 'priority' in col_map:
            self._apply_conditional_formatting_rule(ws, col_map['priority'], 'High', fills['red'], Font(bold=True, color="9C0006"))
            self._apply_conditional_formatting_rule(ws, col_map['priority'], 'Medium', fills['yellow'])

        if 'needs_review' in col_map:
            self._apply_conditional_formatting_rule(ws, col_map['needs_review'], 'YES', fills['red'], Font(bold=True))

    def _process_hyperlinks(self, ws, link_font):
        """Process and convert hyperlinks."""
        # Process regular link columns
        header_row = [cell.value for cell in ws[1]]
        link_columns = [idx for idx, h in enumerate(header_row, 1) if 'Link' in str(h) or 'Documents' in str(h)]

        for col_idx in link_columns:
            for row_idx in range(2, min(ws.max_row + 1, 1002)):
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

        # Process source columns with value|||link format
        for col_idx in range(1, ws.max_column + 1):
            for row_idx in range(2, min(ws.max_row + 1, 1002)):
                cell = ws.cell(row_idx, col_idx)
                cell_value = str(cell.value or '')

                if '|||' in cell_value:
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
    # VBA LINKS
    # ====================

    def add_participant_links(self, filepath: str, master_view_df: pd.DataFrame):
        """Add VBA-powered or simple hyperlinks for participant filtering."""
        try:
            import xlwings as xw
            self._add_links_with_xlwings(filepath, master_view_df)
        except ImportError:
            logger.warning("xlwings not installed - using simple hyperlinks")
            self._add_links_without_xlwings(filepath, master_view_df)

    def _add_links_with_xlwings(self, filepath: str, master_view_df: pd.DataFrame):
        """Add VBA macros for auto-filtering (requires xlwings)."""
        import xlwings as xw

        app = xw.App(visible=False)
        wb = app.books.open(filepath)

        try:
            # Add VBA macro
            vba_code = '''
Sub FilterParticipant(participantID As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Participant Master View")
    ws.Activate
    On Error Resume Next
    ws.AutoFilterMode = False
    On Error GoTo 0
    ws.Range("A1").AutoFilter
    ws.Range("A1").AutoFilter Field:=1, Criteria1:=CStr(participantID)
    ws.Range("A2").Select
    ActiveWindow.ScrollRow = 1
End Sub
'''
            try:
                vb_module = wb.api.VBProject.VBComponents.Add(1)
                vb_module.CodeModule.AddFromString(vba_code)
                logger.info("VBA module added")
            except:
                logger.warning("Could not add VBA (enable 'Trust access to VBA')")

            # Add links
            summary_ws = wb.sheets['Participant Summary']
            header_row = summary_ws.range('1:1').value

            view_col = header_row.index('View Details') + 1
            id_col = header_row.index('Participant ID') + 1

            for row_idx in range(2, summary_ws.range('A1').current_region.last_cell.row + 1):
                pid = summary_ws.range(f'{chr(64 + id_col)}{row_idx}').value
                if pid:
                    cell = summary_ws.range(f'{chr(64 + view_col)}{row_idx}')
                    cell.value = "View Details"
                    cell.color = (5, 99, 193)
                    cell.api.Font.Underline = True
                    cell.note = f"Click to filter to Participant {pid}"

            wb.save()
            logger.info("VBA links added")
        finally:
            wb.close()
            app.quit()

    def _add_links_without_xlwings(self, filepath: str, master_view_df: pd.DataFrame):
        """Add simple hyperlinks without VBA (fallback)."""
        wb = load_workbook(filepath)

        if 'Participant Summary' not in wb.sheetnames or 'Participant Master View' not in wb.sheetnames:
            return

        summary_ws = wb['Participant Summary']
        header_row = [cell.value for cell in summary_ws[1]]

        try:
            view_col = header_row.index('View Details') + 1
            id_col = header_row.index('Participant ID') + 1
        except ValueError:
            return

        # Map participants to first row
        participant_rows = {}
        if not master_view_df.empty:
            for idx, row in master_view_df.iterrows():
                pid = row['Participant ID']
                if pid not in participant_rows:
                    participant_rows[pid] = idx + 2

        # Add hyperlinks
        for row_idx in range(2, summary_ws.max_row + 1):
            pid = summary_ws.cell(row_idx, id_col).value
            if pid in participant_rows:
                cell = summary_ws.cell(row_idx, view_col)
                cell.hyperlink = f"#'Participant Master View'!A{participant_rows[pid]}"
                cell.value = "View Details (manual filter)"
                cell.font = Font(color="0563C1", underline="single")

        wb.save(filepath)
        logger.info(f"Added {len(participant_rows)} simple hyperlinks")

    # ====================
    # INSTRUCTIONS
    # ====================

    def _create_instructions(self, plan_name: str) -> pd.DataFrame:
        """Create instructions sheet."""
        instructions = [
            ['Interactive Extraction Report', f'Plan: {plan_name}'],
            ['', ''],
            ['WORKFLOW:', ''],
            ['1. Participant Summary', 'Start here - overview of all participants'],
            ['2. Click View Details', 'Jump to Master View filtered (with xlwings) or manually filter'],
            ['3. Review elements', 'Check Quality column - Green=Good, Yellow=Review, Red=Conflict'],
            ['4. Click values', 'Values are hyperlinked to source documents'],
            ['5. Resolve conflicts', 'Compare sources and determine correct value'],
            ['', ''],
            ['QUALITY INDICATORS:', ''],
            ['• Good (Green)', 'Single value, high confidence - DATABASE READY'],
            ['• Review (Yellow)', 'Single value but needs verification'],
            ['• Conflict (Red)', 'Multiple different values - MUST RESOLVE'],
            ['• Missing (Light Red)', 'No value extracted'],
            ['', ''],
            ['SOURCE COLUMNS:', ''],
            ['• Value with hyperlink', 'Has data - click to open document'],
            ['• "-"', 'Expected but missing (source was checked)'],
            ['• Blank', 'Not expected (source not relevant for this element)'],
            ['', ''],
            ['TIPS:', ''],
            ['• Filter by Quality = "Conflict"', 'Focus on critical items first'],
            ['• Filter by Participant ID', 'Review one participant at a time'],
            ['', ''],
            ['SharePoint Base URL:', self.sharepoint_base_url],
        ]
        return pd.DataFrame(instructions, columns=['Topic', 'Description'])

    # ====================
    # MAIN REPORT GENERATION
    # ====================

    def generate_interactive_report(self, validated_df: pd.DataFrame, output_path: str, plan_name: str) -> str:
        """Generate comprehensive interactive Excel report."""
        logger.info(f"Generating interactive report for {plan_name}")

        # Generate all tabs
        all_data = self.create_all_data_tab(validated_df)
        comparison = self.create_value_comparison_tab(validated_df)
        participant_summary = self.create_participant_summary(validated_df)
        master_view = self.create_participant_master_view(validated_df)
        review_needed = self.create_review_needed_tab(validated_df, comparison)
        source_stats = self.create_source_statistics_tab(validated_df)
        participant_stats = self.create_participant_statistics_tab(validated_df, comparison)
        instructions = self._create_instructions(plan_name)

        # Write to Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            tabs = [
                (participant_summary, 'Participant Summary'),
                (master_view, 'Participant Master View'),
                (all_data, 'All Extracted Data'),
                (comparison, 'Value Comparison'),
                (review_needed, 'Review Needed'),
                (source_stats, 'Source Statistics'),
                (participant_stats, 'Participant Statistics'),
                (instructions, 'Instructions')
            ]

            for df, sheet_name in tabs:
                if not df.empty or sheet_name == 'Instructions':
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    logger.info(f"Added {sheet_name}: {len(df)} rows")

        # Apply formatting and links
        self.apply_excel_formatting(output_path)
        if not participant_summary.empty and not master_view.empty:
            self.add_participant_links(output_path, master_view)

        logger.info(f"Interactive report generated: {output_path}")
        return output_path


def main():
    """Example usage."""
    pass


if __name__ == "__main__":
    main()
