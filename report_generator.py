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
from openpyxl.formatting.rule import CellIsRule, Rule
from openpyxl.styles.differential import DifferentialStyle
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

        # Store original filename for links, then remove extension for display
        report_df['_original_filename'] = report_df['filename']
        report_df['filename'] = report_df['filename'].apply(
            lambda x: x.rsplit('.', 1)[0] if pd.notna(x) and '.' in x else x
        )

        # Add SharePoint links (using original filename with extension)
        report_df['Document Link'] = report_df.apply(
            lambda row: self.create_sharepoint_link(row['source'], row['_original_filename']),
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
        Redesigned for better readability with many sources.

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

            # Get all unique values for this element and track filenames
            values_by_source = {}
            files_by_source = {}

            for _, row in group.iterrows():
                source = row['source']
                value = row.get('cleaned_value')
                if pd.isna(value):
                    value = row.get('value')

                filename = row.get('filename', '')
                # Remove file extension
                filename_no_ext = filename.rsplit('.', 1)[0] if '.' in filename else filename

                # Track multiple values from same source
                if source not in values_by_source:
                    values_by_source[source] = []
                    files_by_source[source] = []

                if value not in values_by_source[source]:
                    values_by_source[source].append(value)

                # Track unique filenames per source
                if filename_no_ext and filename_no_ext not in files_by_source[source]:
                    files_by_source[source].append(filename_no_ext)

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

            # Build compact values string (only sources with actual values)
            sources_with_values = []
            for source in sorted(values_by_source.keys()):
                values = values_by_source[source]
                # Only include if has non-null values
                non_null = [v for v in values if pd.notna(v)]
                if non_null:
                    value_str = " | ".join([str(v) for v in non_null])
                    sources_with_values.append(f"{source}: {value_str}")

            compact_values = " || ".join(sources_with_values) if sources_with_values else "No values"

            # Build document links string
            all_doc_links = []
            for source in sorted(files_by_source.keys()):
                filenames = files_by_source[source]
                for fname in filenames:
                    link = self.create_sharepoint_link(source, fname)
                    all_doc_links.append(link)

            # Join links with " | " separator
            doc_links_str = " | ".join(all_doc_links) if all_doc_links else ""

            # Build row
            row_data = {
                'Participant ID': participant_id,
                'Element': element,
                'Status': status,
                'Values': compact_values,
                'Sources Checked': len(values_by_source),
                'Unique Values': len(unique_values),
                'Documents': doc_links_str
            }

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
    # TAB 3: PARTICIPANT SUMMARY (Compact Overview for Large Datasets)
    # ====================

    def create_participant_summary(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Create compact participant summary for large datasets (500+ participants).

        Shows one row per participant with summary metrics and action items,
        WITHOUT individual element columns (too wide with 30+ elements).

        Args:
            df: DataFrame with validated extractions

        Returns:
            Compact participant summary DataFrame
        """
        if df.empty or 'participant_id' not in df.columns:
            logger.warning("No participant IDs for summary")
            return pd.DataFrame()

        logger.info("Creating Participant Summary (scaled for large datasets)")

        df_with_id = df[df['participant_id'].notna()].copy()

        if df_with_id.empty:
            logger.warning("No extractions with participant IDs")
            return pd.DataFrame()

        all_elements = sorted(df_with_id['element'].unique())
        total_elements = len(all_elements)

        summary_rows = []

        for participant_id, participant_group in df_with_id.groupby('participant_id'):
            # Count elements by quality
            good_count = 0
            review_count = 0
            conflict_count = 0
            missing_count = 0

            # Track which elements have issues (for action details)
            conflict_elements = []
            review_elements = []
            missing_elements = []

            all_sources = set()
            all_doc_links = []

            for element in all_elements:
                element_data = participant_group[participant_group['element'] == element]

                if element_data.empty:
                    missing_count += 1
                    missing_elements.append(element)
                else:
                    # Track sources
                    for source in element_data['source'].unique():
                        all_sources.add(source)

                    # Get unique values
                    values = element_data['cleaned_value'].dropna()
                    if values.empty:
                        values = element_data['value'].dropna()

                    unique_values = values.unique()

                    if len(unique_values) == 0:
                        missing_count += 1
                        missing_elements.append(element)
                    elif len(unique_values) == 1:
                        # Check confidence
                        has_low_conf = (element_data['confidence'] == 'LOW').any()
                        has_flags = element_data['flags'].notna().any() and (element_data['flags'] != '').any()

                        if has_low_conf or has_flags:
                            review_count += 1
                            review_elements.append(element)
                        else:
                            good_count += 1
                    else:
                        conflict_count += 1
                        conflict_elements.append(element)

            # Collect document links
            for _, row in participant_group.iterrows():
                link = self.create_sharepoint_link(row['source'], row.get('_original_filename', row['filename']))
                all_doc_links.append(link)

            # Calculate percentages
            completeness_pct = (good_count / total_elements * 100) if total_elements > 0 else 0
            quality_pct = (good_count / (good_count + review_count + conflict_count) * 100) if (good_count + review_count + conflict_count) > 0 else 0

            # Determine status and priority
            if conflict_count > 0:
                status = 'Conflicts'
                priority = 'High'
                action = f"Resolve {conflict_count} conflicts: {', '.join(conflict_elements[:3])}" + ("..." if len(conflict_elements) > 3 else "")
            elif review_count > 0:
                status = 'Review'
                priority = 'Medium' if review_count > 5 else 'Low'
                action = f"Review {review_count} items: {', '.join(review_elements[:3])}" + ("..." if len(review_elements) > 3 else "")
            elif missing_count > 0:
                status = 'Incomplete'
                priority = 'Medium' if missing_count > 10 else 'Low'
                action = f"Fill {missing_count} missing: {', '.join(missing_elements[:3])}" + ("..." if len(missing_elements) > 3 else "")
            else:
                status = 'Complete'
                priority = 'None'
                action = 'Database Ready ✓'

            summary_rows.append({
                'Participant ID': participant_id,
                'View Details': f'#Master_View!A1',  # Will be updated with actual row link after Master View is created
                'Status': status,
                'Priority': priority,
                'Completeness %': round(completeness_pct, 1),
                'Quality %': round(quality_pct, 1),
                'Good': good_count,
                'Review': review_count,
                'Conflicts': conflict_count,
                'Missing': missing_count,
                'Total': total_elements,
                'Sources': len(all_sources),
                'Action Required': action,
                'Documents': ' | '.join(set(all_doc_links[:3])) if all_doc_links else ''  # Limit to 3 links
            })

        summary_df = pd.DataFrame(summary_rows)

        # Sort by priority
        priority_order = {'High': 0, 'Medium': 1, 'Low': 2, 'None': 3}
        summary_df['_sort'] = summary_df['Priority'].map(priority_order)
        summary_df = summary_df.sort_values(['_sort', 'Completeness %'])
        summary_df = summary_df.drop('_sort', axis=1)

        logger.info(f"Participant Summary ready with {len(summary_df)} participants")
        return summary_df

    # ====================
    # TAB 4: PARTICIPANT MASTER VIEW (Database Building - Full Detail)
    # ====================

    def create_participant_master_view(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Create participant-centric master view optimized for database building.

        NEW DESIGN: One row per element per participant (not one row per participant)
        - Easy filtering by participant
        - One element at a time for focused review
        - Direct document links for each source that should have this element
        - Much more efficient for reviewing element-by-element

        Args:
            df: DataFrame with validated extractions

        Returns:
            Participant master view DataFrame (one row per participant-element combination)
        """
        if df.empty or 'participant_id' not in df.columns:
            logger.warning("No participant IDs for master view")
            return pd.DataFrame()

        logger.info("Creating Participant Master View (element-per-row design)")

        # Remove rows without participant ID
        df_with_id = df[df['participant_id'].notna()].copy()

        if df_with_id.empty:
            logger.warning("No extractions with participant IDs")
            return pd.DataFrame()

        # Get all unique elements and participants
        all_elements = sorted(df_with_id['element'].unique())
        all_participants = sorted(df_with_id['participant_id'].unique())

        master_rows = []

        for participant_id in all_participants:
            participant_data = df_with_id[df_with_id['participant_id'] == participant_id]

            for element in all_elements:
                element_data = participant_data[participant_data['element'] == element]

                # Build source-specific values and links
                source_values = {}
                source_links = {}

                if not element_data.empty:
                    for _, row in element_data.iterrows():
                        source = row['source']
                        value = row.get('cleaned_value')
                        if pd.isna(value):
                            value = row.get('value')

                        # Track value from this source
                        if source not in source_values:
                            source_values[source] = []
                        if pd.notna(value) and value not in source_values[source]:
                            source_values[source].append(str(value))

                        # Create link for this source
                        if source not in source_links:
                            filename = row.get('_original_filename', row.get('filename', ''))
                            source_links[source] = self.create_sharepoint_link(source, filename)

                # Determine best value and quality
                all_values = []
                for vals in source_values.values():
                    all_values.extend(vals)

                unique_values = list(set(all_values))

                if len(unique_values) == 0:
                    best_value = ''
                    quality = 'Missing'
                elif len(unique_values) == 1:
                    best_value = unique_values[0]
                    # Check confidence
                    if not element_data.empty:
                        has_low_conf = (element_data['confidence'] == 'LOW').any()
                        has_flags = element_data['flags'].notna().any() and (element_data['flags'] != '').any()
                        quality = 'Review' if (has_low_conf or has_flags) else 'Good'
                    else:
                        quality = 'Good'
                else:
                    best_value = ' | '.join(unique_values)
                    quality = 'Conflict'

                # Build source columns (one link per source)
                row_data = {
                    'Participant ID': participant_id,
                    'Element': element,
                    'Value': best_value,
                    'Quality': quality,
                    'Sources with Data': len(source_values),
                    'Unique Values': len(unique_values)
                }

                # Add individual source links as separate columns (up to first 5 sources)
                sorted_sources = sorted(source_links.keys())
                for i, source in enumerate(sorted_sources[:5], 1):  # Limit to 5 source columns
                    link = source_links[source]
                    value = ' | '.join(source_values.get(source, ['-']))
                    row_data[f'Source{i}'] = source
                    row_data[f'Source{i}_Value'] = value
                    row_data[f'Source{i}_Link'] = link

                # If more than 5 sources, add overflow column
                if len(sorted_sources) > 5:
                    other_sources = sorted_sources[5:]
                    row_data['Other Sources'] = ', '.join(other_sources)

                master_rows.append(row_data)

        master_df = pd.DataFrame(master_rows)

        # Sort by participant, then quality priority (Conflict → Review → Missing → Good)
        quality_order = {'Conflict': 0, 'Review': 1, 'Missing': 2, 'Good': 3}
        master_df['_sort'] = master_df['Quality'].map(quality_order)
        master_df = master_df.sort_values(['Participant ID', '_sort', 'Element'])
        master_df = master_df.drop('_sort', axis=1)

        logger.info(f"Participant Master View ready with {len(master_df)} rows ({len(all_participants)} participants × {len(all_elements)} elements)")
        return master_df

    # ====================
    # TAB 4: REVIEW NEEDED
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

    def _apply_conditional_formatting_rule(self, ws, col_letter, value, fill, font=None):
        """
        Helper to apply a conditional formatting rule without duplication.

        Args:
            ws: Worksheet
            col_letter: Column letter
            value: Value to match
            fill: Fill pattern
            font: Optional font style
        """
        try:
            rule = CellIsRule(operator='equal', formula=[f'"{value}"'], fill=fill)
            if font:
                rule.font = font
            ws.conditional_formatting.add(
                f'{col_letter}2:{col_letter}{ws.max_row}',
                rule
            )
        except Exception as e:
            logger.warning(f"Could not apply conditional formatting for {value} in column {col_letter}: {e}")

    def apply_excel_formatting(self, filepath: str):
        """
        Apply comprehensive Excel formatting to all tabs using conditional formatting for performance.

        Args:
            filepath: Path to Excel file
        """
        try:
            logger.info(f"Applying formatting to {filepath}")
            wb = load_workbook(filepath)

            # Define styles
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            link_font = Font(color="0563C1", underline="single")

            # Define fills for conditional formatting
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            light_red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

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

                # Get column mapping
                header_row = [cell.value for cell in ws[1]]
                col_map = {}
                link_columns = []

                for idx, header in enumerate(header_row, start=1):
                    col_letter = ws.cell(1, idx).column_letter
                    header_str = str(header) if header else ''

                    # Map known columns
                    if header == 'Confidence':
                        col_map['confidence'] = col_letter
                    elif header == 'Status':
                        col_map['status'] = col_letter
                    elif header == 'Action' or header == 'Action Required':
                        col_map['action'] = col_letter
                    elif header == 'Priority':
                        col_map['priority'] = col_letter
                    elif header == 'Needs Review':
                        col_map['needs_review'] = col_letter
                    elif header_str.endswith('_Quality'):
                        if 'quality_cols' not in col_map:
                            col_map['quality_cols'] = []
                        col_map['quality_cols'].append(col_letter)

                    # Track link columns
                    if 'Link' in header_str or 'Documents' in header_str:
                        link_columns.append(idx)

                # Apply conditional formatting using helper method
                if 'confidence' in col_map:
                    self._apply_conditional_formatting_rule(ws, col_map['confidence'], 'HIGH', green_fill)
                    self._apply_conditional_formatting_rule(ws, col_map['confidence'], 'LOW', red_fill)

                if 'status' in col_map:
                    self._apply_conditional_formatting_rule(ws, col_map['status'], 'Match', green_fill)
                    self._apply_conditional_formatting_rule(ws, col_map['status'], 'Complete', green_fill)
                    self._apply_conditional_formatting_rule(ws, col_map['status'], 'Missing', red_fill)
                    self._apply_conditional_formatting_rule(ws, col_map['status'], 'Incomplete', yellow_fill)
                    self._apply_conditional_formatting_rule(ws, col_map['status'], 'Conflicts', yellow_fill)
                    self._apply_conditional_formatting_rule(ws, col_map['status'], 'Review', yellow_fill)
                    # Handle "Conflict" variants
                    for conflict_val in ['Conflict (Within Source)', 'Conflict (Cross Source)', 'Conflict (Within + Cross)']:
                        self._apply_conditional_formatting_rule(ws, col_map['status'], conflict_val, yellow_fill)

                if 'quality_cols' in col_map:
                    for qual_col in col_map['quality_cols']:
                        self._apply_conditional_formatting_rule(ws, qual_col, 'Good', green_fill)
                        self._apply_conditional_formatting_rule(ws, qual_col, 'Review', yellow_fill)
                        self._apply_conditional_formatting_rule(ws, qual_col, 'Conflict', red_fill)
                        self._apply_conditional_formatting_rule(ws, qual_col, 'Missing', light_red_fill)

                if 'action' in col_map:
                    self._apply_conditional_formatting_rule(ws, col_map['action'], 'REVIEW', yellow_fill, Font(bold=True))
                    self._apply_conditional_formatting_rule(ws, col_map['action'], 'MISSING', red_fill)

                if 'priority' in col_map:
                    self._apply_conditional_formatting_rule(ws, col_map['priority'], 'High', red_fill, Font(bold=True, color="9C0006"))
                    self._apply_conditional_formatting_rule(ws, col_map['priority'], 'Medium', yellow_fill)

                if 'needs_review' in col_map:
                    self._apply_conditional_formatting_rule(ws, col_map['needs_review'], 'YES', red_fill, Font(bold=True))

                # Convert URLs to hyperlinks
                for col_idx in link_columns:
                    for row_idx in range(2, min(ws.max_row + 1, 1002)):  # Limit to 1000 rows for performance
                        cell = ws.cell(row_idx, col_idx)
                        cell_value = str(cell.value or '')

                        if cell_value and (cell_value.startswith('http://') or cell_value.startswith('https://')):
                            if ' | ' in cell_value:
                                urls = cell_value.split(' | ')
                                cell.hyperlink = urls[0]
                                cell.value = f"Open File ({len(urls)} docs)"
                                cell.font = link_font
                            else:
                                cell.hyperlink = cell_value
                                cell.value = "Open File"
                                cell.font = link_font

                # Auto-adjust column widths (optimized - sample first 100 rows)
                sample_size = min(100, ws.max_row)
                for col_idx in range(1, ws.max_column + 1):
                    column_letter = ws.cell(1, col_idx).column_letter
                    header_len = len(str(ws.cell(1, col_idx).value or ''))
                    max_length = header_len

                    for row_idx in range(2, min(sample_size + 2, ws.max_row + 1)):
                        try:
                            cell_value = str(ws.cell(row_idx, col_idx).value or '')
                            max_length = max(max_length, len(cell_value))
                        except:
                            pass

                    ws.column_dimensions[column_letter].width = min(max_length + 2, 50)

                # Freeze top row and enable autofilter
                ws.freeze_panes = 'A2'
                if ws.max_row > 1:
                    ws.auto_filter.ref = ws.dimensions

            wb.save(filepath)
            logger.info("Formatting applied successfully")

        except Exception as e:
            logger.error(f"Error applying formatting: {e}", exc_info=True)

    # ====================
    # CROSS-TAB HYPERLINKS
    # ====================

    def add_participant_links(self, filepath: str, master_view_df: pd.DataFrame):
        """
        Add hyperlinks from Participant Summary to Master View tab.

        Creates clickable "View Details" links that navigate to the first row
        of each participant in the Master View tab.

        Args:
            filepath: Path to Excel file
            master_view_df: Master View DataFrame to get row positions
        """
        try:
            logger.info("Adding participant hyperlinks between tabs")
            wb = load_workbook(filepath)

            if 'Participant Summary' not in wb.sheetnames or 'Participant Master View' not in wb.sheetnames:
                logger.warning("Cannot add links - required sheets not found")
                return

            summary_ws = wb['Participant Summary']
            link_font = Font(color="0563C1", underline="single")

            # Find the "View Details" column in Summary
            header_row = [cell.value for cell in summary_ws[1]]
            try:
                view_details_col_idx = header_row.index('View Details') + 1
                participant_id_col_idx = header_row.index('Participant ID') + 1
            except ValueError:
                logger.warning("View Details or Participant ID column not found")
                return

            # Build a map of participant_id -> first row in Master View
            participant_row_map = {}
            if not master_view_df.empty:
                for idx, row in master_view_df.iterrows():
                    participant_id = row['Participant ID']
                    if participant_id not in participant_row_map:
                        # Row in Excel = dataframe index + 2 (1 for header, 1 for 0-based indexing)
                        participant_row_map[participant_id] = idx + 2

            # Update links in Summary sheet
            for row_idx in range(2, summary_ws.max_row + 1):
                participant_id = summary_ws.cell(row_idx, participant_id_col_idx).value

                if participant_id in participant_row_map:
                    target_row = participant_row_map[participant_id]
                    cell = summary_ws.cell(row_idx, view_details_col_idx)

                    # Create internal hyperlink to Master View
                    cell.hyperlink = f"#'Participant Master View'!A{target_row}"
                    cell.value = "View Details"
                    cell.font = link_font
                else:
                    # No data in Master View for this participant
                    cell = summary_ws.cell(row_idx, view_details_col_idx)
                    cell.value = "-"

            wb.save(filepath)
            logger.info(f"Added {len(participant_row_map)} participant detail links")

        except Exception as e:
            logger.error(f"Error adding participant links: {e}", exc_info=True)

    # ====================
    # INSTRUCTIONS TAB
    # ====================

    def _create_instructions(self, plan_name: str) -> pd.DataFrame:
        """Create instructions sheet for the report."""
        instructions = [
            ['Interactive Extraction Report', f'Plan: {plan_name}'],
            ['', ''],
            ['RECOMMENDED WORKFLOW - LARGE DATASETS (500+ participants):', ''],
            ['START HERE → Tab 1: Participant Summary', '⭐ COMPACT VIEW - One row per participant, summary metrics only'],
            ['', '1. Sort by Priority (High first) - focus on conflicts and critical items'],
            ['', '2. Review Action Required column - shows exactly what needs attention'],
            ['', '3. Filter Status = "Conflicts" to handle most critical items first'],
            ['', '4. Use Completeness % to track progress (aim for 100%)'],
            ['', '5. Click "View Details" link to jump directly to that participant in Master View'],
            ['', '6. Click "Open File" to verify and resolve issues'],
            ['', '7. For cross-source comparison: Use Tab 4 (Value Comparison)'],
            ['', ''],
            ['TAB GUIDE:', ''],
            ['1. Participant Summary', '⭐ PRIMARY for 500+ participants (START HERE)'],
            ['', 'COMPACT: Only 14 columns - no individual elements shown'],
            ['', 'Shows: Status, Priority, Completeness %, Quality %, Good/Review/Conflicts/Missing counts'],
            ['', 'NEW: "View Details" hyperlink - click to jump directly to that participant in Master View'],
            ['', 'Action Required: Lists specific elements needing attention'],
            ['', 'Best for: Quick overview, prioritizing work, tracking overall progress'],
            ['', 'When to use: Initial review, daily progress tracking, identifying problem participants'],
            ['', ''],
            ['2. Participant Master View', '⭐ DETAILED VIEW - One row per element per participant'],
            ['', 'DESIGN: Many rows per participant (one for each element)'],
            ['', 'Shows: Element name, value, quality, and individual source links'],
            ['', 'Individual source columns: Each source that has data gets its own link'],
            ['', 'Best for: Element-by-element review, direct document verification'],
            ['', 'When to use: After identifying participants in Summary, for detailed review'],
            ['', 'TIP: Filter Participant ID column to focus on one participant at a time'],
            ['', 'TIP: Filter Element column to review one element across all participants'],
            ['', ''],
            ['3. All Extracted Data', 'Raw extraction data - one row per extraction'],
            ['', 'Use for detailed investigation of specific extractions'],
            ['', 'GREEN = High confidence | RED = Low confidence'],
            ['', ''],
            ['3. Value Comparison', 'Element-by-element comparison across sources'],
            ['', 'Values column shows: SourceName: Value || OtherSource: Value'],
            ['', 'Use when you need to see which sources contributed what values'],
            ['', 'GREEN = Match | YELLOW = Conflict | RED = Missing'],
            ['', ''],
            ['4. Review Needed', 'Consolidated list of all items requiring attention'],
            ['', 'Low confidence extractions + conflicts grouped together'],
            ['', 'Prioritized by urgency (High/Medium/Low)'],
            ['', ''],
            ['5. Source Statistics', 'Performance metrics by document source'],
            ['', 'Identifies which sources have extraction quality issues'],
            ['', ''],
            ['6. Participant Statistics', 'Completeness overview by participant'],
            ['', 'Quick view of which participants need the most work'],
            ['', ''],
            ['QUALITY INDICATORS (in Participant Master View):', ''],
            ['• Good', 'Single value, high confidence - DATABASE READY'],
            ['• Review', 'Single value but flagged for manual verification'],
            ['• Conflict', 'Multiple different values found - MUST RESOLVE'],
            ['• Missing', 'No value extracted - needs manual entry or source review'],
            ['', ''],
            ['EFFICIENT WORKFLOW TIPS:', ''],
            ['• Sort by Status', 'Handle Conflicts first (highest priority), then Review, then Incomplete'],
            ['• Use Excel filters', 'Filter to Status = "Conflicts" to focus on critical items'],
            ['• Group similar participants', 'Sort by Action Required to batch similar work'],
            ['• Track progress', 'Filter Status = "Complete" to see what\'s database-ready'],
            ['• Quick verification', 'Ctrl+Click Documents links to open multiple sources at once'],
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
        participant_summary = self.create_participant_summary(validated_df)  # NEW: Compact for large datasets
        master_view = self.create_participant_master_view(validated_df)  # Detailed view
        review_needed = self.create_review_needed_tab(validated_df, comparison)
        source_stats = self.create_source_statistics_tab(validated_df)
        participant_stats = self.create_participant_statistics_tab(validated_df, comparison)
        instructions = self._create_instructions(plan_name)

        # Create Excel writer
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write all tabs in order of priority for database building workflow

            # Tab 1: Participant Summary (PRIMARY for large datasets 500+)
            if not participant_summary.empty:
                participant_summary.to_excel(writer, sheet_name='Participant Summary', index=False)
                logger.info(f"Added Participant Summary: {len(participant_summary)} participants")

            # Tab 2: Participant Master View (Detailed - may be wide with 30+ elements)
            if not master_view.empty:
                master_view.to_excel(writer, sheet_name='Participant Master View', index=False)
                logger.info(f"Added Participant Master View: {len(master_view)} participants")

            # Tab 3: All Extracted Data
            if not all_data.empty:
                all_data.to_excel(writer, sheet_name='All Extracted Data', index=False)
                logger.info(f"Added All Extracted Data: {len(all_data)} rows")

            # Tab 4: Value Comparison
            if not comparison.empty:
                comparison.to_excel(writer, sheet_name='Value Comparison', index=False)
                logger.info(f"Added Value Comparison: {len(comparison)} comparisons")

            # Tab 5: Review Needed
            if not review_needed.empty:
                review_needed.to_excel(writer, sheet_name='Review Needed', index=False)
                logger.info(f"Added Review Needed: {len(review_needed)} items")

            # Tab 6: Source Statistics
            if not source_stats.empty:
                source_stats.to_excel(writer, sheet_name='Source Statistics', index=False)
                logger.info(f"Added Source Statistics: {len(source_stats)} sources")

            # Tab 7: Participant Statistics
            if not participant_stats.empty:
                participant_stats.to_excel(writer, sheet_name='Participant Statistics', index=False)
                logger.info(f"Added Participant Statistics: {len(participant_stats)} participants")

            # Instructions always added
            instructions.to_excel(writer, sheet_name='Instructions', index=False)
            logger.info("Added Instructions tab")

        # Apply formatting
        self.apply_excel_formatting(output_path)

        # Add hyperlinks between Summary and Master View
        if not participant_summary.empty and not master_view.empty:
            self.add_participant_links(output_path, master_view)

        logger.info(f"Redesigned interactive report generated: {output_path}")
        return output_path


def main():
    """Example usage of ReportGenerator."""
    pass


if __name__ == "__main__":
    main()
