"""
Statistics Manager Module
Comprehensive statistics tracking for extraction workflow including:
- Parsing success/failure rates
- Confidence level distribution
- Participant-level statistics
- Module execution timing
"""

import pandas as pd
import numpy as np
import logging
from typing import Dict, List, Optional, Tuple
from datetime import datetime

# Configure module logger
logger = logging.getLogger(__name__)


class StatisticsManager:
    """Manage comprehensive statistics for extraction workflow."""

    def __init__(self):
        """Initialize the statistics manager."""
        self.timing_stats = {}
        logger.info("Initialized StatisticsManager")

    def record_timing(self, module_name: str, start_time: datetime, end_time: datetime):
        """
        Record execution time for a module.

        Args:
            module_name: Name of the module (e.g., "Text Cleaning", "Extraction")
            start_time: Start datetime
            end_time: End datetime
        """
        duration = (end_time - start_time).total_seconds()
        self.timing_stats[module_name] = {
            'start': start_time,
            'end': end_time,
            'duration_seconds': duration,
            'duration_formatted': self._format_duration(duration)
        }
        logger.debug(f"Recorded timing for {module_name}: {duration:.2f}s")

    def _format_duration(self, seconds: float) -> str:
        """
        Format duration in human-readable format.

        Args:
            seconds: Duration in seconds

        Returns:
            Formatted string (e.g., "1m 23s", "45s")
        """
        if seconds < 60:
            return f"{seconds:.2f}s"
        elif seconds < 3600:
            minutes = int(seconds // 60)
            secs = seconds % 60
            return f"{minutes}m {secs:.2f}s"
        else:
            hours = int(seconds // 3600)
            minutes = int((seconds % 3600) // 60)
            secs = seconds % 60
            return f"{hours}h {minutes}m {secs:.2f}s"

    def calculate_parsing_statistics(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate parsing success/failure statistics.

        Args:
            df: DataFrame with 'value' and 'cleaned_value' columns

        Returns:
            DataFrame with parsing statistics by source and element
        """
        if df.empty:
            logger.warning("Empty dataframe provided for parsing statistics")
            return pd.DataFrame()

        logger.info("Calculating parsing statistics")

        parsing_stats = []

        # Group by source and element
        for (source, element), group in df.groupby(['source', 'element']):
            total = len(group)

            # Count raw values present
            raw_present = group['value'].notna().sum()

            # Count successfully parsed (cleaned_value is not None)
            parsed_success = group['cleaned_value'].notna().sum()

            # Count parsing failures (raw value present but cleaned_value is None)
            parse_failures = ((group['value'].notna()) & (group['cleaned_value'].isna())).sum()

            # Calculate rates
            parse_success_rate = (parsed_success / raw_present * 100) if raw_present > 0 else 0

            parsing_stats.append({
                'Source': source,
                'Element': element,
                'Total Extractions': total,
                'Raw Values Present': raw_present,
                'Successfully Parsed': parsed_success,
                'Parse Failures': parse_failures,
                'Parse Success Rate %': round(parse_success_rate, 2)
            })

        result_df = pd.DataFrame(parsing_stats)
        logger.info(f"Calculated parsing statistics for {len(result_df)} source/element combinations")
        return result_df

    def calculate_confidence_statistics(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate confidence level statistics.

        Args:
            df: DataFrame with 'confidence' column

        Returns:
            DataFrame with confidence statistics by source
        """
        if df.empty or 'confidence' not in df.columns:
            logger.warning("Empty dataframe or missing 'confidence' column")
            return pd.DataFrame()

        logger.info("Calculating confidence statistics")

        confidence_stats = []

        # Overall statistics
        total = len(df)
        high_conf = (df['confidence'] == 'HIGH').sum()
        low_conf = (df['confidence'] == 'LOW').sum()

        confidence_stats.append({
            'Source': 'ALL SOURCES',
            'Total Extractions': total,
            'High Confidence': high_conf,
            'Low Confidence': low_conf,
            'High Confidence %': round(high_conf / total * 100, 2) if total > 0 else 0,
            'Low Confidence %': round(low_conf / total * 100, 2) if total > 0 else 0
        })

        # By source
        for source, group in df.groupby('source'):
            total = len(group)
            high_conf = (group['confidence'] == 'HIGH').sum()
            low_conf = (group['confidence'] == 'LOW').sum()

            confidence_stats.append({
                'Source': source,
                'Total Extractions': total,
                'High Confidence': high_conf,
                'Low Confidence': low_conf,
                'High Confidence %': round(high_conf / total * 100, 2) if total > 0 else 0,
                'Low Confidence %': round(low_conf / total * 100, 2) if total > 0 else 0
            })

        result_df = pd.DataFrame(confidence_stats)
        logger.info(f"Calculated confidence statistics for {len(result_df)} groups")
        return result_df

    def calculate_flag_statistics(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate statistics on validation flags.

        Args:
            df: DataFrame with 'flags' column

        Returns:
            DataFrame with flag counts
        """
        if df.empty or 'flags' not in df.columns:
            logger.warning("Empty dataframe or missing 'flags' column")
            return pd.DataFrame()

        logger.info("Calculating flag statistics")

        # Count each flag type
        flag_counts = {}

        for flags_str in df['flags']:
            if pd.isna(flags_str) or flags_str == '':
                continue

            # Split comma-separated flags
            flags = [f.strip() for f in str(flags_str).split(',')]
            for flag in flags:
                if flag:
                    flag_counts[flag] = flag_counts.get(flag, 0) + 1

        # Convert to DataFrame
        flag_stats = [
            {'Flag Type': flag, 'Count': count}
            for flag, count in sorted(flag_counts.items(), key=lambda x: x[1], reverse=True)
        ]

        result_df = pd.DataFrame(flag_stats)
        logger.info(f"Calculated statistics for {len(result_df)} flag types")
        return result_df

    def calculate_participant_statistics(self, df: pd.DataFrame, id_column: str = 'participant_id') -> pd.DataFrame:
        """
        Calculate statistics by participant ID across all sources.

        Args:
            df: DataFrame with participant ID column
            id_column: Name of the ID column

        Returns:
            DataFrame with participant-level statistics
        """
        if df.empty:
            logger.warning("Empty dataframe provided for participant statistics")
            return pd.DataFrame()

        # Check if ID column exists
        if id_column not in df.columns:
            logger.info(f"No '{id_column}' column found - skipping participant statistics")
            return pd.DataFrame()

        logger.info("Calculating participant-level statistics")

        participant_stats = []

        # Group by participant ID
        for participant_id, group in df.groupby(id_column):
            if pd.isna(participant_id):
                continue

            # Basic counts
            total_extractions = len(group)
            sources = group['source'].nunique()
            documents = group['filename'].nunique()

            # Extraction success
            values_present = group['value'].notna().sum()
            values_missing = group['value'].isna().sum()

            # Parsing success
            parsed_success = group['cleaned_value'].notna().sum() if 'cleaned_value' in group.columns else 0
            parse_failures = values_present - parsed_success if 'cleaned_value' in group.columns else 0

            # Confidence levels
            if 'confidence' in group.columns:
                high_conf = (group['confidence'] == 'HIGH').sum()
                low_conf = (group['confidence'] == 'LOW').sum()
                high_conf_pct = round(high_conf / total_extractions * 100, 2) if total_extractions > 0 else 0
            else:
                high_conf = 0
                low_conf = 0
                high_conf_pct = 0

            participant_stats.append({
                'Participant ID': participant_id,
                'Sources': sources,
                'Documents': documents,
                'Total Extractions': total_extractions,
                'Values Present': values_present,
                'Values Missing': values_missing,
                'Successfully Parsed': parsed_success,
                'Parse Failures': parse_failures,
                'High Confidence': high_conf,
                'Low Confidence': low_conf,
                'High Confidence %': high_conf_pct
            })

        result_df = pd.DataFrame(participant_stats)

        # Sort by participant ID
        if not result_df.empty:
            result_df = result_df.sort_values('Participant ID')

        logger.info(f"Calculated statistics for {len(result_df)} participants")
        return result_df

    def calculate_element_statistics(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate detailed element-level statistics.

        Args:
            df: DataFrame with extraction data

        Returns:
            DataFrame with element statistics
        """
        if df.empty:
            logger.warning("Empty dataframe provided for element statistics")
            return pd.DataFrame()

        logger.info("Calculating element-level statistics")

        element_stats = []

        # Group by source and element
        for (source, element), group in df.groupby(['source', 'element']):
            total = len(group)

            # Extraction success
            found = group['value'].notna().sum()
            not_found = group['value'].isna().sum()
            found_pct = round(found / total * 100, 2) if total > 0 else 0

            # Parsing success
            if 'cleaned_value' in group.columns:
                parsed = group['cleaned_value'].notna().sum()
                parse_failed = ((group['value'].notna()) & (group['cleaned_value'].isna())).sum()
                parse_pct = round(parsed / found * 100, 2) if found > 0 else 0
            else:
                parsed = 0
                parse_failed = 0
                parse_pct = 0

            # Confidence
            if 'confidence' in group.columns:
                high_conf = (group['confidence'] == 'HIGH').sum()
                low_conf = (group['confidence'] == 'LOW').sum()
                high_conf_pct = round(high_conf / total * 100, 2) if total > 0 else 0
            else:
                high_conf = 0
                low_conf = 0
                high_conf_pct = 0

            element_stats.append({
                'Source': source,
                'Element': element,
                'Total': total,
                'Found': found,
                'Not Found': not_found,
                'Found %': found_pct,
                'Parsed': parsed,
                'Parse Failed': parse_failed,
                'Parse %': parse_pct,
                'High Confidence': high_conf,
                'Low Confidence': low_conf,
                'High Confidence %': high_conf_pct
            })

        result_df = pd.DataFrame(element_stats)
        logger.info(f"Calculated statistics for {len(result_df)} source/element combinations")
        return result_df

    def get_timing_statistics(self) -> pd.DataFrame:
        """
        Get timing statistics for all modules.

        Returns:
            DataFrame with timing information
        """
        if not self.timing_stats:
            logger.warning("No timing statistics recorded")
            return pd.DataFrame()

        timing_data = []

        for module_name, stats in self.timing_stats.items():
            timing_data.append({
                'Module': module_name,
                'Start Time': stats['start'].strftime('%Y-%m-%d %H:%M:%S'),
                'End Time': stats['end'].strftime('%Y-%m-%d %H:%M:%S'),
                'Duration': stats['duration_formatted'],
                'Duration (seconds)': round(stats['duration_seconds'], 2)
            })

        result_df = pd.DataFrame(timing_data)

        # Add total time
        if timing_data:
            total_duration = sum(stats['duration_seconds'] for stats in self.timing_stats.values())
            total_row = pd.DataFrame([{
                'Module': 'TOTAL',
                'Start Time': '',
                'End Time': '',
                'Duration': self._format_duration(total_duration),
                'Duration (seconds)': round(total_duration, 2)
            }])
            result_df = pd.concat([result_df, total_row], ignore_index=True)

        logger.info(f"Generated timing statistics for {len(self.timing_stats)} modules")
        return result_df

    def generate_comprehensive_statistics(
        self,
        df: pd.DataFrame,
        include_participant_stats: bool = True
    ) -> Dict[str, pd.DataFrame]:
        """
        Generate all statistics in one call.

        Args:
            df: DataFrame with all extraction data
            include_participant_stats: Whether to generate participant-level stats

        Returns:
            Dictionary mapping statistic type to DataFrame
        """
        logger.info("Generating comprehensive statistics")

        stats = {}

        # Element-level statistics
        stats['element_statistics'] = self.calculate_element_statistics(df)

        # Parsing statistics
        if 'cleaned_value' in df.columns:
            stats['parsing_statistics'] = self.calculate_parsing_statistics(df)

        # Confidence statistics
        if 'confidence' in df.columns:
            stats['confidence_statistics'] = self.calculate_confidence_statistics(df)
            stats['flag_statistics'] = self.calculate_flag_statistics(df)

        # Participant statistics
        if include_participant_stats:
            participant_stats = self.calculate_participant_statistics(df)
            if not participant_stats.empty:
                stats['participant_statistics'] = participant_stats

        # Timing statistics
        timing_stats = self.get_timing_statistics()
        if not timing_stats.empty:
            stats['timing_statistics'] = timing_stats

        logger.info(f"Generated {len(stats)} statistic types")
        return stats


def main():
    """Example usage of StatisticsManager."""
    # Example:
    # stats_mgr = StatisticsManager()
    #
    # # Record timing
    # start = datetime.now()
    # # ... do work ...
    # end = datetime.now()
    # stats_mgr.record_timing("Extraction", start, end)
    #
    # # Generate statistics
    # df = pd.DataFrame(extractions)
    # all_stats = stats_mgr.generate_comprehensive_statistics(df)
    pass


if __name__ == "__main__":
    main()
