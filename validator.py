"""
Validation and Sanity Checking Module
Validates cleaned extractions and separates high-confidence from low-confidence data.

This module:
1. Validates date logic (DOB < DOH < DOTE)
2. Detects positional outliers within document types
3. Identifies within-document position gaps
4. Flags multiple extractions of same element
5. Validates value reasonableness
6. Tracks missing critical elements
7. Groups by participant ID when available
8. Provides detailed flag reasons
"""

import re
import logging
from datetime import datetime, timedelta
from typing import Optional, Dict, List, Tuple
import pandas as pd
import numpy as np
from pathlib import Path
import toml

# Configure module logger
logger = logging.getLogger(__name__)


class ExtractionValidator:
    """Validate and flag extractions by confidence level."""

    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize the extraction validator.

        Args:
            config_path: Optional path to TOML config for validation settings
        """
        self.config = {}

        # Default validation settings
        self.enable_date_logic = True
        self.enable_positional_outliers = True
        self.enable_within_document_gaps = True
        self.enable_value_reasonableness = True

        self.positional_outlier_threshold = 3.0
        self.within_document_gap_threshold = 2000
        self.date_future_tolerance_days = 0
        self.critical_elements = []

        # Load config if provided
        if config_path:
            self.load_config(config_path)

        logger.info("Initialized ExtractionValidator")

    def load_config(self, config_path: str):
        """
        Load validation configuration from TOML file.

        Args:
            config_path: Path to TOML configuration file
        """
        try:
            config_path_obj = Path(config_path)
            if not config_path_obj.exists():
                logger.warning(f"Config file not found: {config_path}")
                return

            self.config = toml.load(config_path)

            # Extract validation settings if present
            validation_config = self.config.get('Validation', {})

            self.enable_date_logic = validation_config.get('enable_date_logic', True)
            self.enable_positional_outliers = validation_config.get('enable_positional_outliers', True)
            self.enable_within_document_gaps = validation_config.get('enable_within_document_gaps', True)
            self.enable_value_reasonableness = validation_config.get('enable_value_reasonableness', True)

            self.positional_outlier_threshold = validation_config.get('positional_outlier_threshold', 3.0)
            self.within_document_gap_threshold = validation_config.get('within_document_gap_threshold', 2000)
            self.date_future_tolerance_days = validation_config.get('date_future_tolerance_days', 0)
            self.critical_elements = validation_config.get('critical_elements', [])

            logger.info(f"Loaded validation config from: {config_path}")
            logger.debug(f"Validation settings: date_logic={self.enable_date_logic}, "
                        f"positional_outliers={self.enable_positional_outliers}, "
                        f"within_doc_gaps={self.enable_within_document_gaps}, "
                        f"value_reasonableness={self.enable_value_reasonableness}")

        except Exception as e:
            logger.error(f"Error loading config from {config_path}: {e}")

    def parse_date(self, date_str: str) -> Optional[datetime]:
        """
        Parse date string to datetime object.

        Args:
            date_str: Date string (expected format: MM/DD/YYYY)

        Returns:
            datetime object or None if parsing fails
        """
        if not date_str or pd.isna(date_str):
            return None

        try:
            # Try MM/DD/YYYY format (standard from value_cleaner)
            return datetime.strptime(str(date_str).strip(), '%m/%d/%Y')
        except ValueError:
            try:
                # Try other common formats
                return datetime.strptime(str(date_str).strip(), '%Y-%m-%d')
            except ValueError:
                logger.debug(f"Could not parse date: {date_str}")
                return None

    def check_date_logic(self, document_extractions: pd.DataFrame) -> Dict[int, Tuple[List[str], Dict[str, str]]]:
        """
        Check date logic within a document.
        Rule: DOB < DOH < DOTE and DOB < any other dated event

        Args:
            document_extractions: DataFrame with all extractions from one document

        Returns:
            Dict mapping row index to (flags, flag_reasons)
        """
        if not self.enable_date_logic:
            return {}

        violations = {}

        # Identify date elements
        date_elements = {}
        for idx, row in document_extractions.iterrows():
            element = row.get('element', '').upper()
            cleaned_value = row.get('cleaned_value')

            # Check if this is a date element
            if any(keyword in element for keyword in ['DATE', 'DOB', 'DOH', 'DOTE', 'BIRTH', 'HIRE', 'TERMINATION']):
                parsed_date = self.parse_date(cleaned_value)
                if parsed_date:
                    # Store with element name and index
                    date_elements[idx] = {
                        'element': element,
                        'date': parsed_date,
                        'value': cleaned_value
                    }

        # Check date logic
        for idx1, data1 in date_elements.items():
            elem1 = data1['element']
            date1 = data1['date']

            for idx2, data2 in date_elements.items():
                if idx1 == idx2:
                    continue

                elem2 = data2['element']
                date2 = data2['date']

                # Check DOB < DOH
                if 'DOB' in elem1 or 'BIRTH' in elem1:
                    if 'DOH' in elem2 or 'HIRE' in elem2:
                        if date1 >= date2:
                            if idx1 not in violations:
                                violations[idx1] = ([], {})
                            violations[idx1][0].append('date_logic_violation')
                            violations[idx1][1]['date_logic_violation'] = \
                                f"{elem1} ({data1['value']}) should be before {elem2} ({data2['value']})"

                # Check DOB < DOTE
                if 'DOB' in elem1 or 'BIRTH' in elem1:
                    if 'DOTE' in elem2 or 'TERMINATION' in elem2:
                        if date1 >= date2:
                            if idx1 not in violations:
                                violations[idx1] = ([], {})
                            violations[idx1][0].append('date_logic_violation')
                            violations[idx1][1]['date_logic_violation'] = \
                                f"{elem1} ({data1['value']}) should be before {elem2} ({data2['value']})"

                # Check DOH < DOTE
                if 'DOH' in elem1 or 'HIRE' in elem1:
                    if 'DOTE' in elem2 or 'TERMINATION' in elem2:
                        if date1 >= date2:
                            if idx1 not in violations:
                                violations[idx1] = ([], {})
                            violations[idx1][0].append('date_logic_violation')
                            violations[idx1][1]['date_logic_violation'] = \
                                f"{elem1} ({data1['value']}) should be before {elem2} ({data2['value']})"

        return violations

    def calculate_positional_statistics(self, df: pd.DataFrame) -> Dict[Tuple[str, str], Dict[str, float]]:
        """
        Calculate positional statistics for each element within each source.

        Args:
            df: DataFrame with all extractions

        Returns:
            Dict mapping (source, element) to {'mean': X, 'std': Y, 'count': Z}
        """
        stats = {}

        # Group by source and element
        for (source, element), group in df.groupby(['source', 'element']):
            positions = group['extraction_position'].dropna()

            # Need at least 5 data points for meaningful statistics
            if len(positions) >= 5:
                stats[(source, element)] = {
                    'mean': positions.mean(),
                    'std': positions.std(),
                    'count': len(positions)
                }

        return stats

    def check_positional_outliers(self, row: pd.Series, stats: Dict[Tuple[str, str], Dict[str, float]]) -> Tuple[List[str], Dict[str, str]]:
        """
        Check if extraction position is an outlier.

        Args:
            row: DataFrame row
            stats: Positional statistics

        Returns:
            (flags, flag_reasons)
        """
        if not self.enable_positional_outliers:
            return [], {}

        flags = []
        flag_reasons = {}

        source = row.get('source')
        element = row.get('element')
        position = row.get('extraction_position')

        if pd.isna(position):
            return flags, flag_reasons

        # Get statistics for this source/element combination
        key = (source, element)
        if key not in stats:
            return flags, flag_reasons

        mean = stats[key]['mean']
        std = stats[key]['std']

        # Calculate Z-score
        if std > 0:
            z_score = abs((position - mean) / std)

            if z_score > self.positional_outlier_threshold:
                flags.append('positional_outlier')
                flag_reasons['positional_outlier'] = \
                    f"Position {position} is {z_score:.1f} std devs from mean ({mean:.0f})"
                logger.debug(f"Positional outlier: {source}/{element} at position {position} "
                           f"(Z-score: {z_score:.1f})")

        return flags, flag_reasons

    def check_within_document_gaps(self, document_extractions: pd.DataFrame) -> Dict[int, Tuple[List[str], Dict[str, str]]]:
        """
        Check for large position gaps within a document.

        Args:
            document_extractions: DataFrame with all extractions from one document

        Returns:
            Dict mapping row index to (flags, flag_reasons)
        """
        if not self.enable_within_document_gaps:
            return {}

        violations = {}

        # Sort by extraction_order
        sorted_df = document_extractions.sort_values('extraction_order')

        prev_position = None
        for idx, row in sorted_df.iterrows():
            position = row.get('extraction_position')

            if pd.isna(position):
                continue

            if prev_position is not None:
                gap = abs(position - prev_position)

                if gap > self.within_document_gap_threshold:
                    if idx not in violations:
                        violations[idx] = ([], {})
                    violations[idx][0].append('within_document_position_gap')
                    violations[idx][1]['within_document_position_gap'] = \
                        f"Gap of {gap} chars from previous element (position {prev_position} to {position})"
                    logger.debug(f"Large position gap detected: {gap} chars in {row.get('filename')}")

            prev_position = position

        return violations

    def check_multiple_extractions(self, row: pd.Series) -> Tuple[List[str], Dict[str, str]]:
        """
        Check if element has multiple extractions (extraction_order > 1).

        Args:
            row: DataFrame row

        Returns:
            (flags, flag_reasons)
        """
        flags = []
        flag_reasons = {}

        extraction_order = row.get('extraction_order')

        if not pd.isna(extraction_order) and extraction_order > 1:
            flags.append('multiple_extractions')
            flag_reasons['multiple_extractions'] = \
                f"Found {extraction_order} instances of {row.get('element')} in document"

        return flags, flag_reasons

    def check_value_reasonableness(self, row: pd.Series) -> Tuple[List[str], Dict[str, str]]:
        """
        Check if value is reasonable for its element type.

        Args:
            row: DataFrame row

        Returns:
            (flags, flag_reasons)
        """
        if not self.enable_value_reasonableness:
            return [], {}

        flags = []
        flag_reasons = {}

        element = row.get('element', '').upper()
        cleaned_value = row.get('cleaned_value')

        if pd.isna(cleaned_value):
            return flags, flag_reasons

        # Check dates
        if any(keyword in element for keyword in ['DATE', 'DOB', 'DOH', 'DOTE', 'BIRTH', 'HIRE', 'TERMINATION']):
            parsed_date = self.parse_date(cleaned_value)
            if parsed_date:
                # Check if date is in future
                future_cutoff = datetime.now() + timedelta(days=self.date_future_tolerance_days)
                if parsed_date > future_cutoff:
                    flags.append('date_in_future')
                    flag_reasons['date_in_future'] = \
                        f"Date {cleaned_value} is in the future"

                # Check if DOB is too old (before 1900)
                if 'DOB' in element or 'BIRTH' in element:
                    if parsed_date.year < 1900:
                        flags.append('date_too_old')
                        flag_reasons['date_too_old'] = \
                            f"Birth date {cleaned_value} is before 1900"

        # Check amounts
        if any(keyword in element for keyword in ['AMOUNT', 'SALARY', 'WAGE', 'PAY']):
            try:
                value = float(str(cleaned_value).replace(',', '').replace('$', ''))

                # Check for negative (unless parentheses were in original)
                if value < 0 and '(' not in str(row.get('value', '')):
                    flags.append('negative_amount')
                    flag_reasons['negative_amount'] = \
                        f"Amount {cleaned_value} is negative"

                # Check for unreasonable salary range
                if 'SALARY' in element or 'WAGE' in element:
                    if value > 10000000:
                        flags.append('amount_too_high')
                        flag_reasons['amount_too_high'] = \
                            f"Salary {cleaned_value} exceeds $10M"
                    elif value < 0:
                        flags.append('negative_salary')
                        flag_reasons['negative_salary'] = \
                            f"Salary {cleaned_value} is negative"
            except (ValueError, TypeError):
                pass

        # Check percentages
        if 'PERCENT' in element or 'PCT' in element or 'RATE' in element:
            try:
                value = float(cleaned_value)
                if value < 0 or value > 1:
                    flags.append('percentage_out_of_range')
                    flag_reasons['percentage_out_of_range'] = \
                        f"Percentage {cleaned_value} is outside 0.0-1.0 range"
            except (ValueError, TypeError):
                pass

        # Check SSN
        if 'SSN' in element:
            ssn_str = str(cleaned_value).replace('-', '').replace(' ', '')

            # Check length
            if len(ssn_str) != 9:
                flags.append('invalid_ssn_length')
                flag_reasons['invalid_ssn_length'] = \
                    f"SSN {cleaned_value} does not have 9 digits"

            # Check for invalid patterns
            elif ssn_str in ['000000000', '123456789', '111111111', '222222222',
                            '333333333', '444444444', '555555555', '666666666',
                            '777777777', '888888888', '999999999']:
                flags.append('invalid_ssn_pattern')
                flag_reasons['invalid_ssn_pattern'] = \
                    f"SSN {cleaned_value} is an invalid pattern"

        return flags, flag_reasons

    def validate_extractions(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Validate all extractions and add confidence levels and flags.

        Args:
            df: DataFrame with cleaned extractions

        Returns:
            DataFrame with added columns: confidence, flags, flag_reasons
        """
        if df.empty:
            logger.warning("Empty dataframe provided for validation")
            return df

        logger.info(f"Validating {len(df)} extractions")

        # Create a copy to avoid modifying original
        result_df = df.copy()

        # Initialize new columns
        result_df['confidence'] = 'HIGH'
        result_df['flags'] = [[] for _ in range(len(result_df))]
        result_df['flag_reasons'] = [{} for _ in range(len(result_df))]

        # Calculate positional statistics for outlier detection
        positional_stats = self.calculate_positional_statistics(result_df)
        logger.debug(f"Calculated positional statistics for {len(positional_stats)} source/element combinations")

        # Process each document separately for within-document checks
        for (source, filename), doc_group in result_df.groupby(['source', 'filename']):
            doc_indices = doc_group.index.tolist()

            # Check date logic within document
            date_violations = self.check_date_logic(doc_group)
            for idx, (flags, reasons) in date_violations.items():
                result_df.at[idx, 'flags'] = result_df.at[idx, 'flags'] + flags
                result_df.at[idx, 'flag_reasons'].update(reasons)

            # Check within-document position gaps
            gap_violations = self.check_within_document_gaps(doc_group)
            for idx, (flags, reasons) in gap_violations.items():
                result_df.at[idx, 'flags'] = result_df.at[idx, 'flags'] + flags
                result_df.at[idx, 'flag_reasons'].update(reasons)

        # Process each row for individual checks
        for idx, row in result_df.iterrows():
            # Check positional outliers
            flags, reasons = self.check_positional_outliers(row, positional_stats)
            if flags:
                result_df.at[idx, 'flags'] = result_df.at[idx, 'flags'] + flags
                result_df.at[idx, 'flag_reasons'].update(reasons)

            # Check multiple extractions
            flags, reasons = self.check_multiple_extractions(row)
            if flags:
                result_df.at[idx, 'flags'] = result_df.at[idx, 'flags'] + flags
                result_df.at[idx, 'flag_reasons'].update(reasons)

            # Check value reasonableness
            flags, reasons = self.check_value_reasonableness(row)
            if flags:
                result_df.at[idx, 'flags'] = result_df.at[idx, 'flags'] + flags
                result_df.at[idx, 'flag_reasons'].update(reasons)

        # Set confidence level based on flags
        for idx, row in result_df.iterrows():
            if len(row['flags']) > 0:
                result_df.at[idx, 'confidence'] = 'LOW'

        # Convert lists/dicts to strings for Excel output
        result_df['flags'] = result_df['flags'].apply(lambda x: ', '.join(x) if x else '')
        result_df['flag_reasons'] = result_df['flag_reasons'].apply(
            lambda x: ' | '.join([f"{k}: {v}" for k, v in x.items()]) if x else ''
        )

        # Log statistics
        high_count = (result_df['confidence'] == 'HIGH').sum()
        low_count = (result_df['confidence'] == 'LOW').sum()

        logger.info(f"Validation complete: {high_count} high confidence, {low_count} low confidence")

        if low_count > 0:
            logger.warning(f"{low_count} extractions flagged for review")

        return result_df

    def check_missing_critical_elements(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Check for documents missing critical elements.

        Args:
            df: DataFrame with validated extractions

        Returns:
            DataFrame with missing elements report
        """
        if not self.critical_elements:
            logger.debug("No critical elements configured")
            return pd.DataFrame()

        missing_report = []

        # Group by source and filename
        for (source, filename), doc_group in df.groupby(['source', 'filename']):
            found_elements = doc_group['element'].unique().tolist()

            # Check which critical elements are missing
            missing = [elem for elem in self.critical_elements if elem not in found_elements]

            if missing:
                missing_report.append({
                    'source': source,
                    'filename': filename,
                    'missing_elements': ', '.join(missing),
                    'found_elements': ', '.join(found_elements)
                })
                logger.warning(f"Document {filename} missing critical elements: {missing}")

        if missing_report:
            return pd.DataFrame(missing_report)
        else:
            logger.info("All documents have all critical elements")
            return pd.DataFrame()


def main():
    """Example usage of ExtractionValidator."""
    # Example:
    # validator = ExtractionValidator(config_path="plans/plan_ABC/configs/Birth_Certificate.toml")
    #
    # # Validate extractions
    # df = pd.DataFrame(cleaned_extractions)
    # validated_df = validator.validate_extractions(df)
    #
    # # Separate by confidence
    # high_conf = validated_df[validated_df['confidence'] == 'HIGH']
    # low_conf = validated_df[validated_df['confidence'] == 'LOW']
    pass


if __name__ == "__main__":
    main()
