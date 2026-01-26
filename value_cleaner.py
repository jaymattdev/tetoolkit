"""
Value Cleaning and Parsing Module
Cleans and parses raw extracted values into standardized formats.

This module:
1. Parses dates to MM/DD/YYYY format with 2-year date logic
2. Cleans amounts (dollars, percentages, decimals)
3. Cleans strings (remove invalid chars, uppercase)
4. Parses names into FNAME/LNAME with special logic for BENEFICIARY/SPOUSE

Configuration-based cleaner assignment:
- Cleaner assignments are defined in cleaner_config.py
- Easy to add/modify which elements use which cleaners
- Supports both explicit assignments and keyword-based fallbacks
"""

import re
import logging
from datetime import datetime
from typing import Optional, Dict, Tuple
import pandas as pd
from cleaner_config import get_cleaner_type

# Configure module logger
logger = logging.getLogger(__name__)


class ValueCleaner:
    """Clean and parse raw extracted values."""

    def __init__(self):
        """Initialize the value cleaner."""
        logger.info("Initialized ValueCleaner")

    def clean_date(self, raw_value: str) -> Optional[str]:
        """
        Clean and parse date to MM/DD/YYYY format.

        2-year date logic: If resolved year > 2027, subtract 100 years.

        Args:
            raw_value: Raw date string

        Returns:
            Cleaned date in MM/DD/YYYY format or None if parsing fails
        """
        if not raw_value or pd.isna(raw_value):
            return None

        try:
            # Remove extra whitespace
            date_str = str(raw_value).strip()

            # Common date separators
            date_str = re.sub(r'[/\-\.]', '/', date_str)

            # Try various date formats
            formats = [
                '%m/%d/%Y',    # 01/15/1990
                '%m/%d/%y',    # 01/15/90
                '%d/%m/%Y',    # 15/01/1990
                '%d/%m/%y',    # 15/01/90
                '%Y/%m/%d',    # 1990/01/15
                '%y/%m/%d',    # 90/01/15
                '%B %d, %Y',   # January 15, 1990
                '%b %d, %Y',   # Jan 15, 1990
                '%B %d %Y',    # January 15 1990
                '%b %d %Y',    # Jan 15 1990
                '%d %B %Y',    # 15 January 1990
                '%d %b %Y',    # 15 Jan 1990
            ]

            parsed_date = None
            for fmt in formats:
                try:
                    parsed_date = datetime.strptime(date_str, fmt)
                    break
                except ValueError:
                    continue

            if not parsed_date:
                logger.warning(f"Could not parse date: {raw_value}")
                return None

            # Apply 2-year date logic
            if parsed_date.year > 2027:
                parsed_date = parsed_date.replace(year=parsed_date.year - 100)
                logger.debug(f"Applied 2-year logic: {raw_value} -> {parsed_date.year}")

            # Format as MM/DD/YYYY with leading zeros
            cleaned = parsed_date.strftime('%m/%d/%Y')
            logger.debug(f"Cleaned date: {raw_value} -> {cleaned}")
            return cleaned

        except Exception as e:
            logger.error(f"Error cleaning date '{raw_value}': {e}")
            return None

    def clean_dollar_amount(self, raw_value: str) -> Optional[str]:
        """
        Clean dollar amount to decimal format.

        $100.50 -> 100.50
        $1,234.56 -> 1234.56

        Args:
            raw_value: Raw dollar amount string

        Returns:
            Cleaned amount (no $, always 2 decimal places) or None
        """
        if not raw_value or pd.isna(raw_value):
            return None

        try:
            # Remove dollar sign, commas, and whitespace
            amount_str = str(raw_value).strip()
            amount_str = re.sub(r'[\$,\s]', '', amount_str)

            # Remove any parentheses (negative amounts)
            is_negative = '(' in str(raw_value) or '-' in amount_str
            amount_str = re.sub(r'[()]', '', amount_str)

            # Convert to float
            amount = float(amount_str)

            # Apply negative if needed
            if is_negative:
                amount = -amount

            # Format with 2 decimal places
            cleaned = f"{amount:.2f}"
            logger.debug(f"Cleaned dollar amount: {raw_value} -> {cleaned}")
            return cleaned

        except Exception as e:
            logger.error(f"Error cleaning dollar amount '{raw_value}': {e}")
            return None

    def clean_percentage(self, raw_value: str) -> Optional[str]:
        """
        Clean percentage to decimal format.

        100% -> 1.0
        50% -> 0.5
        12.5% -> 0.125

        Args:
            raw_value: Raw percentage string

        Returns:
            Decimal format or None
        """
        if not raw_value or pd.isna(raw_value):
            return None

        try:
            # Remove % sign and whitespace
            pct_str = str(raw_value).strip()
            pct_str = re.sub(r'[%\s]', '', pct_str)

            # Convert to float and divide by 100
            pct_value = float(pct_str) / 100.0

            cleaned = str(pct_value)
            logger.debug(f"Cleaned percentage: {raw_value} -> {cleaned}")
            return cleaned

        except Exception as e:
            logger.error(f"Error cleaning percentage '{raw_value}': {e}")
            return None

    def clean_decimal(self, raw_value: str) -> Optional[str]:
        """
        Clean decimal number (preserve decimal places).

        123.456 -> 123.456
        100 -> 100

        Args:
            raw_value: Raw decimal string

        Returns:
            Cleaned decimal (preserves original precision) or None
        """
        if not raw_value or pd.isna(raw_value):
            return None

        try:
            # Remove commas and whitespace
            decimal_str = str(raw_value).strip()
            decimal_str = re.sub(r'[,\s]', '', decimal_str)

            # Convert to float then back to string to normalize
            decimal_value = float(decimal_str)

            # Preserve original decimal places - don't round unnecessarily
            cleaned = str(decimal_value)
            logger.debug(f"Cleaned decimal: {raw_value} -> {cleaned}")
            return cleaned

        except Exception as e:
            logger.error(f"Error cleaning decimal '{raw_value}': {e}")
            return None

    def clean_string(self, raw_value: str) -> Optional[str]:
        """
        Clean string value.

        - Remove invalid characters
        - Uppercase the value

        Args:
            raw_value: Raw string

        Returns:
            Cleaned uppercase string or None
        """
        if not raw_value or pd.isna(raw_value):
            return None

        try:
            # Convert to string and strip
            string_val = str(raw_value).strip()

            # Remove characters that normally wouldn't belong
            # Keep: letters, numbers, spaces, basic punctuation
            cleaned = re.sub(r'[^\w\s\-\']', '', string_val)

            # Uppercase
            cleaned = cleaned.upper()

            # Remove excessive whitespace
            cleaned = re.sub(r'\s+', ' ', cleaned).strip()

            if not cleaned:
                return None

            logger.debug(f"Cleaned string: {raw_value} -> {cleaned}")
            return cleaned

        except Exception as e:
            logger.error(f"Error cleaning string '{raw_value}': {e}")
            return None

    def parse_name(self, raw_value: str, start_anchor: str = "", reverse_order: bool = False) -> Dict[str, Optional[str]]:
        """
        Parse name into first and last name components.

        Logic:
        - First word -> FNAME, remaining -> LNAME
        - If reverse_order=True: "Smith, John" -> FNAME: John, LNAME: Smith
        - If start_anchor contains "BENEFICIARY": BFNAME, BLNAME
        - If start_anchor contains "SPOUSE": SFNAME, SLNAME
        - Otherwise: FNAME, LNAME

        Args:
            raw_value: Raw name string
            start_anchor: The start anchor used in extraction
            reverse_order: If True, expect "Last, First" format

        Returns:
            Dictionary with name components (e.g., {'FNAME': 'JOHN', 'LNAME': 'SMITH'})
        """
        if not raw_value or pd.isna(raw_value):
            return {}

        try:
            # Clean the name first
            name_str = str(raw_value).strip()

            # Remove extra whitespace
            name_str = re.sub(r'\s+', ' ', name_str)

            # Uppercase
            name_str = name_str.upper()

            # Determine prefix based on start anchor
            prefix = ""
            if start_anchor:
                anchor_upper = start_anchor.upper()
                if "BENEFICIARY" in anchor_upper or "BENEF" in anchor_upper:
                    prefix = "B"
                    logger.debug(f"Detected beneficiary name: {start_anchor}")
                elif "SPOUSE" in anchor_upper:
                    prefix = "S"
                    logger.debug(f"Detected spouse name: {start_anchor}")

            # Parse name
            if reverse_order and ',' in name_str:
                # Reverse order: "Smith, John Michael"
                parts = name_str.split(',', 1)
                last_name = parts[0].strip()
                first_name = parts[1].strip() if len(parts) > 1 else ""
            else:
                # Normal order: "John Michael Smith"
                parts = name_str.split()
                if len(parts) == 0:
                    return {}
                elif len(parts) == 1:
                    # Only one name - treat as first name
                    first_name = parts[0]
                    last_name = ""
                else:
                    # First word is first name, rest is last name
                    first_name = parts[0]
                    last_name = ' '.join(parts[1:])

            # Build result dictionary
            result = {}
            if first_name:
                result[f'{prefix}FNAME'] = first_name
            if last_name:
                result[f'{prefix}LNAME'] = last_name

            logger.debug(f"Parsed name: {raw_value} -> {result}")
            return result

        except Exception as e:
            logger.error(f"Error parsing name '{raw_value}': {e}")
            return {}

    def clean_value(self, raw_value: str, element: str, start_anchor: str = "",
                    reverse_name_order: bool = False) -> Optional[str]:
        """
        Clean a value based on its element type using configuration-based cleaner assignment.

        Cleaner type is determined from cleaner_config.py based on element name.

        Args:
            raw_value: Raw extracted value
            element: Element type (DOB, SSN, NAME, etc.)
            start_anchor: Start anchor used in extraction (for name parsing)
            reverse_name_order: If True, expect reverse name order

        Returns:
            Cleaned value or None if parsing fails
        """
        if not raw_value or pd.isna(raw_value):
            return None

        try:
            # Get cleaner type from configuration
            cleaner_type = get_cleaner_type(element)

            logger.debug(f"Element '{element}' assigned cleaner type: '{cleaner_type}'")

            # Route to appropriate cleaner based on config
            if cleaner_type == 'date':
                return self.clean_date(raw_value)

            elif cleaner_type == 'dollar':
                # Smart detection: check if value has % or $
                if '%' in str(raw_value):
                    return self.clean_percentage(raw_value)
                elif '$' in str(raw_value):
                    return self.clean_dollar_amount(raw_value)
                else:
                    return self.clean_decimal(raw_value)

            elif cleaner_type == 'percentage':
                if '%' in str(raw_value):
                    return self.clean_percentage(raw_value)
                else:
                    return self.clean_decimal(raw_value)

            elif cleaner_type == 'decimal':
                return self.clean_decimal(raw_value)

            elif cleaner_type == 'name':
                # Names are handled in clean_extractions_dataframe
                return str(raw_value).strip().upper()

            elif cleaner_type == 'passthrough':
                # Keep as-is, just strip whitespace
                return str(raw_value).strip()

            elif cleaner_type == 'string':
                return self.clean_string(raw_value)

            else:
                # Unknown cleaner type - default to string
                logger.warning(f"Unknown cleaner type '{cleaner_type}' for element '{element}', defaulting to string cleaner")
                return self.clean_string(raw_value)

        except Exception as e:
            logger.error(f"Error cleaning value '{raw_value}' for element '{element}': {e}")
            return None

    def clean_extractions_dataframe(self, df: pd.DataFrame, reverse_name_order: bool = False) -> pd.DataFrame:
        """
        Clean all values in an extractions dataframe.

        For NAME elements, creates separate rows for each name component (FNAME, LNAME, BFNAME, etc.)
        with the original raw value preserved in each row.

        Example:
            Input:  1 row with element='NAME', value='John Smith'
            Output: 2 rows:
                    - element='FNAME', value='John Smith', cleaned_value='JOHN'
                    - element='LNAME', value='John Smith', cleaned_value='SMITH'

        Args:
            df: DataFrame with raw extractions
            reverse_name_order: If True, expect reverse name order for names

        Returns:
            DataFrame with cleaned values (names expanded into separate rows)
        """
        if df.empty:
            logger.warning("Empty dataframe provided for cleaning")
            return df

        logger.info(f"Cleaning {len(df)} extracted values")

        # List to collect all rows (original + name component rows)
        all_rows = []

        # Clean each row
        for idx, row in df.iterrows():
            element = row.get('element', '')
            raw_value = row.get('value')
            start_anchor = row.get('start_anchor', '')

            # Handle names separately - create multiple rows
            if 'NAME' in element.upper() and raw_value and not pd.isna(raw_value):
                name_components = self.parse_name(raw_value, start_anchor, reverse_name_order)

                if name_components:
                    # Create a separate row for each name component
                    for component_name, component_value in name_components.items():
                        # Create new row based on original
                        new_row = row.copy()
                        new_row['element'] = component_name  # FNAME, LNAME, BFNAME, etc.
                        new_row['value'] = raw_value  # Keep original raw value
                        new_row['cleaned_value'] = component_value  # Parsed component
                        all_rows.append(new_row)

                    logger.debug(f"Expanded NAME '{raw_value}' into {len(name_components)} components: {list(name_components.keys())}")
                else:
                    # Name parsing failed - keep original row with no cleaned value
                    new_row = row.copy()
                    new_row['cleaned_value'] = None
                    all_rows.append(new_row)
            else:
                # Clean other types - single row
                new_row = row.copy()
                if raw_value and not pd.isna(raw_value):
                    cleaned = self.clean_value(raw_value, element, start_anchor, reverse_name_order)
                    new_row['cleaned_value'] = cleaned
                else:
                    new_row['cleaned_value'] = None
                all_rows.append(new_row)

        # Create result dataframe from all rows
        result_df = pd.DataFrame(all_rows)

        # Reset index
        result_df = result_df.reset_index(drop=True)

        # Log statistics
        original_count = len(df)
        result_count = len(result_df)
        cleaned_count = result_df['cleaned_value'].notna().sum()
        failed_count = result_count - cleaned_count
        name_expansions = result_count - original_count

        logger.info(f"Cleaning complete: {original_count} input rows -> {result_count} output rows")
        if name_expansions > 0:
            logger.info(f"  {name_expansions} additional rows from name component expansion")
        logger.info(f"  {cleaned_count}/{result_count} values cleaned, {failed_count} failed")

        if failed_count > 0:
            logger.warning(f"{failed_count} values could not be cleaned")

        return result_df


def main():
    """Example usage of ValueCleaner."""
    # Example:
    # cleaner = ValueCleaner()
    #
    # # Test date cleaning
    # print(cleaner.clean_date("01/15/90"))  # 01/15/1990
    # print(cleaner.clean_date("01/15/35"))  # 01/15/1935 (2035 > 2027, subtract 100)
    #
    # # Test amount cleaning
    # print(cleaner.clean_dollar_amount("$1,234.56"))  # 1234.56
    # print(cleaner.clean_percentage("50%"))  # 0.5
    #
    # # Test name parsing
    # print(cleaner.parse_name("John Michael Smith"))  # {'FNAME': 'JOHN', 'LNAME': 'MICHAEL SMITH'}
    # print(cleaner.parse_name("Smith, John", reverse_order=True))  # {'FNAME': 'JOHN', 'LNAME': 'SMITH'}
    pass


if __name__ == "__main__":
    main()
