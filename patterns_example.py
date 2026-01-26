"""
Pattern Repository - Main Import File

This file imports and aggregates all patterns from the organized pattern modules.

Pattern Organization:
- patterns/dates.py       - All date-related patterns
- patterns/identifiers.py - All ID and reference number patterns
- patterns/amounts.py     - All monetary and numeric patterns
- patterns/strings.py     - All text-based patterns (names, addresses, etc.)

Usage:
    from patterns_example import PATTERNS

    # Or import specific categories
    from patterns import dates, identifiers, amounts, strings

Each pattern can be:
  - A single regex string
  - A list of regex strings (tried in order until match found)
"""

# Import from the patterns package
from patterns import PATTERNS

# Re-export for backward compatibility
__all__ = ['PATTERNS']

# You can also import individual categories if needed:
# from patterns import dates, identifiers, amounts, strings
#
# Example customization - add your own patterns:
# PATTERNS['CUSTOM_DATE'] = r'\b\d{4}-\d{2}-\d{2}\b'
# PATTERNS['CUSTOM_ID'] = [
#     r'\b[A-Z]{3}\d{6}\b',
#     r'\b\d{8}\b'
# ]
#
# Or modify existing patterns:
# PATTERNS['SSN'] = r'\b\d{3}-\d{2}-\d{4}\b'  # Only accept dashed format
