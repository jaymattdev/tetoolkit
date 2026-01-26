"""
Pattern Repository Package

Centralized pattern repository for text extraction.
Patterns are organized by category for easy maintenance.

Import patterns directly or use the PATTERNS dictionary.

Usage:
    from patterns import PATTERNS
    from patterns import dates, identifiers, amounts, strings
"""

# Import all pattern modules
from patterns import dates
from patterns import identifiers
from patterns import amounts
from patterns import strings

# Aggregate all patterns into a single dictionary
PATTERNS = {}

# Add all date patterns
for attr in dir(dates):
    if not attr.startswith('_'):
        PATTERNS[attr] = getattr(dates, attr)

# Add all identifier patterns
for attr in dir(identifiers):
    if not attr.startswith('_'):
        PATTERNS[attr] = getattr(identifiers, attr)

# Add all amount patterns
for attr in dir(amounts):
    if not attr.startswith('_'):
        PATTERNS[attr] = getattr(amounts, attr)

# Add all string patterns
for attr in dir(strings):
    if not attr.startswith('_'):
        PATTERNS[attr] = getattr(strings, attr)

# Export for convenience
__all__ = ['PATTERNS', 'dates', 'identifiers', 'amounts', 'strings']
