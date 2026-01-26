"""
Date Pattern Repository

All date-related extraction patterns.
Each pattern can be a single regex or a list of regex patterns.
"""

# Date of Birth patterns
DOB = [
    r'\b(?:0?[1-9]|1[0-2])[/-](?:0?[1-9]|[12][0-9]|3[01])[/-](?:\d{2}|\d{4})\b',  # MM/DD/YYYY or MM-DD-YYYY
    r'\b(?:0?[1-9]|[12][0-9]|3[01])[/-](?:0?[1-9]|1[0-2])[/-](?:\d{2}|\d{4})\b',  # DD/MM/YYYY or DD-MM-YYYY
    r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4}\b',  # Month DD, YYYY
    r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b',  # Full month name
]

# Date of Hire patterns
DOH = [
    r'\b(?:0?[1-9]|1[0-2])[/-](?:0?[1-9]|[12][0-9]|3[01])[/-](?:\d{2}|\d{4})\b',  # MM/DD/YYYY
    r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4}\b',  # Month DD, YYYY
    r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b',
]

# Date of Termination/End patterns
DOTE = [
    r'\b(?:0?[1-9]|1[0-2])[/-](?:0?[1-9]|[12][0-9]|3[01])[/-](?:\d{2}|\d{4})\b',  # MM/DD/YYYY
    r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4}\b',  # Month DD, YYYY
    r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b',
]

# Generic date pattern (use for any date field)
DATE = [
    r'\b(?:0?[1-9]|1[0-2])[/-](?:0?[1-9]|[12][0-9]|3[01])[/-](?:\d{2}|\d{4})\b',  # MM/DD/YYYY
    r'\b(?:0?[1-9]|[12][0-9]|3[01])[/-](?:0?[1-9]|1[0-2])[/-](?:\d{2}|\d{4})\b',  # DD/MM/YYYY
    r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4}\b',  # Month DD, YYYY
    r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b',
    r'\b\d{4}[/-](?:0?[1-9]|1[0-2])[/-](?:0?[1-9]|[12][0-9]|3[01])\b',  # YYYY/MM/DD or YYYY-MM-DD (ISO format)
]

# Effective date
EFFECTIVE_DATE = DATE

# Start date
START_DATE = DATE

# End date
END_DATE = DATE

# Application date
APPLICATION_DATE = DATE

# Issue date
ISSUE_DATE = DATE

# Expiration date
EXPIRATION_DATE = DATE

# Year only (for specific use cases)
YEAR = r'\b(?:19|20)\d{2}\b'

# Month and year only (e.g., "March 2023", "03/2023")
MONTH_YEAR = [
    r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(?:19|20)\d{2}\b',
    r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+(?:19|20)\d{2}\b',
    r'\b(?:0?[1-9]|1[0-2])[/-](?:19|20)\d{2}\b',
]
