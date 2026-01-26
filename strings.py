"""
String Pattern Repository

All text-based extraction patterns (names, addresses, contact info, etc.).
Each pattern can be a single regex or a list of regex patterns.
"""

# Email addresses
EMAIL = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

# Phone numbers - various formats
PHONE = [
    r'\b(?:\+?1[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b',  # (555) 123-4567, 555-123-4567, etc.
    r'\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b',  # 555-123-4567 or 555.123.4567
    r'\b\d{10}\b',  # 5551234567 (no separators)
]

# Extension numbers
EXTENSION = [
    r'\b(?:ext\.?|extension)\s*\d{2,5}\b',
    r'\bx\d{2,5}\b',
]

# Fax numbers
FAX = PHONE

# ZIP codes
ZIP = [
    r'\b\d{5}-\d{4}\b',  # Extended: 12345-6789
    r'\b\d{5}\b',  # Basic: 12345
]

# State codes (US)
STATE = r'\b[A-Z]{2}\b'

# Full state names (example - customize as needed)
STATE_FULL = r'\b(?:Alabama|Alaska|Arizona|Arkansas|California|Colorado|Connecticut|Delaware|Florida|Georgia|Hawaii|Idaho|Illinois|Indiana|Iowa|Kansas|Kentucky|Louisiana|Maine|Maryland|Massachusetts|Michigan|Minnesota|Mississippi|Missouri|Montana|Nebraska|Nevada|New Hampshire|New Jersey|New Mexico|New York|North Carolina|North Dakota|Ohio|Oklahoma|Oregon|Pennsylvania|Rhode Island|South Carolina|South Dakota|Tennessee|Texas|Utah|Vermont|Virginia|Washington|West Virginia|Wisconsin|Wyoming)\b'

# City names (generic - may need customization)
CITY = r'\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\b'

# Street addresses
ADDRESS = [
    r'\b\d+\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\s+(?:Street|St\.?|Avenue|Ave\.?|Road|Rd\.?|Boulevard|Blvd\.?|Lane|Ln\.?|Drive|Dr\.?|Court|Ct\.?|Circle|Cir\.?|Way|Parkway|Pkwy\.?)\b',
    r'\b\d+\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\b',  # Simplified address
]

# Street types
STREET_TYPE = r'\b(?:Street|St\.?|Avenue|Ave\.?|Road|Rd\.?|Boulevard|Blvd\.?|Lane|Ln\.?|Drive|Dr\.?|Court|Ct\.?|Circle|Cir\.?|Way|Parkway|Pkwy\.?|Place|Pl\.?)\b'

# Apartment/Unit numbers
UNIT = [
    r'\b(?:Apt\.?|Apartment|Unit|Ste\.?|Suite)\s*#?\s*[A-Z0-9-]+\b',
    r'\b#\s*[A-Z0-9-]+\b',
]

# PO Box
PO_BOX = r'\b(?:P\.?O\.?\s*Box|PO\s*Box)\s+\d+\b'

# Website URLs
URL = [
    r'\b(?:https?://)?(?:www\.)?[a-zA-Z0-9-]+\.[a-zA-Z]{2,}(?:/[^\s]*)?\b',
    r'\b[a-zA-Z0-9-]+\.(?:com|org|net|edu|gov|mil|int|info|biz)\b',
]

# Department names (customize as needed)
DEPARTMENT = r'\b(?:HR|Human Resources|IT|Information Technology|Finance|Accounting|Sales|Marketing|Operations|Engineering|Legal|Administration|Admin|Customer Service|Production|R&D|Research and Development)\b'

# Job titles (common examples - customize as needed)
JOB_TITLE = r'\b(?:Manager|Director|Coordinator|Specialist|Analyst|Assistant|Supervisor|Representative|Rep|Executive|Officer|Administrator|Technician|Engineer|Developer|Designer|Consultant)\b'

# Gender
GENDER = [
    r'\b(?:Male|Female|M|F|Non-binary|Other)\b',
    r'\b[MF]\b',
]

# Marital status
MARITAL_STATUS = r'\b(?:Single|Married|Divorced|Widowed|Separated|Domestic Partnership)\b'

# Yes/No values
YES_NO = [
    r'\b(?:Yes|No|Y|N)\b',
    r'\b[YN]\b',
]

# True/False values
TRUE_FALSE = r'\b(?:True|False|T|F)\b'

# Status values (generic)
STATUS = r'\b(?:Active|Inactive|Pending|Approved|Denied|Complete|Incomplete|Open|Closed|Processing)\b'

# Employment status
EMPLOYMENT_STATUS = r'\b(?:Full[- ]?Time|Part[- ]?Time|Contract|Temporary|Permanent|Seasonal|Intern|Contractor)\b'

# Race/Ethnicity (for demographic forms - handle sensitively)
RACE_ETHNICITY = r'\b(?:White|Black|African American|Hispanic|Latino|Asian|Native American|Pacific Islander|Two or More Races|Other|Prefer not to answer)\b'

# Country names (common examples)
COUNTRY = r'\b(?:United States|USA|US|Canada|Mexico|United Kingdom|UK|Australia|Germany|France|Japan|China|India|Brazil)\b'

# Language codes
LANGUAGE = r'\b(?:English|Spanish|French|German|Chinese|Japanese|Korean|Arabic|Portuguese|Russian|Italian)\b'

# License plate (US format example)
LICENSE_PLATE = r'\b[A-Z]{2,3}\s*\d{3,4}\b'

# VIN (Vehicle Identification Number)
VIN = r'\b[A-HJ-NPR-Z0-9]{17}\b'

# Generic text field (any text between quotes)
QUOTED_TEXT = r'"([^"]*)"'

# Capitalized words (likely proper nouns)
PROPER_NOUN = r'\b[A-Z][a-z]+\b'

# All caps words
ALL_CAPS = r'\b[A-Z]{2,}\b'

# Signatures (text indicating signature)
SIGNATURE = r'\b(?:Signed|Signature|X)\b'

# Date signed text
DATE_SIGNED = r'\b(?:Date Signed|Signature Date|Signed On)\b'
