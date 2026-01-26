"""
Amount and Number Pattern Repository

All monetary amounts, percentages, and numeric value patterns.
Each pattern can be a single regex or a list of regex patterns.
"""

# Dollar amounts - various formats
AMOUNT = [
    r'\$\s*\d+(?:,\d{3})*(?:\.\d{2})?',  # $1,234.56
    r'\$\s*\d+(?:\.\d{2})?',  # $1234.56 (no commas)
    r'\b\d+(?:,\d{3})*(?:\.\d{2})?\s*(?:USD|dollars?)\b',  # 1234.56 USD or dollars
]

# Salary amounts
SALARY = AMOUNT

# Wage amounts
WAGE = AMOUNT

# Pay rate
PAY_RATE = AMOUNT

# Bonus amounts
BONUS = AMOUNT

# Commission amounts
COMMISSION = AMOUNT

# Benefits amount
BENEFITS = AMOUNT

# Compensation
COMPENSATION = AMOUNT

# Hourly rate
HOURLY_RATE = [
    r'\$\s*\d+(?:\.\d{2})?\s*(?:/hr|per hour|hourly)',
    r'\$\s*\d+(?:\.\d{2})?',
]

# Annual salary
ANNUAL_SALARY = [
    r'\$\s*\d+(?:,\d{3})*(?:\.\d{2})?\s*(?:/yr|per year|annually|annual)',
    r'\$\s*\d+(?:,\d{3})*(?:\.\d{2})?',
]

# Percentage values
PERCENTAGE = [
    r'\b\d+(?:\.\d{1,2})?\s*%',  # 50% or 50.5%
    r'\b\d+(?:\.\d{1,2})?\s*percent',  # 50 percent
]

# Interest rate
INTEREST_RATE = PERCENTAGE

# Tax rate
TAX_RATE = PERCENTAGE

# Discount percentage
DISCOUNT = PERCENTAGE

# Generic numeric values
NUMBER = r'\b\d+(?:,\d{3})*(?:\.\d+)?\b'

# Integer only
INTEGER = r'\b\d+\b'

# Decimal number
DECIMAL = r'\b\d+\.\d+\b'

# Negative numbers
NEGATIVE_NUMBER = r'-\s*\d+(?:,\d{3})*(?:\.\d+)?'

# Range of numbers (e.g., "25-30", "100 to 150")
NUMBER_RANGE = [
    r'\b\d+\s*-\s*\d+\b',  # 25-30
    r'\b\d+\s+to\s+\d+\b',  # 25 to 30
    r'\b\d+\s*–\s*\d+\b',  # 25–30 (em dash)
]

# Currency codes (for international amounts)
CURRENCY_CODE = r'\b[A-Z]{3}\b'  # USD, EUR, GBP, etc.

# Account numbers (generic)
ACCOUNT_NUMBER = r'\b\d{8,16}\b'

# Check numbers
CHECK_NUMBER = r'\b\d{4,8}\b'

# Invoice numbers
INVOICE_NUMBER = [
    r'\b(?:INV|INVOICE)[#-]?\s*\d{4,10}\b',
    r'\b\d{4,10}\b',
]

# PO (Purchase Order) numbers
PO_NUMBER = [
    r'\b(?:PO|P\.O\.|PURCHASE ORDER)[#-]?\s*\d{4,10}\b',
    r'\b\d{4,10}\b',
]

# Quantity
QUANTITY = [
    r'\b\d+(?:,\d{3})*\s*(?:units?|pcs?|pieces?|items?)',
    r'\b\d+(?:,\d{3})*\b',
]

# Hours (work hours, project hours, etc.)
HOURS = [
    r'\b\d+(?:\.\d{1,2})?\s*(?:hrs?|hours?)',
    r'\b\d+(?:\.\d{1,2})?\b',
]

# Days (PTO, sick days, etc.)
DAYS = [
    r'\b\d+\s*(?:days?)',
    r'\b\d+\b',
]

# Weeks
WEEKS = [
    r'\b\d+\s*(?:weeks?|wks?)',
    r'\b\d+\b',
]

# Months (duration)
MONTHS = [
    r'\b\d+\s*(?:months?|mos?)',
    r'\b\d+\b',
]

# Years (duration)
YEARS_DURATION = [
    r'\b\d+\s*(?:years?|yrs?)',
    r'\b\d+\b',
]
