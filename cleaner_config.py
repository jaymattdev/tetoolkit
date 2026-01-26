"""
Value Cleaner Configuration

Define which elements use which cleaning functions.
This makes it easy to manage cleaner assignments without touching the core logic.
"""

# Cleaner type assignments
# Format: 'ELEMENT_NAME': 'cleaner_type'
#
# Available cleaner types:
#   'date'       - Parse and format dates (MM/DD/YYYY with 2-year logic)
#   'dollar'     - Clean dollar amounts ($1,234.56 -> 1234.56)
#   'percentage' - Convert percentages (50% -> 0.5)
#   'decimal'    - Clean decimal numbers (preserve precision)
#   'string'     - Clean strings (uppercase, remove invalid chars)
#   'name'       - Parse names into FNAME/LNAME components
#   'passthrough'- Keep as-is (no cleaning)

CLEANER_ASSIGNMENTS = {
    # ===== DATE CLEANERS =====
    # All date-related elements
    'DOB': 'date',
    'DOH': 'date',
    'DOTE': 'date',
    'DATE': 'date',
    'EFFECTIVE_DATE': 'date',
    'START_DATE': 'date',
    'END_DATE': 'date',
    'APPLICATION_DATE': 'date',
    'ISSUE_DATE': 'date',
    'EXPIRATION_DATE': 'date',
    'BIRTH_DATE': 'date',
    'HIRE_DATE': 'date',
    'TERMINATION_DATE': 'date',
    'DATE_SIGNED': 'date',

    # ===== DOLLAR AMOUNT CLEANERS =====
    # Monetary values
    'AMOUNT': 'dollar',
    'SALARY': 'dollar',
    'WAGE': 'dollar',
    'PAY_RATE': 'dollar',
    'BONUS': 'dollar',
    'COMMISSION': 'dollar',
    'BENEFITS': 'dollar',
    'COMPENSATION': 'dollar',
    'HOURLY_RATE': 'dollar',
    'ANNUAL_SALARY': 'dollar',

    # ===== PERCENTAGE CLEANERS =====
    # Percentage values
    'PERCENTAGE': 'percentage',
    'INTEREST_RATE': 'percentage',
    'TAX_RATE': 'percentage',
    'DISCOUNT': 'percentage',

    # ===== DECIMAL CLEANERS =====
    # Numeric values (preserve precision)
    'NUMBER': 'decimal',
    'HOURS': 'decimal',
    'DAYS': 'decimal',
    'WEEKS': 'decimal',
    'MONTHS': 'decimal',
    'YEARS_DURATION': 'decimal',
    'QUANTITY': 'decimal',

    # ===== STRING CLEANERS =====
    # IDs and structured strings
    'SSN': 'string',
    'EMPLOYEE_ID': 'string',
    'STAFF_ID': 'string',
    'PERSONNEL_ID': 'string',
    'PAYROLL_ID': 'string',
    'MEMBER_ID': 'string',
    'CUSTOMER_ID': 'string',
    'PARTICIPANT_ID': 'string',
    'STUDENT_ID': 'string',
    'TAX_ID': 'string',
    'EIN': 'string',
    'DRIVERS_LICENSE': 'string',
    'PASSPORT': 'string',
    'MRN': 'string',
    'INSURANCE_ID': 'string',
    'POLICY_NUMBER': 'string',
    'GROUP_NUMBER': 'string',
    'BADGE_NUMBER': 'string',
    'CASE_NUMBER': 'string',
    'REFERENCE_NUMBER': 'string',
    'CONFIRMATION_NUMBER': 'string',
    'TRACKING_NUMBER': 'string',
    'ORDER_NUMBER': 'string',
    'TRANSACTION_ID': 'string',
    'CLAIM_NUMBER': 'string',
    'APPLICATION_NUMBER': 'string',
    'CERTIFICATE_NUMBER': 'string',
    'LICENSE_NUMBER': 'string',
    'PERMIT_NUMBER': 'string',
    'REGISTRATION_NUMBER': 'string',
    'SERIAL_NUMBER': 'string',

    # Contact info
    'PHONE': 'string',
    'FAX': 'string',
    'EMAIL': 'passthrough',  # Keep email as-is
    'ZIP': 'string',
    'STATE': 'string',

    # ===== NAME CLEANERS =====
    # Names are handled specially
    'NAME': 'name',
    'FULL_NAME': 'name',
    'APPLICANT_NAME': 'name',
    'EMPLOYEE_NAME': 'name',

    # ===== PASSTHROUGH =====
    # Keep these as-is
    'URL': 'passthrough',
    'ADDRESS': 'passthrough',
    'NOTES': 'passthrough',
}


# Keyword-based cleaner assignment
# If element name contains these keywords, assign this cleaner
# This provides a fallback for elements not explicitly listed above
KEYWORD_ASSIGNMENTS = {
    'date': ['DATE', 'DOB', 'DOH', 'DOTE', 'BIRTH', 'HIRE', 'TERMINATION', 'EFFECTIVE', 'EXPIRATION', 'ISSUE'],
    'dollar': ['AMOUNT', 'SALARY', 'WAGE', 'PAY', 'BONUS', 'COMMISSION', 'COMPENSATION', 'BENEFIT'],
    'percentage': ['PERCENT', 'PCT', 'RATE', '%'],
    'decimal': ['NUMBER', 'HOURS', 'DAYS', 'WEEKS', 'MONTHS', 'QUANTITY', 'COUNT'],
    'name': ['NAME'],
    'string': ['SSN', 'PHONE', 'ZIP', 'ID', 'NUMBER', 'CODE'],
}


def get_cleaner_type(element: str) -> str:
    """
    Determine which cleaner to use for an element.

    Args:
        element: Element name (e.g., 'DOB', 'SSN', 'SALARY')

    Returns:
        Cleaner type ('date', 'dollar', 'percentage', 'decimal', 'string', 'name', 'passthrough')
    """
    element_upper = element.upper()

    # Check explicit assignments first
    if element_upper in CLEANER_ASSIGNMENTS:
        return CLEANER_ASSIGNMENTS[element_upper]

    # Fall back to keyword matching
    for cleaner_type, keywords in KEYWORD_ASSIGNMENTS.items():
        if any(keyword in element_upper for keyword in keywords):
            return cleaner_type

    # Default to string cleaner
    return 'string'


def add_cleaner_assignment(element: str, cleaner_type: str):
    """
    Add a custom cleaner assignment.

    Args:
        element: Element name
        cleaner_type: Type of cleaner to use

    Example:
        add_cleaner_assignment('CUSTOM_DATE', 'date')
        add_cleaner_assignment('BADGE_ID', 'string')
    """
    CLEANER_ASSIGNMENTS[element.upper()] = cleaner_type


def remove_cleaner_assignment(element: str):
    """
    Remove a cleaner assignment (will fall back to keyword matching).

    Args:
        element: Element name
    """
    element_upper = element.upper()
    if element_upper in CLEANER_ASSIGNMENTS:
        del CLEANER_ASSIGNMENTS[element_upper]


def get_all_assignments() -> dict:
    """
    Get all current cleaner assignments.

    Returns:
        Dictionary of all cleaner assignments
    """
    return CLEANER_ASSIGNMENTS.copy()


def print_cleaner_summary():
    """Print a summary of cleaner assignments by type."""
    summary = {}
    for element, cleaner_type in CLEANER_ASSIGNMENTS.items():
        if cleaner_type not in summary:
            summary[cleaner_type] = []
        summary[cleaner_type].append(element)

    print("Cleaner Assignment Summary:")
    print("=" * 60)
    for cleaner_type, elements in sorted(summary.items()):
        print(f"\n{cleaner_type.upper()} ({len(elements)} elements):")
        for element in sorted(elements):
            print(f"  - {element}")
    print("=" * 60)


if __name__ == "__main__":
    # Example usage
    print_cleaner_summary()

    # Test get_cleaner_type
    print("\nExamples:")
    test_elements = ['DOB', 'SSN', 'SALARY', 'CUSTOM_ELEMENT']
    for elem in test_elements:
        cleaner = get_cleaner_type(elem)
        print(f"  {elem} -> {cleaner}")
