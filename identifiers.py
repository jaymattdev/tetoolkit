"""
Identifier Pattern Repository

All ID-based extraction patterns (SSN, employee IDs, reference numbers, etc.).
Each pattern can be a single regex or a list of regex patterns.
"""

# Social Security Number - various formats
SSN = [
    r'\b\d{3}-\d{2}-\d{4}\b',  # With dashes: 123-45-6789
    r'\b\d{3}\s+\d{2}\s+\d{4}\b',  # With spaces: 123 45 6789
    r'\b\d{9}\b',  # No separators: 123456789 (use cautiously - may match other numbers)
]

# Employee ID - various formats
EMPLOYEE_ID = [
    r'\b(?:EMP|EMPL)[-:]?\s*\d{4,6}\b',  # EMP-12345, EMPL:12345
    r'\b(?:ID|EMPLOYEE\s*ID)[-:]?\s*\d{4,6}\b',  # ID-12345, Employee ID: 12345
    r'\b[A-Z]{2,3}\d{4,6}\b',  # XX12345 (two or three letter prefix)
    r'\bE\d{5,6}\b',  # E12345
]

# Staff ID
STAFF_ID = EMPLOYEE_ID

# Personnel ID
PERSONNEL_ID = EMPLOYEE_ID

# Payroll ID
PAYROLL_ID = EMPLOYEE_ID

# Member ID
MEMBER_ID = [
    r'\b(?:MEM|MEMBER)[-:]?\s*\d{4,8}\b',
    r'\bM\d{5,8}\b',
]

# Customer ID
CUSTOMER_ID = [
    r'\b(?:CUST|CUSTOMER)[-:]?\s*\d{4,8}\b',
    r'\bC\d{5,8}\b',
]

# Participant ID (for studies, programs, etc.)
PARTICIPANT_ID = [
    r'\b(?:PART|PARTICIPANT)[-:]?\s*\d{4,8}\b',
    r'\bP\d{5,8}\b',
]

# Student ID
STUDENT_ID = [
    r'\b(?:STU|STUDENT)[-:]?\s*\d{4,8}\b',
    r'\bS\d{5,8}\b',
]

# Tax ID / EIN (Employer Identification Number)
TAX_ID = [
    r'\b\d{2}-\d{7}\b',  # 12-3456789
    r'\b\d{9}\b',
]

EIN = TAX_ID

# Driver's License (US format - varies by state)
DRIVERS_LICENSE = [
    r'\b[A-Z]\d{7,8}\b',  # A1234567
    r'\b\d{8,9}\b',  # 12345678
    r'\b[A-Z]{2}\d{6}\b',  # AB123456
]

# Passport number
PASSPORT = [
    r'\b[A-Z]{1,2}\d{6,9}\b',  # US format
    r'\b\d{9}\b',
]

# Medical Record Number (MRN)
MRN = [
    r'\b(?:MRN|MEDICAL RECORD)[-:]?\s*\d{6,10}\b',
    r'\bMR\d{6,10}\b',
]

# Insurance ID
INSURANCE_ID = [
    r'\b[A-Z]{3}\d{9}\b',  # ABC123456789
    r'\b[A-Z0-9]{8,15}\b',  # Alphanumeric
]

# Policy number
POLICY_NUMBER = [
    r'\b(?:POL|POLICY)[-:]?\s*[A-Z0-9]{6,12}\b',
    r'\b[A-Z0-9]{8,12}\b',
]

# Group number
GROUP_NUMBER = [
    r'\b(?:GRP|GROUP)[-:]?\s*[A-Z0-9]{4,10}\b',
    r'\b[A-Z0-9]{6,10}\b',
]

# Badge number
BADGE_NUMBER = [
    r'\b(?:BADGE|BADGE #)[-:]?\s*\d{4,6}\b',
    r'\b\d{4,6}\b',
]

# Case number
CASE_NUMBER = [
    r'\b(?:CASE|CASE #)[-:]?\s*[A-Z0-9]{6,12}\b',
    r'\b[A-Z]{2}\d{8}\b',
]

# Reference number
REFERENCE_NUMBER = [
    r'\b(?:REF|REFERENCE)[-:]?\s*[A-Z0-9]{6,12}\b',
    r'\b[A-Z0-9]{8,12}\b',
]

# Confirmation number
CONFIRMATION_NUMBER = [
    r'\b(?:CONF|CONFIRMATION)[-:]?\s*[A-Z0-9]{6,12}\b',
    r'\b[A-Z0-9]{6,10}\b',
]

# Tracking number
TRACKING_NUMBER = [
    r'\b(?:TRACK|TRACKING)[-:]?\s*[A-Z0-9]{8,20}\b',
    r'\b[A-Z0-9]{10,20}\b',
]

# Order number
ORDER_NUMBER = [
    r'\b(?:ORDER|ORD)[-:]?\s*[A-Z0-9]{6,12}\b',
    r'\b[A-Z0-9]{8,12}\b',
]

# Transaction ID
TRANSACTION_ID = [
    r'\b(?:TXN|TRANS|TRANSACTION)[-:]?\s*[A-Z0-9]{8,16}\b',
    r'\b[A-Z0-9]{10,16}\b',
]

# Claim number
CLAIM_NUMBER = [
    r'\b(?:CLAIM|CLM)[-:]?\s*[A-Z0-9]{6,12}\b',
    r'\b[A-Z0-9]{8,12}\b',
]

# Application number
APPLICATION_NUMBER = [
    r'\b(?:APP|APPLICATION)[-:]?\s*[A-Z0-9]{6,12}\b',
    r'\b[A-Z0-9]{8,12}\b',
]

# Certificate number
CERTIFICATE_NUMBER = [
    r'\b(?:CERT|CERTIFICATE)[-:]?\s*[A-Z0-9]{6,12}\b',
    r'\b[A-Z0-9]{8,12}\b',
]

# License number (professional licenses)
LICENSE_NUMBER = [
    r'\b(?:LIC|LICENSE)[-:]?\s*[A-Z0-9]{6,12}\b',
    r'\b[A-Z]{2}\d{6}\b',
]

# Permit number
PERMIT_NUMBER = [
    r'\b(?:PERMIT|PER)[-:]?\s*[A-Z0-9]{6,12}\b',
    r'\b[A-Z0-9]{6,10}\b',
]

# Registration number
REGISTRATION_NUMBER = [
    r'\b(?:REG|REGISTRATION)[-:]?\s*[A-Z0-9]{6,12}\b',
    r'\b[A-Z0-9]{6,10}\b',
]

# Serial number
SERIAL_NUMBER = [
    r'\b(?:SN|SERIAL|S/N)[-:]?\s*[A-Z0-9]{6,16}\b',
    r'\b[A-Z0-9]{8,16}\b',
]

# UUID (Universally Unique Identifier)
UUID = r'\b[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}\b'

# GUID (Global Unique Identifier) - same as UUID
GUID = UUID

# Barcode (generic)
BARCODE = r'\b\d{12,14}\b'  # UPC/EAN format

# ISBN (International Standard Book Number)
ISBN = [
    r'\b(?:ISBN[-:]?\s*)?(?:\d{9}[\dX]|\d{13})\b',  # ISBN-10 or ISBN-13
]

# Credit card (basic pattern - use cautiously for security reasons)
CREDIT_CARD = r'\b\d{4}[-\s]?\d{4}[-\s]?\d{4}[-\s]?\d{4}\b'

# Routing number (US bank)
ROUTING_NUMBER = r'\b\d{9}\b'

# Swift code (international bank)
SWIFT_CODE = r'\b[A-Z]{6}[A-Z0-9]{2}(?:[A-Z0-9]{3})?\b'

# IBAN (International Bank Account Number)
IBAN = r'\b[A-Z]{2}\d{2}[A-Z0-9]{4,30}\b'

# IP Address
IP_ADDRESS = r'\b(?:\d{1,3}\.){3}\d{1,3}\b'

# MAC Address
MAC_ADDRESS = r'\b(?:[0-9A-Fa-f]{2}[:-]){5}[0-9A-Fa-f]{2}\b'

# Version number
VERSION = r'\bv?\d+\.\d+(?:\.\d+)?(?:\.\d+)?\b'
