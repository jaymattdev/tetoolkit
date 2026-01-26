# Pattern Organization Guide

## Overview

Patterns are now organized into separate files by category, making them much easier to manage, especially when you have lots of patterns.

## New Structure

```
extraction-tool/
├── patterns/                    # ← Pattern repository directory
│   ├── __init__.py             # Aggregates all patterns
│   ├── README.md               # Pattern documentation
│   ├── dates.py                # 12 date patterns
│   ├── identifiers.py          # 42 identifier patterns
│   ├── amounts.py              # 30 amount/number patterns
│   └── strings.py              # 31 text patterns
│
└── patterns_example.py         # Main import file (115 total patterns)
```

## Benefits

### 1. **Easy to Find Patterns**
Instead of scrolling through one huge file, patterns are organized by type:
- Need a date pattern? → `patterns/dates.py`
- Need an ID pattern? → `patterns/identifiers.py`
- Need a dollar amount? → `patterns/amounts.py`
- Need a phone or email? → `patterns/strings.py`

### 2. **Easy to Maintain**
- Each category in its own file
- Add/modify patterns without searching through hundreds of lines
- Clear organization reduces errors

### 3. **Reusable**
- Import specific categories when needed
- Share pattern files between projects
- Build your own custom categories

### 4. **No Breaking Changes**
- Existing code continues to work
- Same `PATTERNS` dictionary
- Same usage pattern

## Quick Reference

### Total Patterns Available: 115

**Dates (12 patterns):**
- DOB, DOH, DOTE, DATE, EFFECTIVE_DATE, START_DATE, END_DATE, APPLICATION_DATE, ISSUE_DATE, EXPIRATION_DATE, YEAR, MONTH_YEAR

**Identifiers (42 patterns):**
- SSN, EMPLOYEE_ID, STAFF_ID, PERSONNEL_ID, PAYROLL_ID, MEMBER_ID, CUSTOMER_ID, PARTICIPANT_ID, STUDENT_ID, TAX_ID, EIN, DRIVERS_LICENSE, PASSPORT, MRN, INSURANCE_ID, POLICY_NUMBER, GROUP_NUMBER, BADGE_NUMBER, CASE_NUMBER, REFERENCE_NUMBER, CONFIRMATION_NUMBER, TRACKING_NUMBER, ORDER_NUMBER, TRANSACTION_ID, CLAIM_NUMBER, APPLICATION_NUMBER, CERTIFICATE_NUMBER, LICENSE_NUMBER, PERMIT_NUMBER, REGISTRATION_NUMBER, SERIAL_NUMBER, UUID, GUID, BARCODE, ISBN, CREDIT_CARD, ROUTING_NUMBER, SWIFT_CODE, IBAN, IP_ADDRESS, MAC_ADDRESS, VERSION

**Amounts (30 patterns):**
- AMOUNT, SALARY, WAGE, PAY_RATE, BONUS, COMMISSION, BENEFITS, COMPENSATION, HOURLY_RATE, ANNUAL_SALARY, PERCENTAGE, INTEREST_RATE, TAX_RATE, DISCOUNT, NUMBER, INTEGER, DECIMAL, NEGATIVE_NUMBER, NUMBER_RANGE, CURRENCY_CODE, ACCOUNT_NUMBER, CHECK_NUMBER, INVOICE_NUMBER, PO_NUMBER, QUANTITY, HOURS, DAYS, WEEKS, MONTHS, YEARS_DURATION

**Strings (31 patterns):**
- EMAIL, PHONE, EXTENSION, FAX, ZIP, STATE, STATE_FULL, CITY, ADDRESS, STREET_TYPE, UNIT, PO_BOX, URL, DEPARTMENT, JOB_TITLE, GENDER, MARITAL_STATUS, YES_NO, TRUE_FALSE, STATUS, EMPLOYMENT_STATUS, RACE_ETHNICITY, COUNTRY, LANGUAGE, LICENSE_PLATE, VIN, QUOTED_TEXT, PROPER_NOUN, ALL_CAPS, SIGNATURE, DATE_SIGNED

## Usage Examples

### Standard Usage (No Changes Required)

```python
from patterns_example import PATTERNS

# Use PATTERNS dictionary as before
dob = PATTERNS['DOB']
ssn = PATTERNS['SSN']
email = PATTERNS['EMAIL']
```

### Import Specific Categories

```python
from patterns import dates, identifiers, amounts, strings

# Access patterns from specific categories
birth_date = dates.DOB
employee_id = identifiers.EMPLOYEE_ID
salary = amounts.SALARY
phone = strings.PHONE
```

### Browse Available Patterns

```python
from patterns import dates

# See all available date patterns
for name in dir(dates):
    if not name.startswith('_'):
        print(f"{name}: {getattr(dates, name)}")
```

## Adding Custom Patterns

### Option 1: Add to Existing Category

Edit the appropriate file in `patterns/`:

```python
# File: patterns/dates.py

# Add your custom date pattern
CUSTOM_DATE = r'\b\d{4}-\d{2}-\d{2}\b'

# Or with multiple patterns
FLEXIBLE_DATE = [
    r'\b\d{2}/\d{2}/\d{4}\b',
    r'\b\d{4}-\d{2}-\d{2}\b',
]
```

### Option 2: Override in patterns_example.py

```python
from patterns_example import PATTERNS

# Add new patterns
PATTERNS['MY_CUSTOM_ID'] = r'\b[A-Z]{3}\d{6}\b'

# Override existing patterns
PATTERNS['SSN'] = r'\b\d{3}-\d{2}-\d{4}\b'  # Only dashed format
```

### Option 3: Create New Category

1. Create `patterns/custom.py`
2. Add your patterns
3. Import in `patterns/__init__.py`

## Migration from Old Structure

If you had custom patterns in the old `patterns_example.py`:

### Before (Old Structure)
```python
PATTERNS = {
    'DOB': r'...',
    'SSN': r'...',
    'CUSTOM_PATTERN': r'...',
    # ... hundreds of patterns
}
```

### After (New Structure)
```python
# Most patterns are now in organized files
from patterns_example import PATTERNS

# Add your custom patterns
PATTERNS['CUSTOM_PATTERN'] = r'...'

# Or add them to appropriate category file
# In patterns/identifiers.py:
# CUSTOM_PATTERN = r'...'
```

## Best Practices

1. **Add patterns to the right category:**
   - Dates → `dates.py`
   - IDs/Numbers → `identifiers.py`
   - Money/Amounts → `amounts.py`
   - Text/Contact → `strings.py`

2. **Use descriptive names:**
   ```python
   # Good
   EMPLOYEE_BADGE_ID = r'\bBDG\d{6}\b'

   # Bad
   ID1 = r'\bBDG\d{6}\b'
   ```

3. **Comment complex patterns:**
   ```python
   COMPLEX_DATE = [
       r'\b\d{2}/\d{2}/\d{4}\b',  # MM/DD/YYYY
       r'\b\w+\s+\d{1,2},\s+\d{4}\b',  # Month DD, YYYY
   ]
   ```

4. **Test before deploying:**
   ```python
   import re
   pattern = r'\b\d{3}-\d{2}-\d{4}\b'
   test = "SSN: 123-45-6789"
   match = re.search(pattern, test)
   print(match.group() if match else "No match")
   ```

## Troubleshooting

### "Pattern not found" error
- Check spelling (case-sensitive)
- Verify pattern exists in category file
- Ensure `__init__.py` imports the category

### Import errors
- Make sure `patterns/__init__.py` exists
- Check for syntax errors in pattern files
- Try `python -c "from patterns_example import PATTERNS; print(len(PATTERNS))"`

### Pattern not matching
- Test at [regex101.com](https://regex101.com)
- Check for special characters needing escape
- Try making it the first pattern in a list

## Performance Notes

- **No performance impact:** Patterns are only loaded once at import
- **Memory efficient:** Only patterns you use are accessed
- **Fast lookups:** Dictionary-based access is O(1)

## Summary

✅ **115 pre-built patterns** organized by category
✅ **Easy to find** what you need
✅ **Easy to maintain** with separate files
✅ **Backward compatible** - existing code works
✅ **Extensible** - add your own easily

See [patterns/README.md](patterns/README.md) for detailed documentation on each pattern.
