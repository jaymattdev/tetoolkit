# Cleaner Configuration Guide

## Overview

The value cleaner now uses a **configuration-based approach** where you simply specify which elements use which cleaners. No more digging through if/else chains!

## How It Works

### Before (Old Approach)
```python
# Had to modify value_cleaner.py code
if 'DATE' in element or 'DOB' in element or 'DOH' in element:
    return self.clean_date(raw_value)
elif 'AMOUNT' in element or 'SALARY' in element:
    return self.clean_dollar_amount(raw_value)
# ... many more conditions
```

### After (New Approach)
```python
# Just edit cleaner_config.py
CLEANER_ASSIGNMENTS = {
    'DOB': 'date',
    'SALARY': 'dollar',
    'SSN': 'string',
    # ... easy to read and maintain!
}
```

## Available Cleaner Types

| Cleaner Type | Description | Example Input | Example Output |
|--------------|-------------|---------------|----------------|
| `date` | Parse and format dates | `01/15/90` | `01/15/1990` |
| `dollar` | Clean dollar amounts | `$1,234.56` | `1234.56` |
| `percentage` | Convert percentages | `50%` | `0.5` |
| `decimal` | Clean decimal numbers | `123.456` | `123.456` |
| `string` | Clean strings (uppercase, remove invalid chars) | `abc-123` | `ABC-123` |
| `name` | Parse names into FNAME/LNAME | `John Smith` | `FNAME: JOHN, LNAME: SMITH` |
| `passthrough` | Keep as-is (no cleaning) | `test@example.com` | `test@example.com` |

## Configuration File: cleaner_config.py

### Explicit Assignments

Define exact element-to-cleaner mappings:

```python
CLEANER_ASSIGNMENTS = {
    # Dates
    'DOB': 'date',
    'DOH': 'date',
    'DOTE': 'date',

    # Dollar amounts
    'SALARY': 'dollar',
    'WAGE': 'dollar',
    'BONUS': 'dollar',

    # IDs and strings
    'SSN': 'string',
    'EMPLOYEE_ID': 'string',

    # Passthroughs
    'EMAIL': 'passthrough',
    'URL': 'passthrough',
}
```

### Keyword-Based Fallback

For elements not explicitly listed, keywords are matched:

```python
KEYWORD_ASSIGNMENTS = {
    'date': ['DATE', 'DOB', 'DOH', 'BIRTH', 'HIRE'],
    'dollar': ['AMOUNT', 'SALARY', 'WAGE', 'PAY'],
    'string': ['SSN', 'PHONE', 'ID', 'NUMBER'],
    # ...
}
```

**Example:**
- `CUSTOM_DATE_FIELD` → matches keyword `DATE` → uses `date` cleaner
- `PAY_RATE` → matches keyword `PAY` → uses `dollar` cleaner

## Adding Custom Elements

### Option 1: Edit cleaner_config.py

```python
# Add your element to CLEANER_ASSIGNMENTS
CLEANER_ASSIGNMENTS = {
    # ... existing assignments ...

    # Your custom elements
    'BADGE_ID': 'string',
    'REVIEW_DATE': 'date',
    'DISCOUNT_AMOUNT': 'dollar',
}
```

### Option 2: Programmatically Add Assignments

```python
from cleaner_config import add_cleaner_assignment

# Add custom assignments at runtime
add_cleaner_assignment('CUSTOM_DATE', 'date')
add_cleaner_assignment('BADGE_NUMBER', 'string')
```

## Pre-Configured Elements

### Dates (14 elements)
```
DOB, DOH, DOTE, DATE, EFFECTIVE_DATE, START_DATE, END_DATE,
APPLICATION_DATE, ISSUE_DATE, EXPIRATION_DATE, BIRTH_DATE,
HIRE_DATE, TERMINATION_DATE, DATE_SIGNED
```

### Dollar Amounts (10 elements)
```
AMOUNT, SALARY, WAGE, PAY_RATE, BONUS, COMMISSION,
BENEFITS, COMPENSATION, HOURLY_RATE, ANNUAL_SALARY
```

### Percentages (4 elements)
```
PERCENTAGE, INTEREST_RATE, TAX_RATE, DISCOUNT
```

### Decimals (7 elements)
```
NUMBER, HOURS, DAYS, WEEKS, MONTHS, YEARS_DURATION, QUANTITY
```

### Strings (45 elements)
```
SSN, EMPLOYEE_ID, STAFF_ID, PERSONNEL_ID, PAYROLL_ID,
MEMBER_ID, CUSTOMER_ID, PARTICIPANT_ID, STUDENT_ID,
TAX_ID, EIN, DRIVERS_LICENSE, PASSPORT, MRN, INSURANCE_ID,
POLICY_NUMBER, GROUP_NUMBER, BADGE_NUMBER, CASE_NUMBER,
REFERENCE_NUMBER, CONFIRMATION_NUMBER, TRACKING_NUMBER,
ORDER_NUMBER, TRANSACTION_ID, CLAIM_NUMBER, APPLICATION_NUMBER,
CERTIFICATE_NUMBER, LICENSE_NUMBER, PERMIT_NUMBER,
REGISTRATION_NUMBER, SERIAL_NUMBER, PHONE, FAX, ZIP, STATE
```

### Names (4 elements)
```
NAME, FULL_NAME, APPLICANT_NAME, EMPLOYEE_NAME
```

### Passthroughs (3 elements)
```
EMAIL, URL, ADDRESS, NOTES
```

**Total: 87 pre-configured elements**

## Usage Examples

### Basic Usage (Automatic)

The system works automatically when you run extractions:

```python
from value_cleaner import ValueCleaner

cleaner = ValueCleaner()

# Clean a date
result = cleaner.clean_value('01/15/90', 'DOB')
# Returns: '01/15/1990'

# Clean a salary
result = cleaner.clean_value('$50,000', 'SALARY')
# Returns: '50000.00'

# Clean an SSN
result = cleaner.clean_value('123-45-6789', 'SSN')
# Returns: '123-45-6789'
```

### Check Cleaner Assignment

```python
from cleaner_config import get_cleaner_type

# Check what cleaner will be used
cleaner_type = get_cleaner_type('DOB')
print(cleaner_type)  # 'date'

cleaner_type = get_cleaner_type('CUSTOM_DATE_FIELD')
print(cleaner_type)  # 'date' (matches 'DATE' keyword)
```

### View All Assignments

```python
from cleaner_config import print_cleaner_summary

# Print a summary of all assignments
print_cleaner_summary()
```

Output:
```
Cleaner Assignment Summary:
============================================================

DATE (14 elements):
  - APPLICATION_DATE
  - BIRTH_DATE
  - DATE
  - DOB
  ...

DOLLAR (10 elements):
  - AMOUNT
  - ANNUAL_SALARY
  - BENEFITS
  ...
```

### Add Custom Assignment

```python
from cleaner_config import add_cleaner_assignment

# Add a custom element
add_cleaner_assignment('REVIEW_DATE', 'date')
add_cleaner_assignment('PROMO_CODE', 'string')

# Now these will use the specified cleaners
```

## Cleaner Behaviors

### Date Cleaner
- Parses various date formats
- Outputs: `MM/DD/YYYY`
- **2-year logic:** If year > 2027, subtracts 100
  - `01/15/90` → `01/15/1990`
  - `01/15/35` → `01/15/1935` (2035 > 2027, so subtract 100)

### Dollar Cleaner
- Removes `$`, `,`, whitespace
- Handles negative amounts (parentheses or minus)
- Always outputs 2 decimal places
  - `$1,234.56` → `1234.56`
  - `($500)` → `-500.00`

### Percentage Cleaner
- Converts to decimal (divides by 100)
  - `50%` → `0.5`
  - `12.5%` → `0.125`

### Decimal Cleaner
- Removes commas
- Preserves original precision
  - `1,234.567` → `1234.567`
  - `100` → `100`

### String Cleaner
- Removes invalid characters (keeps letters, numbers, spaces, `-`, `'`)
- Converts to uppercase
- Normalizes whitespace
  - `abc-123 XYZ` → `ABC-123 XYZ`
  - `Test  Value` → `TEST VALUE`

### Name Cleaner
- Parses into FNAME/LNAME components
- **Creates separate rows** for each name component
- Handles BENEFICIARY (BFNAME/BLNAME) and SPOUSE (SFNAME/SLNAME)
- Supports reverse order (`Last, First`)
- **Original raw value preserved in each component row**

**Example:**
```
Input:  1 row: element='NAME', value='John Smith'
Output: 2 rows:
        - element='FNAME', value='John Smith', cleaned_value='JOHN'
        - element='LNAME', value='John Smith', cleaned_value='SMITH'
```

**Beneficiary Example:**
```
Input:  1 row: element='NAME', value='Jane Doe', start_anchor='BENEFICIARY'
Output: 2 rows:
        - element='BFNAME', value='Jane Doe', cleaned_value='JANE'
        - element='BLNAME', value='Jane Doe', cleaned_value='DOE'
```

### Passthrough Cleaner
- Keeps value as-is
- Only strips whitespace
  - `test@example.com` → `test@example.com`

## Smart Detection

Dollar and percentage cleaners have smart detection:

```python
# Element is configured as 'dollar', but value has %
cleaner.clean_value('50%', 'SALARY')
# Detects % and uses percentage cleaner → '0.5'

# Element is configured as 'dollar', value has $
cleaner.clean_value('$1,000', 'SALARY')
# Uses dollar cleaner → '1000.00'

# Element is configured as 'dollar', no special chars
cleaner.clean_value('1000', 'SALARY')
# Uses decimal cleaner → '1000'
```

## Troubleshooting

### Element not cleaning correctly

1. Check what cleaner is assigned:
   ```python
   from cleaner_config import get_cleaner_type
   print(get_cleaner_type('YOUR_ELEMENT'))
   ```

2. Add explicit assignment if needed:
   ```python
   # In cleaner_config.py
   CLEANER_ASSIGNMENTS['YOUR_ELEMENT'] = 'date'  # or other type
   ```

### Want to change default behavior

Edit [cleaner_config.py](cleaner_config.py):

```python
# Change existing assignment
CLEANER_ASSIGNMENTS['SSN'] = 'passthrough'  # Keep SSN as-is

# Or remove to use keyword matching
del CLEANER_ASSIGNMENTS['SSN']
```

### Custom cleaner not recognized

Make sure cleaner type is valid:
- `date`, `dollar`, `percentage`, `decimal`, `string`, `name`, `passthrough`

Invalid types will default to `string` cleaner.

## Benefits

✅ **Easy to maintain** - All assignments in one place
✅ **No code changes** - Just edit configuration
✅ **Clear mapping** - See exactly what uses what
✅ **Extensible** - Add new elements easily
✅ **Keyword fallback** - Handles unexpected elements gracefully
✅ **87 pre-configured elements** - Ready to use

## Migration from Old System

The new system is **100% backward compatible**. No changes needed to existing code - it just works better now!

If you had custom logic in the old `clean_value` method, move it to [cleaner_config.py](cleaner_config.py):

**Before:**
```python
# Custom logic buried in value_cleaner.py
if 'CUSTOM' in element:
    return self.custom_clean(raw_value)
```

**After:**
```python
# Clear configuration in cleaner_config.py
CLEANER_ASSIGNMENTS['CUSTOM_ELEMENT'] = 'string'
```

## See Also

- [value_cleaner.py](value_cleaner.py) - Core cleaning logic
- [cleaner_config.py](cleaner_config.py) - Configuration file
- [docs/modules/VALUE_CLEANING.md](docs/modules/VALUE_CLEANING.md) - Detailed cleaning documentation
