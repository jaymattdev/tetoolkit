# Text Extraction Tool

A modular Python-based tool for extracting structured data from text documents using configurable regex patterns and custom extraction rules.

## Features

- **Interactive Terminal UI**: User-friendly interface for easy setup and execution
- **OCR Text Cleaning**: Automatic correction of OCR errors and spell-checking of critical keywords using fuzzy matching
- **Excel-to-TOML Configuration**: Convert extraction specifications from Excel to TOML config files
- **Multiple Pattern Support**: Each element can have multiple regex patterns to account for format variations
- **Organized Pattern Repository**: 115+ patterns organized by category (dates, identifiers, amounts, strings)
- **Flexible Pattern Matching**: Uses regex patterns from a centralized pattern repository
- **Participant ID Extraction**: Automatically extract participant IDs from filenames (customizable)
- **Custom Name Extraction**: Anchor-based name extraction with configurable start/stop anchors
- **Configuration-Based Value Cleaning**: Simple config file to assign cleaners - 87 elements pre-configured
- **Value Cleaning & Parsing**: Standardize dates, amounts, names, and other extracted values
- **Validation & Confidence Flagging**: Separate high-confidence from low-confidence extractions based on logical consistency and positional analysis
- **Interactive Excel Reports**: Database-building reports with SharePoint links and filtering
- **Batch Processing**: Process multiple document types in a single run
- **Comprehensive Logging**: Debug, info, warning, and error logs for monitoring and troubleshooting
- **Comprehensive Statistics**: Track extraction success rates and missing data
- **Multiple Output Formats**: Save results as Excel (human-readable) and Pickle (machine-readable)
- **Text Files Only**: Supports .txt files only (no PDF or DOC support)

## Project Structure

```
extraction-tool/
├── run.py                       # ⭐ Interactive terminal UI (recommended)
├── main.py                      # Main execution script
├── config_generator.py          # Excel to TOML configuration generator
├── text_cleaner.py              # OCR error correction and spell-checking
├── extractor.py                 # Text extraction engine (with participant ID extraction)
├── orchestrator.py              # Multi-source extraction coordinator
├── value_cleaner.py             # Value parsing and cleaning
├── validator.py                 # Validation and confidence flagging
├── statistics_manager.py        # Comprehensive statistics calculation
├── report_generator.py          # Interactive Excel report generation
├── output_manager.py            # Output file management
├── patterns_example.py          # Example pattern repository (supports multiple patterns)
├── requirements.txt             # Python dependencies
│
├── docs/                        # 📚 Complete documentation
│   ├── README.md                # Documentation index
│   ├── modules/                 # Module-specific guides
│   │   ├── CLEANING_MODULE.md
│   │   ├── VALUE_CLEANING.md
│   │   ├── VALIDATION.md
│   │   ├── STATISTICS.md
│   │   └── REPORT_GUIDE.md
│   └── reference/               # Reference material
│       ├── FOLDER_STRUCTURE.md
│       ├── ORGANIZATION.md
│       ├── OUTPUT_EXAMPLE.md
│       └── SUMMARY.md
│
├── examples/                    # Example files
│   ├── extraction_specs_example.xlsx
│   └── patterns_example.py
│
└── plans/                       # ← All plans organized here
    ├── plan_ABC/                # ← Each plan is self-contained
    │   ├── extraction_specs.xlsx        # Plan-specific specs
    │   ├── Birth_Certificate/           # Document folders (only .txt files)
    │   │   └── cleaned/                 # Cleaned documents (auto-generated)
    │   ├── Employment_Application/
    │   ├── configs/                     # Plan-specific configs (auto-generated)
    │   ├── cleaning_reports/            # Cleaning change logs (auto-generated)
    │   ├── logs/                        # Execution logs (auto-generated)
    │   └── output/                      # Plan-specific results (auto-generated)
    │
    └── plan_XYZ/                # ← Another independent plan
        ├── extraction_specs.xlsx
        ├── Document_Type_A/
        ├── configs/
        └── output/
```

**Key Point:** Each plan folder contains its own `configs/` directory, ensuring configurations never get mixed up between different plans.

## Requirements

All packages are included in Anaconda base installation:
- pandas
- openpyxl
- toml
- jellyfish (for fuzzy string matching)
- logging
- pathlib
- re

## Installation

```bash
# If using Anaconda (recommended)
# All dependencies are already installed!

# If NOT using Anaconda, install requirements
pip install -r requirements.txt
```

No additional packages needed with Anaconda!

## Quick Start

### Method 1: Interactive Mode (Recommended for Teams)

The easiest way to run the extraction tool:

```bash
python run.py
```

The interactive interface will guide you through:
1. Selecting a plan
2. Choosing a pattern repository
3. Configuring options (cleaning, log level)
4. Validating your setup
5. Running the extraction

**Perfect for team members who just want to get started quickly!**

### Method 2: Command Line (For Automation)

```bash
# Standard workflow (with OCR cleaning)
python main.py --plan plan_ABC --patterns patterns_example.py

# Skip cleaning (extract from original docs)
python main.py --plan plan_ABC --patterns patterns_example.py --skip-cleaning

# Enable debug logging
python main.py --plan plan_ABC --patterns patterns_example.py --log-level DEBUG
```

---

## Setup Guide

### 1. Prepare Your Data Structure

Organize your plans folder like this (only .txt files supported):

```
plans/
└── plan_ABC/
    ├── extraction_specs.xlsx          # Your extraction specifications
    ├── Birth_Certificate/             # Document type folders (only .txt files)
    │   ├── doc1.txt
    │   ├── doc2.txt
    │   └── ...
    ├── Employment_Application/
    │   ├── app1.txt
    │   ├── app2.txt
    │   └── ...
    ├── Termination_Letter/
    │   ├── term1.txt
    │   └── ...
    ├── configs/                       # Auto-generated (don't create manually)
    └── output/                        # Auto-generated (don't create manually)
```

**Important:** Only `.txt` files are supported. PDF and DOC files are not supported.

### 2. Create Extraction Specifications

Create `extraction_specs.xlsx` in your plan folder with these columns:

| Source Name | Elements to Extract | Name Start Anchor | Name Stop Anchor | ID in File |
|-------------|---------------------|-------------------|------------------|------------|
| Birth_Certificate | ["DOB", "SSN"] | ["NAME", "FULL NAME"] | ["DATE OF BIRTH", "DOB"] | FALSE |
| Employment_Application | ["DOH", "SSN", "PHONE"] | ["APPLICANT NAME"] | ["DATE"] | TRUE |

**Location:** `plans/plan_ABC/extraction_specs.xlsx`

See `examples/extraction_specs_example.xlsx` for a complete example.

### 3. Configure Pattern Repository

Patterns are now organized in the `patterns/` directory by category for easy management:

```
patterns/
├── dates.py          # Date patterns (DOB, DOH, DOTE, etc.)
├── identifiers.py    # ID patterns (SSN, Employee ID, etc.)
├── amounts.py        # Monetary patterns (salary, amounts, etc.)
└── strings.py        # Text patterns (email, phone, address, etc.)
```

**Option 1: Use pre-built patterns (115+ patterns included)**

```python
from patterns_example import PATTERNS
# All patterns from all categories automatically loaded!
```

**Option 2: Add custom patterns to existing categories**

Edit the appropriate file in `patterns/`:

```python
# In patterns/dates.py
CUSTOM_DATE = r'\b\d{4}-\d{2}-\d{2}\b'  # Add your pattern

# In patterns/identifiers.py
BADGE_ID = r'\bBDG\d{6}\b'  # Add your pattern
```

**Option 3: Customize in patterns_example.py**

```python
from patterns_example import PATTERNS

# Add or override patterns
PATTERNS['CUSTOM_ID'] = r'\b[A-Z]{3}\d{6}\b'
PATTERNS['DOB'] = r'\b\d{2}/\d{2}/\d{4}\b'  # Override existing
```

**Each pattern can be a single string or list of patterns:**

```python
# Single pattern
EMAIL = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

# Multiple patterns (tries each until match found)
DOB = [
    r'\b(?:0?[1-9]|1[0-2])[/-](?:0?[1-9]|[12][0-9]|3[01])[/-](?:\d{2}|\d{4})\b',  # MM/DD/YYYY
    r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4}\b',  # Month DD, YYYY
]
```

See [patterns/README.md](patterns/README.md) for complete pattern documentation.

### 4. Customize Participant ID Extraction (Optional)

If you need to extract participant IDs from filenames, edit the `extract_id_from_filename` function in [extractor.py](extractor.py:99):

```python
def extract_id_from_filename(self, filename: str) -> Optional[str]:
    # Example: Extract SSN pattern (###-##-####)
    match = re.search(r'(\d{3}-\d{2}-\d{4})', filename)
    if match:
        return match.group(1)

    # Add your custom pattern here
    return None
```

## Workflow Details

### Step 1: Configuration Generation

The tool reads your Excel specifications and generates TOML configuration files for each document type. These are saved in `plans/{plan_name}/configs/` to keep everything organized per plan:

```toml
[Document]
document_source = "Birth_Certificate"
number_of_elements = 2
name_extraction = true
id_in_file = false

[Extraction]
active_patterns = ["DOB", "SSN"]
name_start_anchors = ["NAME", "FULL NAME"]
name_stop_anchors = ["DATE OF BIRTH", "DOB"]

[Output]
output_file_name = "Birth_Certificate_extractions.xlsx"
```

### Step 2: OCR Text Cleaning (Optional but Recommended)

Before extraction, documents are cleaned to fix OCR errors:

**Safe OCR Cleanup:**
- Removes zero-width spaces and invisible characters
- Normalizes smart quotes and special characters
- Fixes spacing issues around punctuation
- **Does NOT** do character substitutions (to prevent data corruption)

**Spell-Checking (Primary Cleaning Method):**
- Uses **fuzzy matching** (Jaro-Winkler similarity)
- Safely corrects misspelled keywords from config
- Example: `SOCI AL` → `SOCIAL`, `SECUR1TY` → `SECURITY`
- Preserves data integrity while fixing anchor text

**Output:**
- Cleaned documents saved to `{source_folder}/cleaned/`
- Detailed change log in `cleaning_reports/`
- All changes tracked for review

See [CLEANING_MODULE.md](CLEANING_MODULE.md) for full documentation.

**Skip cleaning:**
```bash
python main.py --plan plan_ABC --patterns patterns.py --skip-cleaning
```

### Step 3: Pattern-Based Extraction

For each document (using cleaned version if available):
- Extracts all elements using regex patterns from your pattern repository
- Uses `finditer()` to capture **all instances** of each element
- Records extraction order and position in the document

### Step 4: Name Extraction

For documents configured with name extraction:
- Creates dynamic regex patterns using start/stop anchors
- Extracts capitalized name sequences between anchors
- Example: Between "NAME:" and "DATE OF BIRTH", extract "John Michael Smith"

### Step 5: Missing Data Handling

If an element is not found in a document:
- Still creates a row in the output
- Sets value, extraction_order, and extraction_position to `None`
- Allows tracking of missing data in statistics

### Step 6: Statistics Calculation

Generates comprehensive statistics:
- **Summary**: Documents processed, total elements, found/not found counts
- **Detailed**: Per-element breakdown by document source

### Step 7: Output Generation

Creates two timestamped output files in `plans/{plan_name}/output/`:

1. **Excel file**: `{plan_name}_extractions_{YYYYMMDD}_{HHMMSS}.xlsx`
   - Sheet 1: Raw Extractions (all extracted data)
   - Sheet 2: Summary Statistics (by document source)
   - Sheet 3: Detailed Statistics (by element type)

2. **Pickle file**: `{plan_name}_extractions_{YYYYMMDD}_{HHMMSS}.pkl`
   - Raw Python list of dictionaries for downstream processing

**Example filenames:**
- `plan_ABC_extractions_20260125_143022.xlsx`
- `plan_ABC_extractions_20260125_143022.pkl`

## Output Format

### Raw Extractions DataFrame

| source | filename | element | value | extraction_order | extraction_position |
|--------|----------|---------|-------|------------------|---------------------|
| Birth_Certificate | doc1.txt | DOB | 01/15/1990 | 1 | 245 |
| Birth_Certificate | doc1.txt | SSN | 123-45-6789 | 1 | 312 |
| Birth_Certificate | doc1.txt | NAME | John Smith | 1 | 45 |
| Birth_Certificate | doc2.txt | DOB | None | None | None |

### Summary Statistics

| Source | Documents Processed | Total Elements | Found | Not Found | Found % |
|--------|---------------------|----------------|-------|-----------|---------|
| Birth_Certificate | 50 | 150 | 142 | 8 | 94.67 |
| Employment_Application | 30 | 90 | 88 | 2 | 97.78 |

## Advanced Features

### Custom ID Extraction from Filename

If `id_in_file` is set to `true` in the Excel specs, the tool expects you to implement a custom function for extracting IDs from filenames. Add this to `extractor.py`:

```python
def extract_id_from_filename(self, filename):
    """Extract ID from filename - customize as needed."""
    # Example: "participant_12345_form.txt" -> "12345"
    import re
    match = re.search(r'_(\d+)_', filename)
    return match.group(1) if match else None
```

### Adding New Pattern Types

Simply add patterns to your pattern repository:

```python
PATTERNS = {
    # ... existing patterns
    'CUSTOM_FIELD': r'your-regex-pattern-here',
}
```

Then add "CUSTOM_FIELD" to the "Elements to Extract" column in your Excel specs.

### Processing Non-Text Files

The extractor supports multiple file types. To add PDF or DOCX support:

1. Install additional packages (included in Anaconda):
   ```bash
   conda install -c conda-forge PyPDF2 python-docx
   ```

2. Add text extraction logic to `extractor.py`:
   ```python
   def load_text_file(self, file_path):
       ext = Path(file_path).suffix.lower()

       if ext == '.pdf':
           # Add PDF extraction logic
           pass
       elif ext == '.docx':
           # Add DOCX extraction logic
           pass
       else:
           # Default text file handling
           with open(file_path, 'r') as f:
               return f.read()
   ```

## Troubleshooting

### No extractions found
- Check that your pattern repository is correctly loaded
- Verify regex patterns match your document format
- Enable case-insensitive matching if needed

### Name extraction not working
- Ensure start and stop anchors exist in the document
- Check that anchors are spelled exactly as they appear
- Verify there's text between the anchors

### Missing statistics
- Ensure `calculate_statistics()` is called before saving output
- Check that extractions list is not empty

## Future Enhancements

Consider adding:
- Multi-line pattern support for complex extractions
- Confidence scores for extractions
- Validation rules for extracted data
- Machine learning-based extraction for unstructured names
- Support for image-based documents (OCR)

## License

This tool uses only standard Python and Anaconda packages - no external dependencies required!

## Support

For issues or questions, refer to the inline documentation in each module or contact your team's data engineering support.
