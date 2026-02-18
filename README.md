# Text Extraction Tool

A modular Python-based tool for extracting structured data from text documents using configurable regex patterns and custom extraction rules.

## Features

- **Interactive Terminal UI**: User-friendly interface with workflow mode selection
- **Master Configuration**: Centralized settings via `master_config.toml` for SharePoint, cleaning, validation, and output options
- **OCR Text Cleaning**: Automatic correction of OCR errors and spell-checking of critical keywords using fuzzy matching
- **Excel-to-TOML Configuration**: Convert extraction specifications from Excel to TOML config files
- **Multiple Pattern Support**: Each element can have multiple regex patterns to account for format variations
- **Organized Pattern Repository**: 115+ patterns organized by category (dates, identifiers, amounts, strings)
- **Pre-Compiled Pattern Support**: Works with both raw string patterns and pre-compiled `re.compile()` patterns
- **Duplicate Mappings**: Automatically rename second occurrences of elements (e.g., DOB -> SDOB for spouse)
- **Name Prefixes**: Configurable name extraction with prefix support (Name, Spouse, Beneficiary, AP)
- **Participant ID Extraction**: Automatically extract participant IDs from filenames (customizable)
- **Custom Name Extraction**: Anchor-based name extraction with configurable start/stop anchors
- **Value Cleaning & Parsing**: Standardize dates, amounts, names, and other extracted values
- **Validation & Confidence Flagging**: Separate high-confidence from low-confidence extractions
- **Interactive Excel Reports**: Database-building reports with SharePoint links and filtering
- **Workflow Modes**: Full extraction, extract-only, or report-only modes
- **Batch Processing**: Process multiple document types in a single run
- **Comprehensive Statistics**: Track extraction success rates and missing data
- **Multiple Output Formats**: Save results as Excel (human-readable) and Pickle (machine-readable)

## Project Structure

```
extraction-tool/
├── run.py                       # Interactive terminal UI (recommended)
├── main.py                      # Main execution script
├── master_config.toml           # Global configuration settings
├── config_loader.py             # Master configuration loader
├── config_generator.py          # Excel to TOML configuration generator
├── text_cleaner.py              # OCR error correction and spell-checking
├── extractor.py                 # Text extraction engine
├── orchestrator.py              # Multi-source extraction coordinator
├── value_cleaner.py             # Value parsing and cleaning
├── validator.py                 # Validation and confidence flagging
├── statistics_manager.py        # Comprehensive statistics calculation
├── report_generator.py          # Interactive Excel report generation
├── output_manager.py            # Output file management
├── patterns_example.py          # Example pattern repository
├── requirements.txt             # Python dependencies
│
├── docs/                        # Complete documentation
│   ├── README.md                # Documentation index
│   ├── SOP.md                   # Standard Operating Procedures
│   ├── modules/                 # Module-specific guides
│   └── reference/               # Reference material
│
├── patterns/                    # Pattern repository
│   ├── dates.py                 # Date patterns
│   ├── identifiers.py           # ID patterns
│   ├── amounts.py               # Monetary patterns
│   └── strings.py               # Text patterns
│
└── plans/                       # All plans organized here
    └── plan_ABC/                # Each plan is self-contained
        ├── extraction_specs.xlsx        # Plan-specific specs
        ├── Clusters/                    # Document type folders
        │   ├── Birth_Certificate/       # Document folders (.txt files)
        │   │   └── cleaned/             # Cleaned documents (auto-generated)
        │   └── Employment_Application/
        ├── configs/                     # Plan-specific configs (auto-generated)
        ├── cleaning_reports/            # Cleaning change logs (auto-generated)
        └── output/                      # Plan-specific results (auto-generated)
```

## Requirements

All packages are included in Anaconda base installation:
- pandas
- openpyxl
- toml
- jellyfish (for fuzzy string matching)

## Installation

```bash
# If using Anaconda (recommended) - all dependencies are already installed!

# If NOT using Anaconda, install requirements
pip install -r requirements.txt
```

## Quick Start

### Method 1: Interactive Mode (Recommended)

```bash
python run.py
```

The interactive interface will guide you through:
1. **Workflow Mode Selection**:
   - Full Extraction - Extract data and generate report
   - Extract Only - Extract data, skip report generation
   - Report Only - Generate report from existing extraction data
2. Selecting a plan
3. Choosing a pattern repository
4. Configuring options (cleaning, log level)
5. Running the extraction

### Method 2: Command Line

```bash
# Full workflow (extraction + report)
python main.py --plan plan_ABC --patterns patterns_example.py

# Extract only (skip report generation)
python main.py --plan plan_ABC --patterns patterns_example.py --skip-report

# Report only (from existing extraction data)
python main.py --plan plan_ABC --report-only

# Skip text cleaning
python main.py --plan plan_ABC --patterns patterns_example.py --skip-cleaning

# Enable debug logging
python main.py --plan plan_ABC --patterns patterns_example.py --log-level DEBUG
```

---

## Master Configuration

The `master_config.toml` file provides centralized settings that apply to all extraction runs:

```toml
[SharePoint]
# Base URL for document links in reports
base_url = "https://yourcompany.sharepoint.com/sites/yoursite/"

# Hyperlink display style: "full" (complete URL) or "short" ("Open File" text)
hyperlink_style = "short"

# File extension for output document links
output_link_extension = ".pdf"

[Cleaning]
# Jaro-Winkler similarity threshold for spell-checking (0.0 - 1.0)
spell_check_threshold = 0.85

[Validation]
# Z-score threshold for positional outlier detection
positional_outlier_threshold = 3.0

# Maximum character gap between sequential elements
within_document_gap_threshold = 2000

[FileTypes]
# Supported file extensions for extraction
supported_extensions = [".txt"]

[Output]
# Include these columns in report output
include_extraction_order = true
include_extraction_position = false
include_flags = true
include_flag_reasons = true
```

### Configuration Loading

Settings are automatically loaded by modules that need them:
- **report_generator.py** - Uses SharePoint and Output settings
- **config_generator.py** - Uses Cleaning and Validation settings
- **orchestrator.py** - Uses FileTypes settings

---

## Setup Guide

### 1. Prepare Your Data Structure

Organize your plans folder with the Clusters subdirectory:

```
plans/
└── plan_ABC/
    ├── extraction_specs.xlsx          # Your extraction specifications
    └── Clusters/                       # Document type folders
        ├── Birth_Certificate/
        │   ├── doc1.txt
        │   ├── doc2.txt
        │   └── cleaned/               # Auto-generated
        ├── Employment_Application/
        │   ├── app1.txt
        │   └── app2.txt
        └── Termination_Letter/
            └── term1.txt
```

### 2. Create Extraction Specifications

Create `extraction_specs.xlsx` in your plan folder with these columns:

| Source Name | Elements to Extract | Name Start Anchor | Name Stop Anchor | ID in File | Duplicate Mappings | Name Prefixes | Reverse Name Order |
|-------------|---------------------|-------------------|------------------|------------|--------------------|--------------|--------------------|
| Birth_Certificate | ["DOB", "SSN"] | ["NAME"] | ["DATE OF BIRTH"] | FALSE | {} | ["Name"] | FALSE |
| Employment_Application | ["DOH", "SSN", "DOB"] | ["EMPLOYEE NAME", "SPOUSE NAME"] | ["DATE", "SSN"] | TRUE | {"DOB": "SDOB"} | ["Name", "Spouse"] | FALSE |

**Column Descriptions:**

- **Source Name**: Document type folder name (must match folder in Clusters/)
- **Elements to Extract**: List of element names to extract using patterns
- **Name Start Anchor**: List of text anchors that precede names
- **Name Stop Anchor**: List of text anchors that follow names (paired with start anchors by index)
- **ID in File**: Whether participant ID is embedded in filename
- **Duplicate Mappings**: Dictionary mapping element to renamed element for 2nd+ occurrences
  - Example: `{"DOB": "SDOB", "SSN": "SSSN"}` - Second DOB becomes SDOB, second SSN becomes SSSN
- **Name Prefixes**: List of prefix types for each name anchor pair
  - Options: `"Name"` (FNAME/LNAME), `"Spouse"` (SFNAME/SLNAME), `"Beneficiary"` (BFNAME/BLNAME), `"AP"` (APFNAME/APLNAME)
- **Reverse Name Order**: If true, expects "Last, First" name format

### 3. Configure Pattern Repository

Use the organized pattern repository in `patterns/`:

```python
from patterns_example import PATTERNS

# All 115+ patterns from all categories automatically loaded!
# Categories: dates, identifiers, amounts, strings
```

**Adding Custom Patterns:**

```python
# Single pattern
PATTERNS['CUSTOM_ID'] = r'\b[A-Z]{3}\d{6}\b'

# Multiple patterns (tries each until match found)
PATTERNS['DOB'] = [
    r'\b(?:0?[1-9]|1[0-2])[/-](?:0?[1-9]|[12][0-9]|3[01])[/-](?:\d{2}|\d{4})\b',
    r'\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4}\b',
]
```

---

## Workflow Details

### Step 1: Configuration Generation

Reads Excel specifications and generates TOML configs in `plans/{plan_name}/configs/`:

```toml
[Document]
document_source = "Birth_Certificate"
name_extraction = true

[Extraction]
active_patterns = ["DOB", "SSN"]
name_start_anchors = ["NAME"]
name_stop_anchors = ["DATE OF BIRTH"]
duplicate_mappings = {}
name_prefixes = ["Name"]

[Cleaning]
spell_check_threshold = 0.85  # From master config

[Validation]
positional_outlier_threshold = 3.0  # From master config
within_document_gap_threshold = 2000  # From master config
```

### Step 2: OCR Text Cleaning (Optional)

Cleans documents before extraction:
- Removes zero-width spaces and invisible characters
- Normalizes smart quotes and special characters
- Spell-checks keywords using fuzzy matching (Jaro-Winkler)

Output saved to `{source_folder}/cleaned/` with change logs in `cleaning_reports/`.

### Step 3: Pattern-Based Extraction

For each document:
- Extracts all elements using regex patterns
- Uses `finditer()` to capture all instances
- Records extraction order and position
- Applies duplicate mappings to rename 2nd+ occurrences

### Step 4: Name Extraction

For documents with name anchors:
- Creates dynamic regex using start/stop anchors
- Extracts names between anchors
- Applies name prefixes (Name, Spouse, Beneficiary, AP)
- Splits into FNAME/LNAME components

### Step 5: Value Cleaning

Standardizes extracted values:
- Dates: Converts to consistent MM/DD/YYYY format
- Amounts: Removes symbols, standardizes decimals
- Names: Uppercase, removes extra spaces
- SSN: Validates format

### Step 6: Validation

Flags potential issues:
- Date logic violations (DOB < DOH < DOTE)
- Positional outliers (unusual extraction positions)
- Multiple extractions of same element
- Value reasonableness checks

### Step 7: Report Generation

Creates interactive Excel report with:
- Participant Summary (quality metrics)
- All Extracted Data (with SharePoint links)
- Source Statistics
- Participant Statistics

---

## Output Files

All outputs saved to `plans/{plan_name}/output/`:

1. **Validated Pickle**: `{plan_name}_VALIDATED_{timestamp}.pkl`
   - Complete extraction data for downstream processing

2. **Interactive Report**: `{plan_name}_INTERACTIVE_REPORT.xlsx`
   - Participant Summary tab (conflicts, quality %)
   - All Extracted Data tab (with document links)
   - Source Statistics tab
   - Participant Statistics tab

3. **Raw Extractions**: `{plan_name}_extractions_{timestamp}.xlsx`
   - Raw extraction data
   - Summary statistics
   - Detailed statistics

---

## Troubleshooting

### No extractions found
- Verify regex patterns match your document format
- Check that pattern names in specs match pattern repository keys
- Enable DEBUG logging to see pattern matching details

### Duplicate mappings not working
- Ensure `Duplicate Mappings` column contains valid JSON dictionary
- Example: `{"DOB": "SDOB"}` (not `{DOB: SDOB}`)

### Name extraction not working
- Verify start/stop anchors exist in documents
- Check anchor spelling matches exactly
- Ensure Name Prefixes list length matches anchor pairs

### Report shows wrong hyperlink style
- Check `hyperlink_style` in `master_config.toml`
- Options: `"short"` (Open File) or `"full"` (complete URL)

---

## Documentation

See the `docs/` folder for complete documentation:
- [SOP.md](docs/SOP.md) - Standard Operating Procedures
- [modules/](docs/modules/) - Module-specific guides
- [reference/](docs/reference/) - Reference material

## Support

For issues or questions, refer to the inline documentation in each module.
