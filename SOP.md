# Standard Operating Procedures (SOP)

## Text Extraction Tool - Complete Guide

This document provides a comprehensive overview of all modules, their responsibilities, and standard procedures for both users and developers.

---

## Table of Contents

1. [System Overview](#system-overview)
2. [Module Reference](#module-reference)
3. [User Procedures](#user-procedures)
4. [Developer Procedures](#developer-procedures)
5. [Configuration Reference](#configuration-reference)
6. [Troubleshooting Guide](#troubleshooting-guide)

---

## System Overview

### Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           TEXT EXTRACTION TOOL                               │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  ┌──────────────┐    ┌─────────────────┐    ┌──────────────────────────┐   │
│  │   run.py     │───▶│    main.py      │───▶│  master_config.toml      │   │
│  │ (Interactive)│    │ (Orchestration) │    │  config_loader.py        │   │
│  └──────────────┘    └────────┬────────┘    └──────────────────────────┘   │
│                               │                                              │
│         ┌─────────────────────┼─────────────────────┐                       │
│         ▼                     ▼                     ▼                        │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────────┐              │
│  │config_generator│    │text_cleaner  │    │patterns_example  │              │
│  │    .py       │    │    .py       │    │    .py           │              │
│  └──────────────┘    └──────────────┘    └──────────────────┘              │
│         │                     │                     │                        │
│         └─────────────────────┼─────────────────────┘                       │
│                               ▼                                              │
│                      ┌──────────────┐                                       │
│                      │orchestrator.py│                                       │
│                      └───────┬──────┘                                       │
│                              │                                               │
│                              ▼                                               │
│                      ┌──────────────┐                                       │
│                      │ extractor.py │                                       │
│                      └───────┬──────┘                                       │
│                              │                                               │
│         ┌────────────────────┼────────────────────┐                         │
│         ▼                    ▼                    ▼                          │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────┐                  │
│  │value_cleaner │    │ validator.py │    │statistics_   │                  │
│  │    .py       │    │              │    │manager.py    │                  │
│  └──────────────┘    └──────────────┘    └──────────────┘                  │
│         │                    │                    │                          │
│         └────────────────────┼────────────────────┘                         │
│                              ▼                                               │
│         ┌────────────────────┴────────────────────┐                         │
│         ▼                                         ▼                          │
│  ┌──────────────┐                        ┌──────────────┐                   │
│  │output_manager│                        │report_       │                   │
│  │    .py       │                        │generator.py  │                   │
│  └──────────────┘                        └──────────────┘                   │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Data Flow

```
Input Documents (.txt)
        │
        ▼
┌───────────────────┐
│  Text Cleaning    │  ──▶  Cleaned Documents + Change Reports
│  (text_cleaner)   │
└───────────────────┘
        │
        ▼
┌───────────────────┐
│  Extraction       │  ──▶  Raw Extractions (element, value, position)
│  (extractor)      │
└───────────────────┘
        │
        ▼
┌───────────────────┐
│  Value Cleaning   │  ──▶  Cleaned Values (standardized formats)
│  (value_cleaner)  │
└───────────────────┘
        │
        ▼
┌───────────────────┐
│  Validation       │  ──▶  Validated Data (with flags, confidence)
│  (validator)      │
└───────────────────┘
        │
        ▼
┌───────────────────┐
│  Output           │  ──▶  Excel + Pickle + Interactive Report
│  (output_manager) │
└───────────────────┘
```

---

## Module Reference

### 1. run.py - Interactive Terminal UI

**Purpose:** Provides user-friendly interface for running extractions.

**Key Functions:**
| Function | Description |
|----------|-------------|
| `main()` | Main interactive loop |
| `print_header()` | Display header with master config info |
| `list_plans()` | Find available plans in plans/ directory |
| `list_pattern_files()` | Find pattern repository files |
| `validate_plan_structure()` | Verify plan has required structure |
| `select_from_list()` | Interactive list selection |

**Workflow Modes:**
- **Full Extraction**: Complete workflow (clean → extract → validate → report)
- **Extract Only**: Skip report generation (`--skip-report`)
- **Report Only**: Generate report from existing pickle data (`--report-only`)

---

### 2. main.py - Main Execution Script

**Purpose:** Orchestrates the complete extraction workflow.

**Key Functions:**
| Function | Description |
|----------|-------------|
| `run_extraction_workflow()` | Execute full extraction pipeline |
| `run_report_only()` | Generate report from existing data |
| `setup_logging()` | Configure logging system |
| `load_patterns()` | Import pattern repository |

**Command Line Arguments:**
| Argument | Description |
|----------|-------------|
| `--plan` | Plan folder name (required) |
| `--patterns` | Pattern repository file |
| `--skip-cleaning` | Skip OCR text cleaning |
| `--skip-report` | Skip report generation |
| `--report-only` | Generate report only |
| `--log-level` | Logging level (DEBUG/INFO/WARNING/ERROR) |

---

### 3. config_loader.py - Master Configuration

**Purpose:** Loads and provides access to master_config.toml settings.

**Key Class: `MasterConfig`**

| Property | Section | Description |
|----------|---------|-------------|
| `sharepoint_base_url` | SharePoint | Base URL for document links |
| `hyperlink_style` | SharePoint | Display style: "full" or "short" |
| `output_link_extension` | SharePoint | File extension for links |
| `spell_check_threshold` | Cleaning | Jaro-Winkler threshold (0.0-1.0) |
| `positional_outlier_threshold` | Validation | Z-score threshold |
| `within_document_gap_threshold` | Validation | Max character gap |
| `supported_extensions` | FileTypes | List of file extensions |
| `include_extraction_order` | Output | Include in report |
| `include_extraction_position` | Output | Include in report |
| `include_flags` | Output | Include in report |
| `include_flag_reasons` | Output | Include in report |

**Usage:**
```python
from config_loader import get_master_config
config = get_master_config()
threshold = config.spell_check_threshold
```

---

### 4. config_generator.py - Configuration Generator

**Purpose:** Converts Excel specifications to TOML config files.

**Key Class: `ConfigGenerator`**

| Method | Description |
|--------|-------------|
| `load_excel()` | Load and validate Excel specifications |
| `parse_list_column()` | Parse JSON list strings |
| `parse_dict_column()` | Parse JSON dict strings |
| `create_toml_config()` | Generate TOML config from Excel row |
| `generate_all_configs()` | Process all rows in Excel |

**Generated Config Structure:**
```toml
[Document]
document_source = "Birth_Certificate"
name_extraction = true
id_in_file = false

[Extraction]
active_patterns = ["DOB", "SSN"]
name_start_anchors = ["NAME"]
name_stop_anchors = ["DATE OF BIRTH"]
duplicate_mappings = {"DOB" = "SDOB"}
name_prefixes = ["Name", "Spouse"]

[Cleaning]
spell_check_threshold = 0.85  # From master config

[Validation]
positional_outlier_threshold = 3.0  # From master config
within_document_gap_threshold = 2000  # From master config

[Output]
output_file_name = "Birth_Certificate_extractions.xlsx"
```

---

### 5. text_cleaner.py - OCR Text Cleaning

**Purpose:** Corrects OCR errors in source documents before extraction.

**Key Class: `TextCleaner`**

| Method | Description |
|--------|-------------|
| `clean_text()` | Apply all cleaning transformations |
| `remove_invisible_chars()` | Remove zero-width spaces |
| `normalize_quotes()` | Convert smart quotes |
| `fix_spacing()` | Fix spacing around punctuation |
| `spell_check_keywords()` | Fuzzy match keyword corrections |
| `clean_folder()` | Process all files in a folder |

**Outputs:**
- Cleaned files: `{source_folder}/cleaned/`
- Change reports: `{plan_folder}/cleaning_reports/`

---

### 6. extractor.py - Text Extraction Engine

**Purpose:** Extracts data elements from text using regex patterns.

**Key Class: `TextExtractor`**

| Method | Description |
|--------|-------------|
| `extract_from_file()` | Extract all elements from a file |
| `extract_element()` | Extract single element with pattern |
| `extract_name()` | Extract names using anchors |
| `extract_id_from_filename()` | Get participant ID from filename |
| `apply_duplicate_mappings()` | Rename 2nd+ occurrences |
| `deduplicate_extractions()` | Remove duplicate values |
| `extract_from_folder()` | Process all files in folder |

**Duplicate Mappings:**
When a document has intentionally repeated elements (e.g., participant DOB and spouse DOB):
```python
duplicate_mappings = {"DOB": "SDOB", "SSN": "SSSN"}
# DOB extraction_order=1 stays as "DOB"
# DOB extraction_order=2 becomes "SDOB"
```

**Name Prefixes:**
Controls prefix for name elements:
- `"Name"` → FNAME, LNAME
- `"Spouse"` → SFNAME, SLNAME
- `"Beneficiary"` → BFNAME, BLNAME
- `"AP"` → APFNAME, APLNAME

---

### 7. orchestrator.py - Multi-Source Coordinator

**Purpose:** Coordinates extraction across multiple document types.

**Key Class: `ExtractionOrchestrator`**

| Method | Description |
|--------|-------------|
| `process_all_sources()` | Process all folders in Clusters/ |
| `process_single_source()` | Extract from one document type |
| `get_config_for_source()` | Find TOML config for source |
| `calculate_statistics()` | Compute extraction statistics |
| `get_extractions_dataframe()` | Convert results to DataFrame |

**Folder Structure:**
```
plans/plan_ABC/
└── Clusters/                    # Document type folders
    ├── Birth_Certificate/       # Source folder
    │   ├── doc1.txt
    │   └── cleaned/             # Cleaned versions
    └── Employment_Application/
```

---

### 8. value_cleaner.py - Value Parsing and Cleaning

**Purpose:** Standardizes extracted values into consistent formats.

**Key Class: `ValueCleaner`**

| Method | Description |
|--------|-------------|
| `clean_date()` | Parse and standardize dates |
| `clean_ssn()` | Validate and format SSN |
| `clean_amount()` | Parse monetary amounts |
| `clean_string()` | Clean general text |
| `parse_name()` | Split names into FNAME/LNAME |
| `clean_extractions_dataframe()` | Process all extractions |

**Name Parsing:**
```python
# Normal order: "John Michael Smith"
# reverse_order=False: FNAME=JOHN, LNAME=MICHAEL SMITH

# Reverse order: "Smith, John"
# reverse_order=True: FNAME=JOHN, LNAME=SMITH
```

---

### 9. validator.py - Validation and Flagging

**Purpose:** Validates extractions and assigns confidence levels.

**Key Class: `ExtractionValidator`**

| Method | Description |
|--------|-------------|
| `validate_extractions()` | Run all validation checks |
| `check_date_logic()` | Verify DOB < DOH < DOTE |
| `check_positional_outliers()` | Flag unusual positions |
| `check_multiple_extractions()` | Flag duplicate elements |
| `check_conflicting_values()` | Flag different values for same element |
| `check_value_reasonableness()` | Validate value ranges |

**Validation Flags:**
| Flag | Description |
|------|-------------|
| `date_logic_violation` | Dates not in expected order |
| `positional_outlier` | Unusual extraction position |
| `multiple_extractions` | Same element extracted multiple times |
| `within_document_conflict` | Different values for same element |
| `within_document_gap` | Large gap between extractions |

**Confidence Levels:**
- **HIGH**: No flags, value present
- **LOW**: Has flags or issues
- **None**: Missing value

---

### 10. statistics_manager.py - Statistics Calculation

**Purpose:** Calculates comprehensive extraction statistics.

**Key Class: `StatisticsManager`**

| Method | Description |
|--------|-------------|
| `calculate_all_statistics()` | Compute all metrics |
| `get_summary_statistics()` | High-level summary |
| `get_detailed_statistics()` | Per-element breakdown |
| `get_timing_statistics()` | Performance timing |

---

### 11. report_generator.py - Interactive Report Generation

**Purpose:** Creates Excel reports for data review and database building.

**Key Class: `ReportGenerator`**

| Method | Description |
|--------|-------------|
| `generate_interactive_report()` | Create complete report |
| `create_participant_summary()` | Quality metrics per participant |
| `create_all_data_tab()` | All extractions with links |
| `create_source_statistics_tab()` | Stats by document source |
| `apply_excel_formatting()` | Apply styles and formatting |
| `_process_hyperlinks()` | Create clickable links |

**Report Tabs:**
1. **Participant Summary**: Quality %, conflicts, missing elements
2. **All Extracted Data**: Full extraction data with SharePoint links
3. **Source Statistics**: Extraction success by document type
4. **Participant Statistics**: Completeness by participant

**Hyperlink Styles:**
- `"short"`: Displays "Open File" as clickable text
- `"full"`: Displays complete URL in cell

---

### 12. output_manager.py - Output File Management

**Purpose:** Manages saving extraction results to files.

**Key Class: `OutputManager`**

| Method | Description |
|--------|-------------|
| `save_excel()` | Save DataFrame to Excel |
| `save_pickle()` | Save to pickle format |
| `generate_timestamp()` | Create timestamp string |

---

## User Procedures

### Procedure 1: Running a New Extraction

**Prerequisites:**
- Plan folder created in `plans/`
- Document folders in `plans/{plan}/Clusters/`
- `extraction_specs.xlsx` in plan folder
- Pattern repository file available

**Steps:**

1. **Start the interactive interface:**
   ```bash
   python run.py
   ```

2. **Select workflow mode:**
   - Choose "1" for Full Extraction (recommended for new runs)

3. **Select your plan:**
   - Choose from the list of available plans

4. **Select pattern repository:**
   - Choose from available pattern files

5. **Configure options:**
   - Skip cleaning? (No for first run)
   - Log level? (INFO recommended)

6. **Confirm and run:**
   - Review summary
   - Press Enter to start

7. **Check outputs:**
   - Results in `plans/{plan}/output/`
   - Interactive report: `{plan}_INTERACTIVE_REPORT.xlsx`

---

### Procedure 2: Regenerating Report Only

**Prerequisites:**
- Previous extraction completed
- Validated pickle file exists in output folder

**Steps:**

1. **Start interactive interface:**
   ```bash
   python run.py
   ```

2. **Select "Report Only" mode (option 3)**

3. **Select plan with existing data**

4. **Run report generation**

**Or via command line:**
```bash
python main.py --plan plan_ABC --report-only
```

---

### Procedure 3: Adding New Document Types

1. **Create folder in Clusters:**
   ```
   plans/plan_ABC/Clusters/New_Document_Type/
   ```

2. **Add text files:**
   - Place .txt files in the folder

3. **Update extraction_specs.xlsx:**
   - Add new row with Source Name = "New_Document_Type"
   - Define elements, anchors, mappings

4. **Run extraction:**
   - The tool will auto-generate TOML config

---

### Procedure 4: Configuring Duplicate Mappings

For documents with intentionally repeated elements (e.g., participant and spouse data):

1. **In extraction_specs.xlsx:**
   ```
   Duplicate Mappings: {"DOB": "SDOB", "SSN": "SSSN"}
   ```

2. **How it works:**
   - First DOB found → element stays "DOB"
   - Second DOB found → element renamed to "SDOB"
   - First SSN found → element stays "SSN"
   - Second SSN found → element renamed to "SSSN"

---

### Procedure 5: Configuring Name Extraction

1. **In extraction_specs.xlsx:**
   ```
   Name Start Anchor: ["EMPLOYEE NAME", "SPOUSE NAME", "BENEFICIARY"]
   Name Stop Anchor: ["DATE", "SSN", "ADDRESS"]
   Name Prefixes: ["Name", "Spouse", "Beneficiary"]
   ```

2. **Anchor pairing:**
   - Index 0: "EMPLOYEE NAME" → "DATE" (prefix: Name → FNAME, LNAME)
   - Index 1: "SPOUSE NAME" → "SSN" (prefix: Spouse → SFNAME, SLNAME)
   - Index 2: "BENEFICIARY" → "ADDRESS" (prefix: Beneficiary → BFNAME, BLNAME)

---

## Developer Procedures

### Procedure D1: Adding New Validation Checks

1. **Edit validator.py:**
   ```python
   def check_custom_validation(self, document_extractions):
       """Custom validation logic."""
       violations = {}
       # Add validation logic
       return violations
   ```

2. **Call from validate_extractions():**
   ```python
   # In the document-level loop:
   custom_violations = self.check_custom_validation(doc_group)
   for idx, (flags, reasons) in custom_violations.items():
       result_df.at[idx, 'flags'] = result_df.at[idx, 'flags'] + flags
       result_df.at[idx, 'flag_reasons'].update(reasons)
   ```

---

### Procedure D2: Adding New Master Config Settings

1. **Add to master_config.toml:**
   ```toml
   [NewSection]
   new_setting = "value"
   ```

2. **Add default in config_loader.py:**
   ```python
   DEFAULTS = {
       # ...
       'NewSection': {
           'new_setting': 'default_value',
       },
   }
   ```

3. **Add property in MasterConfig class:**
   ```python
   @property
   def new_setting(self) -> str:
       """Description of setting."""
       return self._get_nested('NewSection', 'new_setting')
   ```

4. **Use in modules:**
   ```python
   from config_loader import get_master_config
   config = get_master_config()
   value = config.new_setting
   ```

---

### Procedure D3: Adding New Pattern Categories

1. **Create new file in patterns/:**
   ```python
   # patterns/custom_category.py
   CUSTOM_PATTERN = r'\bPATTERN\b'
   ```

2. **Import in patterns_example.py:**
   ```python
   from patterns.custom_category import *

   PATTERNS['CUSTOM_PATTERN'] = CUSTOM_PATTERN
   ```

---

### Procedure D4: Customizing Participant ID Extraction

1. **Edit extractor.py:**
   ```python
   def extract_id_from_filename(self, filename):
       """Extract participant ID from filename."""
       # Example: "ABC_12345_form.txt" -> "12345"
       match = re.search(r'_(\d{5})_', filename)
       if match:
           return match.group(1)

       # Add more patterns as needed
       return None
   ```

---

## Configuration Reference

### master_config.toml

```toml
# =============================================================================
# MASTER CONFIGURATION - Text Extraction Tool
# =============================================================================

[SharePoint]
# Base URL for document links in reports
base_url = "https://yourcompany.sharepoint.com/sites/yoursite/"

# Hyperlink display style:
#   "full"  - Shows complete URL in cell
#   "short" - Shows "Open File" as clickable text
hyperlink_style = "short"

# File extension for output document links
output_link_extension = ".pdf"


[Cleaning]
# Jaro-Winkler similarity threshold for spell-checking (0.0 - 1.0)
# Higher = more strict, Lower = more lenient
spell_check_threshold = 0.85


[Validation]
# Z-score threshold for positional outlier detection
# Higher = less strict (fewer outliers flagged)
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

### extraction_specs.xlsx Columns

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Source Name | String | Yes | Document type folder name |
| Elements to Extract | List | Yes | Element names to extract |
| Name Start Anchor | List | No | Anchors before names |
| Name Stop Anchor | List | No | Anchors after names |
| ID in File | Boolean | No | Extract ID from filename |
| Duplicate Mappings | Dict | No | Rename 2nd+ occurrences |
| Name Prefixes | List | No | Prefix types for names |
| Reverse Name Order | Boolean | No | Expect "Last, First" |
| Additional Keywords | List | No | Extra spell-check words |

---

## Troubleshooting Guide

### Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| No extractions found | Pattern doesn't match | Enable DEBUG logging; test pattern |
| Missing names | Anchor not found | Check anchor spelling in documents |
| Wrong hyperlink style | Config not loaded | Verify master_config.toml exists |
| Duplicate mappings ignored | Invalid JSON | Use proper format: `{"DOB": "SDOB"}` |
| Report generation fails | No pickle file | Run full extraction first |

### Debug Commands

```bash
# Enable debug logging
python main.py --plan plan_ABC --patterns patterns.py --log-level DEBUG

# Check master config loading
python config_loader.py

# Validate pattern file
python -c "from patterns_example import PATTERNS; print(len(PATTERNS))"
```

### Log Locations

- Extraction logs: `plans/{plan}/logs/`
- Cleaning reports: `plans/{plan}/cleaning_reports/`

---

## Version History

See [CHANGELOG.md](CHANGELOG.md) for version history and recent updates.
