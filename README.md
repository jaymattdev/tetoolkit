# Extraction Tool - Documentation Index

## Quick Links

- **[Main README](../README.md)** - Project overview and quick start
- **[SOP](SOP.md)** - Standard Operating Procedures (complete guide)
- **[Changelog](CHANGELOG.md)** - Version history and recent updates

## Module Guides

Complete documentation for each module:

- **[Text Cleaning](modules/CLEANING_MODULE.md)** - OCR error correction and spell-checking
- **[Value Cleaning](modules/VALUE_CLEANING.md)** - Value parsing and standardization
- **[Validation](modules/VALIDATION.md)** - Confidence flagging and sanity checks
- **[Statistics](modules/STATISTICS.md)** - Comprehensive statistics and timing
- **[Interactive Reports](modules/REPORT_GUIDE.md)** - Database building reports

## Reference Material

Advanced topics and reference guides:

- **[Folder Structure](reference/FOLDER_STRUCTURE.md)** - Project organization
- **[Organization Guide](reference/ORGANIZATION.md)** - Plan management
- **[Output Examples](reference/OUTPUT_EXAMPLE.md)** - Sample outputs
- **[Complete Summary](reference/SUMMARY.md)** - Full system overview

## Configuration

- **[Master Config](../master_config.toml)** - Global settings (SharePoint, cleaning, validation, output)

## Other Documentation

- **[Reorganization Plan](REORGANIZATION_PLAN.md)** - Structure improvements

## Workflow Overview

```
┌─────────────────────────────────────────────────────────────────┐
│  1. Config Generation  →  2. OCR Cleaning  →  3. Extraction    │
│           ↓                                                      │
│  4. Value Cleaning  →  5. Validation  →  6. Statistics          │
│           ↓                                                      │
│  7. Output Generation  →  8. Interactive Report                 │
└─────────────────────────────────────────────────────────────────┘

Workflow Modes:
- Full Extraction: Steps 1-8
- Extract Only: Steps 1-7 (skip report)
- Report Only: Step 8 only (from existing data)
```

## Getting Help

Each module documentation includes:
- Overview and purpose
- Configuration options
- Usage examples
- Troubleshooting
- Best practices

Start with the [Main README](../README.md) for quick start instructions.
