# Quick Start Guide - Text Extraction Tool

## For Team Members - Easiest Method

### Step 1: Run the Interactive Tool

```bash
python run.py
```

That's it! The tool will guide you through everything.

---

## What You'll Be Asked

### 1. Select a Plan
- Choose from existing plans in the `plans/` folder
- Or create a new plan name

### 2. Choose Pattern Repository
- Usually: `patterns_example.py` (includes 115+ pre-built patterns!)
- Just press Enter to use the default
- Patterns are organized by category in the `patterns/` folder

### 3. Options
- **Skip cleaning?** → Usually say "No" (unless docs are already clean)
- **Log level?** → Usually select "1" for INFO (standard output)

### 4. Confirm and Run
- Review the settings
- Press Enter to start

---

## What the Tool Does

The extraction workflow has **9 automated steps**:

1. ✅ **Config Generation** - Creates configuration files from Excel specs
2. ✅ **Text Cleaning** - Fixes OCR errors and spell-checks documents
3. ✅ **Pattern Loading** - Loads extraction patterns
4. ✅ **Data Extraction** - Extracts data from all documents
5. ✅ **Value Cleaning** - Standardizes dates, names, SSNs, etc.
6. ✅ **Validation** - Flags questionable extractions for review
7. ✅ **Statistics** - Calculates success rates and completion metrics
8. ✅ **Output Generation** - Creates Excel and pickle files
9. ✅ **Interactive Report** - Generates database-building report with links

---

## Where to Find Results

After running, check:

```
plans/your_plan_name/output/
```

You'll find:
- **`plan_ABC_VALIDATED.xlsx`** - All extractions with stats
- **`plan_ABC_HIGH_CONFIDENCE.xlsx`** - Ready for database import
- **`plan_ABC_LOW_CONFIDENCE.xlsx`** - Needs manual review
- **`plan_ABC_INTERACTIVE_REPORT.xlsx`** - Database building report

---

## Supported File Types

- ✅ **Text files (.txt)** - Fully supported
- ❌ **PDF files** - Not supported
- ❌ **Word documents** - Not supported

Convert PDFs and Word docs to .txt before processing.

---

## Common Issues

### "No plans found"
- Make sure you have a folder in `plans/` (e.g., `plans/plan_ABC/`)
- Put your documents and `extraction_specs.xlsx` in that folder

### "No pattern file found"
- Make sure `patterns_example.py` exists in the root directory
- Or specify the correct path to your pattern file

### "No document folders found"
- Each document type needs its own folder inside your plan folder
- Example: `plans/plan_ABC/Birth_Certificate/`

### "No .txt files found"
- Only .txt files are supported
- Check that your documents have the `.txt` extension

---

## Need More Details?

Check the full documentation:
- **[README.md](README.md)** - Complete guide
- **[docs/README.md](docs/README.md)** - Module documentation
- **[docs/modules/](docs/modules/)** - Detailed module guides

---

## Advanced Usage (Command Line)

For automation or scripting:

```bash
# Standard run
python main.py --plan plan_ABC --patterns patterns_example.py

# Skip cleaning
python main.py --plan plan_ABC --patterns patterns_example.py --skip-cleaning

# Debug mode
python main.py --plan plan_ABC --patterns patterns_example.py --log-level DEBUG
```

---

## Tips for Success

1. **Start Small** - Test with a few documents first
2. **Review Low Confidence** - Always check the low confidence report
3. **Update Patterns** - Improve patterns based on what you find
4. **Customize ID Extraction** - Edit `extractor.py` to match your filename pattern
5. **Use Interactive Reports** - The interactive Excel report is perfect for database building

---

## Questions?

- Check the [README.md](README.md)
- Review the [documentation](docs/README.md)
- Look at the examples in `examples/`
