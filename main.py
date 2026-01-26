"""
Main Execution Script for Text Extraction Tool

This script orchestrates the entire extraction workflow:
1. Generate TOML configurations from Excel specification
2. Clean and spell-check OCR'd documents
3. Extract data from all documents in the plan folder
4. Clean and parse extracted values
5. Validate and flag extractions by confidence level
6. Calculate statistics
7. Save outputs to Excel and pickle formats

Usage:
    python main.py --plan plan_ABC --patterns patterns_example.py [--skip-cleaning] [--log-level INFO]
"""

import argparse
import sys
import logging
import pandas as pd
from pathlib import Path
from datetime import datetime
from config_generator import ConfigGenerator
from text_cleaner import TextCleaner
from extractor import TextExtractor
from orchestrator import ExtractionOrchestrator
from value_cleaner import ValueCleaner
from validator import ExtractionValidator
from statistics_manager import StatisticsManager
from report_generator import ReportGenerator
from output_manager import OutputManager

# Configure module logger
logger = logging.getLogger(__name__)


def load_pattern_repository(patterns_module_path=None):
    """
    Load pattern repository from external module.

    Args:
        patterns_module_path: Path to Python file containing pattern dictionary

    Returns:
        Dictionary mapping element names to regex patterns
    """
    # BOILERPLATE: Connect to your pattern repository here
    # This is a placeholder for your existing pattern repository

    if patterns_module_path:
        try:
            import importlib.util
            spec = importlib.util.spec_from_file_location("patterns", patterns_module_path)
            patterns_module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(patterns_module)

            # Assuming your patterns module has a PATTERNS dictionary
            if hasattr(patterns_module, 'PATTERNS'):
                return patterns_module.PATTERNS
            else:
                print("Warning: patterns module does not have 'PATTERNS' dictionary")
                return {}

        except Exception as e:
            print(f"Error loading pattern repository: {e}")
            return {}
    else:
        print("No pattern repository specified. Using empty patterns.")
        return {}


def setup_logging(log_level=logging.INFO, plan_folder=None):
    """
    Configure logging for the entire application.

    Args:
        log_level: Logging level (DEBUG, INFO, WARNING, ERROR)
        plan_folder: Optional plan folder path for log file
    """
    # Create formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)

    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)
    root_logger.addHandler(console_handler)

    # File handler if plan folder specified
    if plan_folder:
        log_dir = Path(plan_folder) / "logs"
        log_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = log_dir / f"extraction_{timestamp}.log"

        file_handler = logging.FileHandler(log_file)
        file_handler.setLevel(logging.DEBUG)  # Always log everything to file
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)

        logger.info(f"Logging to file: {log_file}")


def run_extraction_workflow(plan_name, plan_folder_path, patterns_path=None, skip_cleaning=False):
    """
    Run the complete extraction workflow.

    Args:
        plan_name: Name of the plan being processed
        plan_folder_path: Path to the plan folder (e.g., plans/plan_ABC)
        patterns_path: Path to Python file containing pattern repository
        skip_cleaning: If True, skip the text cleaning step

    Returns:
        Tuple of (excel_path, pickle_path)
    """
    # Initialize statistics manager for timing
    stats_mgr = StatisticsManager()
    workflow_start = datetime.now()

    plan_folder = Path(plan_folder_path)

    logger.info(f"Starting extraction workflow for plan: {plan_name}")
    logger.debug(f"Plan folder path: {plan_folder}")

    print("\n" + "="*80)
    print(f"TEXT EXTRACTION WORKFLOW - {plan_name}")
    print("="*80 + "\n")
    print(f"Plan folder: {plan_folder}")
    print(f"Expected structure:")
    print(f"  - {plan_folder}/extraction_specs.xlsx")
    print(f"  - {plan_folder}/[Document_Type_Folders]")
    print(f"  - {plan_folder}/configs/ (will be created)")
    print(f"  - {plan_folder}/output/ (will be created)")
    if not skip_cleaning:
        print(f"  - {plan_folder}/cleaning_reports/ (will be created)")
    print()

    # Define paths within the plan folder
    specs_excel = plan_folder / "extraction_specs.xlsx"
    configs_folder = plan_folder / "configs"
    output_folder = plan_folder / "output"

    # Create configs and output folders if they don't exist
    configs_folder.mkdir(parents=True, exist_ok=True)
    output_folder.mkdir(parents=True, exist_ok=True)

    # Step 1: Generate configurations from specs
    step_start = datetime.now()
    if specs_excel.exists():
        print("STEP 1: Generating TOML configurations from Excel specifications")
        print("-" * 80)
        logger.info("Generating configurations from Excel specifications")

        config_gen = ConfigGenerator(str(specs_excel), output_dir=str(configs_folder))
        generated_configs = config_gen.generate_all_configs()

        if not generated_configs:
            logger.error("Failed to generate configurations")
            print("Error: Failed to generate configurations")
            return None, None

        logger.info(f"Successfully generated {len(generated_configs)} configuration files")
        print(f"✓ Generated {len(generated_configs)} configuration files\n")

        stats_mgr.record_timing("1. Config Generation", step_start, datetime.now())
    else:
        print("STEP 1: Using existing TOML configurations")
        print("-" * 80)
        logger.warning(f"extraction_specs.xlsx not found at {specs_excel}")
        print(f"⚠ Warning: extraction_specs.xlsx not found at {specs_excel}")
        print(f"Configuration folder: {configs_folder}\n")

    # Step 2: Clean OCR'd documents (if not skipped)
    step_start = datetime.now()
    if not skip_cleaning:
        print("STEP 2: Cleaning and spell-checking OCR'd documents")
        print("-" * 80)
        logger.info("Starting text cleaning process")

        # Get all config files
        config_files = list(configs_folder.glob("*.toml"))

        if not config_files:
            logger.warning("No configuration files found for cleaning")
            print("⚠ Warning: No configuration files found. Skipping cleaning.\n")
        else:
            all_cleaning_stats = []

            for config_path in config_files:
                source_name = config_path.stem
                source_folder = plan_folder / source_name

                if not source_folder.exists():
                    logger.warning(f"Source folder not found: {source_folder}")
                    continue

                logger.info(f"Cleaning documents for source: {source_name}")
                print(f"  Cleaning: {source_name}")

                try:
                    cleaner = TextCleaner(
                        config_path=str(config_path),
                        plan_folder=str(plan_folder)
                    )

                    stats = cleaner.clean_folder(str(source_folder))
                    all_cleaning_stats.append(stats)

                    # Generate cleaning report
                    report_path = cleaner.generate_cleaning_report([stats])

                    print(f"    ✓ Cleaned {stats['files_processed']} files, "
                          f"{stats['total_changes']} changes")
                    logger.info(f"Cleaning complete for {source_name}: "
                              f"{stats['total_changes']} changes")

                except Exception as e:
                    logger.error(f"Error cleaning {source_name}: {e}", exc_info=True)
                    print(f"    ✗ Error cleaning {source_name}: {e}")

            print(f"✓ Cleaning complete for all sources\n")
            logger.info("Text cleaning process completed")

        stats_mgr.record_timing("2. Text Cleaning", step_start, datetime.now())
    else:
        logger.info("Skipping text cleaning (--skip-cleaning flag set)")
        print("STEP 2: Skipping text cleaning")
        print("-" * 80)
        print("✓ Cleaning skipped\n")

    # Step 3: Load pattern repository
    step_start = datetime.now()
    step_num = 2 if skip_cleaning else 3
    print(f"STEP {step_num}: Loading pattern repository")
    print("-" * 80)
    logger.info("Loading pattern repository")

    patterns_dict = load_pattern_repository(patterns_path)
    logger.info(f"Loaded {len(patterns_dict)} pattern definitions")
    print(f"✓ Loaded {len(patterns_dict)} pattern definitions\n")

    stats_mgr.record_timing("3. Pattern Loading", step_start, datetime.now())

    # Step 4: Extract data from all documents
    step_start = datetime.now()
    step_num += 1
    print(f"STEP {step_num}: Extracting data from documents")
    print("-" * 80)
    logger.info("Starting data extraction")

    orchestrator = ExtractionOrchestrator(
        plan_folder=str(plan_folder),
        configs_folder=str(configs_folder),
        patterns_dict=patterns_dict
    )

    all_extractions = orchestrator.process_all_sources()

    if not all_extractions:
        logger.warning("No extractions were performed")
        print("Warning: No extractions were performed")
        return None, None

    logger.info(f"Completed extractions: {len(all_extractions)} total items")
    print(f"✓ Completed extractions: {len(all_extractions)} total items\n")

    stats_mgr.record_timing("4. Data Extraction", step_start, datetime.now())

    # Step 5: Clean and parse extracted values
    step_start = datetime.now()
    step_num += 1
    print(f"STEP {step_num}: Cleaning and parsing extracted values")
    print("-" * 80)
    logger.info("Starting value cleaning and parsing")

    # Convert extractions to DataFrame
    extractions_df = orchestrator.get_extractions_dataframe()

    # Get reverse name order setting from first config (assume same for all in plan)
    config_files = list(configs_folder.glob("*.toml"))
    reverse_name_order = False
    if config_files:
        import toml
        first_config = toml.load(config_files[0])
        reverse_name_order = first_config.get('Parsing', {}).get('reverse_name_order', False)
        logger.debug(f"Reverse name order setting: {reverse_name_order}")

    # Clean values
    value_cleaner = ValueCleaner()
    cleaned_df = value_cleaner.clean_extractions_dataframe(extractions_df, reverse_name_order)

    # Convert back to list for compatibility
    cleaned_extractions = cleaned_df.to_dict('records')

    logger.info(f"Value cleaning complete: {len(cleaned_extractions)} records processed")
    print(f"✓ Cleaned and parsed {len(cleaned_extractions)} values\n")

    stats_mgr.record_timing("5. Value Cleaning", step_start, datetime.now())

    # Step 6: Validate and flag extractions
    step_start = datetime.now()
    step_num += 1
    print(f"STEP {step_num}: Validating extractions and separating by confidence")
    print("-" * 80)
    logger.info("Starting validation and sanity checking")

    # Use first config for validation settings (assume same for all in plan)
    validator = ExtractionValidator(config_path=str(config_files[0]) if config_files else None)
    validated_df = validator.validate_extractions(cleaned_df)

    # Separate by confidence level
    high_confidence_df = validated_df[validated_df['confidence'] == 'HIGH'].copy()
    low_confidence_df = validated_df[validated_df['confidence'] == 'LOW'].copy()

    high_count = len(high_confidence_df)
    low_count = len(low_confidence_df)

    logger.info(f"Validation complete: {high_count} high confidence, {low_count} low confidence")
    print(f"✓ High confidence extractions: {high_count}")
    print(f"✓ Low confidence extractions: {low_count}")

    if low_count > 0:
        print(f"  ⚠ {low_count} extractions flagged for review\n")
    else:
        print()

    # Convert back to lists for output manager
    all_validated_extractions = validated_df.to_dict('records')
    high_confidence_extractions = high_confidence_df.to_dict('records')
    low_confidence_extractions = low_confidence_df.to_dict('records')

    stats_mgr.record_timing("6. Validation", step_start, datetime.now())

    # Step 7: Calculate statistics
    step_start = datetime.now()
    step_num += 1
    print(f"STEP {step_num}: Calculating comprehensive statistics")
    print("-" * 80)
    logger.info("Calculating comprehensive statistics")

    # Generate comprehensive statistics using new stats manager
    all_stats = stats_mgr.generate_comprehensive_statistics(validated_df, include_participant_stats=True)

    # Also get legacy statistics for backward compatibility
    stats = orchestrator.calculate_statistics()
    summary_df, detailed_df = orchestrator.get_statistics_dataframe()

    # Extract individual stat DataFrames
    element_stats_df = all_stats.get('element_statistics', pd.DataFrame())
    parsing_stats_df = all_stats.get('parsing_statistics', pd.DataFrame())
    confidence_stats_df = all_stats.get('confidence_statistics', pd.DataFrame())
    flag_stats_df = all_stats.get('flag_statistics', pd.DataFrame())
    participant_stats_df = all_stats.get('participant_statistics', pd.DataFrame())
    timing_stats_df = all_stats.get('timing_statistics', pd.DataFrame())

    logger.info(f"Comprehensive statistics calculated")
    print(f"✓ Element statistics: {len(element_stats_df)} records")
    print(f"✓ Parsing statistics: {len(parsing_stats_df)} records")
    print(f"✓ Confidence statistics: {len(confidence_stats_df)} records")
    if not participant_stats_df.empty:
        print(f"✓ Participant statistics: {len(participant_stats_df)} participants")
    print()

    stats_mgr.record_timing("7. Statistics Calculation", step_start, datetime.now())

    # Step 8: Save outputs
    step_start = datetime.now()
    step_num += 1
    print(f"STEP {step_num}: Saving output files")
    print("-" * 80)
    logger.info("Saving output files")

    output_mgr = OutputManager(output_folder=str(output_folder))

    # Prepare additional statistics for output
    additional_stats_dict = {
        'element_statistics': element_stats_df,
        'parsing_statistics': parsing_stats_df,
        'confidence_statistics': confidence_stats_df,
        'flag_statistics': flag_stats_df,
        'participant_statistics': participant_stats_df,
        'timing_statistics': timing_stats_df
    }

    # Save combined validated output (with all statistics)
    excel_path, pickle_path = output_mgr.save_all_outputs(
        extractions_list=all_validated_extractions,
        summary_stats_df=summary_df,
        detailed_stats_df=detailed_df,
        plan_name=plan_name,
        suffix="VALIDATED",
        additional_stats=additional_stats_dict
    )

    # Save high confidence output
    high_excel, high_pickle = output_mgr.save_all_outputs(
        extractions_list=high_confidence_extractions,
        summary_stats_df=summary_df,
        detailed_stats_df=detailed_df,
        plan_name=plan_name,
        suffix="HIGH_CONFIDENCE",
        additional_stats=additional_stats_dict
    )

    # Save low confidence output (if any)
    if low_count > 0:
        low_excel, low_pickle = output_mgr.save_all_outputs(
            extractions_list=low_confidence_extractions,
            summary_stats_df=summary_df,
            detailed_stats_df=detailed_df,
            plan_name=plan_name,
            suffix="LOW_CONFIDENCE",
            additional_stats=additional_stats_dict
        )
        logger.info(f"Low confidence output saved: {low_excel}")
        print(f"  - Low confidence: {low_excel}")

    logger.info(f"Combined validated output saved: {excel_path}")
    logger.info(f"High confidence output saved: {high_excel}")
    print(f"  - Combined validated: {excel_path}")
    print(f"  - High confidence: {high_excel}")

    stats_mgr.record_timing("8. Output Generation", step_start, datetime.now())

    # Step 9: Generate interactive report
    step_start = datetime.now()
    step_num += 1
    print(f"STEP {step_num}: Generating interactive Excel report for database building")
    print("-" * 80)
    logger.info("Generating interactive report")

    report_gen = ReportGenerator(
        sharepoint_base_url="https://yourcompany.sharepoint.com/sites/yoursite/"
    )

    interactive_report_path = output_folder / f"{plan_name}_INTERACTIVE_REPORT.xlsx"
    report_path = report_gen.generate_interactive_report(
        validated_df=validated_df,
        output_path=str(interactive_report_path),
        plan_name=plan_name
    )

    logger.info(f"Interactive report generated: {report_path}")
    print(f"✓ Interactive report: {report_path}\n")

    stats_mgr.record_timing("9. Interactive Report", step_start, datetime.now())

    # Display summary report
    output_mgr.create_extraction_summary_report(summary_df)

    # Record total workflow time
    workflow_end = datetime.now()
    stats_mgr.record_timing("TOTAL WORKFLOW", workflow_start, workflow_end)

    # Display timing summary
    print("\n" + "="*80)
    print("TIMING SUMMARY")
    print("="*80)
    for module, timing in stats_mgr.timing_stats.items():
        print(f"{module}: {timing['duration_formatted']}")
    print("="*80 + "\n")

    logger.info("Workflow completed successfully")
    print("="*80)
    print("WORKFLOW COMPLETE")
    print("="*80 + "\n")

    return excel_path, pickle_path


def main():
    """Main entry point for the extraction tool."""
    parser = argparse.ArgumentParser(
        description="Text Extraction Tool - Extract structured data from documents",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process plan_ABC with pattern repository
  python main.py --plan plan_ABC --patterns patterns_example.py

  # Process plan using default structure
  python main.py --plan plan_XYZ --patterns patterns_example.py

Expected folder structure:
  plans/
    plan_ABC/
      extraction_specs.xlsx          ← Excel specifications
      Birth_Certificate/             ← Document folders
      Employment_Application/
      configs/                       ← Generated configs (auto-created)
      output/                        ← Results (auto-created)
        """
    )

    parser.add_argument(
        '--plan',
        required=True,
        help='Name of the plan (e.g., plan_ABC). Will look in plans/{plan_name}/'
    )

    parser.add_argument(
        '--patterns',
        help='Path to Python file containing pattern repository (e.g., patterns_example.py)'
    )

    parser.add_argument(
        '--skip-cleaning',
        action='store_true',
        help='Skip the text cleaning step (extract from original documents)'
    )

    parser.add_argument(
        '--log-level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help='Logging level (default: INFO)'
    )

    args = parser.parse_args()

    # Construct plan folder path
    plan_folder = Path("plans") / args.plan

    # Validate plan folder exists
    if not plan_folder.exists():
        print(f"Error: Plan folder not found: {plan_folder}")
        print(f"\nExpected structure:")
        print(f"  plans/")
        print(f"    {args.plan}/")
        print(f"      extraction_specs.xlsx")
        print(f"      [Document_Type_Folders]/")
        sys.exit(1)

    # Setup logging
    log_level = getattr(logging, args.log_level)
    setup_logging(log_level=log_level, plan_folder=str(plan_folder))

    logger.info(f"Starting extraction tool for plan: {args.plan}")
    logger.debug(f"Arguments: {args}")

    # Run the workflow
    excel_path, pickle_path = run_extraction_workflow(
        plan_name=args.plan,
        plan_folder_path=str(plan_folder),
        patterns_path=args.patterns,
        skip_cleaning=args.skip_cleaning
    )

    if excel_path and pickle_path:
        print("Extraction completed successfully!")
        sys.exit(0)
    else:
        print("Extraction completed with warnings or errors.")
        sys.exit(1)


if __name__ == "__main__":
    main()
