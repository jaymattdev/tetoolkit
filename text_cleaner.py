"""
Text Cleaning Module
Cleans OCR'd text documents to improve pattern matching accuracy.

This module:
1. Fixes common OCR errors
2. Spell-checks critical keywords using fuzzy matching
3. Normalizes whitespace and formatting
4. Logs all changes for review
"""

import re
import os
import logging
from pathlib import Path
from typing import Dict, List, Tuple
import jellyfish
import toml
import pandas as pd
from datetime import datetime

# Configure module logger
logger = logging.getLogger(__name__)


class TextCleaner:
    """Clean and correct OCR'd text documents."""

    # Safe OCR corrections - removing problematic special characters only
    # Note: Character substitutions (0->O, l->1, etc.) are NOT included as they can corrupt data
    # Spell-checking handles misspelled keywords more safely with fuzzy matching
    OCR_CORRECTIONS = {
        # Remove zero-width characters and other invisible characters
        r'[\u200B-\u200D\uFEFF]': '',  # Zero-width spaces
        r'\u00A0': ' ',  # Non-breaking space to regular space

        # Clean up problematic special characters that OCR adds
        r'[`´]': "'",  # Smart/backticks to regular apostrophe
        r'["""]': '"',  # Smart quotes to regular quotes
        r'[–—]': '-',  # Em/en dashes to regular hyphen

        # Remove excessive special characters (keep important ones: ,.%$()[]/)
        r'[^\w\s,.%$()[\]\/\-:;\'\"@#&+=<>]': '',  # Remove unusual special chars

        # Fix common spacing issues around colons (but don't change content)
        r'\s+:': ':',  # Remove space before colon
        r':\s{2,}': ': ',  # Normalize multiple spaces after colon to single space
    }

    def __init__(self, config_path, plan_folder, output_dir=None):
        """
        Initialize the text cleaner.

        Args:
            config_path: Path to TOML configuration file
            plan_folder: Path to plan folder containing documents
            output_dir: Directory to save cleaning reports (defaults to plan_folder/cleaning_reports)
        """
        self.config_path = config_path
        self.config = self.load_config()
        self.plan_folder = Path(plan_folder)

        # Set up output directory for cleaning reports
        if output_dir:
            self.output_dir = Path(output_dir)
        else:
            self.output_dir = self.plan_folder / "cleaning_reports"

        self.output_dir.mkdir(parents=True, exist_ok=True)

        # Track all changes made during cleaning
        self.cleaning_log = []

        logger.info(f"Initialized TextCleaner for config: {config_path}")

    def load_config(self):
        """Load TOML configuration file."""
        try:
            with open(self.config_path, 'r') as f:
                config = toml.load(f)
            logger.debug(f"Loaded config from {self.config_path}")
            return config
        except Exception as e:
            logger.error(f"Failed to load config {self.config_path}: {e}")
            raise

    def get_spell_check_keywords(self):
        """
        Get keywords that should be spell-checked from config.

        Returns:
            List of keywords to spell-check
        """
        keywords = []

        # Get element names (e.g., DOB, SSN, PHONE)
        if 'Extraction' in self.config and 'active_patterns' in self.config['Extraction']:
            keywords.extend(self.config['Extraction']['active_patterns'])

        # Get name anchors
        if 'Extraction' in self.config:
            if 'name_start_anchors' in self.config['Extraction']:
                keywords.extend(self.config['Extraction']['name_start_anchors'])
            if 'name_stop_anchors' in self.config['Extraction']:
                keywords.extend(self.config['Extraction']['name_stop_anchors'])

        # Add common document keywords if specified in config
        if 'Cleaning' in self.config and 'additional_keywords' in self.config['Cleaning']:
            keywords.extend(self.config['Cleaning']['additional_keywords'])

        logger.debug(f"Spell-check keywords: {keywords}")
        return list(set(keywords))  # Remove duplicates

    def apply_ocr_corrections(self, text, filename):
        """
        Apply common OCR error corrections.

        Args:
            text: Original text
            filename: Name of file being cleaned

        Returns:
            Corrected text
        """
        original_text = text
        changes = []

        for pattern, replacement in self.OCR_CORRECTIONS.items():
            matches = list(re.finditer(pattern, text))
            if matches:
                text = re.sub(pattern, replacement, text)
                for match in matches:
                    changes.append({
                        'type': 'OCR_CORRECTION',
                        'original': match.group(0),
                        'corrected': replacement,
                        'position': match.start(),
                        'pattern': pattern
                    })

        if changes:
            logger.debug(f"Applied {len(changes)} OCR corrections to {filename}")
            for change in changes:
                self.cleaning_log.append({
                    'filename': filename,
                    **change
                })

        return text

    def spell_check_keywords(self, text, filename, keywords, threshold=0.85):
        """
        Spell-check critical keywords using fuzzy matching.

        Args:
            text: Text to check
            filename: Name of file being cleaned
            keywords: List of keywords to check for
            threshold: Jaro-Winkler similarity threshold (0-1)

        Returns:
            Corrected text
        """
        changes = []

        # Find all capitalized words (potential keywords)
        words = re.findall(r'\b[A-Z][A-Z]+\b', text)

        for word in set(words):
            # Check if this word is close to any keyword
            for keyword in keywords:
                similarity = jellyfish.jaro_winkler_similarity(word.upper(), keyword.upper())

                if similarity >= threshold and word.upper() != keyword.upper():
                    # Found a likely misspelling
                    logger.debug(f"Spell-check: '{word}' -> '{keyword}' (similarity: {similarity:.2f})")

                    # Replace all occurrences
                    text = re.sub(r'\b' + re.escape(word) + r'\b', keyword.upper(), text)

                    changes.append({
                        'type': 'SPELL_CHECK',
                        'original': word,
                        'corrected': keyword.upper(),
                        'similarity': round(similarity, 3),
                        'keyword': keyword
                    })
                    break  # Only match to first keyword

        if changes:
            logger.info(f"Applied {len(changes)} spell corrections to {filename}")
            for change in changes:
                self.cleaning_log.append({
                    'filename': filename,
                    **change
                })

        return text

    def normalize_whitespace(self, text):
        """
        Normalize whitespace and formatting.

        Args:
            text: Text to normalize

        Returns:
            Normalized text
        """
        # Remove multiple spaces
        text = re.sub(r' +', ' ', text)

        # Normalize line endings
        text = re.sub(r'\r\n', '\n', text)

        # Remove trailing whitespace from lines
        text = re.sub(r' +\n', '\n', text)

        # Remove excessive blank lines (more than 2)
        text = re.sub(r'\n{3,}', '\n\n', text)

        return text

    def clean_document(self, file_path):
        """
        Clean a single document.

        Args:
            file_path: Path to document file

        Returns:
            Tuple of (cleaned_text, original_text, num_changes)
        """
        filename = os.path.basename(file_path)
        logger.info(f"Cleaning document: {filename}")

        try:
            # Read original text
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                original_text = f.read()

            if not original_text.strip():
                logger.warning(f"Empty document: {filename}")
                return "", original_text, 0

            # Track changes before cleaning
            initial_log_size = len(self.cleaning_log)

            # Apply cleaning steps
            text = original_text

            # Step 1: OCR corrections
            text = self.apply_ocr_corrections(text, filename)

            # Step 2: Spell-check keywords
            keywords = self.get_spell_check_keywords()
            if keywords:
                text = self.spell_check_keywords(text, filename, keywords)
            else:
                logger.warning(f"No keywords configured for spell-checking")

            # Step 3: Normalize whitespace
            text = self.normalize_whitespace(text)

            # Count changes made
            num_changes = len(self.cleaning_log) - initial_log_size

            if num_changes > 0:
                logger.info(f"Made {num_changes} changes to {filename}")
            else:
                logger.debug(f"No changes needed for {filename}")

            return text, original_text, num_changes

        except Exception as e:
            logger.error(f"Error cleaning {filename}: {e}", exc_info=True)
            return None, None, 0

    def clean_folder(self, source_folder, file_extensions=None, save_cleaned=True):
        """
        Clean all documents in a folder.

        Args:
            source_folder: Path to folder containing documents
            file_extensions: List of file extensions to process
            save_cleaned: Whether to save cleaned documents

        Returns:
            Dictionary with cleaning statistics
        """
        if file_extensions is None:
            file_extensions = ['.txt']

        folder = Path(source_folder)
        source_name = folder.name

        logger.info(f"Cleaning folder: {source_name}")

        # Get all files
        files = []
        for ext in file_extensions:
            files.extend(list(folder.glob(f"*{ext}")))

        if not files:
            logger.warning(f"No files found in {source_folder}")
            return {'files_processed': 0, 'total_changes': 0}

        logger.info(f"Found {len(files)} files to clean")

        # Create cleaned documents folder if saving
        if save_cleaned:
            cleaned_folder = folder / "cleaned"
            cleaned_folder.mkdir(exist_ok=True)
            logger.debug(f"Cleaned documents will be saved to: {cleaned_folder}")

        # Clean each file
        total_changes = 0
        files_with_changes = 0

        for file_path in files:
            cleaned_text, original_text, num_changes = self.clean_document(str(file_path))

            if cleaned_text is None:
                continue

            total_changes += num_changes
            if num_changes > 0:
                files_with_changes += 1

            # Save cleaned version
            if save_cleaned and cleaned_text:
                cleaned_file_path = cleaned_folder / file_path.name
                with open(cleaned_file_path, 'w', encoding='utf-8') as f:
                    f.write(cleaned_text)

        stats = {
            'source': source_name,
            'files_processed': len(files),
            'files_with_changes': files_with_changes,
            'total_changes': total_changes
        }

        logger.info(f"Cleaning complete for {source_name}: {total_changes} total changes across {files_with_changes} files")

        return stats

    def generate_cleaning_report(self, stats_list):
        """
        Generate Excel report of cleaning changes.

        Args:
            stats_list: List of statistics dictionaries from clean_folder

        Returns:
            Path to saved report
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        source_name = self.config['Document']['document_source']
        report_path = self.output_dir / f"{source_name}_cleaning_report_{timestamp}.xlsx"

        logger.info(f"Generating cleaning report: {report_path}")

        try:
            with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
                # Sheet 1: Detailed changes log
                if self.cleaning_log:
                    changes_df = pd.DataFrame(self.cleaning_log)
                    changes_df.to_excel(writer, sheet_name='Detailed Changes', index=False)
                    logger.debug(f"Wrote {len(self.cleaning_log)} changes to report")
                else:
                    # Empty sheet if no changes
                    pd.DataFrame({'Message': ['No changes were made']}).to_excel(
                        writer, sheet_name='Detailed Changes', index=False
                    )

                # Sheet 2: Summary statistics
                if stats_list:
                    stats_df = pd.DataFrame(stats_list)
                    stats_df.to_excel(writer, sheet_name='Summary', index=False)

                # Auto-adjust column widths
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width

            logger.info(f"Cleaning report saved: {report_path}")
            return str(report_path)

        except Exception as e:
            logger.error(f"Failed to generate cleaning report: {e}", exc_info=True)
            return None


def setup_logging(log_level=logging.INFO, log_file=None):
    """
    Configure logging for the cleaning module.

    Args:
        log_level: Logging level (DEBUG, INFO, WARNING, ERROR)
        log_file: Optional file path to save logs
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

    # File handler if specified
    if log_file:
        file_handler = logging.FileHandler(log_file)
        file_handler.setLevel(logging.DEBUG)  # Always log everything to file
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)


def main():
    """Example usage of TextCleaner."""
    # Example:
    # setup_logging(log_level=logging.INFO, log_file="cleaning.log")
    #
    # cleaner = TextCleaner(
    #     config_path="plans/plan_ABC/configs/Birth_Certificate.toml",
    #     plan_folder="plans/plan_ABC"
    # )
    #
    # stats = cleaner.clean_folder("plans/plan_ABC/Birth_Certificate")
    # cleaner.generate_cleaning_report([stats])
    pass


if __name__ == "__main__":
    main()
