"""
Text Extraction Module
Performs raw extraction of elements from text documents using regex patterns.
"""

import re
import os
import logging
import glob
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import toml

# Configure module logger
logger = logging.getLogger(__name__)


class TextExtractor:
    """Extract elements from text documents using regex patterns."""

    def __init__(self, config_path, patterns_dict=None, pdf_source_path=None):
        """
        Initialize the extractor with a configuration file.

        Args:
            config_path: Path to TOML configuration file
            patterns_dict: Dictionary mapping element names to regex patterns.
                          Can be a single pattern string or a list of pattern strings.
            pdf_source_path: Path to PDF source files for filename matching.
                            If None, loads from master config.
        """
        self.config_path = config_path
        self.config = self.load_config()
        self.patterns_dict = patterns_dict or {}

        # Load PDF source path from master config if not provided
        if pdf_source_path is None:
            from config_loader import get_master_config
            master_config = get_master_config()
            self.pdf_source_path = master_config.pdf_source_path
        else:
            self.pdf_source_path = pdf_source_path

        logger.debug(f"Initialized TextExtractor with config: {config_path}")

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

    def load_text_file(self, file_path):
        """
        Load text content from a file.

        Args:
            file_path: Path to text file

        Returns:
            String content of the file
        """
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            logger.debug(f"Loaded {len(content)} characters from {file_path}")
            return content
        except Exception as e:
            logger.error(f"Error reading {file_path}: {e}")
            print(f"Error reading {file_path}: {e}")
            return ""

    def _is_compiled_pattern(self, pattern):
        """
        Check if a pattern is a compiled regex pattern.

        Args:
            pattern: Pattern to check (string or compiled regex)

        Returns:
            True if pattern is compiled, False otherwise
        """
        return hasattr(pattern, 'pattern') and hasattr(pattern, 'finditer')

    def _normalize_pattern(self, pattern):
        """
        Normalize pattern to compiled regex.

        If pattern is already compiled, return as-is.
        If pattern is a string, compile it with IGNORECASE and MULTILINE flags.

        Args:
            pattern: Raw string pattern or compiled regex pattern

        Returns:
            Compiled regex pattern
        """
        if self._is_compiled_pattern(pattern):
            # Already compiled - return as-is
            return pattern
        elif isinstance(pattern, str):
            # Compile with default flags
            return re.compile(pattern, re.IGNORECASE | re.MULTILINE)
        else:
            raise ValueError(f"Invalid pattern type: {type(pattern)}. Expected string or compiled regex.")

    def extract_with_pattern(self, text, pattern, element_name):
        """
        Extract all instances of a pattern from text.

        Supports both raw string patterns and pre-compiled regex patterns.

        Args:
            text: Source text to search
            pattern: Regex pattern to use (string or compiled regex)
            element_name: Name of the element being extracted

        Returns:
            List of tuples: (value, start_position, end_position, order)
        """
        results = []

        try:
            # Normalize pattern to compiled regex
            compiled_pattern = self._normalize_pattern(pattern)

            # Use finditer to get all matches with positions
            for order, match in enumerate(compiled_pattern.finditer(text)):
                # Check if pattern has capture groups
                if match.lastindex and match.lastindex > 0:
                    # Use first capture group if available
                    value = match.group(1)
                else:
                    # Use full match if no capture groups
                    value = match.group(0)

                start_pos = match.start()
                end_pos = match.end()

                results.append({
                    'element': element_name,
                    'value': value,
                    'extraction_order': order + 1,
                    'extraction_position': start_pos,
                    'end_position': end_pos
                })

        except Exception as e:
            logger.error(f"Error extracting {element_name} with pattern {pattern}: {e}")
            print(f"Error extracting {element_name}: {e}")

        return results

    def extract_id_from_filename(self, filename: str) -> Optional[str]:
        """
        Extract participant ID (SSN) from filename.

        Supports two filename formats:
        - PDF format: 2_SSN_DOCID_SOURCENAME_Page_N.pdf -> SSN (2nd part)
        - Text format: SSN_DOCID_1_Page_N.txt -> SSN (1st part)

        Args:
            filename: Name of the file

        Returns:
            Participant ID (SSN) as string, or None if not found
        """
        # Remove extension
        base_name = filename.rsplit('.', 1)[0]
        parts = base_name.split('_')

        if len(parts) < 2:
            logger.debug(f"Filename '{filename}' has too few parts")
            return None

        # PDF format: 2_SSN_DOCID_SOURCENAME_Page_N.pdf
        # First part is "2", second part is SSN
        if parts[0] == '2' and len(parts) >= 3:
            ssn = parts[1]
            # Validate SSN is 9 digits
            if re.match(r'^\d{9}$', ssn):
                logger.debug(f"Extracted SSN from PDF filename: {ssn}")
                return ssn

        # Text format: SSN_DOCID_1_Page_N.txt
        # First part is SSN
        ssn = parts[0]
        if re.match(r'^\d{9}$', ssn):
            logger.debug(f"Extracted SSN from text filename: {ssn}")
            return ssn

        logger.debug(f"No valid SSN pattern matched for filename: {filename}")
        return None

    def find_matching_pdf(self, text_filename: str, source: str) -> Optional[str]:
        """
        Find the matching PDF file for a text file.

        Text filename format: SSN_DOCID_1_Page_PAGENUMBER.txt
        PDF filename format:  2_SSN_DOCID_SOURCENAME_Page_PAGENUMBER.pdf

        Args:
            text_filename: Name of the text file (e.g., "000000000_00000000_1_Page_5.txt")
            source: Document source/type name (e.g., "Birth_Certificate")

        Returns:
            PDF filename if found, None otherwise
        """
        if not self.pdf_source_path:
            logger.debug("No PDF source path configured, skipping PDF matching")
            return None

        # Parse text filename: SSN_DOCID_1_Page_PAGENUMBER.txt
        # Split by underscore and extract parts
        base_name = text_filename.rsplit('.', 1)[0]  # Remove extension
        parts = base_name.split('_')

        if len(parts) < 5:
            logger.warning(f"Text filename '{text_filename}' doesn't match expected format SSN_DOCID_1_Page_N")
            return None

        # Extract SSN and DOCID (first two parts)
        ssn = parts[0]
        docid = parts[1]
        # parts[2] is the "1" we skip
        # parts[3] should be "Page"
        # parts[4] should be page number

        # Find page info (everything from "Page" onwards)
        try:
            page_idx = parts.index('Page')
            page_info = '_'.join(parts[page_idx:])  # "Page_5"
        except ValueError:
            logger.warning(f"Could not find 'Page' in filename '{text_filename}'")
            return None

        # Build expected PDF filename pattern: 2_SSN_DOCID_SOURCENAME_Page_PAGENUMBER.pdf
        expected_pdf_name = f"2_{ssn}_{docid}_{source}_{page_info}.pdf"

        # Check if PDF exists
        pdf_folder = Path(self.pdf_source_path) / source
        pdf_path = pdf_folder / expected_pdf_name

        if pdf_path.exists():
            logger.debug(f"Found matching PDF: {expected_pdf_name}")
            return expected_pdf_name
        else:
            # Try to find PDF with glob pattern in case of slight naming variations
            search_pattern = str(pdf_folder / f"2_{ssn}_{docid}_{source}_Page_*.pdf")
            matches = glob.glob(search_pattern)
            if matches:
                # Return just the filename, not full path
                pdf_name = os.path.basename(matches[0])
                logger.debug(f"Found matching PDF via glob: {pdf_name}")
                return pdf_name
            else:
                logger.warning(f"No matching PDF found for '{text_filename}' in {pdf_folder}")
                return None

    def extract_name(self, text, start_anchors, stop_anchors, name_prefixes=None):
        """
        Extract names using start and stop anchors.

        Anchors are paired by index:
        - start_anchors[0] pairs with stop_anchors[0]
        - start_anchors[1] pairs with stop_anchors[1]
        - etc.

        Args:
            text: Source text to search
            start_anchors: List of start anchor phrases
            stop_anchors: List of stop anchor phrases (must match length of start_anchors)
            name_prefixes: List of prefix types corresponding to each anchor pair
                          Options: "Name" (no prefix), "Spouse" (S), "Beneficiary" (B), "AP" (AP)

        Returns:
            List of extraction dictionaries
        """
        results = []
        extraction_order = 1

        # Validate that start and stop anchors have matching lengths
        if len(start_anchors) != len(stop_anchors):
            logger.warning(
                f"Name anchor mismatch: {len(start_anchors)} start anchors but {len(stop_anchors)} stop anchors. "
                f"Anchors should be paired by index."
            )
            print(
                f"Warning: {len(start_anchors)} start anchors but {len(stop_anchors)} stop anchors. "
                f"Using minimum length for pairing."
            )

        # Default prefixes if not provided
        if name_prefixes is None:
            name_prefixes = []

        # Pair anchors by index (zip stops at shortest list)
        for idx, (start_anchor, stop_anchor) in enumerate(zip(start_anchors, stop_anchors)):
            # Get the prefix for this anchor pair (default to "Name" if not specified)
            prefix_type = name_prefixes[idx] if idx < len(name_prefixes) else "Name"

            # Create dynamic pattern for name extraction
            # Pattern finds text between start anchor and stop anchor
            # Captures capitalized words that form a name
            # [:\.\,;]* handles optional punctuation after the start anchor

            pattern = rf'{re.escape(start_anchor)}[:\.\,;]*\s*([A-Z][a-zA-Z\'\-]*(?:\s+[A-Z][a-zA-Z\'\-]*)*)\s*(?={re.escape(stop_anchor)})'

            try:
                for match in re.finditer(pattern, text, re.MULTILINE):
                    name_value = match.group(1).strip()

                    # Basic validation: name should have at least 2 characters
                    if len(name_value) >= 2:
                        results.append({
                            'element': 'NAME',
                            'value': name_value,
                            'extraction_order': extraction_order,
                            'extraction_position': match.start(),
                            'end_position': match.end(),
                            'start_anchor': start_anchor,
                            'stop_anchor': stop_anchor,
                            'name_prefix': prefix_type
                        })
                        extraction_order += 1

            except Exception as e:
                print(f"Error in name extraction with anchors '{start_anchor}' -> '{stop_anchor}': {e}")
                continue

        return results

    def apply_duplicate_mappings(self, extractions):
        """
        Apply duplicate mappings to rename extraction_order >= 2 elements.

        When a document deliberately has multiple fields for the same element
        (e.g., participant DOB and spouse DOB), this method renames the second
        and subsequent occurrences based on the configured mappings.

        Config example:
            duplicate_mappings = {"DOB": "SDOB", "SSN": "SSSN"}

        This means:
            - DOB with extraction_order 1 stays as "DOB"
            - DOB with extraction_order 2 becomes "SDOB"
            - DOB with extraction_order 3 becomes "SDOB" (same as order 2)
            - DOB with extraction_order 4+ becomes "SDOB" (same as order 2)
            - SSN with extraction_order 1 stays as "SSN"
            - SSN with extraction_order 2+ becomes "SSSN"

        Args:
            extractions: List of extraction dictionaries

        Returns:
            List of extractions with duplicate mappings applied
        """
        # Get duplicate mappings from config
        duplicate_mappings = self.config.get('Extraction', {}).get('duplicate_mappings', {})

        if not duplicate_mappings:
            return extractions

        logger.debug(f"Applying duplicate mappings: {duplicate_mappings}")

        for extraction in extractions:
            element = extraction.get('element')
            extraction_order = extraction.get('extraction_order')

            # Check if this element has a duplicate mapping and is extraction_order >= 2
            if element in duplicate_mappings and extraction_order is not None and extraction_order >= 2:
                new_element = duplicate_mappings[element]
                logger.info(
                    f"Duplicate mapping: {element} (order {extraction_order}) -> {new_element} "
                    f"in {extraction.get('filename')}"
                )
                extraction['element'] = new_element
                extraction['original_element'] = element  # Keep track of original for reference

        return extractions

    def deduplicate_extractions(self, extractions):
        """
        Deduplicate extractions within the same document.

        For each element:
        - If all extractions have the same value, keep only one
        - If extractions have different values, keep all unique values

        Args:
            extractions: List of extraction dictionaries from same document

        Returns:
            Deduplicated list of extractions
        """
        from collections import defaultdict

        # Group extractions by element
        by_element = defaultdict(list)
        for extraction in extractions:
            element = extraction.get('element')
            by_element[element].append(extraction)

        deduplicated = []

        for element, element_extractions in by_element.items():
            # Get all non-null values
            values_with_extractions = []
            for ext in element_extractions:
                value = ext.get('value')
                if value is not None:
                    values_with_extractions.append((value, ext))

            # If no values found, keep one blank entry
            if not values_with_extractions:
                if element_extractions:
                    deduplicated.append(element_extractions[0])
                continue

            # Get unique values (case-sensitive comparison)
            unique_values = {}
            for value, ext in values_with_extractions:
                if value not in unique_values:
                    unique_values[value] = ext

            # Keep all unique values (validator will flag conflicts if multiple)
            deduplicated.extend(unique_values.values())

            # Log if multiple unique values found
            if len(unique_values) > 1:
                logger.info(
                    f"Multiple values for {element} in {element_extractions[0].get('filename')}: "
                    f"{list(unique_values.keys())}"
                )

        return deduplicated

    def extract_from_document(self, file_path):
        """
        Extract all configured elements from a single document.

        Args:
            file_path: Path to the document file

        Returns:
            List of extraction dictionaries
        """
        text = self.load_text_file(file_path)
        if not text:
            return []

        all_extractions = []
        filename = os.path.basename(file_path)
        source = self.config['Document']['document_source']

        # Find matching PDF file
        pdf_filename = self.find_matching_pdf(filename, source)

        # Extract participant ID from PDF filename (preferred) or fall back to text filename
        if pdf_filename:
            participant_id = self.extract_id_from_filename(pdf_filename)
        else:
            participant_id = self.extract_id_from_filename(filename)

        # Extract regular elements using patterns
        active_patterns = self.config['Extraction']['active_patterns']

        for element_name in active_patterns:
            # Get pattern(s) from repository - can be single pattern or list of patterns
            pattern_or_patterns = self.patterns_dict.get(element_name)

            if pattern_or_patterns:
                # Normalize to list of patterns (handles strings, compiled patterns, and lists)
                if isinstance(pattern_or_patterns, str) or self._is_compiled_pattern(pattern_or_patterns):
                    # Single pattern (string or compiled)
                    patterns = [pattern_or_patterns]
                elif isinstance(pattern_or_patterns, list):
                    # List of patterns (can be strings or compiled)
                    patterns = pattern_or_patterns
                else:
                    logger.warning(f"Invalid pattern type for '{element_name}': {type(pattern_or_patterns)}")
                    print(f"Warning: Invalid pattern type for '{element_name}': {type(pattern_or_patterns)}")
                    patterns = []

                # Try each pattern until we find matches
                element_results = []
                for pattern in patterns:
                    results = self.extract_with_pattern(text, pattern, element_name)
                    element_results.extend(results)
                    # If we found matches with this pattern, we can stop trying other patterns
                    # Comment out the break below if you want to try ALL patterns regardless
                    if results:
                        break

                if element_results:
                    for result in element_results:
                        result['filename'] = filename
                        result['pdf_filename'] = pdf_filename
                        result['source'] = source
                        result['participant_id'] = participant_id
                        all_extractions.append(result)
                else:
                    # No matches found with any pattern - add blank entry
                    all_extractions.append({
                        'filename': filename,
                        'pdf_filename': pdf_filename,
                        'source': source,
                        'participant_id': participant_id,
                        'element': element_name,
                        'value': None,
                        'extraction_order': None,
                        'extraction_position': None,
                        'end_position': None
                    })
            else:
                print(f"Warning: No pattern found for element '{element_name}'")
                # Add blank entry for missing pattern
                all_extractions.append({
                    'filename': filename,
                    'pdf_filename': pdf_filename,
                    'source': source,
                    'participant_id': participant_id,
                    'element': element_name,
                    'value': None,
                    'extraction_order': None,
                    'extraction_position': None,
                    'end_position': None
                })

        # Extract names if configured
        if self.config['Document']['name_extraction']:
            start_anchors = self.config['Extraction']['name_start_anchors']
            stop_anchors = self.config['Extraction']['name_stop_anchors']
            name_prefixes = self.config['Extraction'].get('name_prefixes', [])

            name_results = self.extract_name(text, start_anchors, stop_anchors, name_prefixes)

            if name_results:
                for result in name_results:
                    result['filename'] = filename
                    result['pdf_filename'] = pdf_filename
                    result['source'] = source
                    result['participant_id'] = participant_id
                    all_extractions.append(result)
            else:
                # No name found - add blank entry
                all_extractions.append({
                    'filename': filename,
                    'pdf_filename': pdf_filename,
                    'source': source,
                    'participant_id': participant_id,
                    'element': 'NAME',
                    'value': None,
                    'extraction_order': None,
                    'extraction_position': None,
                    'end_position': None
                })

        # Apply duplicate mappings BEFORE deduplication
        # This renames extraction_order >= 2 elements to their mapped names
        all_extractions = self.apply_duplicate_mappings(all_extractions)

        # Deduplicate extractions - keep only unique values per element
        all_extractions = self.deduplicate_extractions(all_extractions)

        return all_extractions

    def extract_from_folder(self, folder_path, file_extensions=None, use_cleaned=True):
        """
        Extract from all documents in a folder.

        Args:
            folder_path: Path to folder containing documents
            file_extensions: List of file extensions to process (e.g., ['.txt', '.pdf'])
                           If None, processes all .txt files
            use_cleaned: If True, use cleaned documents from 'cleaned' subfolder

        Returns:
            List of all extraction dictionaries
        """
        if file_extensions is None:
            file_extensions = ['.txt']

        all_extractions = []
        folder = Path(folder_path)

        if not folder.exists():
            logger.error(f"Folder not found: {folder_path}")
            print(f"Folder not found: {folder_path}")
            return all_extractions

        # Check if cleaned folder exists and use it if available
        cleaned_folder = folder / "cleaned"
        if use_cleaned and cleaned_folder.exists():
            extraction_folder = cleaned_folder
            logger.info(f"Using cleaned documents from: {cleaned_folder}")
            print(f"  Using cleaned documents")
        else:
            extraction_folder = folder
            if use_cleaned:
                logger.warning(f"Cleaned folder not found at {cleaned_folder}, using original documents")
            logger.info(f"Extracting from: {extraction_folder}")

        # Get all files with specified extensions
        files = []
        for ext in file_extensions:
            files.extend(list(extraction_folder.glob(f"*{ext}")))

        logger.info(f"Processing {len(files)} files from {extraction_folder}")
        print(f"Processing {len(files)} files from {folder_path}")

        for file_path in files:
            logger.debug(f"Extracting from: {file_path.name}")
            extractions = self.extract_from_document(str(file_path))
            all_extractions.extend(extractions)

        logger.info(f"Extracted {len(all_extractions)} items from {len(files)} files")
        return all_extractions


def main():
    """Example usage of TextExtractor."""
    # Example pattern dictionary (to be replaced with your pattern repository)
    # patterns = {
    #     'DOB': r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b',
    #     'SSN': r'\b\d{3}-\d{2}-\d{4}\b'
    # }
    #
    # extractor = TextExtractor("configs/Birth_Certificate.toml", patterns)
    # results = extractor.extract_from_folder("plan_ABC/Birth_Certificate")
    pass


if __name__ == "__main__":
    main()
