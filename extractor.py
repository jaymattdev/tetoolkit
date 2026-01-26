"""
Text Extraction Module
Performs raw extraction of elements from text documents using regex patterns.
"""

import re
import os
import logging
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import toml

# Configure module logger
logger = logging.getLogger(__name__)


class TextExtractor:
    """Extract elements from text documents using regex patterns."""

    def __init__(self, config_path, patterns_dict=None):
        """
        Initialize the extractor with a configuration file.

        Args:
            config_path: Path to TOML configuration file
            patterns_dict: Dictionary mapping element names to regex patterns.
                          Can be a single pattern string or a list of pattern strings.
        """
        self.config_path = config_path
        self.config = self.load_config()
        self.patterns_dict = patterns_dict or {}
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

    def extract_with_pattern(self, text, pattern, element_name):
        """
        Extract all instances of a pattern from text.

        Args:
            text: Source text to search
            pattern: Regex pattern to use
            element_name: Name of the element being extracted

        Returns:
            List of tuples: (value, start_position, end_position, order)
        """
        results = []

        try:
            # Use finditer to get all matches with positions
            for order, match in enumerate(re.finditer(pattern, text, re.IGNORECASE | re.MULTILINE)):
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
            print(f"Error extracting {element_name}: {e}")

        return results

    def extract_id_from_filename(self, filename: str) -> Optional[str]:
        """
        Extract participant ID from filename.

        BOILERPLATE: Customize this function to match your filename pattern.

        Args:
            filename: Name of the file (e.g., "participant_123_document.txt")

        Returns:
            Participant ID as string, or None if not found

        Example patterns to implement:
            - "SSN_123-45-6789_document.txt" -> "123-45-6789"
            - "participant_001234_birth_cert.txt" -> "001234"
            - "EMP-12345_application.txt" -> "EMP-12345"
            - "john_doe_123456789.txt" -> "123456789"
        """
        # BOILERPLATE: Replace this with your actual filename pattern
        # Example 1: Extract SSN pattern (###-##-####)
        # match = re.search(r'(\d{3}-\d{2}-\d{4})', filename)
        # if match:
        #     return match.group(1)

        # Example 2: Extract participant number after "participant_"
        # match = re.search(r'participant[_-](\d+)', filename, re.IGNORECASE)
        # if match:
        #     return match.group(1)

        # Example 3: Extract employee ID (EMP-#####)
        # match = re.search(r'(EMP-\d+)', filename, re.IGNORECASE)
        # if match:
        #     return match.group(1)

        # Default: return None if no pattern matches
        logger.debug(f"No participant ID pattern matched for filename: {filename}")
        return None

    def extract_name(self, text, start_anchors, stop_anchors):
        """
        Extract names using start and stop anchors.

        Args:
            text: Source text to search
            start_anchors: List of possible start anchor phrases
            stop_anchors: List of possible stop anchor phrases

        Returns:
            List of extraction dictionaries
        """
        results = []
        extraction_order = 1

        # Try each combination of start and stop anchors
        for start_anchor in start_anchors:
            for stop_anchor in stop_anchors:
                # Create dynamic pattern for name extraction
                # Pattern finds text between start anchor and stop anchor
                # Captures capitalized words that form a name

                pattern = rf'{re.escape(start_anchor)}\s*[:]*\s*([A-Z][a-zA-Z\'\-]*(?:\s+[A-Z][a-zA-Z\'\-]*)*)\s*(?={re.escape(stop_anchor)})'

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
                                'stop_anchor': stop_anchor
                            })
                            extraction_order += 1

                except Exception as e:
                    print(f"Error in name extraction with anchors '{start_anchor}' -> '{stop_anchor}': {e}")
                    continue

        return results

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

        # Extract participant ID from filename (if pattern is defined)
        participant_id = self.extract_id_from_filename(filename)

        # Extract regular elements using patterns
        active_patterns = self.config['Extraction']['active_patterns']

        for element_name in active_patterns:
            # Get pattern(s) from repository - can be single pattern or list of patterns
            pattern_or_patterns = self.patterns_dict.get(element_name)

            if pattern_or_patterns:
                # Normalize to list of patterns
                if isinstance(pattern_or_patterns, str):
                    patterns = [pattern_or_patterns]
                elif isinstance(pattern_or_patterns, list):
                    patterns = pattern_or_patterns
                else:
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
                        result['source'] = source
                        result['participant_id'] = participant_id
                        all_extractions.append(result)
                else:
                    # No matches found with any pattern - add blank entry
                    all_extractions.append({
                        'filename': filename,
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

            name_results = self.extract_name(text, start_anchors, stop_anchors)

            if name_results:
                for result in name_results:
                    result['filename'] = filename
                    result['source'] = source
                    result['participant_id'] = participant_id
                    all_extractions.append(result)
            else:
                # No name found - add blank entry
                all_extractions.append({
                    'filename': filename,
                    'source': source,
                    'participant_id': participant_id,
                    'element': 'NAME',
                    'value': None,
                    'extraction_order': None,
                    'extraction_position': None,
                    'end_position': None
                })

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
