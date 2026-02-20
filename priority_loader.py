"""
Source Priority Loader

Loads the element-level source priority mapping from source_priority.xlsx.
This file defines which source takes precedence for each element when
choosing the "best" value across multiple document sources.
"""

import json
import logging
import pandas as pd
from pathlib import Path
from typing import Dict, List

logger = logging.getLogger(__name__)

PRIORITY_FILENAME = "source_priority.xlsx"


def load_source_priority(plan_folder: str) -> Dict[str, List[str]]:
    """
    Load source priority mapping from source_priority.xlsx.

    Excel format:
        Element          | Source Priority
        DOB              | ["Birth_Certificate", "Employment_App"]
        SSN              | ["SSN_Card", "Employment_App"]

    Args:
        plan_folder: Path to the plan folder containing source_priority.xlsx

    Returns:
        Dict mapping element name (uppercase) -> ordered list of sources,
        highest priority first. Returns empty dict if file not found.
    """
    priority_path = Path(plan_folder) / PRIORITY_FILENAME

    if not priority_path.exists():
        logger.warning(
            f"No {PRIORITY_FILENAME} found in {plan_folder}. "
            "Best Data tab will use first available source when values conflict."
        )
        return {}

    try:
        df = pd.read_excel(priority_path)

        if 'Element' not in df.columns or 'Source Priority' not in df.columns:
            logger.error(
                f"{PRIORITY_FILENAME} must have columns 'Element' and 'Source Priority'"
            )
            return {}

        priority_map = {}

        for _, row in df.iterrows():
            element = str(row['Element']).strip().upper()
            sources_raw = row['Source Priority']

            try:
                if isinstance(sources_raw, str):
                    sources = json.loads(sources_raw)
                elif isinstance(sources_raw, list):
                    sources = sources_raw
                else:
                    sources = []
            except (json.JSONDecodeError, ValueError):
                logger.warning(
                    f"Could not parse Source Priority for element '{element}': {sources_raw}"
                )
                sources = []

            if element and sources:
                priority_map[element] = [str(s).strip() for s in sources]

        logger.info(f"Loaded source priority for {len(priority_map)} elements from {priority_path.name}")
        return priority_map

    except Exception as e:
        logger.error(f"Error loading {PRIORITY_FILENAME}: {e}", exc_info=True)
        return {}


def get_highest_priority_source(sources_present: list, element: str, priority_map: Dict[str, List[str]]) -> str:
    """
    Given a list of sources that have a value for this element,
    return whichever is highest priority according to the map.

    Falls back to the first source in the list if no priority defined.

    Args:
        sources_present: List of source names that have extracted this element
        element: Element name (will be uppercased for lookup)
        priority_map: Dict from load_source_priority()

    Returns:
        The highest priority source name
    """
    priority_list = priority_map.get(element.upper(), [])

    for source in priority_list:
        if source in sources_present:
            return source

    # Fallback: no priority defined, use first
    return sources_present[0]
