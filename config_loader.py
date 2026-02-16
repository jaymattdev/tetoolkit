"""
Master Configuration Loader

Loads and validates the master configuration file for the extraction tool.
Provides default values for all settings if config file is missing or incomplete.
"""

import logging
from pathlib import Path
from typing import Any, Dict, Optional
import toml

logger = logging.getLogger(__name__)

# Default configuration values
DEFAULTS = {
    'SharePoint': {
        'base_url': 'https://yourcompany.sharepoint.com/sites/yoursite/',
        'hyperlink_style': 'short',
    },
    'Defaults': {
        'patterns_file': 'patterns_example.py',
        'log_level': 'INFO',
        'skip_cleaning': False,
        'skip_report': False,
    },
    'Cleaning': {
        'spell_check_threshold': 0.85,
    },
    'Validation': {
        'positional_outlier_threshold': 3.0,
        'within_document_gap_threshold': 2000,
    },
    'FileTypes': {
        'supported_extensions': ['.txt'],
        'output_link_extension': '.pdf',
    },
    'Output': {
        'include_extraction_order': True,
        'include_extraction_position': False,
        'include_flags': True,
        'include_flag_reasons': True,
    },
}


class MasterConfig:
    """
    Master configuration manager for the extraction tool.

    Loads settings from master_config.toml and provides default values
    for any missing settings.
    """

    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize the configuration manager.

        Args:
            config_path: Path to master_config.toml. If None, searches in
                        current directory and parent directories.
        """
        self._config: Dict[str, Any] = {}
        self._config_path: Optional[Path] = None

        # Load configuration
        self._load_config(config_path)

    def _find_config_file(self) -> Optional[Path]:
        """Search for master_config.toml in current and parent directories."""
        search_paths = [
            Path.cwd() / 'master_config.toml',
            Path(__file__).parent / 'master_config.toml',
        ]

        for path in search_paths:
            if path.exists():
                return path

        return None

    def _load_config(self, config_path: Optional[str] = None):
        """Load configuration from TOML file."""
        if config_path:
            path = Path(config_path)
        else:
            path = self._find_config_file()

        if path and path.exists():
            try:
                self._config = toml.load(path)
                self._config_path = path
                logger.info(f"Loaded master config from: {path}")
            except Exception as e:
                logger.warning(f"Error loading config from {path}: {e}")
                logger.info("Using default configuration values")
                self._config = {}
        else:
            logger.info("No master_config.toml found. Using default values.")
            self._config = {}

    def _get_nested(self, section: str, key: str, default: Any = None) -> Any:
        """Get a nested configuration value with fallback to defaults."""
        # Try to get from loaded config
        if section in self._config and key in self._config[section]:
            return self._config[section][key]

        # Fall back to defaults
        if section in DEFAULTS and key in DEFAULTS[section]:
            return DEFAULTS[section][key]

        return default

    # ====================
    # SHAREPOINT SETTINGS
    # ====================

    @property
    def sharepoint_base_url(self) -> str:
        """Base URL for SharePoint document links."""
        return self._get_nested('SharePoint', 'base_url')

    @property
    def hyperlink_style(self) -> str:
        """Hyperlink display style: 'full' or 'short'."""
        style = self._get_nested('SharePoint', 'hyperlink_style')
        if style not in ('full', 'short'):
            logger.warning(f"Invalid hyperlink_style '{style}', using 'short'")
            return 'short'
        return style

    # ====================
    # DEFAULT SETTINGS
    # ====================

    @property
    def default_patterns_file(self) -> str:
        """Default patterns file path."""
        return self._get_nested('Defaults', 'patterns_file')

    @property
    def default_log_level(self) -> str:
        """Default logging level."""
        return self._get_nested('Defaults', 'log_level')

    @property
    def default_skip_cleaning(self) -> bool:
        """Default value for skip_cleaning flag."""
        return self._get_nested('Defaults', 'skip_cleaning')

    @property
    def default_skip_report(self) -> bool:
        """Default value for skip_report flag."""
        return self._get_nested('Defaults', 'skip_report')

    # ====================
    # CLEANING SETTINGS
    # ====================

    @property
    def spell_check_threshold(self) -> float:
        """Jaro-Winkler similarity threshold for spell-checking."""
        return self._get_nested('Cleaning', 'spell_check_threshold')

    # ====================
    # VALIDATION SETTINGS
    # ====================

    @property
    def positional_outlier_threshold(self) -> float:
        """Z-score threshold for positional outlier detection."""
        return self._get_nested('Validation', 'positional_outlier_threshold')

    @property
    def within_document_gap_threshold(self) -> int:
        """Maximum character gap between sequential elements."""
        return self._get_nested('Validation', 'within_document_gap_threshold')

    # ====================
    # FILE TYPE SETTINGS
    # ====================

    @property
    def supported_extensions(self) -> list:
        """List of supported file extensions for extraction."""
        return self._get_nested('FileTypes', 'supported_extensions')

    @property
    def output_link_extension(self) -> str:
        """File extension for output document links."""
        return self._get_nested('FileTypes', 'output_link_extension')

    # ====================
    # OUTPUT SETTINGS
    # ====================

    @property
    def include_extraction_order(self) -> bool:
        """Include extraction_order column in output."""
        return self._get_nested('Output', 'include_extraction_order')

    @property
    def include_extraction_position(self) -> bool:
        """Include extraction_position column in output."""
        return self._get_nested('Output', 'include_extraction_position')

    @property
    def include_flags(self) -> bool:
        """Include flags column in output."""
        return self._get_nested('Output', 'include_flags')

    @property
    def include_flag_reasons(self) -> bool:
        """Include flag_reasons column in output."""
        return self._get_nested('Output', 'include_flag_reasons')

    # ====================
    # UTILITY METHODS
    # ====================

    @property
    def config_path(self) -> Optional[Path]:
        """Path to the loaded configuration file."""
        return self._config_path

    def get_raw_config(self) -> Dict[str, Any]:
        """Get the raw configuration dictionary."""
        return self._config.copy()

    def print_summary(self):
        """Print a summary of current configuration."""
        print("\nMaster Configuration:")
        print("-" * 40)

        if self._config_path:
            print(f"  Config file: {self._config_path}")
        else:
            print("  Config file: (using defaults)")

        print(f"\n  SharePoint:")
        print(f"    Base URL: {self.sharepoint_base_url}")
        print(f"    Hyperlink style: {self.hyperlink_style}")

        print(f"\n  Defaults:")
        print(f"    Patterns file: {self.default_patterns_file}")
        print(f"    Log level: {self.default_log_level}")

        print(f"\n  Cleaning:")
        print(f"    Spell check threshold: {self.spell_check_threshold}")

        print(f"\n  Validation:")
        print(f"    Positional outlier threshold: {self.positional_outlier_threshold}")
        print(f"    Within-document gap threshold: {self.within_document_gap_threshold}")

        print("-" * 40)


# Global singleton instance
_master_config: Optional[MasterConfig] = None


def get_master_config(config_path: Optional[str] = None) -> MasterConfig:
    """
    Get the master configuration instance.

    Uses a singleton pattern - first call loads the config,
    subsequent calls return the same instance.

    Args:
        config_path: Optional path to config file (only used on first call)

    Returns:
        MasterConfig instance
    """
    global _master_config
    if _master_config is None:
        _master_config = MasterConfig(config_path)
    return _master_config


def reload_master_config(config_path: Optional[str] = None) -> MasterConfig:
    """
    Force reload of master configuration.

    Args:
        config_path: Optional path to config file

    Returns:
        New MasterConfig instance
    """
    global _master_config
    _master_config = MasterConfig(config_path)
    return _master_config


def main():
    """Test configuration loading."""
    config = get_master_config()
    config.print_summary()


if __name__ == "__main__":
    main()
