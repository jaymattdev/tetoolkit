#!/usr/bin/env python3
"""
Interactive Terminal UI for Text Extraction Tool

This provides a simplified, interactive interface for running the extraction workflow.
Makes setup as easy as possible for team members without needing to remember command-line arguments.
"""

import os
import sys
from pathlib import Path
from typing import Optional, List


def clear_screen():
    """Clear the terminal screen."""
    os.system('clear' if os.name != 'nt' else 'cls')


def print_header():
    """Print the application header."""
    clear_screen()
    print("=" * 80)
    print(" " * 20 + "TEXT EXTRACTION TOOL - INTERACTIVE MODE")
    print("=" * 80)
    print()


def print_section(title: str):
    """Print a section header."""
    print()
    print("-" * 80)
    print(f"  {title}")
    print("-" * 80)


def get_input(prompt: str, default: Optional[str] = None) -> str:
    """
    Get user input with optional default value.

    Args:
        prompt: The prompt to display
        default: Default value if user presses Enter

    Returns:
        User input or default value
    """
    if default:
        full_prompt = f"{prompt} [{default}]: "
    else:
        full_prompt = f"{prompt}: "

    value = input(full_prompt).strip()

    if not value and default:
        return default

    return value


def get_yes_no(prompt: str, default: bool = True) -> bool:
    """
    Get yes/no input from user.

    Args:
        prompt: The prompt to display
        default: Default value (True for yes, False for no)

    Returns:
        True for yes, False for no
    """
    default_str = "Y/n" if default else "y/N"
    response = input(f"{prompt} [{default_str}]: ").strip().lower()

    if not response:
        return default

    return response in ['y', 'yes', 'true', '1']


def list_plans() -> List[str]:
    """
    List all available plans in the plans/ directory.

    Returns:
        List of plan names
    """
    plans_dir = Path("plans")

    if not plans_dir.exists():
        return []

    plans = [p.name for p in plans_dir.iterdir() if p.is_dir() and not p.name.startswith('.')]
    return sorted(plans)


def list_pattern_files() -> List[str]:
    """
    List all Python files that could be pattern repositories.

    Returns:
        List of pattern file paths
    """
    pattern_files = []

    # Look in current directory
    for file in Path(".").glob("*pattern*.py"):
        if file.name != "__pycache__":
            pattern_files.append(str(file))

    return sorted(pattern_files)


def select_from_list(items: List[str], item_type: str, allow_new: bool = False) -> str:
    """
    Let user select from a list or create new.

    Args:
        items: List of items to choose from
        item_type: Type of item (for display)
        allow_new: Allow user to enter a new value

    Returns:
        Selected or entered value
    """
    if not items and not allow_new:
        print(f"  ⚠ No {item_type}s found!")
        return ""

    if items:
        print(f"\n  Available {item_type}s:")
        for i, item in enumerate(items, 1):
            print(f"    {i}. {item}")

        if allow_new:
            print(f"    0. Enter custom {item_type}")

    if allow_new and not items:
        return get_input(f"  Enter {item_type}")

    while True:
        choice = input(f"\n  Select {item_type} [1-{len(items)}]: ").strip()

        if allow_new and choice == "0":
            return get_input(f"  Enter custom {item_type}")

        try:
            index = int(choice) - 1
            if 0 <= index < len(items):
                return items[index]
            else:
                print(f"  ⚠ Please enter a number between 1 and {len(items)}")
        except ValueError:
            print(f"  ⚠ Please enter a valid number")


def validate_plan_structure(plan_name: str) -> bool:
    """
    Check if the plan has the expected structure.

    Args:
        plan_name: Name of the plan

    Returns:
        True if structure looks good, False otherwise
    """
    plan_path = Path("plans") / plan_name

    if not plan_path.exists():
        print(f"  ⚠ Plan folder not found: {plan_path}")
        return False

    # Check for extraction_specs.xlsx or configs folder
    specs_file = plan_path / "extraction_specs.xlsx"
    configs_folder = plan_path / "configs"

    has_specs = specs_file.exists()
    has_configs = configs_folder.exists() and any(configs_folder.glob("*.toml"))

    if not has_specs and not has_configs:
        print(f"  ⚠ Warning: No extraction_specs.xlsx or config files found in {plan_name}")
        print(f"     Expected either:")
        print(f"       - {specs_file}")
        print(f"       - Config files in {configs_folder}")
        return False

    # Check for document folders
    doc_folders = [f for f in plan_path.iterdir()
                   if f.is_dir() and f.name not in ['configs', 'output', 'cleaned', '.git']]

    if not doc_folders:
        print(f"  ⚠ Warning: No document folders found in {plan_name}")
        print(f"     Expected folders like: Birth_Certificate, Employment_Application, etc.")
        return False

    print(f"  ✓ Plan structure looks good!")
    print(f"    - Document folders: {', '.join([f.name for f in doc_folders])}")
    if has_specs:
        print(f"    - Specifications: extraction_specs.xlsx")
    if has_configs:
        config_count = len(list(configs_folder.glob("*.toml")))
        print(f"    - Configurations: {config_count} TOML files")

    return True


def show_summary(plan: str, patterns: str, skip_cleaning: bool, log_level: str):
    """Show configuration summary before running."""
    print_section("CONFIGURATION SUMMARY")
    print(f"  Plan:           {plan}")
    print(f"  Patterns:       {patterns}")
    print(f"  Skip cleaning:  {'Yes' if skip_cleaning else 'No'}")
    print(f"  Log level:      {log_level}")
    print()


def main():
    """Main interactive interface."""
    print_header()

    print("  Welcome! This interactive tool will help you run the text extraction workflow.")
    print("  Just answer a few questions and we'll handle the rest.")
    print()

    # Step 1: Select plan
    print_section("STEP 1: Select Plan")

    available_plans = list_plans()

    if available_plans:
        plan_name = select_from_list(available_plans, "plan", allow_new=True)
    else:
        print("  ⚠ No plans found in plans/ directory")
        plan_name = get_input("  Enter plan name")

    if not plan_name:
        print("\n  ⚠ Error: Plan name is required")
        sys.exit(1)

    # Validate plan structure
    print()
    print(f"  Checking plan structure for '{plan_name}'...")
    if not validate_plan_structure(plan_name):
        if not get_yes_no("\n  Continue anyway?", default=False):
            print("\n  Exiting...")
            sys.exit(0)

    # Step 2: Select pattern repository
    print_section("STEP 2: Select Pattern Repository")

    available_patterns = list_pattern_files()

    if available_patterns:
        patterns_file = select_from_list(available_patterns, "pattern file", allow_new=True)
    else:
        patterns_file = get_input("  Enter path to patterns file", default="patterns_example.py")

    if not patterns_file:
        print("  ⚠ Warning: No patterns file specified. Using default: patterns_example.py")
        patterns_file = "patterns_example.py"

    # Validate patterns file exists
    if not Path(patterns_file).exists():
        print(f"\n  ⚠ Warning: Pattern file not found: {patterns_file}")
        if not get_yes_no("  Continue anyway?", default=False):
            print("\n  Exiting...")
            sys.exit(0)

    # Step 3: Additional options
    print_section("STEP 3: Additional Options")

    skip_cleaning = get_yes_no("  Skip text cleaning step? (useful if already cleaned)", default=False)

    print("\n  Log level options:")
    print("    1. INFO    - Standard output (recommended)")
    print("    2. DEBUG   - Detailed output (for troubleshooting)")
    print("    3. WARNING - Minimal output (warnings and errors only)")
    print("    4. ERROR   - Errors only")

    log_choice = get_input("  Select log level", default="1")
    log_levels = {"1": "INFO", "2": "DEBUG", "3": "WARNING", "4": "ERROR"}
    log_level = log_levels.get(log_choice, "INFO")

    # Step 4: Confirm and run
    print_section("READY TO RUN")

    show_summary(plan_name, patterns_file, skip_cleaning, log_level)

    if not get_yes_no("  Start extraction?", default=True):
        print("\n  Cancelled by user. Exiting...")
        sys.exit(0)

    # Build command
    cmd_parts = [
        sys.executable,  # Python interpreter
        "main.py",
        "--plan", plan_name,
        "--patterns", patterns_file,
        "--log-level", log_level
    ]

    if skip_cleaning:
        cmd_parts.append("--skip-cleaning")

    command = " ".join(cmd_parts)

    print()
    print("=" * 80)
    print("  STARTING EXTRACTION WORKFLOW")
    print("=" * 80)
    print()
    print(f"  Running: {command}")
    print()
    print("=" * 80)
    print()

    # Run the command
    exit_code = os.system(command)

    # Show completion message
    print()
    print("=" * 80)
    if exit_code == 0:
        print("  ✓ EXTRACTION COMPLETED SUCCESSFULLY!")
        print(f"  Check results in: plans/{plan_name}/output/")
    else:
        print("  ⚠ EXTRACTION COMPLETED WITH ERRORS")
        print(f"  Review the output above for details")
    print("=" * 80)
    print()

    sys.exit(exit_code)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n  ⚠ Cancelled by user (Ctrl+C)")
        print("\n  Exiting...\n")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n  ⚠ Unexpected error: {e}")
        print("\n  Please report this issue if it persists.\n")
        sys.exit(1)
