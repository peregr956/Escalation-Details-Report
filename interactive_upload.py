"""Interactive file upload module for the Escalation Details Report.

This module provides interactive CLI prompts for selecting Excel data files
and configuration files when the report generator is launched without
command-line arguments.
"""
from pathlib import Path
from typing import List, Optional, Tuple
import os
import shlex


# Valid Excel extensions
VALID_EXCEL_EXTENSIONS = {'.xlsx', '.xls'}

# Default config file name
DEFAULT_CONFIG_FILE = 'client_config.yaml'


def validate_file_path(file_path: str, valid_extensions: Optional[set] = None) -> Tuple[bool, str]:
    """Validate that a file path exists and has a valid extension.
    
    Args:
        file_path: The file path to validate
        valid_extensions: Optional set of valid extensions (e.g., {'.xlsx', '.xls'})
        
    Returns:
        Tuple of (is_valid, error_message)
    """
    path = Path(file_path).expanduser().resolve()
    
    if not path.exists():
        return False, f"File not found: {path}"
    
    if not path.is_file():
        return False, f"Not a file: {path}"
    
    if valid_extensions:
        ext = path.suffix.lower()
        if ext not in valid_extensions:
            valid_str = ', '.join(valid_extensions)
            return False, f"Invalid file type '{ext}'. Expected: {valid_str}"
    
    return True, ""


def browse_file_dialog(
    title: str = "Select file",
    filetypes: Optional[List[Tuple[str, str]]] = None,
    multiple: bool = False
) -> Optional[List[str]]:
    """Open a file browser dialog using tkinter.
    
    Args:
        title: Dialog window title
        filetypes: List of (description, pattern) tuples for file filters
        multiple: Whether to allow multiple file selection
        
    Returns:
        List of selected file paths, or None if cancelled/unavailable
    """
    try:
        import tkinter as tk
        from tkinter import filedialog
        
        # Create hidden root window
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)  # Bring dialog to front
        
        if filetypes is None:
            filetypes = [("All files", "*.*")]
        
        if multiple:
            files = filedialog.askopenfilenames(
                title=title,
                filetypes=filetypes
            )
            result = list(files) if files else None
        else:
            file = filedialog.askopenfilename(
                title=title,
                filetypes=filetypes
            )
            result = [file] if file else None
        
        root.destroy()
        return result
        
    except ImportError:
        return None
    except Exception:
        return None


def parse_file_paths(input_str: str) -> List[str]:
    """Parse a string containing one or more file paths.
    
    Supports:
    - Space-separated paths
    - Comma-separated paths
    - Quoted paths with spaces
    - Mix of the above
    
    Args:
        input_str: Raw input string from user
        
    Returns:
        List of individual file paths
    """
    # First try to split by comma
    if ',' in input_str:
        paths = [p.strip() for p in input_str.split(',')]
        paths = [p for p in paths if p]  # Remove empty strings
    else:
        # Use shlex to handle quoted paths with spaces
        try:
            paths = shlex.split(input_str)
        except ValueError:
            # Fallback to simple split if shlex fails
            paths = input_str.split()
    
    # Expand user home directory and resolve paths
    resolved_paths = []
    for p in paths:
        path = Path(p).expanduser()
        resolved_paths.append(str(path))
    
    return resolved_paths


def prompt_for_excel_files() -> Optional[List[str]]:
    """Interactive prompt to select one or more Excel files.
    
    Presents the user with options to:
    1. Type file path(s) directly
    2. Open a file browser dialog (if tkinter available)
    3. Cancel
    
    Returns:
        List of validated Excel file paths, or None if cancelled
    """
    print("\n" + "=" * 60)
    print("  Escalation Details Report - Interactive File Selection")
    print("=" * 60)
    print("\nNo data files specified. Please select your Excel report file(s).")
    print("\nYou can provide 1-3 Excel files for the report:")
    print("  - 1 file: Current period only")
    print("  - 2-3 files: Include trend comparison (oldest to newest)")
    print()
    
    while True:
        print("Options:")
        print("  1. Type file path(s)")
        print("  2. Browse for files")
        print("  3. Cancel")
        print()
        
        choice = input("Select option [1/2/3]: ").strip()
        
        if choice == '3' or choice.lower() in ('q', 'quit', 'exit', 'cancel'):
            print("\nCancelled.")
            return None
        
        if choice == '2':
            # Try to open file browser
            print("\nOpening file browser...")
            files = browse_file_dialog(
                title="Select Excel Data File(s)",
                filetypes=[
                    ("Excel files", "*.xlsx"),
                    ("Excel files", "*.xls"),
                    ("All files", "*.*")
                ],
                multiple=True
            )
            
            if files is None:
                print("File browser not available. Please type the file path(s) instead.")
                choice = '1'  # Fall through to path input
            elif len(files) == 0:
                print("No files selected. Please try again.")
                continue
            else:
                # Validate selected files
                validated_files = []
                all_valid = True
                for f in files:
                    is_valid, error = validate_file_path(f, VALID_EXCEL_EXTENSIONS)
                    if is_valid:
                        validated_files.append(str(Path(f).resolve()))
                    else:
                        print(f"  Error: {error}")
                        all_valid = False
                
                if all_valid and validated_files:
                    return validated_files
                elif validated_files:
                    # Some files valid, some not
                    print(f"\nOnly {len(validated_files)} of {len(files)} files are valid.")
                    use_valid = input("Use valid files only? [y/n]: ").strip().lower()
                    if use_valid == 'y':
                        return validated_files
                continue
        
        if choice == '1' or choice == '':
            print("\nEnter path(s) to Excel file(s).")
            print("  - For multiple files, separate with commas or spaces")
            print("  - Quote paths containing spaces")
            print("  - Example: data/aug.xlsx data/sep.xlsx data/oct.xlsx")
            print()
            
            path_input = input("File path(s): ").strip()
            
            if not path_input:
                print("No input provided. Please try again.")
                continue
            
            # Parse and validate paths
            paths = parse_file_paths(path_input)
            
            if not paths:
                print("Could not parse file paths. Please try again.")
                continue
            
            validated_files = []
            all_valid = True
            
            for p in paths:
                is_valid, error = validate_file_path(p, VALID_EXCEL_EXTENSIONS)
                if is_valid:
                    validated_files.append(str(Path(p).resolve()))
                else:
                    print(f"  Error: {error}")
                    all_valid = False
            
            if all_valid and validated_files:
                return validated_files
            elif validated_files:
                print(f"\nOnly {len(validated_files)} of {len(paths)} files are valid.")
                use_valid = input("Use valid files only? [y/n]: ").strip().lower()
                if use_valid == 'y':
                    return validated_files
            else:
                print("\nNo valid files found. Please try again.")
            continue
        
        print("Invalid option. Please enter 1, 2, or 3.")


def prompt_for_config_file() -> Optional[str]:
    """Interactive prompt to select a configuration file.
    
    Checks for default config file and offers to use it.
    
    Returns:
        Path to config file, or None to use defaults
    """
    default_config = Path(DEFAULT_CONFIG_FILE)
    
    print("\n" + "-" * 60)
    print("  Configuration File Selection")
    print("-" * 60)
    
    # Check if default config exists
    if default_config.exists():
        print(f"\nDefault config file found: {DEFAULT_CONFIG_FILE}")
        use_default = input("Use this config file? [Y/n]: ").strip().lower()
        
        if use_default == '' or use_default == 'y' or use_default == 'yes':
            return str(default_config.resolve())
    else:
        print(f"\nNo default config file ({DEFAULT_CONFIG_FILE}) found.")
    
    print("\nOptions:")
    print("  1. Type config file path")
    print("  2. Browse for config file")
    print("  3. Use built-in defaults (no config file)")
    print()
    
    while True:
        choice = input("Select option [1/2/3]: ").strip()
        
        if choice == '3' or choice == '':
            print("Using built-in defaults.")
            return None
        
        if choice == '2':
            print("\nOpening file browser...")
            files = browse_file_dialog(
                title="Select Configuration File",
                filetypes=[
                    ("YAML files", "*.yaml"),
                    ("YAML files", "*.yml"),
                    ("All files", "*.*")
                ],
                multiple=False
            )
            
            if files is None:
                print("File browser not available. Please type the file path instead.")
                choice = '1'
            elif len(files) == 0:
                print("No file selected.")
                continue
            else:
                config_path = files[0]
                is_valid, error = validate_file_path(config_path)
                if is_valid:
                    return str(Path(config_path).resolve())
                else:
                    print(f"  Error: {error}")
                continue
        
        if choice == '1':
            path_input = input("Config file path: ").strip()
            
            if not path_input:
                print("No input provided.")
                continue
            
            is_valid, error = validate_file_path(path_input)
            if is_valid:
                return str(Path(path_input).expanduser().resolve())
            else:
                print(f"  Error: {error}")
                continue
        
        print("Invalid option. Please enter 1, 2, or 3.")


def confirm_selection(excel_files: List[str], config_file: Optional[str]) -> bool:
    """Display selected files and ask for confirmation.
    
    Args:
        excel_files: List of Excel file paths
        config_file: Config file path or None
        
    Returns:
        True if user confirms, False otherwise
    """
    print("\n" + "=" * 60)
    print("  Confirm File Selection")
    print("=" * 60)
    
    print("\nExcel data file(s):")
    for i, f in enumerate(excel_files, 1):
        period_label = "current period" if i == len(excel_files) else f"period -{len(excel_files) - i}"
        print(f"  {i}. {f} ({period_label})")
    
    print("\nConfiguration:")
    if config_file:
        print(f"  {config_file}")
    else:
        print("  Using built-in defaults")
    
    print()
    confirm = input("Proceed with these files? [Y/n]: ").strip().lower()
    
    return confirm == '' or confirm == 'y' or confirm == 'yes'


def interactive_file_selection() -> Tuple[Optional[List[str]], Optional[str]]:
    """Main entry point for interactive file selection.
    
    Orchestrates the full interactive flow:
    1. Prompt for Excel files
    2. Prompt for config file
    3. Confirm selection
    
    Returns:
        Tuple of (excel_files, config_file) or (None, None) if cancelled
    """
    # Get Excel files
    excel_files = prompt_for_excel_files()
    if not excel_files:
        return None, None
    
    # Get config file
    config_file = prompt_for_config_file()
    
    # Confirm selection
    if not confirm_selection(excel_files, config_file):
        print("\nCancelled. Run again to select different files.")
        return None, None
    
    return excel_files, config_file


if __name__ == "__main__":
    # Test the interactive selection
    print("Testing interactive file selection...")
    excel_files, config_file = interactive_file_selection()
    
    if excel_files:
        print("\n\nSelected files:")
        print(f"  Excel: {excel_files}")
        print(f"  Config: {config_file}")
    else:
        print("\n\nNo files selected.")
