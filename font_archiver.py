#!/usr/bin/env python3
"""
Font Archiver Script

This script scans for fonts on a Windows system, groups them by family,
zips them, and uploads them to a GitHub repository.
"""

import concurrent.futures
import datetime
import getpass
import glob
import logging
import multiprocessing
import os
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import zipfile
from logging import Logger
from typing import Dict, List, Set, Tuple, Optional, Any

import fontTools.ttLib as ttLib
from colorama import Fore, Style, init
from github import Github, GithubException

# Initialize colorama
init(autoreset=True)

# Verify Python 3
if sys.version_info.major < 3:
    print("Error: This script requires Python 3.")
    sys.exit(1)


# Custom formatter without microseconds
class NoMicrosecondsFormatter(logging.Formatter):
    def formatTime(self, record, _date_format=None):
        """
        Format the time without microseconds.

        Args:
            record: The log record to format
            _date_format: Unused parameter required by the parent class

        Returns:
            Formatted time string without microseconds
        """
        return datetime.datetime.fromtimestamp(record.created).strftime('%Y-%m-%d %H:%M:%S')


# Custom colored console handler
class ColoredConsoleHandler(logging.StreamHandler):
    def emit(self, record):
        # Add color based on the log level
        msg = self.format(record)
        if record.levelno >= logging.ERROR:
            print(f"{Fore.RED}{msg}{Style.RESET_ALL}")
        elif record.levelno >= logging.WARNING:
            print(f"{Fore.YELLOW}{msg}{Style.RESET_ALL}")
        elif record.levelno >= logging.INFO:
            print(f"{Fore.GREEN}{msg}{Style.RESET_ALL}")
        else:
            print(msg)


# Configure the root logger
logger: Logger = logging.getLogger()
logger.setLevel(logging.INFO)
# Remove default handlers
logger.propagate = False


# Function to sanitize paths for logging
def _sanitize_path(path: str) -> str:
    """
    Sanitize a path for logging by replacing sensitive parts with placeholders.

    Args:
        path: The path to sanitize

    Returns:
        Sanitized path
    """
    if not path:
        return path

    # Replace user home directory with placeholder
    home_dir = os.path.expanduser("~")
    if home_dir in path:
        path = path.replace(home_dir, "<HOME>")

    # Replace Windows user profile directory with placeholder
    if "Users" in path and "\\" in path:
        parts = path.split("\\")
        for i, part in enumerate(parts):
            if part == "Users" and i + 1 < len(parts):
                # Replace the username after "Users" with a placeholder
                parts[i + 1] = "<USER>"
                path = "\\".join(parts)
                break

    # Replace common sensitive directories
    replacements = {
        os.environ.get("LOCALAPPDATA", ""): "<LOCALAPPDATA>",
        os.environ.get("APPDATA", ""): "<APPDATA>",
        os.environ.get("TEMP", ""): "<TEMP>",
        os.environ.get("TMP", ""): "<TMP>",
        tempfile.gettempdir(): "<TEMPDIR>"
    }

    for original, replacement in replacements.items():
        if original and original in path:
            path = path.replace(original, replacement)

    return path

# Function to get a temporary directory
def get_temp_dir():
    """
    Get a path to a temporary directory for storing files.

    Returns:
        Path to a temporary directory
    """
    # Create a base temporary directory for this application
    base_temp_dir = os.path.join(tempfile.gettempdir(), "Font-Archiver")
    os.makedirs(base_temp_dir, exist_ok=True)

    # Create a unique subdirectory for this run
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_dir = os.path.join(base_temp_dir, f"run_{timestamp}")
    os.makedirs(temp_dir, exist_ok=True)

    logger.info(f"Using temporary directory: {_sanitize_path(temp_dir)}")
    return temp_dir


# Constants
WINDOWS_FONTS_DIR = r"C:\Windows\Fonts"
LOCAL_FONTS_DIR = os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Microsoft\Windows\Fonts")
TEMP_DIR = get_temp_dir()
OUTPUT_DIR = os.path.join(TEMP_DIR, "Font-Storage")
REPO_NAME = "Font-Storage"
TOKEN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "github_token.txt")

# Now set up file logging in the temporary directory
log_file_path = os.path.join(TEMP_DIR, "font-upload.log")
file_handler = logging.FileHandler(log_file_path)
file_formatter = NoMicrosecondsFormatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)
start_time = datetime.datetime.now()
logger.info(f"Script started at: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")

# Global flag to track if Ctrl+C was pressed
exit_flag = False


# Function to delete temporary directory
def delete_temp_directory(ask_confirmation=True):
    """
    Delete the temporary directory, optionally asking for confirmation.

    Args:
        ask_confirmation: Whether to ask for confirmation before deleting (default: True)
    """
    try:
        if ask_confirmation:
            print(f"\nDo you want to delete the temporary directory? ({_sanitize_path(TEMP_DIR)}) (y/n)")
            if input().lower() == 'y':
                logger.info(f"Deleting temporary directory: {_sanitize_path(TEMP_DIR)}")
                shutil.rmtree(TEMP_DIR, ignore_errors=True)
                print(f"Temporary directory deleted: {_sanitize_path(TEMP_DIR)}")
            else:
                logger.info("User chose not to delete the temporary directory")
                print(f"Temporary directory preserved: {_sanitize_path(TEMP_DIR)}")
        else:
            # Delete it without asking
            logger.info(f"Deleting temporary directory without confirmation: {_sanitize_path(TEMP_DIR)}")
            shutil.rmtree(TEMP_DIR, ignore_errors=True)
    except Exception as e:
        logger.error(f"Error while trying to delete temporary directory: {e}")


# Signal handler for Ctrl+C
def signal_handler(_: int, __: Any) -> None:
    """
    Signal handler for interruption signals (Ctrl+C).

    Args:
        _: Signal number (required by signal module but unused)
        __: Current stack frame (required by signal module but unused)
    """
    global exit_flag
    if not exit_flag:
        logger.info("Ctrl+C received. Will exit after current zip operation completes.")
        print("\nCtrl+C received. The program will exit after the current operation completes.")
        exit_flag = True
        # Ask if the user wants to delete the temp directory
        delete_temp_directory(ask_confirmation=True)
    else:
        logger.info("Ctrl+C received again. Forcing exit.")
        # Delete the temp directory without asking to avoid readline re-entry issues
        delete_temp_directory(ask_confirmation=False)
        sys.exit(1)


# Register the signal handler
signal.signal(signal.SIGINT, signal_handler)

# Default Windows fonts to exclude
# noinspection SpellCheckingInspection
DEFAULT_WINDOWS_FONTS = {
    "Arial", "Calibri", "Cambria", "Candara", "Comic Sans MS", "Consolas", "Constantia",
    "Corbel", "Courier New", "Ebrima", "Franklin Gothic", "Gabriola", "Gadugi",
    "Georgia", "Impact", "Javanese Text", "Leelawadee UI", "Lucida Console",
    "Lucida Sans Unicode", "Malgun Gothic", "Microsoft Sans Serif", "MingLiU",
    "MS Gothic", "MS PGothic", "MS UI Gothic", "MV Boli", "Myanmar Text", "Nirmala UI",
    "Palatino Linotype", "Segoe MDL2 Assets", "Segoe Print", "Segoe Script",
    "Segoe UI", "SimSun", "Sitka", "Sylfaen", "Symbol", "Tahoma", "Times New Roman",
    "Trebuchet MS", "Verdana", "Webdings", "Wingdings", "Yu Gothic"
}


def _clean_font_family_name(family_name: str) -> str:
    """Remove weight/style indicators from the font family name."""
    return re.sub(
        r'\s*(Bold|Italic|Light|Regular|Medium|Thin|Black|Oblique|Condensed|Extended'
        r'|Narrow|Wide|Semi|Extra|Ultra|Demi|Heavy)$',
        '', family_name, flags=re.IGNORECASE)


def _clean_font_filename(filename: str) -> str:
    """Extract and clean base name from font filename."""
    base_name = os.path.splitext(os.path.basename(filename))[0]
    return re.sub(
        r'[-_]*(Bold|Italic|Light|Regular|Medium|Thin|Black|Oblique|Condensed|Extended|'
        r'Narrow|Wide|Semi|Extra|Ultra|Demi|Heavy)$',
        '', base_name, flags=re.IGNORECASE)


def _extract_name_from_record(record) -> Optional[str]:
    """Extract a Unicode name from a font name record."""
    if record.isUnicode():
        try:
            return record.toUnicode()
        except UnicodeDecodeError:
            return None
    return None


def get_font_family(font_path: str) -> Optional[str]:
    """Extract the font family name from a font file."""
    try:
        font = ttLib.TTFont(font_path)
        name_records = font['name'].names

        # Try to get the typographic family name first (nameID 16)
        for record in name_records:
            if record.nameID == 16:
                name = _extract_name_from_record(record)
                if name:
                    return name

        # Fall back to the font family name (nameID 1)
        for record in name_records:
            if record.nameID == 1:
                name = _extract_name_from_record(record)
                if name:
                    return _clean_font_family_name(name)

        # If we can't get the family name from the font, use the filename
        return _clean_font_filename(font_path)

    except Exception as e:
        logger.warning(f"Could not extract family name from {_sanitize_path(font_path)}: {e}")
        # Use filename as a fallback
        return _clean_font_filename(font_path)


def is_default_windows_font(family_name: str) -> bool:
    """Check if a font family is a default Windows font."""
    for default_font in DEFAULT_WINDOWS_FONTS:
        if default_font.lower() in family_name.lower():
            return True
    return False


def _process_font_file(font_path: str, font_families: Dict[str, List[str]]) -> None:
    """Process a single font file and add it to the appropriate family."""
    family_name = get_font_family(font_path)

    if not family_name or is_default_windows_font(family_name):
        return

    family_name = normalize_nerd_font_name(family_name)

    if family_name not in font_families:
        font_families[family_name] = []
    font_families[family_name].append(font_path)


def _scan_directory(directory: str, font_families: Dict[str, List[str]],
                    processed_fonts: Set[str] = None) -> Set[str]:
    """
    Scan a directory for font files and add them to font_families.

    Args:
        directory: Directory to scan for fonts
        font_families: Dictionary to update with found fonts
        processed_fonts: Set of already processed font filenames (optional)

    Returns:
        Set of processed font filenames
    """
    if processed_fonts is None:
        processed_fonts = set()

    if not os.path.exists(directory):
        return processed_fonts

    logger.info(f"Scanning fonts directory: {_sanitize_path(directory)}")

    for ext in ['*.ttf', '*.otf']:
        for font_path in glob.glob(os.path.join(directory, ext)):
            font_name = os.path.basename(font_path).lower()

            # Skip if already processed
            if font_name in processed_fonts:
                continue

            processed_fonts.add(font_name)
            _process_font_file(font_path, font_families)

    return processed_fonts


def scan_fonts() -> Dict[str, List[str]]:
    """
    Scan for fonts in Windows directories and group them by family.

    Returns:
        Dict mapping font family names to lists of font file paths.
    """
    logger.info("Scanning for fonts...")

    # Dictionary to store font families
    font_families: Dict[str, List[str]] = {}

    # Track processed fonts to handle duplicates
    processed_fonts: Set[str] = set()

    # Scan user fonts first (preferred over system fonts)
    if os.path.exists(LOCAL_FONTS_DIR):
        logger.info(f"Scanning user fonts directory: {_sanitize_path(LOCAL_FONTS_DIR)}")
        process_fonts_directory(LOCAL_FONTS_DIR, font_families, processed_fonts, add_to_processed=True)

    # Then scan system fonts
    if os.path.exists(WINDOWS_FONTS_DIR):
        logger.info(f"Scanning system fonts directory: {_sanitize_path(WINDOWS_FONTS_DIR)}")
        process_fonts_directory(WINDOWS_FONTS_DIR, font_families, processed_fonts, add_to_processed=False)

    # Remove any families with no fonts (shouldn't happen, but just in case)
    font_families = {k: v for k, v in font_families.items() if v}

    logger.info(f"Found {len(font_families)} font families")
    return font_families


# noinspection GrazieInspection
def process_fonts_directory(
        directory: str,
        font_families: Dict[str, List[str]],
        processed_fonts: Set[str],
        add_to_processed: bool
) -> None:
    """
    Process fonts in a directory and add them to the font_families dictionary.

    Args:
        directory: Directory to scan for fonts
        font_families: Dictionary to store font families
        processed_fonts: Set of already processed font filenames
        add_to_processed: Whether to add processed fonts to the processed_fonts set
    """
    for ext in ['*.ttf', '*.otf']:
        for font_path in glob.glob(os.path.join(directory, ext)):
            font_name = os.path.basename(font_path).lower()

            # Skip if we already processed this font from user directory
            if not add_to_processed and font_name in processed_fonts:
                continue

            if add_to_processed:
                processed_fonts.add(font_name)

            add_font_to_families(font_path, font_families)


# noinspection GrazieInspection
def add_font_to_families(font_path: str, font_families: Dict[str, List[str]]) -> None:
    """
    Process a single font file and add it to the appropriate family.

    Args:
        font_path: Path to the font file
        font_families: Dictionary to store font families
    """
    family_name = get_font_family(font_path)
    if family_name and not is_default_windows_font(family_name):
        family_name = normalize_nerd_font_name(family_name)

        if family_name not in font_families:
            font_families[family_name] = []
        font_families[family_name].append(font_path)


# noinspection GrazieInspection
def normalize_nerd_font_name(family_name: str) -> str:
    """
    Normalize "Nerd Font" family names to ensure consistency.

    Args:
        family_name: The original font family name

    Returns:
        Normalized font family name
    """
    if "nerd font" in family_name.lower() and not family_name.lower().endswith("nerd font"):
        base_name = family_name.split("Nerd Font")[0].strip()
        return f"{base_name} Nerd Font"
    return family_name


# noinspection GrazieInspection
def limit_font_families(font_families: Dict[str, List[str]], limit: int) -> Dict[str, List[str]]:
    """
    Limit the number of font families for debugging purposes.

    Args:
        font_families: Dictionary of font families
        limit: Maximum number of families to keep

    Returns:
        Limited dictionary of font families
    """
    if len(font_families) > limit:
        logger.info(f"Limiting to first {limit} font families for debugging (out of {len(font_families)} total)")
        limited_families = {}
        for i, (family, paths) in enumerate(font_families.items()):
            if i >= limit:
                break
            limited_families[family] = paths
        return limited_families
    return font_families


def _prepare_temp_directory(font_paths: List[str], temp_dir: str) -> None:
    """
    Copy font files to a temporary directory.

    Args:
        font_paths: List of paths to font files
        temp_dir: Path to the temporary directory
    """
    for font_path in font_paths:
        shutil.copy2(font_path, os.path.join(temp_dir, os.path.basename(font_path)))


def _cleanup_temp_directory(temp_dir: str) -> None:
    """
    Clean up the temporary directory with proper error handling.

    Args:
        temp_dir: Path to the temporary directory
    """
    if not os.path.exists(temp_dir):
        return

    logger.debug(f"Removing directory: {_sanitize_path(temp_dir)}")

    # Use shutil.rmtree with ignore_errors=True for simplicity
    try:
        shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception as e:
        logger.warning(f"Failed to remove directory {_sanitize_path(temp_dir)}: {e}")


def _verify_temp_directory(temp_dir: str, family_name: str) -> bool:
    """
    Verify that the temporary directory exists.

    Args:
        temp_dir: Path to the temporary directory
        family_name: The name of the font family for logging

    Returns:
        True if the directory exists, False otherwise
    """
    if not os.path.exists(temp_dir):
        logger.error(f"Temporary directory {_sanitize_path(temp_dir)} does not exist for {family_name}")
        return False
    return True


def _change_to_directory(directory: str, family_name: str) -> bool:
    """
    Change to the specified directory.

    Args:
        directory: Path to the directory to change to
        family_name: The name of the font family for logging

    Returns:
        True if successful, False otherwise
    """
    try:
        os.chdir(directory)
        return True
    except Exception as e:
        logger.error(f"Error changing to directory {_sanitize_path(directory)} for {family_name}: {str(e)}")
        return False


def _ensure_output_directory(zip_path: str, family_name: str) -> bool:
    """
    Ensure the output directory for the zip file exists.

    Args:
        zip_path: Path to the output zip file
        family_name: The name of the font family for logging

    Returns:
        True if successful, False otherwise
    """
    try:
        os.makedirs(os.path.dirname(os.path.abspath(zip_path)), exist_ok=True)
        return True
    except Exception as e:
        logger.error(f"Error creating output directory for {family_name}: {str(e)}")
        return False


def _handle_existing_zip(zip_path: str) -> str:
    """
    Handle an existing archive file by removing it or generating a new name.

    Args:
        zip_path: Path to the output archive file

    Returns:
        The path to use for the archive file
    """
    if os.path.exists(zip_path):
        try:
            os.remove(zip_path)
            return zip_path
        except Exception as e:
            logger.warning(f"Could not remove existing archive file {_sanitize_path(zip_path)}: {e}")
            # If we can't remove it, use a different name
            # Preserve the original file extension (.zip or .7z)
            file_ext = os.path.splitext(zip_path)[1]
            return os.path.join(
                os.path.dirname(zip_path),
                f"{os.path.splitext(os.path.basename(zip_path))[0]}_{int(time.time())}{file_ext}"
            )
    return zip_path


def _move_zip_file(rel_zip_path: str, zip_path: str, family_name: str) -> bool:
    """
    Move the archive file to the correct location.

    Args:
        rel_zip_path: Relative path to the output archive file
        zip_path: Absolute path to the output archive file
        family_name: The name of the font family for logging

    Returns:
        True if successful, False otherwise
    """
    try:
        shutil.move(rel_zip_path, zip_path)
        return True
    except Exception as e:
        logger.error(f"Error moving archive file for {family_name}: {str(e)}")
        return False


def _verify_font_paths(font_paths: List[str], zip_path: str) -> bool:
    """
    Verify that all font paths exist.

    Args:
        font_paths: List of paths to font files
        zip_path: Path to the output zip file for logging

    Returns:
        True if all font paths exist, False otherwise
    """
    missing_fonts = [path for path in font_paths if not os.path.exists(path)]
    if missing_fonts:
        logger.error(f"Missing font files for {_sanitize_path(zip_path)}: {missing_fonts}")
        return False
    return True


def _create_zip_file(font_paths: List[str], zip_path: str) -> bool:
    """
    Create a zip file using Python's zipfile module.

    Args:
        font_paths: List of paths to font files
        zip_path: Path to the output zip file

    Returns:
        True if successful, False otherwise
    """
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for font_path in font_paths:
                try:
                    zipf.write(font_path, os.path.basename(font_path))
                except Exception as e:
                    logger.error(f"Error adding {_sanitize_path(font_path)} to zip: {str(e)}")
                    # Continue with other files
        return True
    except zipfile.BadZipFile as e:
        logger.error(f"Bad zip file error for {_sanitize_path(zip_path)}: {str(e)}")
        return False
    except PermissionError as e:
        logger.error(f"Permission error creating zip file {_sanitize_path(zip_path)}: {str(e)}")
        return False
    except Exception as e:
        logger.error(f"Error creating zip file {_sanitize_path(zip_path)}: {str(e)}")
        return False


def _verify_zip_file(zip_path: str) -> bool:
    """
    Verify that the zip file was created and is valid.

    Args:
        zip_path: Path to the output zip file

    Returns:
        True if the zip file exists and is valid, False otherwise
    """
    # Verify the zip file was created
    if not os.path.exists(zip_path):
        logger.error(f"Zip file {_sanitize_path(zip_path)} was not created")
        return False

    # Verify the zip file is valid
    try:
        with zipfile.ZipFile(zip_path, 'r') as zipf:
            # Test the integrity of the zip file
            if zipf.testzip() is not None:
                logger.error(f"Zip file {_sanitize_path(zip_path)} is corrupted")
                return False
        return True
    except Exception as e:
        logger.error(f"Error verifying zip file {_sanitize_path(zip_path)}: {str(e)}")
        return False


def _verify_7z_file(zip_path: str) -> bool:
    """
    Verify that the 7z file was created and is valid.

    Args:
        zip_path: Path to the output 7z file

    Returns:
        True if the 7z file exists and is valid, False otherwise
    """
    # Verify the 7z file was created
    if not os.path.exists(zip_path):
        logger.error(f"7z file {_sanitize_path(zip_path)} was not created")
        return False

    # Verify the 7z file is valid using the 7z command-line tool
    try:
        # Use 7z t (test) command to verify the archive
        process = subprocess.run(
            ["7z", "t", zip_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False
        )

        # Check if the test was successful
        if process.returncode != 0:
            logger.error(f"7z file {_sanitize_path(zip_path)} is corrupted: {process.stderr}")
            return False

        return True
    except FileNotFoundError:
        logger.error("7zip command-line tool (7z) not found. Cannot verify 7z file.")
        return False
    except Exception as e:
        logger.error(f"Error verifying 7z file {_sanitize_path(zip_path)}: {str(e)}")
        return False


def _create_zip_with_zipfile(font_paths: List[str], zip_path: str) -> bool:
    """
    Create a zip file using Python's zipfile module.

    Args:
        font_paths: List of paths to font files
        zip_path: Path to the output zip file

    Returns:
        True if successful, False if failed
    """
    try:
        # Ensure the output directory exists
        if not _ensure_output_directory(zip_path, os.path.basename(zip_path)):
            return False

        # If the destination file already exists, handle it
        zip_path = _handle_existing_zip(zip_path)

        # Verify all font paths exist before attempting to create the zip
        if not _verify_font_paths(font_paths, zip_path):
            return False

        # Create the zip file
        if not _create_zip_file(font_paths, zip_path):
            return False

        # Verify the zip file was created and is valid
        if not _verify_zip_file(zip_path):
            return False

        return True
    except Exception as e:
        logger.error(f"Unexpected error creating zip with zipfile: {str(e)}")
        return False


def _sanitize_name(family_name: str) -> str:
    """
    Sanitize a family name for use in filenames.

    Args:
        family_name: The name of the font family

    Returns:
        Sanitized name
    """
    return re.sub(r'[^\w\-.]', '_', family_name)


def _prepare_zip_path(family_name: str, output_dir: str) -> Tuple[str, str]:
    """
    Prepare the archive file path for a font family.

    Args:
        family_name: The name of the font family
        output_dir: Directory to save the archive file

    Returns:
        Tuple of (absolute_archive_path, sanitized_name)
    """
    # Sanitize family name for filename
    safe_name = _sanitize_name(family_name)

    # Use .7z extension for 7zip archives
    archive_path = os.path.join(output_dir, f"{safe_name}.7z")

    # Convert to an absolute path to avoid issues when changing directories
    return os.path.abspath(archive_path), safe_name


def _setup_temp_directory(output_dir: str, safe_name: str) -> str:
    """
    Set up a temporary directory for font files.

    Args:
        output_dir: Directory to create the temporary directory in
        safe_name: Sanitized name to use for the temporary directory

    Returns:
        Absolute path to the temporary directory
    """
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Create a unique temporary directory name using a timestamp
    timestamp = int(time.time() * 1000)
    temp_dir = os.path.join(output_dir, f"temp_{safe_name}_{timestamp}")

    # Convert to an absolute path to avoid issues
    temp_dir = os.path.abspath(temp_dir)

    # Create the directory (don't use exist_ok to ensure it's new)
    try:
        os.makedirs(temp_dir)
    except FileExistsError:
        # If it somehow exists, add another random component
        import random
        temp_dir = os.path.join(output_dir, f"temp_{safe_name}_{timestamp}_{random.randint(1000, 9999)}")
        temp_dir = os.path.abspath(temp_dir)
        os.makedirs(temp_dir)

    return temp_dir


def _get_retry_path(original_path: str, retry_count: int) -> str:
    """
    Generate a retry path for an archive file.

    Args:
        original_path: Original path to the archive file
        retry_count: Current retry count

    Returns:
        New path with retry suffix
    """
    # Preserve the original file extension (.zip or .7z)
    file_ext = os.path.splitext(original_path)[1]
    return os.path.join(
        os.path.dirname(original_path),
        f"{os.path.splitext(os.path.basename(original_path))[0]}_retry{retry_count}{file_ext}"
    )


def _get_zip_size(zip_path: str) -> int:
    """
    Get the size of an archive file.

    Args:
        zip_path: Path to the archive file

    Returns:
        Size of the archive file in bytes, or 0 if an error occurs
    """
    try:
        return os.path.getsize(zip_path)
    except Exception as e:
        logger.error(f"Error getting size of archive file {_sanitize_path(zip_path)}: {str(e)}")
        return 0


def _display_progress_bar(progress: float, width: int = 50, prefix: str = '', suffix: str = '') -> None:
    """
    Display a progress bar in the console.

    Args:
        progress: Progress as a percentage (0-100)
        width: Width of the progress bar in characters
        prefix: Text to display before the progress bar
        suffix: Text to display after the progress bar
    """
    # Ensure progress is between 0 and 100
    progress = min(max(progress, 0), 100)

    # Calculate the number of filled blocks
    filled_length = int(width * progress / 100)

    # Create the progress bar
    bar = '█' * filled_length + '-' * (width - filled_length)

    # Store the previous suffix length in a function attribute
    # This allows us to track it between calls
    if not hasattr(_display_progress_bar, 'prev_suffix_len'):
        _display_progress_bar.prev_suffix_len = 0

    # Calculate how many spaces we need to clear the previous suffix
    # We only want to clear the suffix text, not the entire line
    clear_suffix = ' ' * max(0, _display_progress_bar.prev_suffix_len - len(suffix))

    # Update the previous suffix length for the next call
    _display_progress_bar.prev_suffix_len = len(suffix)

    # Print the progress bar with clearing spaces for the suffix only
    print(f'\r{prefix} |{bar}| {progress:.1f}% {suffix}{clear_suffix}', end='', flush=True)

    # Print a newline when progress is complete
    if progress == 100:
        print()


def _get_cpu_core_count() -> int:
    """
    Get the number of CPU cores on the host machine.

    Returns:
        Number of CPU cores
    """
    try:
        return multiprocessing.cpu_count()
    except Exception as e:
        logger.warning(f"Error getting CPU core count: {str(e)}. Using default value of 4.")
        return 4


def _create_zip_with_7zip(font_paths: List[str], zip_path: str) -> bool:
    """
    Create a 7z archive using the 7zip command-line tool.

    Args:
        font_paths: List of paths to font files
        zip_path: Path to the output 7z file

    Returns:
        True if successful, False otherwise
    """
    try:
        # Ensure the file extension is .7z
        if not zip_path.lower().endswith('.7z'):
            zip_path = os.path.splitext(zip_path)[0] + '.7z'

        # Ensure the output directory exists
        if not _ensure_output_directory(zip_path, os.path.basename(zip_path)):
            return False

        # If the destination file already exists, handle it
        zip_path = _handle_existing_zip(zip_path)

        # Verify all font paths exist before attempting to create the archive
        if not _verify_font_paths(font_paths, zip_path):
            return False

        # Get CPU core count to determine compression level
        cpu_cores = _get_cpu_core_count()
        # Use a compression level based on CPU cores but cap it at 9
        compression_level = min(cpu_cores, 9)

        # Create a temporary directory to store the files with just their basename
        with tempfile.TemporaryDirectory() as temp_dir:
            # Copy each font file to the temporary directory with just its basename
            temp_font_paths = []
            for font_path in font_paths:
                basename = os.path.basename(font_path)
                temp_font_path = os.path.join(temp_dir, basename)
                shutil.copy2(font_path, temp_font_path)
                temp_font_paths.append(basename)  # Store the basename

            # Prepare the 7zip command with required switches
            cmd = [
                "7z", "a",  # Add to archive
                "-t7z",  # 7z archive type
                f"-mx={compression_level}",  # Compression level based on CPU cores
                "-m0=lzma2",  # LZMA2 compression method
                zip_path  # Output file
            ]

            # Set the working directory to the temporary directory
            # This ensures only the basename are included in the archive

            # Add all font files to the command (just the basename)
            cmd.extend(temp_font_paths)

            # Execute the 7zip command from the temporary directory
            process = subprocess.run(
                cmd, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE,
                text=True,
                check=False,
                cwd=temp_dir  # Set the working directory to the temporary directory
            )

            # Check if the command was successful
            if process.returncode != 0:
                logger.error(f"7zip command failed with return code {process.returncode}: {process.stderr}")
                return False

        # Verify the 7z file was created and is valid
        if not _verify_7z_file(zip_path):
            return False

        logger.info(
            f"Successfully created 7z archive {_sanitize_path(zip_path)} with compression level {compression_level}")
        return True

    except FileNotFoundError:
        logger.error("7zip command-line tool (7z) not found. Please install 7zip and ensure it's in your PATH.")
        return False
    except Exception as e:
        logger.error(f"Error creating 7z archive {_sanitize_path(zip_path)}: {str(e)}")
        return False


def _create_zip_with_strategy(font_paths: List[str], zip_path: str) -> bool:
    """
    Create an archive using the appropriate strategy (7zip or zipfile).

    Args:
        font_paths: List of paths to font files
        zip_path: Path to the output archive file

    Returns:
        True if successful, False otherwise
    """
    # Try to use 7zip first
    try:
        # Change the extension to .7z
        seven_zip_path = os.path.splitext(zip_path)[0] + '.7z'

        # Use 7zip to create the archive
        if _create_zip_with_7zip(font_paths, seven_zip_path):
            return True

        # If 7zip fails, fall back to the zipfile
        logger.warning("7zip compression failed, falling back to zipfile")
    except Exception as e:
        logger.warning(f"Error using 7zip: {str(e)}. Falling back to zipfile.")

    # Ensure the file extension is .zip for the fallback method
    if not zip_path.lower().endswith('.zip'):
        zip_path = os.path.splitext(zip_path)[0] + '.zip'

    # Use Python's zipfile module as a fallback
    return _create_zip_with_zipfile(font_paths, zip_path)


def _attempt_zip_creation_with_retry(family_name: str, font_paths: List[str], original_zip_path: str) -> Tuple[
    bool, str]:
    """
    Attempt to create a zip file with retries.

    Args:
        family_name: The name of the font family
        font_paths: List of paths to font files
        original_zip_path: Original path to the output zip file

    Returns:
        Tuple of (success, final_zip_path)
    """
    success = False
    max_retries = 3
    retry_count = 0
    zip_path = original_zip_path

    while not success and retry_count < max_retries:
        if retry_count > 0:
            # Add a retry suffix to the zip path to avoid conflicts
            zip_path = _get_retry_path(original_zip_path, retry_count)
            logger.info(f"Retrying zip creation for {family_name} (attempt {retry_count + 1}/{max_retries})")

        try:
            # Try to create zip with the appropriate strategy
            success = _create_zip_with_strategy(font_paths, zip_path)

            # If we've succeeded, break out of the retry loop
            if success:
                break

        except Exception as e:
            logger.warning(f"Attempt {retry_count + 1}/{max_retries} failed for {family_name}: {str(e)}")
            # Sleep briefly before retrying to allow any file locks to be released
            time.sleep(0.5)

        retry_count += 1

    return success, zip_path


def _verify_and_get_zip_size(zip_path: str) -> int:
    """
    Verify that the zip file exists and get its size.

    Args:
        zip_path: Path to the zip file

    Returns:
        Size of the zip file in bytes
    """
    # Verify the zip file exists
    if not os.path.exists(zip_path):
        raise FileNotFoundError(f"Zip file {zip_path} was not created")

    # Get the size of the zip file
    return _get_zip_size(zip_path)


def create_zip_for_family(family_name: str, font_paths: List[str], output_dir: str) -> Tuple[str, int]:
    """
    Create an archive file for a font family.

    Args:
        family_name: The name of the font family
        font_paths: List of paths to font files
        output_dir: Directory to save the archive file

    Returns:
        Tuple of (archive_path, archive_size_in_bytes)
    """
    try:
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)

        # Prepare the archive file path and get the sanitized name
        archive_path, safe_name = _prepare_zip_path(family_name, output_dir)

        # Attempt to create the archive file with retries
        success, final_archive_path = _attempt_zip_creation_with_retry(family_name, font_paths, archive_path
                                                                   )

        if not success:
            raise Exception(f"Failed to create archive for {family_name} after multiple attempts")

        # Verify the archive file exists and get its size
        archive_size = _verify_and_get_zip_size(final_archive_path)

        return final_archive_path, archive_size

    except Exception as e:
        logger.error(f"Error creating archive for {family_name}: {str(e)}")
        # Return a fake path and size to avoid breaking the caller
        return os.path.join(output_dir, f"{_sanitize_name(family_name)}.7z"), 0


def create_zips(font_families: Dict[str, List[str]], output_dir: str) -> Tuple[List[str], int]:
    """
    Create archive files for each font family using parallel processing.

    Args:
        font_families: Dictionary mapping font family names to lists of font file paths
        output_dir: Directory to save the archive files

    Returns:
        Tuple of (list_of_archive_paths, total_size_in_bytes)
    """
    logger.info("Creating archive files for font families...")

    # Create an output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Using 7zip for compression with fallback to Python's built-in zipfile module
    logger.info("Using 7zip for compression with LZMA2 method")

    archive_paths = []
    total_size = 0
    total_families = len(font_families)

    # Process all font families using parallel processing
    families_list = list(font_families.items())
    logger.info(f"Processing all {total_families} font families")

    # Use ThreadPoolExecutor for parallel compression
    with concurrent.futures.ThreadPoolExecutor() as executor:
        # Submit all compression tasks
        future_to_family = {}
        for family, paths in families_list:
            logger.info(f"Starting archive task for {family}")
            future = executor.submit(create_zip_for_family, family, paths, output_dir)
            future_to_family[future] = family

        # Process results as they complete
        for future in concurrent.futures.as_completed(future_to_family):
            family = future_to_family[future]
            try:
                archive_path, archive_size = future.result()
                archive_paths.append(archive_path)
                total_size += archive_size

                # Log task completion
                logger.info(f"Finished archive task for {family}")

                # Log progress to file only
                completed = len(archive_paths)
                progress = completed / total_families * 100
                logger.info(
                    f"Progress: {progress:.1f}% - Created archive for {family} ({archive_size / 1024 / 1024:.2f} MB)")

                # Display progress bar
                suffix = f"Created archive for {family} ({archive_size / 1024 / 1024:.2f} MB)"
                _display_progress_bar(progress, prefix="Creating archives:", suffix=suffix)

                # Check if Ctrl+C was pressed
                if exit_flag:
                    logger.info("Exiting after completing current archive operation due to Ctrl+C")
                    return archive_paths, total_size

            except Exception as e:
                logger.error(f"Error creating archive for {family}: {e}")

    return archive_paths, total_size


def _create_readme_file(repo_dir: str, total_families: int, total_size: int) -> None:
    """
    Create README.md with statistics and a disclaimer.

    Args:
        repo_dir: Directory for the repository
        total_families: Number of font families
        total_size: Total size of all zip files in bytes
    """
    readme_path = os.path.join(repo_dir, "README.md")
    with open(readme_path, 'w') as f:
        f.write(f"""# Font Storage

A collection of fonts organized by family.

## Statistics

- Font family count: {total_families}
- Total zip size: {total_size / 1024 / 1024:.2f} MB
- Upload date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## Disclaimer

The commercial status of these fonts is unknown. The repository owner makes no claim to ownership of these items.
These fonts are provided "as is" without warranty of any kind, either expressed or implied.

In the event that the contents of the repository fall under copyright, the repository owner makes no claim to its contents. 
All fonts were obtained from openly available locations.
""")
    logger.info("Created README.md with statistics and disclaimer")


def _initialize_git_and_lfs(repo_dir: str) -> bool:
    """
    Initialize Git and Git LFS in the repository directory.

    Args:
        repo_dir: Directory for the repository

    Returns:
        True if Git and Git LFS were initialized successfully, False otherwise
    """
    try:
        # Check if git is installed
        subprocess.run(["git", "--version"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        # Initialize a Git repository
        subprocess.run(["git", "init"], check=True, cwd=repo_dir)
        logger.info("Git repository initialized")

        # Check if Git LFS is installed
        try:
            subprocess.run(["git", "lfs", "version"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            logger.info("Git LFS is installed")
        except subprocess.CalledProcessError:
            logger.error("Git LFS is not installed. Please install Git LFS: https://git-lfs.github.com/")
            print("\nGit LFS is not installed. Please install Git LFS: https://git-lfs.github.com/")
            print("Continuing without Git LFS support...")
            return False

        # Initialize Git LFS
        subprocess.run(["git", "lfs", "install"], check=True, cwd=repo_dir)
        logger.info("Git LFS initialized")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"Error initializing Git or Git LFS: {e}")
        print(f"\nError initializing Git or Git LFS: {e}")
        print("Continuing without Git LFS support...")
        return False


def _configure_git_lfs(repo_dir: str) -> bool:
    """
    Configure Git LFS by creating and committing a .gitattributes file.

    Args:
        repo_dir: Directory for the repository

    Returns:
        True if Git LFS was configured successfully, False otherwise
    """
    try:
        # Add .gitattributes file for Git LFS
        with open(os.path.join(repo_dir, ".gitattributes"), 'w') as f:
            # Track specific file types with Git LFS
            f.write("*.zip filter=lfs diff=lfs merge=lfs -text\n")
            f.write("*.7z filter=lfs diff=lfs merge=lfs -text\n")
            # Track other binary files with Git LFS (excluding .git files)
            f.write("*.[!g][!i][!t]* filter=lfs diff=lfs merge=lfs -text\n")

        # Add .gitattributes to Git
        subprocess.run(["git", "add", ".gitattributes"], check=True, cwd=repo_dir)
        subprocess.run(["git", "commit", "-m", "Initialize Git LFS"], check=True, cwd=repo_dir)
        logger.info("Git LFS configured to track zip, 7z, and other binary files")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"Error configuring Git LFS: {e}")
        return False


def _copy_gitignore_file(repo_dir: str) -> None:
    """
    Copy the .gitignore file to the repository if it exists.

    Args:
        repo_dir: Directory for the repository
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    gitignore_path = os.path.join(script_dir, ".gitignore")
    if os.path.exists(gitignore_path):
        shutil.copy2(gitignore_path, os.path.join(repo_dir, ".gitignore"))
        logger.info("Copied .gitignore file to repository")


def _copy_single_file(source_path: str, dest_path: str) -> None:
    """
    Copy a single file from source to destination.

    Args:
        source_path: Path to the source file
        dest_path: Path to the destination file
    """
    try:
        shutil.copy2(source_path, dest_path)
    except Exception as e:
        logger.error(f"Error copying file {_sanitize_path(source_path)} to {_sanitize_path(dest_path)}: {e}")


def _copy_directory_contents(dir_path: str, repo_dir: str) -> None:
    """
    Copy the contents of a directory to the repository root.

    Args:
        dir_path: Path to the source directory
        repo_dir: Path to the repository directory
    """
    for subitem in os.listdir(dir_path):
        subitem_path = os.path.join(dir_path, subitem)
        subdest_path = os.path.join(repo_dir, subitem)

        # Skip if the subitem already exists at the destination
        if os.path.exists(subdest_path):
            continue

        # Copy the subitem to the root of the repository
        if os.path.isdir(subitem_path):
            try:
                shutil.copytree(subitem_path, subdest_path)
            except Exception as e:
                logger.error(
                    f"Error copying directory {_sanitize_path(subitem_path)} to {_sanitize_path(subdest_path)}: {e}")
        else:
            _copy_single_file(subitem_path, subdest_path)


def _copy_log_file(repo_dir: str) -> None:
    """
    Copy the log file to the repository directory.

    Args:
        repo_dir: Path to the repository directory
    """
    log_path = os.path.join(TEMP_DIR, "font-upload.log")
    if os.path.exists(log_path):
        dest_log_path = os.path.join(repo_dir, "font-upload.log")
        _copy_single_file(log_path, dest_log_path)
        logger.info(f"Copied log file to repository: {_sanitize_path(dest_log_path)}")


def _copy_files_to_repository(output_dir: str, repo_dir: str) -> None:
    """
    Copy files from the output directory to the repository directory.

    Args:
        output_dir: Source directory containing files to copy
        repo_dir: Destination directory for the repository
    """
    # Copy the font archives directly to the root of the repository
    for item in os.listdir(output_dir):
        item_path = os.path.join(output_dir, item)
        dest_path = os.path.join(repo_dir, item)

        # Skip if the item already exists at the destination
        if os.path.exists(dest_path):
            continue

        # Copy the item to the repository directory
        # Only copy files, not directories, to avoid duplication
        if not os.path.isdir(item_path):
            _copy_single_file(item_path, dest_path)

    logger.info("Copied files to repository")

    # Copy the log file to the repository directory
    _copy_log_file(repo_dir)


def create_git_repo(output_dir: str, total_families: int, total_size: int) -> None:
    """
    Prepare files for the GitHub repository using Git and Git LFS.

    Args:
        output_dir: Directory containing the font archives
        total_families: Number of font families
        total_size: Total size of all zip files in bytes
    """
    logger.info("Preparing files for GitHub repository...")

    # Create a new directory for the repository (sibling to the output directory)
    repo_dir = os.path.join(os.path.dirname(output_dir), REPO_NAME)

    # Create the repository directory if it doesn't exist
    os.makedirs(repo_dir, exist_ok=True)

    os.chdir(repo_dir)

    # Create README.md with statistics and disclaimer
    _create_readme_file(repo_dir, total_families, total_size)

    # Initialize Git and Git LFS
    lfs_initialized = _initialize_git_and_lfs(repo_dir)

    # Configure Git LFS if it was initialized successfully
    if lfs_initialized:
        _configure_git_lfs(repo_dir)

    # Copy the .gitignore file to the repository
    _copy_gitignore_file(repo_dir)

    # Copy files to the repository
    _copy_files_to_repository(output_dir, repo_dir)

    logger.info("Files prepared for GitHub repository successfully")


def get_github_token() -> str:
    """
    Get GitHub personal access token from file or user input.

    Returns:
        GitHub personal access token
    """
    # Use the token file in the temporary directory
    token_file_path = TOKEN_FILE

    # Check if a token file exists
    if os.path.exists(token_file_path):
        with open(token_file_path, 'r') as f:
            token = f.read().strip()
            if token:
                logger.info("GitHub token loaded from file")
                return token

    # Prompt user for token
    print("\nGitHub personal access token not found or empty.")
    print("Please enter your GitHub personal access token:")
    token = getpass.getpass(prompt="GitHub Token: ")

    # Save token to file
    with open(token_file_path, 'w') as f:
        f.write(token)

    logger.info("GitHub token saved to file")
    return token


def check_github_repo_exists(token: str, repo_name: str) -> bool:
    """
    Check if a GitHub repository exists using PyGithub.

    Args:
        token: GitHub personal access token
        repo_name: Name of the repository

    Returns:
        True if the repository exists, False otherwise
    """
    try:
        g = Github(token)
        user = g.get_user()

        try:
            user.get_repo(repo_name)
            return True
        except GithubException as e:
            if e.status == 404:  # 404 means repo doesn't exist
                return False
            else:
                logger.error(f"Error checking if repository exists: {e}")
                # Delete the temporary directory before exiting
                delete_temp_directory(ask_confirmation=False)
                sys.exit(1)
    except GithubException as e:
        logger.error(f"Error checking if repository exists: {e}")
        # Delete the temporary directory before exiting
        delete_temp_directory(ask_confirmation=False)
        sys.exit(1)


def get_github_username(token: str) -> str:
    """
    Get the GitHub username associated with the token using PyGithub.

    Args:
        token: GitHub personal access token

    Returns:
        GitHub username
    """
    try:
        g = Github(token)
        user = g.get_user()
        return user.login
    except GithubException as e:
        logger.error(f"Error getting GitHub username: {e}")
        # Delete the temporary directory before exiting
        delete_temp_directory(ask_confirmation=False)
        sys.exit(1)


# These functions have been replaced by PyGithub implementations in create_github_repo


def _check_if_repo_exists(user: Any, repo_name: str) -> Optional[Any]:
    """
    Check if a repository exists.

    Args:
        user: GitHub user object
        repo_name: Name of the repository

    Returns:
        Repository object if it exists, None otherwise
    """
    try:
        repo = user.get_repo(repo_name)
        logger.info(f"Repository '{repo_name}' already exists")
        return repo
    except GithubException as e:
        if e.status == 404:  # 404 means repo doesn't exist, which is fine
            logger.info(f"Repository '{repo_name}' does not exist")
            return None
        # For other errors, re-raise to be handled by the caller
        raise


def _handle_existing_repo(repo: Any, repo_name: str) -> Tuple[bool, bool]:
    """
    Handle an existing repository by asking the user if they want to delete it.

    Args:
        repo: GitHub repository object
        repo_name: Name of the repository

    Returns:
        Tuple of (success, should_append):
        - success: True if the operation was successful, False if there was an error
        - should_append: True if the user wants to append to existing repo, False if deleted or error
    """
    print(f"\nRepository '{repo_name}' already exists.")
    print("Choose an option:")
    print("1. Fresh (delete existing repository and start fresh)")
    print("2. Append (add to existing repository)")

    choice = input("Enter your choice (1 or 2): ").strip()

    if choice == "2" or choice.lower() == "append":
        logger.info("Will append to existing repository")
        return True, True  # Success should append

    try:
        repo.delete()
        logger.info(f"Deleted repository '{repo_name}'")
        # Wait a moment for the deletion to complete
        time.sleep(2)
        return True, False  # Success should not append (create new)
    except GithubException as e:
        logger.error(f"Failed to delete repository: {e}")
        return False, False  # Error should not append


def _create_new_repo(user: Any, repo_name: str) -> bool:
    """
    Create a new GitHub repository.

    Args:
        user: GitHub user object
        repo_name: Name of the repository

    Returns:
        True if the repository was created successfully, False otherwise
    """
    try:
        user.create_repo(
            name=repo_name,
            description="Collection of fonts organized by family",
            private=False,
            has_issues=True,
            has_projects=False,
            has_wiki=False
        )
        logger.info(f"Created GitHub repository '{repo_name}'")
        return True
    except GithubException as e:
        logger.error(f"Failed to create repository: {e}")
        return False


def create_github_repo(token: str, repo_name: str) -> None:
    """
    Create a GitHub repository using PyGithub.

    Args:
        token: GitHub personal access token
        repo_name: Name of the repository
    """
    try:
        # Create a GitHub instance with the token
        g = Github(token)
        user = g.get_user()

        # Check if repo exists
        repo = _check_if_repo_exists(user, repo_name)

        if repo:
            # Handle existing repository
            success, should_append = _handle_existing_repo(repo, repo_name)

            if not success:
                logger.error("Failed to handle existing repository")
                # Delete the temporary directory before exiting
                delete_temp_directory(ask_confirmation=False)
                sys.exit(1)

            # If a user chose to append to an existing repo, we're done
            if should_append:
                return

        # Create a new repo
        if not _create_new_repo(user, repo_name):
            logger.error("Failed to create new repository")
            # Delete the temporary directory before exiting
            delete_temp_directory(ask_confirmation=False)
            sys.exit(1)

    except GithubException as e:
        logger.error(f"GitHub API error: {e}")
        # Delete the temporary directory before exiting
        delete_temp_directory(ask_confirmation=False)
        sys.exit(1)
    except Exception as e:
        logger.error(f"Error in create_github_repo: {e}")
        # Delete the temporary directory before exiting
        delete_temp_directory(ask_confirmation=False)
        sys.exit(1)


def _get_github_user_plan(token: str) -> Tuple[bool, str]:
    """
    Get the GitHub user's plan information.

    Args:
        token: GitHub personal access token

    Returns:
        Tuple of (success, plan_name)
    """
    try:
        g = Github(token)
        user = g.get_user()

        # Get plan information
        plan_name = user.plan.name if hasattr(user, 'plan') and hasattr(user.plan, 'name') else "unknown"
        logger.info(f"Retrieved GitHub account plan: {plan_name}")
        return True, plan_name
    except GithubException as e:
        logger.error(f"Error getting GitHub user plan: {e}")
        return False, "unknown"


def _prompt_user_for_confirmation(message: str, warning: str) -> bool:
    """
    Prompt the user for confirmation with a warning message.

    Args:
        message: The message to display to the user
        warning: The warning to log and display

    Returns:
        True if the user confirms, False otherwise
    """
    logger.warning(warning)

    print(f"\n{message}")
    print("Do you want to continue? (y/n)")

    if input().lower() != 'y':
        logger.info("User chose not to proceed")
        return False

    return True


def check_github_lfs_storage(token: str) -> bool:
    """
    Check if there is enough GitHub LFS storage for the upload using PyGithub.

    Args:
        token: GitHub personal access token

    Returns:
        True if there is enough storage, False otherwise
    """
    # Get user plan information
    success, plan_name = _get_github_user_plan(token)
    if not success:
        return False

    logger.info(f"Checking GitHub LFS storage limits... (Account plan: {plan_name})")

    # For demonstration purposes, we'll check if the account plan allows LFS
    # Free accounts have limited LFS storage
    if plan_name.lower() == "free":
        message = "You are using a free GitHub account with limited LFS storage.\nLarge uploads may fail if you exceed your storage quota."
        warning = "You are using a free GitHub account with limited LFS storage. Large uploads may fail if you exceed your storage quota."

        return _prompt_user_for_confirmation(message, warning)

    return True


def _check_github_api_rate_limit(token: str) -> bool:
    """
    Check the GitHub API rate limit.

    Args:
        token: GitHub personal access token

    Returns:
        True if successful, False otherwise
    """
    try:
        g = Github(token)
        rate_limit = g.get_rate_limit()
        logger.info(f"GitHub API rate limit: {rate_limit.core.remaining}/{rate_limit.core.limit}")
        return True
    except GithubException as e:
        logger.error(f"Error checking GitHub API rate limit: {e}")
        return False


def check_github_data_limits(token: str, total_size: int) -> bool:
    """
    Check if there is enough GitHub data transfer quota for the upload using PyGithub.

    Args:
        token: GitHub personal access token
        total_size: Total size of all zip files in bytes

    Returns:
        True if there is enough quota, False otherwise
    """
    # Check API rate limit
    if not _check_github_api_rate_limit(token):
        return False

    # Calculate size in MB
    size_mb = total_size / 1024 / 1024
    logger.info(f"Checking GitHub data transfer limits... (Upload size: {size_mb:.2f} MB)")

    # For demonstration purposes, we'll check if the size is reasonable
    if size_mb > 1000:  # 1 GB
        message = f"Upload size is large: {size_mb:.2f} MB\nGitHub has monthly data transfer limits that may affect your upload."
        warning = f"Upload size is large: {size_mb:.2f} MB. GitHub has monthly data transfer limits that may affect your upload."

        return _prompt_user_for_confirmation(message, warning)

    return True


def _is_file_too_large(file_path: str) -> bool:
    """
    Check if a file is too large for direct GitHub API upload.

    Args:
        file_path: Path to the file

    Returns:
        True if the file is too large (>70MB), False otherwise
    """
    # Use Git LFS for files larger than 70MB as per requirements
    return os.path.getsize(file_path) > 70 * 1024 * 1024  # 70MB


def _read_file_content(file_path: str) -> bytes:
    """
    Read binary content from a file.

    Args:
        file_path: Path to the file

    Returns:
        Binary content of the file
    """
    with open(file_path, 'rb') as f:
        return f.read()


def _upload_file_to_github(repo, file_path: str, rel_path: str) -> bool:
    """
    Upload a file to a GitHub repository (create or update).

    Args:
        repo: GitHub repository object
        file_path: Local path to the file
        rel_path: Relative path in the repository

    Returns:
        True if successful, False otherwise
    """
    try:
        # Check if a file already exists in the repo
        try:
            contents = repo.get_contents(rel_path)
            # Update the file
            content = _read_file_content(file_path)
            repo.update_file(
                path=rel_path,
                message=f"Update {rel_path}",
                content=content,
                sha=contents.sha
            )
            logger.info(f"Updated file {_sanitize_path(rel_path)} in repository")
        except GithubException as e:
            if e.status == 404:
                # File doesn't exist, create it
                content = _read_file_content(file_path)
                repo.create_file(
                    path=rel_path,
                    message=f"Add {rel_path}",
                    content=content
                )
                logger.info(f"Added file {_sanitize_path(rel_path)} to repository")
            else:
                raise
        return True
    except Exception as e:
        logger.error(f"Error processing file {_sanitize_path(rel_path)}: {e}")
        return False


def _process_file(repo, repo_dir: str, file_path: str) -> None:
    """
    Process a single file for GitHub upload.

    Args:
        repo: GitHub repository object
        repo_dir: Local repository directory
        file_path: Path to the file
    """
    # Get the relative path to use as the file path in the repo
    rel_path = os.path.relpath(file_path, repo_dir)

    # Check if the file is too large for direct API upload
    if _is_file_too_large(file_path):
        logger.info(f"File {_sanitize_path(rel_path)} is larger than 70MB. Using Git LFS for this file.")
        try:
            # Add the file to Git
            subprocess.run(["git", "add", rel_path], check=True, cwd=repo_dir)
            logger.info(f"Added large file {_sanitize_path(rel_path)} to Git (will be handled by LFS)")
        except subprocess.CalledProcessError as e:
            logger.error(f"Error adding large file {_sanitize_path(rel_path)} to Git: {e}")
            # Fall back to direct API upload with a warning
            logger.warning(
                f"Falling back to direct API upload for {_sanitize_path(rel_path)}. This may fail if the file is too large.")
            _upload_file_to_github(repo, file_path, rel_path)
    else:
        # For smaller files, we can use either Git or direct API upload
        # Using Git for consistency
        try:
            # Add the file to Git
            subprocess.run(["git", "add", rel_path], check=True, cwd=repo_dir)
            logger.info(f"Added file {_sanitize_path(rel_path)} to Git")
        except subprocess.CalledProcessError as e:
            logger.error(f"Error adding file {_sanitize_path(rel_path)} to Git: {e}")
            # Fall back to direct API upload
            logger.warning(f"Falling back to direct API upload for {_sanitize_path(rel_path)}")
            _upload_file_to_github(repo, file_path, rel_path)


def _process_directory(repo, repo_dir: str, directory: str) -> None:
    """
    Process all files in a directory for GitHub upload.

    Args:
        repo: GitHub repository object
        repo_dir: Local repository directory
        directory: Directory to process
    """
    # First, count the total number of files to process
    total_files = 0
    file_list = []

    for root, dirs, files in os.walk(directory):
        # Skip .git directory
        if '.git' in dirs:
            dirs.remove('.git')

        for file in files:
            file_path = os.path.join(root, file)
            file_list.append(file_path)
            total_files += 1

    logger.info(f"Found {total_files} files to process")

    # Process each file with progress tracking
    for i, file_path in enumerate(file_list):
        _process_file(repo, repo_dir, file_path)

        # Update progress bar
        progress = (i + 1) / total_files * 100
        file_name = os.path.basename(file_path)
        _display_progress_bar(progress, prefix="Uploading files:", suffix=f"File {i + 1}/{total_files}: {file_name}")

        # Log progress to file
        logger.info(f"Processed file {i + 1}/{total_files}: {_sanitize_path(file_path)} ({progress:.1f}%)")


def _connect_to_github(token: str, repo_name: str) -> Tuple[Any, str, str]:
    """
    Connect to GitHub and get repository information.

    Args:
        token: GitHub personal access token
        repo_name: Name of the repository

    Returns:
        Tuple of (repository_object, repository_directory, username)
    """
    # Create a GitHub instance with the token
    g = Github(token)

    # Get the repository
    repo = g.get_user().get_repo(repo_name)
    logger.info(f"Connected to GitHub repository '{repo_name}'")

    # Get the current directory (where the local repo is)
    repo_dir = os.getcwd()

    # Get the GitHub username
    username = get_github_username(token)
    if not username:
        raise ValueError("Failed to get GitHub username")

    return repo, repo_dir, username


def _process_repository_files(repo: Any, repo_dir: str) -> None:
    """
    Process all files in the repository directory.

    Args:
        repo: GitHub repository object
        repo_dir: Local repository directory
    """
    _process_directory(repo, repo_dir, repo_dir)


def _configure_git_user(username: str, repo_dir: str) -> None:
    """
    Configure Git user information.

    Args:
        username: GitHub username
        repo_dir: Local repository directory
    """
    # Configure a Git user if not already configured
    subprocess.run(
        ["git", "config", "user.email", f"{username}@users.noreply.github.com"],
        check=False,
        cwd=repo_dir
    )
    subprocess.run(
        ["git", "config", "user.name", username],
        check=False,
        cwd=repo_dir
    )
    logger.info("Configured Git user")


def _commit_changes(repo_dir: str) -> bool:
    """
    Commit changes to the local repository.

    Args:
        repo_dir: Local repository directory

    Returns:
        True if changes were committed, False if no changes to commit
    """
    # Check if there are changes to commit
    status_result = subprocess.run(
        ["git", "status", "--porcelain"],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        cwd=repo_dir
    )

    if not status_result.stdout.strip():
        logger.info("No changes to commit")
        return False

    # Try to commit changes
    try:
        subprocess.run(
            ["git", "commit", "-m", "Upload font archives"],
            check=True,
            cwd=repo_dir
        )
        logger.info("Committed changes to local repository")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"Git commit failed: {e}")

        # Try with a --allow-empty flag as a fallback
        subprocess.run(
            ["git", "commit", "--allow-empty", "-m", "Upload font archives"],
            check=True,
            cwd=repo_dir
        )
        logger.info("Committed changes to local repository (with --allow-empty)")
        return True


def _setup_git_remote(username: str, token: str, repo_name: str, repo_dir: str) -> None:
    """
    Set up Git remote for the repository.

    Args:
        username: GitHub username
        token: GitHub personal access token
        repo_name: Name of the repository
        repo_dir: Local repository directory
    """
    # First, check if the remote already exists
    try:
        subprocess.run(
            ["git", "remote", "remove", "origin"],
            check=False,  # Don't fail if remote doesn't exist
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            cwd=repo_dir
        )
    except subprocess.SubprocessError:
        pass  # Ignore errors when removing a remote

    remote_url = f"https://{username}:{token}@github.com/{username}/{repo_name}.git"
    subprocess.run(
        ["git", "remote", "add", "origin", remote_url],
        check=True,
        cwd=repo_dir
    )
    logger.info("Added GitHub repository as remote")


def _get_or_create_branch(repo_dir: str) -> str:
    """
    Get the current branch name or create a new branch.

    Args:
        repo_dir: Local repository directory

    Returns:
        Name of the current branch
    """
    # Get the current branch name
    branch_result = subprocess.run(
        ["git", "branch", "--show-current"],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        cwd=repo_dir
    )

    current_branch = branch_result.stdout.strip()
    if current_branch:
        return current_branch

    # Default to 'main' if the branch name can't be determined
    for branch_name in ["main", "master"]:
        try:
            subprocess.run(
                ["git", "checkout", "-b", branch_name],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                cwd=repo_dir
            )
            logger.info(f"Created and checked out branch: {branch_name}")
            return branch_name
        except subprocess.CalledProcessError:
            logger.warning(f"Failed to create branch: {branch_name}")
            continue

    # If we get here, we couldn't create either branch
    raise RuntimeError("Failed to determine or create a Git branch")


def _push_to_github_with_lfs(repo_name: str, current_branch: str, repo_dir: str) -> bool:
    """
    Push to GitHub with Git LFS.

    Args:
        repo_name: Name of the repository
        current_branch: Name of the current branch
        repo_dir: Local repository directory

    Returns:
        True if the push was successful, False otherwise
    """
    logger.info(f"Pushing to branch: {current_branch}")

    # Try a normal push first
    try:
        subprocess.run(
            ["git", "push", "-u", "origin", current_branch],
            check=True,
            cwd=repo_dir
        )
        logger.info(f"Pushed to GitHub repository '{repo_name}' with Git LFS")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"Git push failed: {e}")

        # Try force push as a last resort
        try:
            logger.warning("Attempting force push...")
            subprocess.run(
                ["git", "push", "-u", "-f", "origin", current_branch],
                check=True,
                cwd=repo_dir
            )
            logger.info(f"Force pushed to GitHub repository '{repo_name}' with Git LFS")
            return True
        except subprocess.CalledProcessError as push_e:
            logger.error(f"Git force push failed: {push_e}")
            return False


def push_to_github(token: str, repo_name: str) -> None:
    """
    Push the local repository to GitHub using Git commands and PyGithub.
    Uses Git LFS for files larger than 70MB.

    Args:
        token: GitHub personal access token
        repo_name: Name of the repository
    """
    try:
        # Connect to GitHub and get repository information
        repo, repo_dir, username = _connect_to_github(token, repo_name)

        # Process all files in the repository directory
        _process_repository_files(repo, repo_dir)

        try:
            # Configure Git user
            _configure_git_user(username, repo_dir)

            # Commit changes
            _commit_changes(repo_dir)

            # Set up Git remote
            _setup_git_remote(username, token, repo_name, repo_dir)

            # Get or create a branch
            current_branch = _get_or_create_branch(repo_dir)

            # Push to GitHub with LFS
            if not _push_to_github_with_lfs(repo_name, current_branch, repo_dir):
                # Fall back to direct API upload if push fails
                logger.warning("Falling back to direct API upload")
                _process_repository_files(repo, repo_dir)
                logger.info(f"Pushed to GitHub repository '{repo_name}' using API")

        except subprocess.CalledProcessError as e:
            logger.error(f"Git command failed: {e}")
            logger.warning("Falling back to direct API upload for remaining files")

            # Fall back to direct API upload for any remaining files
            _process_repository_files(repo, repo_dir)
            logger.info(f"Pushed to GitHub repository '{repo_name}' using API")

    except GithubException as e:
        logger.error(f"GitHub API error: {e}")
    except ValueError as e:
        logger.error(str(e))
    except Exception as e:
        logger.error(f"Error in push_to_github: {e}")
        # Delete the temporary directory before exiting
        delete_temp_directory(ask_confirmation=False)
        sys.exit(1)


def main():
    """Main function to execute the font archiving process."""
    logger.info("Starting font archiving process")

    # Scan for fonts
    font_families = scan_fonts()

    if not font_families:
        logger.error("No non-default fonts found")
        # Delete the temporary directory before exiting
        delete_temp_directory(ask_confirmation=False)
        sys.exit(1)

    # Create an output directory
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Create archive files
    archive_paths, total_size = create_zips(font_families, OUTPUT_DIR)

    # Calculate and display statistics
    total_fonts = sum(len(paths) for paths in font_families.values())
    total_families = len(font_families)

    logger.info(f"Total fonts: {total_fonts}")
    logger.info(f"Total font families: {total_families}")
    logger.info(f"Total archive size: {total_size / 1024 / 1024:.2f} MB")

    # Confirm with user
    print(f"\nFound {total_fonts} fonts in {total_families} families.")
    print(f"Total archive size: {total_size / 1024 / 1024:.2f} MB")
    print("Do you want to proceed with creating the Git repository and uploading to GitHub? (y/n)")

    if input().lower() != 'y':
        logger.info("User chose not to proceed")
        sys.exit(0)

    # Create a local Git repository
    create_git_repo(OUTPUT_DIR, total_families, total_size)

    # Get GitHub token
    token = get_github_token()

    try:
        # Check GitHub LFS storage limits
        if not check_github_lfs_storage(token):
            logger.error("Aborting due to GitHub LFS storage concerns")
            # Delete the temporary directory before exiting
            delete_temp_directory(ask_confirmation=False)
            sys.exit(1)

        # Check GitHub data transfer limits
        if not check_github_data_limits(token, total_size):
            logger.error("Aborting due to GitHub data transfer concerns")
            # Delete the temporary directory before exiting
            delete_temp_directory(ask_confirmation=False)
            sys.exit(1)

        # Create a GitHub repository
        create_github_repo(token, REPO_NAME)

        # Push to GitHub
        push_to_github(token, REPO_NAME)

        logger.info("Font archiving process completed successfully")
        print(
            f"\nFont archiving completed successfully. Repository: https://github.com/{get_github_username(token)}/{REPO_NAME}")

        # Clean up the temporary directory
        delete_temp_directory()
    except Exception as e:
        logger.error(f"Error during GitHub processing: {e}")
        print(f"\nAn error occurred during GitHub processing: {e}")
        print(f"The local repository has been created and can be found in the {REPO_NAME} directory.")

        # Clean up the temporary directory even if there was an error
        delete_temp_directory()
        sys.exit(1)


if __name__ == "__main__":
    main()
    end_time = datetime.datetime.now()
    logger.info(f"Script ended at: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
