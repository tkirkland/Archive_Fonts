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


# Configure console logging first
console_handler = ColoredConsoleHandler()
console_formatter = NoMicrosecondsFormatter('%(asctime)s - %(levelname)s - %(message)s')
console_handler.setFormatter(console_formatter)

# Configure the root logger with just a console handler initially
logger: Logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(console_handler)
# Remove default handlers
logger.propagate = False


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

    logger.info(f"Using temporary directory: {temp_dir}")
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
logger.info(f"Log file created at: {log_file_path}")

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
            print(f"\nDo you want to delete the temporary directory? ({TEMP_DIR}) (y/n)")
            if input().lower() == 'y':
                logger.info(f"Deleting temporary directory: {TEMP_DIR}")
                shutil.rmtree(TEMP_DIR, ignore_errors=True)
                print(f"Temporary directory deleted: {TEMP_DIR}")
            else:
                logger.info("User chose not to delete the temporary directory")
                print(f"Temporary directory preserved: {TEMP_DIR}")
        else:
            # Delete it without asking
            logger.info(f"Deleting temporary directory without confirmation: {TEMP_DIR}")
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
        logger.warning(f"Could not extract family name from {font_path}: {e}")
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

    logger.info(f"Scanning fonts directory: {directory}")

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
        logger.info(f"Scanning user fonts directory: {LOCAL_FONTS_DIR}")
        process_fonts_directory(LOCAL_FONTS_DIR, font_families, processed_fonts, add_to_processed=True)

    # Then scan system fonts
    if os.path.exists(WINDOWS_FONTS_DIR):
        logger.info(f"Scanning system fonts directory: {WINDOWS_FONTS_DIR}")
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

    logger.debug(f"Removing directory: {temp_dir}")

    # Use shutil.rmtree with ignore_errors=True for simplicity
    try:
        shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception as e:
        logger.warning(f"Failed to remove directory {temp_dir}: {e}")


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
        logger.error(f"Temporary directory {temp_dir} does not exist for {family_name}")
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
        logger.error(f"Error changing to directory {directory} for {family_name}: {str(e)}")
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
            logger.warning(f"Could not remove existing archive file {zip_path}: {e}")
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
        logger.error(f"Missing font files for {zip_path}: {missing_fonts}")
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
                    logger.error(f"Error adding {font_path} to zip: {str(e)}")
                    # Continue with other files
        return True
    except zipfile.BadZipFile as e:
        logger.error(f"Bad zip file error for {zip_path}: {str(e)}")
        return False
    except PermissionError as e:
        logger.error(f"Permission error creating zip file {zip_path}: {str(e)}")
        return False
    except Exception as e:
        logger.error(f"Error creating zip file {zip_path}: {str(e)}")
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
        logger.error(f"Zip file {zip_path} was not created")
        return False

    # Verify the zip file is valid
    try:
        with zipfile.ZipFile(zip_path, 'r') as zipf:
            # Test the integrity of the zip file
            if zipf.testzip() is not None:
                logger.error(f"Zip file {zip_path} is corrupted")
                return False
        return True
    except Exception as e:
        logger.error(f"Error verifying zip file {zip_path}: {str(e)}")
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
        logger.error(f"7z file {zip_path} was not created")
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
            logger.error(f"7z file {zip_path} is corrupted: {process.stderr}")
            return False

        return True
    except FileNotFoundError:
        logger.error("7zip command-line tool (7z) not found. Cannot verify 7z file.")
        return False
    except Exception as e:
        logger.error(f"Error verifying 7z file {zip_path}: {str(e)}")
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
        logger.error(f"Error getting size of archive file {zip_path}: {str(e)}")
        return 0


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

        # Prepare the 7zip command with required switches
        cmd = [
            "7z", "a",  # Add to archive
            "-t7z",  # 7z archive type
            f"-mx={compression_level}",  # Compression level based on CPU cores
            "-m0=lzma2",  # LZMA2 compression method
            zip_path  # Output file
        ]

        # Add all font files to the command
        cmd.extend(font_paths)

        # Execute the 7zip command
        process = subprocess.run(
            cmd, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE,
            text=True,
            check=False
        )

        # Check if the command was successful
        if process.returncode != 0:
            logger.error(f"7zip command failed with return code {process.returncode}: {process.stderr}")
            return False

        # Verify the 7z file was created and is valid
        if not _verify_7z_file(zip_path):
            return False

        logger.info(f"Successfully created 7z archive {zip_path} with compression level {compression_level}")
        return True

    except FileNotFoundError:
        logger.error("7zip command-line tool (7z) not found. Please install 7zip and ensure it's in your PATH.")
        return False
    except Exception as e:
        logger.error(f"Error creating 7z archive {zip_path}: {str(e)}")
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

                # Log progress
                completed = len(archive_paths)
                progress = completed / total_families * 100
                logger.info(
                    f"Progress: {progress:.1f}% - Created archive for {family} ({archive_size / 1024 / 1024:.2f} MB)")

                # Check if Ctrl+C was pressed
                if exit_flag:
                    logger.info("Exiting after completing current archive operation due to Ctrl+C")
                    return archive_paths, total_size

            except Exception as e:
                logger.error(f"Error creating archive for {family}: {e}")

    return archive_paths, total_size


def create_git_repo(output_dir: str, total_families: int, total_size: int) -> None:
    """
    Prepare files for the GitHub repository without using local Git commands.

    Args:
        output_dir: Directory for the repository
        total_families: Number of font families
        total_size: Total size of all zip files in bytes
    """
    logger.info("Preparing files for GitHub repository...")

    # Create a parent directory for the repository
    repo_dir = os.path.dirname(output_dir)
    os.chdir(repo_dir)

    # Create README.md with a disclaimer
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

    # Add .gitattributes file for Git LFS
    with open(os.path.join(repo_dir, ".gitattributes"), 'w') as f:
        f.write("*.zip filter=lfs diff=lfs merge=lfs -text\n")
        f.write("*.7z filter=lfs diff=lfs merge=lfs -text\n")

    # Copy the .gitignore file to the repository if it exists
    script_dir = os.path.dirname(os.path.abspath(__file__))
    gitignore_path = os.path.join(script_dir, ".gitignore")
    if os.path.exists(gitignore_path):
        shutil.copy2(gitignore_path, os.path.join(repo_dir, ".gitignore"))

    # Copy the contents of the font-archive directory to the repository directory
    for item in os.listdir(output_dir):
        item_path = os.path.join(output_dir, item)
        dest_path = os.path.join(repo_dir, item)

        # Skip if the item already exists at the destination
        if os.path.exists(dest_path):
            continue

        # Copy the item to the repository directory
        if os.path.isdir(item_path):
            shutil.copytree(item_path, dest_path)
        else:
            shutil.copy2(item_path, dest_path)

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
                sys.exit(1)
    except GithubException as e:
        logger.error(f"Error checking if repository exists: {e}")
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
        sys.exit(1)


# These functions have been replaced by PyGithub implementations in create_github_repo


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
        try:
            repo = user.get_repo(repo_name)
            print(f"\nRepository '{repo_name}' already exists.")
            print("Do you want to delete it and start fresh? (y/n)")
            if input().lower() == 'y':
                try:
                    repo.delete()
                    logger.info(f"Deleted repository '{repo_name}'")
                    # Wait a moment for the deletion to complete
                    time.sleep(2)
                except GithubException as e:
                    logger.error(f"Failed to delete repository: {e}")
                    sys.exit(1)
            else:
                logger.info("Will append to existing repository")
                return
        except GithubException as e:
            if e.status != 404:  # 404 means repo doesn't exist, which is fine
                logger.error(f"Error checking if repository exists: {e}")
                sys.exit(1)

        # Create a new repo
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
        except GithubException as e:
            logger.error(f"Failed to create repository: {e}")
            sys.exit(1)
    except Exception as e:
        logger.error(f"Error in create_github_repo: {e}")
        sys.exit(1)


def check_github_lfs_storage(token: str) -> bool:
    """
    Check if there is enough GitHub LFS storage for the upload using PyGithub.

    Args:
        token: GitHub personal access token

    Returns:
        True if there is enough storage, False otherwise
    """
    try:
        g = Github(token)
        user = g.get_user()

        # Get plan information
        plan_name = user.plan.name if hasattr(user, 'plan') and hasattr(user.plan, 'name') else "unknown"
        logger.info(f"Checking GitHub LFS storage limits... (Account plan: {plan_name})")

        # For demonstration purposes, we'll check if the account plan allows LFS
        # Free accounts have limited LFS storage
        if plan_name.lower() == "free":
            logger.warning("You are using a free GitHub account with limited LFS storage.")
            logger.warning("Large uploads may fail if you exceed your storage quota.")

            # Ask a user if they want to continue
            print("\nYou are using a free GitHub account with limited LFS storage.")
            print("Large uploads may fail if you exceed your storage quota.")
            print("Do you want to continue? (y/n)")

            if input().lower() != 'y':
                logger.info("User chose not to proceed due to LFS storage concerns")
                return False

        return True
    except GithubException as e:
        logger.error(f"Error checking GitHub LFS storage: {e}")
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
    try:
        g = Github(token)

        # Get rate limit information
        rate_limit = g.get_rate_limit()
        logger.info(f"GitHub API rate limit: {rate_limit.core.remaining}/{rate_limit.core.limit}")

        logger.info(f"Checking GitHub data transfer limits... (Upload size: {total_size / 1024 / 1024:.2f} MB)")

        # For demonstration purposes, we'll check if the size is reasonable
        size_mb = total_size / 1024 / 1024

        if size_mb > 1000:  # 1 GB
            logger.warning(f"Upload size is large: {size_mb:.2f} MB")
            logger.warning("GitHub has monthly data transfer limits that may affect your upload.")

            # Ask a user if they want to continue
            print(f"\nUpload size is large: {size_mb:.2f} MB")
            print("GitHub has monthly data transfer limits that may affect your upload.")
            print("Do you want to continue? (y/n)")

            if input().lower() != 'y':
                logger.info("User chose not to proceed due to data transfer concerns")
                return False

        return True
    except GithubException as e:
        logger.error(f"Error checking GitHub data limits: {e}")
        return False


def _is_file_too_large(file_path: str) -> bool:
    """
    Check if a file is too large for direct GitHub API upload.

    Args:
        file_path: Path to the file

    Returns:
        True if the file is too large (>100MB), False otherwise
    """
    # GitHub recommends using Git LFS for files larger than 100MB
    return os.path.getsize(file_path) > 100 * 1024 * 1024  # 100MB


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
            logger.info(f"Updated file {rel_path} in repository")
        except GithubException as e:
            if e.status == 404:
                # File doesn't exist, create it
                content = _read_file_content(file_path)
                repo.create_file(
                    path=rel_path,
                    message=f"Add {rel_path}",
                    content=content
                )
                logger.info(f"Added file {rel_path} to repository")
            else:
                raise
        return True
    except Exception as e:
        logger.error(f"Error processing file {rel_path}: {e}")
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

    # Skip files that are too large for direct API upload
    if _is_file_too_large(file_path):
        logger.warning(f"File {rel_path} is too large for direct API upload. "
                      f"Please use Git LFS for this file.")
        return

    # Upload the file to GitHub
    _upload_file_to_github(repo, file_path, rel_path)


def _process_directory(repo, repo_dir: str, directory: str) -> None:
    """
    Process all files in a directory for GitHub upload.

    Args:
        repo: GitHub repository object
        repo_dir: Local repository directory
        directory: Directory to process
    """
    # Get all files in the directory
    for root, dirs, files in os.walk(directory):
        # Skip .git directory
        if '.git' in dirs:
            dirs.remove('.git')

        # Process each file
        for file in files:
            file_path = os.path.join(root, file)
            _process_file(repo, repo_dir, file_path)


def push_to_github(token: str, repo_name: str) -> None:
    """
    Push the local repository to GitHub using PyGithub.

    Args:
        token: GitHub personal access token
        repo_name: Name of the repository
    """
    try:
        # Create a GitHub instance with the token
        g = Github(token)

        # Get the repository
        repo = g.get_user().get_repo(repo_name)
        logger.info(f"Connected to GitHub repository '{repo_name}'")

        # Get the current directory (where the local repo is)
        repo_dir = os.getcwd()

        # Process all files in the repository directory
        _process_directory(repo, repo_dir, repo_dir)

        logger.info(f"Pushed to GitHub repository '{repo_name}'")
    except GithubException as e:
        logger.error(f"GitHub API error: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Error in push_to_github: {e}")
        sys.exit(1)


def main():
    """Main function to execute the font archiving process."""
    logger.info("Starting font archiving process")

    # Scan for fonts
    font_families = scan_fonts()

    if not font_families:
        logger.error("No non-default fonts found")
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
            sys.exit(1)

        # Check GitHub data transfer limits
        if not check_github_data_limits(token, total_size):
            logger.error("Aborting due to GitHub data transfer concerns")
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
        print("The local repository has been created and can be found in the Font-Storage directory.")

        # Clean up the temporary directory even if there was an error
        delete_temp_directory()
        sys.exit(1)


if __name__ == "__main__":
    main()
