# Font Archiver

A Python script to scan, organize, and archive font files from Windows systems to GitHub.

## Features

- Scans Windows font directories (system and user)
- Excludes default Windows fonts
- Groups fonts by family
- Creates zip archives for each font family
- Uses Python's built-in zipfile module for compression
- Creates a local Git repository
- Uploads to GitHub with LFS support (automatically uses LFS for files larger than 70MB)
- Provides detailed logging and progress indicators

## Requirements

- Python 3.x
- Git with LFS support (install from https://git-lfs.github.com/)
- GitHub personal access token (for repository creation)

## Installation

1. Clone this repository
2. Install Git LFS if you haven't already:
   ```bash
   git lfs install
   ```
3. Install dependencies using Poetry:
   ```bash
   poetry install
   ```

## Usage

Run the script using Poetry:

```bash
poetry run python font_archiver.py
```

The script will:

1. Scan for fonts in Windows directories
2. Group them by family
3. Create zip archives
4. Show statistics and ask for confirmation
5. Create a local Git repository
6. Upload to GitHub (requires a personal access token)

## GitHub Authentication

The script will look for a file named `github_token.txt` in the script directory. If found, it will use the token
from that file. Otherwise, it will prompt you to enter your GitHub personal access token.

## Output

- A local directory named "Font-Storage" containing zip files for each font family
- A Git repository in the parent directory containing:
  - The contents of the Font-Storage directory (zip files)
  - README.md with statistics and disclaimer
  - .gitattributes for LFS configuration (tracks .zip, .7z, and any file >70MB)
  - .gitignore file
- A GitHub repository named "Font-Storage" with the contents of the local Git repository
- A log file named "font-upload.log"

## Disclaimer

The commercial status of fonts processed by this script is unknown. The script makes no claim to ownership of
these items. Fonts are provided "as is" without warranty of any kind, either expressed or implied.

In the event that the contents of the repository fall under copyright, the repository owner makes no claim to 
its contents. All fonts were obtained from openly available locations.
