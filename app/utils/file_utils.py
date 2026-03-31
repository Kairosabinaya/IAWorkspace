"""Shared file I/O helpers."""

import os
from pathlib import Path

from app.utils.constants import SUPPORTED_EXTENSIONS


def is_valid_excel(file_path: str) -> bool:
    """Check if file exists and has a supported extension."""
    p = Path(file_path)
    return p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS


def get_file_size_display(file_path: str) -> str:
    """Return human-readable file size string."""
    size = os.path.getsize(file_path)
    if size < 1024:
        return f"{size} B"
    elif size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    else:
        return f"{size / (1024 * 1024):.1f} MB"


def ensure_output_folder(folder_path: str) -> Path:
    """Create the output folder if it doesn't exist and return its Path."""
    p = Path(folder_path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def get_default_output_folder(first_file_path: str) -> str:
    """Return default output folder path based on the first dropped file."""
    from app.utils.constants import DEFAULT_OUTPUT_FOLDER
    parent = Path(first_file_path).parent
    return str(parent / DEFAULT_OUTPUT_FOLDER)
