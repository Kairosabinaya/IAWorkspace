"""Data classes for per-file configuration."""

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from app.modules.excel_formatter.models.column_info import ColumnInfo


@dataclass
class SheetConfig:
    """Configuration for a single sheet."""

    name: str
    selected: bool = True
    header_row: int = 1  # 1-based
    header_auto_detected: bool = True
    date_columns: List[ColumnInfo] = field(default_factory=list)
    numeric_columns: List[ColumnInfo] = field(default_factory=list)
    all_columns: List[ColumnInfo] = field(default_factory=list)
    total_rows: int = 0


@dataclass
class FileConfig:
    """Configuration for a single Excel file."""

    file_path: str
    file_name: str
    file_size: str  # Human-readable
    relative_dir: str = ""  # Subdirectory relative to output folder (preserves folder structure)
    sheet_configs: Dict[str, SheetConfig] = field(default_factory=dict)
    date_format: str = "DD-MMM-YY"
    separator_style: str = ","  # "," = 1,000.00 | "." = 1.000,00
    freeze_pane: bool = True
    status: str = "Ready"  # Ready | Analyzing... | Processing... | Done | Error
    error_message: str = ""
    progress: float = 0.0  # 0.0 to 1.0
    analyzed: bool = False
