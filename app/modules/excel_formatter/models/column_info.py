"""Data classes describing column analysis results."""

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


class ColumnType(Enum):
    TEXT = "text"
    DATE = "date"
    NUMERIC_AMOUNT = "amount"
    NUMERIC_ID = "id_code"
    UNKNOWN = "unknown"


@dataclass
class ColumnInfo:
    """Analysis result for a single column."""

    index: int  # 1-based column index
    letter: str  # Column letter (A, B, C, ...)
    header_name: str  # Text in the header row
    detected_type: ColumnType = ColumnType.UNKNOWN
    confidence: float = 0.0  # 0.0 to 1.0
    sample_values: List[str] = field(default_factory=list)  # Up to 3 display samples
    has_decimals: bool = False  # For numeric columns: needs decimal places?
    is_recommended: bool = False  # Pre-checked in the UI?
    user_selected: bool = False  # User manually toggled?
    user_format_type: Optional[str] = None  # None = no format, "date", "number"
    format_preview: str = ""  # How data will look after formatting
