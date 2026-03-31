"""Heuristic logic for detecting column types (date, amount, ID/code)."""

import re
import statistics
from datetime import datetime
from typing import List, Optional, Tuple

from openpyxl.cell.cell import Cell

from app.utils.constants import (
    AMOUNT_COLUMN_KEYWORDS,
    DATE_COLUMN_KEYWORDS,
    ID_COLUMN_KEYWORDS,
)


# ---------------------------------------------------------------------------
# Header row detection
# ---------------------------------------------------------------------------

def compute_header_score(row_cells: List[Cell], next_row_cells: Optional[List[Cell]]) -> float:
    """Score a row on how likely it is to be the table header (0.0–1.0)."""
    if not row_cells:
        return 0.0

    filled = [c for c in row_cells if c.value is not None and str(c.value).strip()]
    if len(filled) < 3:
        return 0.0

    total = len(row_cells)
    fill_ratio = len(filled) / total if total else 0

    # Check for wide merged cells (title rows) — penalise
    for c in row_cells:
        if isinstance(getattr(c, "value", None), str):
            pass  # can't easily check merge from cell alone; handled in analyzer

    # Percentage of text (non-numeric) cells
    text_count = 0
    for c in filled:
        val = c.value
        if isinstance(val, str) and not _looks_numeric(val):
            text_count += 1
    text_ratio = text_count / len(filled) if filled else 0

    # Next row has different types (numbers/dates) → strong header signal
    type_diff_bonus = 0.0
    if next_row_cells:
        next_filled = [c for c in next_row_cells if c.value is not None]
        next_numeric = sum(1 for c in next_filled if _is_numeric_cell(c))
        if next_filled and next_numeric / len(next_filled) > 0.3:
            type_diff_bonus = 0.25

    score = (text_ratio * 0.50) + (fill_ratio * 0.25) + type_diff_bonus
    return min(score, 1.0)


# ---------------------------------------------------------------------------
# Date column detection
# ---------------------------------------------------------------------------

_DATE_PATTERNS = [
    re.compile(r"\d{4}[/\-\.]\d{1,2}[/\-\.]\d{1,2}[T ]\d{2}:\d{2}:\d{2}"),  # YYYY-MM-DDTHH:MM:SS
    re.compile(r"\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4}"),      # DD/MM/YYYY etc.
    re.compile(r"\d{4}[/\-\.]\d{1,2}[/\-\.]\d{1,2}"),          # YYYY-MM-DD
    re.compile(r"\d{1,2}\s+\w{3,9}\s+\d{2,4}", re.I),          # 01 August 2024
    re.compile(r"\w{3,9}\s+\d{1,2},?\s+\d{2,4}", re.I),        # Aug 01, 2024
]

_DATE_FORMAT_CHARS = {"y", "m", "d"}


def is_date_cell(cell: Cell) -> bool:
    """Heuristic: does this cell contain a date-like value?"""
    val = cell.value
    if val is None:
        return False
    if isinstance(val, datetime):
        return True

    # Check number format for date patterns
    nf = (cell.number_format or "").lower()
    if nf and nf != "general" and any(ch in nf for ch in _DATE_FORMAT_CHARS):
        return True

    # String matching
    if isinstance(val, str):
        s = val.strip()
        for pat in _DATE_PATTERNS:
            if pat.fullmatch(s) or pat.match(s):
                return True

    # Excel serial date (number in a plausible date range with date format)
    if isinstance(val, (int, float)) and 1 <= val <= 60000:
        if nf and nf != "general" and any(ch in nf for ch in _DATE_FORMAT_CHARS):
            return True

    return False


def header_name_suggests_date(name: str) -> bool:
    """Check if column header text hints at date content."""
    words = set(re.split(r"[\s_\-./]+", name.lower()))
    return bool(words & DATE_COLUMN_KEYWORDS)


# ---------------------------------------------------------------------------
# Numeric column detection
# ---------------------------------------------------------------------------

def header_name_suggests_amount(name: str) -> bool:
    words = set(re.split(r"[\s_\-./]+", name.lower()))
    return bool(words & AMOUNT_COLUMN_KEYWORDS)


def header_name_suggests_id(name: str) -> bool:
    words = set(re.split(r"[\s_\-./]+", name.lower()))
    return bool(words & ID_COLUMN_KEYWORDS)


def classify_numeric_column(
    values: list, header_name: str
) -> Tuple[str, bool]:
    """Classify a numeric column.

    Returns:
        ("amount" | "id_code", has_decimals)
    """
    if not values:
        return "id_code", False

    # Keyword hints
    name_says_id = header_name_suggests_id(header_name)
    name_says_amount = header_name_suggests_amount(header_name)

    nums = []
    has_leading_zero = False
    has_meaningful_decimal = False

    for v in values:
        if v is None:
            continue
        if isinstance(v, str):
            s = v.strip()
            if s.startswith("0") and len(s) > 1 and s.isdigit():
                has_leading_zero = True
            try:
                v = float(s.replace(",", ""))
            except ValueError:
                continue
        if isinstance(v, (int, float)):
            nums.append(float(v))
            frac = v - int(v)
            if abs(frac) > 0.001:
                has_meaningful_decimal = True

    if not nums:
        return "id_code", False

    abs_vals = [abs(n) for n in nums]
    max_val = max(abs_vals)
    has_large = max_val > 999

    # Sequential check (id-like)
    is_sequential = False
    if len(nums) >= 3:
        diffs = [nums[i + 1] - nums[i] for i in range(len(nums) - 1)]
        non_zero = [d for d in diffs if d != 0]
        if non_zero and all(d == non_zero[0] for d in non_zero):
            is_sequential = True

    # Same digit count (code-like)
    digit_counts = set()
    for n in nums:
        if n == int(n) and n >= 0:
            digit_counts.add(len(str(int(n))))
    same_digits = len(digit_counts) == 1 and len(nums) >= 3

    # Low variance (id-like)
    low_variance = False
    if len(nums) >= 3:
        mean = statistics.mean(abs_vals)
        if mean > 0:
            std = statistics.stdev(abs_vals)
            low_variance = (std / mean) < 0.5

    # Scoring
    id_score = 0
    amount_score = 0

    if name_says_id:
        id_score += 3
    if name_says_amount:
        amount_score += 3
    if is_sequential:
        id_score += 2
    if has_leading_zero:
        id_score += 2
    if same_digits:
        id_score += 1
    if low_variance and not name_says_amount:
        id_score += 1
    if has_large:
        amount_score += 1
    if has_meaningful_decimal:
        amount_score += 1

    # Magnitude variation → amount-like
    if len(nums) >= 3 and max_val > 0:
        min_nonzero = min((v for v in abs_vals if v > 0), default=max_val)
        if max_val / min_nonzero > 10:
            amount_score += 1

    if id_score > amount_score:
        return "id_code", has_meaningful_decimal
    return "amount", has_meaningful_decimal


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _looks_numeric(s: str) -> bool:
    s = s.strip().replace(",", "").replace("%", "")
    try:
        float(s)
        return True
    except ValueError:
        return False


def _is_numeric_cell(cell: Cell) -> bool:
    val = cell.value
    if isinstance(val, (int, float)):
        return True
    if isinstance(val, str):
        return _looks_numeric(val)
    if isinstance(val, datetime):
        return True
    return False
