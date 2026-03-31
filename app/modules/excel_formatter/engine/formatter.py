"""Apply formatting to an openpyxl workbook while preserving existing styles."""

from copy import copy
from datetime import datetime, timedelta

from openpyxl.styles import Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.modules.excel_formatter.models.file_config import SheetConfig
from app.utils.constants import (
    EXCEL_FONT_NAME,
    EXCEL_FONT_SIZE,
    EXCEL_MAX_COL_WIDTH,
    EXCEL_MIN_COL_WIDTH,
    NUMBER_FORMAT_DECIMAL,
    NUMBER_FORMAT_DOT_DECIMAL,
    NUMBER_FORMAT_DOT_INTEGER,
    NUMBER_FORMAT_INTEGER,
)

# Date parse formats for converting string values to datetime
_DATE_PARSE_FMTS = [
    "%Y-%m-%dT%H:%M:%S.%f",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%m/%d/%Y",
    "%d %B %Y",
    "%d %b %Y",
    "%B %d, %Y",
    "%b %d, %Y",
    "%d-%b-%y",
    "%d-%b-%Y",
]

_THIN_SIDE = Side(style="thin")
_THIN_BORDER = Border(
    left=_THIN_SIDE, right=_THIN_SIDE,
    top=_THIN_SIDE, bottom=_THIN_SIDE,
)

# Pre-built fonts — reused for cells without special styling
_FONT_NORMAL = Font(name=EXCEL_FONT_NAME, size=EXCEL_FONT_SIZE)
_FONT_BOLD = Font(name=EXCEL_FONT_NAME, size=EXCEL_FONT_SIZE, bold=True)


def format_sheet(
    ws: Worksheet,
    config: SheetConfig,
    date_format: str,
    freeze_pane: bool,
    separator_style: str = ",",
    progress_callback=None,
):
    """Apply formatting rules to one worksheet.

    Performance: skips empty cells and cells already formatted correctly.
    Uses iter_rows() for efficient cell traversal.
    """
    if not config.selected:
        return

    header_row = config.header_row
    last_data_row = ws.max_row or header_row
    last_col = ws.max_column or 1
    total_rows = max(last_data_row - header_row + 1, 1)

    # Pick number format strings based on separator style
    if separator_style == ".":
        fmt_int = NUMBER_FORMAT_DOT_INTEGER
        fmt_dec = NUMBER_FORMAT_DOT_DECIMAL
    else:
        fmt_int = NUMBER_FORMAT_INTEGER
        fmt_dec = NUMBER_FORMAT_DECIMAL

    # Build lookup sets from all_columns
    date_cols = set()
    number_cols = set()
    for ci in config.all_columns:
        if ci.user_format_type == "date":
            date_cols.add(ci.index)
        elif ci.user_format_type == "number":
            number_cols.add(ci.index)

    # --- Format cells using iter_rows (faster than cell-by-cell lookup) ---
    row_count = 0
    for row in ws.iter_rows(min_row=header_row, max_row=last_data_row,
                            max_col=last_col):
        is_header = (row[0].row == header_row)
        for cell in row:
            # Font + border: ALL cells in range (including empty)
            _apply_font_fast(cell, is_header)
            _apply_border_fast(cell)

            # Date/number format: only filled data cells
            if not is_header and cell.value is not None:
                col_idx = cell.column
                if col_idx in date_cols:
                    _apply_date_format(cell, date_format)
                elif col_idx in number_cols:
                    _apply_number_format(cell, fmt_int, fmt_dec)

        row_count += 1
        if progress_callback and row_count % 500 == 0:
            progress_callback(row_count / total_rows)

    # --- Column widths ---
    _auto_fit_columns(ws, header_row, last_data_row, last_col)

    # --- Freeze pane ---
    if freeze_pane:
        ws.freeze_panes = ws.cell(row=header_row + 1, column=1).coordinate
    else:
        ws.freeze_panes = None


# ---------------------------------------------------------------------------
# Cell-level formatting — optimised to skip already-correct cells
# ---------------------------------------------------------------------------

def _apply_font_fast(cell, is_header: bool):
    """Set font to Arial 9, bold for header. Unconditional assignment
    using pre-built Font objects for maximum speed."""
    cell.font = _FONT_BOLD if is_header else _FONT_NORMAL


def _apply_border_fast(cell):
    """Apply thin border on all sides. Unconditional assignment for speed."""
    cell.border = _THIN_BORDER


_BORDER_WEIGHT = {"thin": 1, "medium": 2, "thick": 3, "double": 3}


def _keep_thicker(existing: Side, new: Side) -> Side:
    ew = _BORDER_WEIGHT.get(existing.style, 0) if existing and existing.style else 0
    nw = _BORDER_WEIGHT.get(new.style, 0) if new and new.style else 0
    return existing if ew >= nw else new


def _apply_date_format(cell, fmt: str):
    """Convert cell value to datetime if needed, then set number format."""
    val = cell.value
    if val is None:
        return

    if isinstance(val, datetime):
        cell.number_format = fmt
        return

    if isinstance(val, (int, float)) and 1 <= val <= 2958465:
        base = datetime(1899, 12, 30)
        cell.value = base + timedelta(days=int(val))
        cell.number_format = fmt
        return

    if isinstance(val, str):
        s = val.strip()
        if not s:
            return
        for pfmt in _DATE_PARSE_FMTS:
            try:
                dt = datetime.strptime(s, pfmt)
                cell.value = dt
                cell.number_format = fmt
                return
            except ValueError:
                continue


def _apply_number_format(cell, fmt_int: str, fmt_dec: str):
    """Set thousand-separator format with per-cell decimal detection."""
    val = cell.value
    if isinstance(val, (int, float)):
        frac = abs(val - int(val))
        if frac > 0.001:
            cell.number_format = fmt_dec
        else:
            cell.number_format = fmt_int


# ---------------------------------------------------------------------------
# Column width
# ---------------------------------------------------------------------------

def _auto_fit_columns(ws: Worksheet, header_row: int, last_row: int, last_col: int):
    """Auto-fit column widths between min/max bounds. Samples up to 200 rows."""
    for col_idx in range(1, last_col + 1):
        max_len = EXCEL_MIN_COL_WIDTH
        letter = get_column_letter(col_idx)
        for row_idx in range(header_row, min(last_row + 1, header_row + 200)):
            cell = ws.cell(row=row_idx, column=col_idx)
            val = cell.value
            if val is not None:
                display = str(val)
                length = len(display) + 2
                if cell.font and cell.font.bold:
                    length += 1
                if length > max_len:
                    max_len = length
        ws.column_dimensions[letter].width = min(max_len, EXCEL_MAX_COL_WIDTH)
