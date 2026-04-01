"""Fast XML-based Excel formatting engine.

Instead of openpyxl's Python-level cell iteration (which triggers expensive
hash/dedup/StyleArray overhead per cell), this engine manipulates the OOXML
XML directly via lxml — achieving ~8-15x speedup for large files.

Architecture:
  1. Read .xlsx as ZIP
  2. Parse xl/styles.xml, add font/border/numFmt entries
  3. Build remap table: (old_style_index, format_type) → new_style_index
  4. Parse each sheet's XML, update cell style indices via remap
  5. Convert date strings → serial numbers where needed
  6. Set column widths and freeze pane
  7. Repack ZIP

Preserves: fills/colors, merged cells, conditional formatting, formulas,
data validation, charts, images — anything stored outside styles.xml and
the cell `s` attribute is untouched.
"""

import copy
import os
import re
import zipfile
from datetime import datetime, timedelta
from typing import Callable, Dict, List, Optional, Set, Tuple

from lxml import etree

from app.modules.excel_formatter.models.file_config import FileConfig, SheetConfig
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

# ---------------------------------------------------------------------------
# OOXML namespaces
# ---------------------------------------------------------------------------

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _tag(name: str) -> str:
    """Qualified element tag in the spreadsheetml namespace."""
    return f"{{{_NS}}}{name}"


# Built-in number format IDs — these do NOT need numFmts entries.
_BUILTIN_NUMFMT: Dict[str, int] = {
    "General": 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4,
    "#,##0;-#,##0": 5, "#,##0;[Red]-#,##0": 6,
    "#,##0.00;-#,##0.00": 7, "#,##0.00;[Red]-#,##0.00": 8,
    "0%": 9, "0.00%": 10, "0.00E+00": 11,
    "mm-dd-yy": 14, "d-mmm-yy": 15, "d-mmm": 16, "mmm-yy": 17,
    "h:mm AM/PM": 18, "h:mm:ss AM/PM": 19, "h:mm": 20, "h:mm:ss": 21,
    "m/d/yy h:mm": 22,
}

# ---------------------------------------------------------------------------
# Date parsing — ported verbatim from formatter.py (3-tier strategy)
# ---------------------------------------------------------------------------

_DATE_NUM_RE = re.compile(r"(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{1,4})")
_DATE_NAMED_RE = re.compile(
    r"(\d{1,2})[/\-.\s]+([A-Za-z]{3,9})[/\-.\s]+(\d{2,4})", re.I,
)
_DATE_NAMED_REV_RE = re.compile(
    r"([A-Za-z]{3,9})[/\-.\s]+(\d{1,2}),?\s*(\d{2,4})", re.I,
)
_MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
    "january": 1, "february": 2, "march": 3, "april": 4,
    "june": 6, "july": 7, "august": 8, "september": 9,
    "october": 10, "november": 11, "december": 12,
}
_DATE_PARSE_FMTS = [
    "%Y-%m-%dT%H:%M:%S.%f",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
]
_EXCEL_EPOCH = datetime(1899, 12, 30)

# ---------------------------------------------------------------------------
# Column-reference helpers
# ---------------------------------------------------------------------------


def _col_index(ref: str) -> int:
    """1-based column index from a cell reference like 'AB12'."""
    col = 0
    for ch in ref:
        if ch.isalpha():
            col = col * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return col


def _col_letter(idx: int) -> str:
    """Column letter(s) from 1-based index.  1→A, 27→AA."""
    result = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def _cell_ref(col_idx: int, row_num: int) -> str:
    return f"{_col_letter(col_idx)}{row_num}"


def _sort_row_cells(row_el, c_tag: str) -> None:
    """Sort <c> elements within a row by column reference (OOXML requirement)."""
    cells = [c for c in row_el if c.tag == c_tag]
    row_el[:] = []
    cells.sort(key=lambda c: _col_index(c.get("r", "")))
    for c in cells:
        row_el.append(c)


# ---------------------------------------------------------------------------
# Shared strings
# ---------------------------------------------------------------------------


def _parse_shared_strings(xml_bytes: Optional[bytes]) -> List[str]:
    """Parse xl/sharedStrings.xml into a flat list of strings."""
    if not xml_bytes:
        return []
    tree = etree.fromstring(xml_bytes)
    strings: List[str] = []
    for si in tree.findall(_tag("si")):
        # Simple <si><t>text</t></si>
        t_el = si.find(_tag("t"))
        if t_el is not None and t_el.text is not None:
            strings.append(t_el.text)
            continue
        # Rich text <si><r><t>…</t></r>…</si>
        parts = []
        for r in si.findall(_tag("r")):
            rt = r.find(_tag("t"))
            if rt is not None and rt.text:
                parts.append(rt.text)
        strings.append("".join(parts))
    return strings


# ---------------------------------------------------------------------------
# Sheet-name → ZIP-path resolution
# ---------------------------------------------------------------------------


def _resolve_sheet_paths(zip_contents: Dict[str, bytes]) -> Dict[str, str]:
    """Return {sheet_name: zip_path} from workbook.xml + rels."""
    wb_xml = zip_contents.get("xl/workbook.xml")
    rels_xml = zip_contents.get("xl/_rels/workbook.xml.rels")
    if not wb_xml or not rels_xml:
        return {}

    # Relationship Id → Target
    rels_tree = etree.fromstring(rels_xml)
    rid_target: Dict[str, str] = {}
    for rel in rels_tree:
        rid_target[rel.get("Id", "")] = rel.get("Target", "")

    # Sheet name → rId → Target
    wb_tree = etree.fromstring(wb_xml)
    sheets_el = wb_tree.find(_tag("sheets"))
    if sheets_el is None:
        return {}

    result: Dict[str, str] = {}
    for s in sheets_el.findall(_tag("sheet")):
        name = s.get("name", "")
        rid = s.get(f"{{{_NS_R}}}id", "")
        target = rid_target.get(rid, "")
        if target:
            # Target is relative to xl/
            path = f"xl/{target}" if not target.startswith("/") else target.lstrip("/")
            result[name] = path
    return result


# ---------------------------------------------------------------------------
# Cell value helpers
# ---------------------------------------------------------------------------


def _get_cell_text(cell_el, shared_strings: List[str]) -> Optional[str]:
    """Return the string representation of a cell's value."""
    t = cell_el.get("t", "")
    v_el = cell_el.find(_tag("v"))

    if t == "s" and v_el is not None and v_el.text:
        try:
            idx = int(v_el.text)
            if 0 <= idx < len(shared_strings):
                return shared_strings[idx]
        except (ValueError, IndexError):
            pass
        return None

    if t == "inlineStr":
        is_el = cell_el.find(_tag("is"))
        if is_el is not None:
            t_el = is_el.find(_tag("t"))
            if t_el is not None and t_el.text is not None:
                return t_el.text
            parts = []
            for r in is_el.findall(_tag("r")):
                rt = r.find(_tag("t"))
                if rt is not None and rt.text:
                    parts.append(rt.text)
            return "".join(parts) if parts else None
        return None

    if v_el is not None and v_el.text:
        return v_el.text
    return None


# ---------------------------------------------------------------------------
# Date conversion (same 3-tier logic as formatter.py)
# ---------------------------------------------------------------------------


def _try_parse_date(s: str) -> Optional[datetime]:
    """Parse a string as a date.  Returns datetime or None."""
    s = s.strip()
    if not s:
        return None

    # Tier 1: fromisoformat
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        pass

    # Tier 2a: numeric date regex (DD/MM/YYYY, YYYY-MM-DD, …)
    m = _DATE_NUM_RE.fullmatch(s)
    if m:
        a, b, c = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a > 31:
            yr, mo, dy = a, b, c
        else:
            dy, mo, yr = a, b, c
        if yr < 100:
            yr += 2000 if yr < 50 else 1900
        if 1 <= mo <= 12 and 1 <= dy <= 31:
            try:
                return datetime(yr, mo, dy)
            except ValueError:
                pass

    # Tier 2b: named-month (01-Aug-24, 01 August 2024)
    m = _DATE_NAMED_RE.fullmatch(s)
    if m:
        dy_s, mon_str, yr_s = m.group(1), m.group(2).lower(), m.group(3)
        mo = _MONTH_MAP.get(mon_str)
        if mo:
            dy, yr = int(dy_s), int(yr_s)
            if yr < 100:
                yr += 2000 if yr < 50 else 1900
            try:
                return datetime(yr, mo, dy)
            except ValueError:
                pass

    # Tier 2c: reverse named-month (Aug 01, 2024)
    m = _DATE_NAMED_REV_RE.fullmatch(s)
    if m:
        mon_str, dy_s, yr_s = m.group(1).lower(), m.group(2), m.group(3)
        mo = _MONTH_MAP.get(mon_str)
        if mo:
            dy, yr = int(dy_s), int(yr_s)
            if yr < 100:
                yr += 2000 if yr < 50 else 1900
            try:
                return datetime(yr, mo, dy)
            except ValueError:
                pass

    # Tier 3: strptime fallback (timestamps)
    for pfmt in _DATE_PARSE_FMTS:
        try:
            return datetime.strptime(s, pfmt)
        except ValueError:
            continue
    return None


def _datetime_to_serial(dt: datetime) -> float:
    """Convert a datetime to an Excel serial date number.

    Uses the same epoch (1899-12-30) as openpyxl, which inherently
    compensates for the Lotus 1-2-3 leap-year bug for modern dates.
    No manual +1 correction needed.
    """
    delta = dt - _EXCEL_EPOCH
    return delta.days + delta.seconds / 86400.0


def _convert_date_cell(cell_el, shared_strings: List[str]) -> None:
    """Convert a string/inlineStr cell value to a serial date in-place."""
    t = cell_el.get("t", "")
    v_el = cell_el.find(_tag("v"))

    # Already numeric — assume serial date, nothing to convert
    if t in ("", "n"):
        if v_el is not None and v_el.text:
            try:
                val = float(v_el.text)
                if 1 <= val <= 2958465:
                    return
            except ValueError:
                pass

    # Skip formula cells — their cached value is already a serial number
    if cell_el.find(_tag("f")) is not None:
        return

    val_str = _get_cell_text(cell_el, shared_strings)
    if not val_str:
        return

    dt = _try_parse_date(val_str)
    if dt is None:
        return

    serial = _datetime_to_serial(dt)

    # Remove string type
    if "t" in cell_el.attrib:
        del cell_el.attrib["t"]

    # Remove inline string element
    is_el = cell_el.find(_tag("is"))
    if is_el is not None:
        cell_el.remove(is_el)

    # Set (or create) <v> with serial date
    if v_el is None:
        v_el = etree.SubElement(cell_el, _tag("v"))
    v_el.text = str(int(serial)) if serial == int(serial) else f"{serial:.10g}"


# ---------------------------------------------------------------------------
# Number-format detection
# ---------------------------------------------------------------------------


def _cell_has_decimal(cell_el) -> bool:
    """Does this numeric cell have meaningful decimal places (>0.001)?"""
    v_el = cell_el.find(_tag("v"))
    if v_el is None or not v_el.text:
        return False
    try:
        val = float(v_el.text)
        return abs(val - int(val)) > 0.001
    except ValueError:
        return False


# ===================================================================
# StyleContext — parse styles.xml, add entries, build remap table
# ===================================================================


class StyleContext:
    """Manage OOXML styles and build a (old_xf, fmt_type) → new_xf remap."""

    def __init__(self, styles_xml: bytes, date_format: str, separator_style: str):
        self._tree = etree.fromstring(styles_xml)

        self._fonts_el = self._tree.find(_tag("fonts"))
        self._borders_el = self._tree.find(_tag("borders"))
        self._fills_el = self._tree.find(_tag("fills"))
        self._cellxfs_el = self._tree.find(_tag("cellXfs"))

        # numFmts may not exist — create before <fonts> if missing
        self._numfmts_el = self._tree.find(_tag("numFmts"))
        if self._numfmts_el is None:
            idx = 0
            for i, child in enumerate(self._tree):
                if child.tag == _tag("fonts"):
                    idx = i
                    break
            self._numfmts_el = etree.Element(_tag("numFmts"), count="0")
            self._tree.insert(idx, self._numfmts_el)

        # Snapshot existing xf entries before we append new ones
        self._existing_xfs = list(self._cellxfs_el.findall(_tag("xf")))
        self._existing_count = len(self._existing_xfs)

        # Add our font / border entries (reuses existing if identical)
        self._normal_font_id = self._find_or_add_font(
            EXCEL_FONT_NAME, EXCEL_FONT_SIZE, bold=False,
        )
        self._bold_font_id = self._find_or_add_font(
            EXCEL_FONT_NAME, EXCEL_FONT_SIZE, bold=True,
        )
        self._border_id = self._find_or_add_thin_border()

        # Resolve number-format IDs
        self._date_nf = self._resolve_numfmt(date_format)
        if separator_style == ".":
            self._int_nf = self._resolve_numfmt(NUMBER_FORMAT_DOT_INTEGER)
            self._dec_nf = self._resolve_numfmt(NUMBER_FORMAT_DOT_DECIMAL)
        else:
            self._int_nf = self._resolve_numfmt(NUMBER_FORMAT_INTEGER)
            self._dec_nf = self._resolve_numfmt(NUMBER_FORMAT_DECIMAL)

        # Build remap: (old_xf_index, format_type) → new_xf_index
        self.remap: Dict[Tuple[int, str], int] = {}
        self._build_remap()

    # --- fonts ---

    def _find_or_add_font(self, name: str, size: int, bold: bool) -> int:
        for i, f in enumerate(self._fonts_el.findall(_tag("font"))):
            n_el = f.find(_tag("name"))
            s_el = f.find(_tag("sz"))
            b_el = f.find(_tag("b"))
            if (n_el is not None and n_el.get("val") == name
                    and s_el is not None and s_el.get("val") == str(size)
                    and (b_el is not None) == bold):
                return i

        el = etree.SubElement(self._fonts_el, _tag("font"))
        if bold:
            etree.SubElement(el, _tag("b"))
        etree.SubElement(el, _tag("sz")).set("val", str(size))
        etree.SubElement(el, _tag("name")).set("val", name)
        cnt = len(self._fonts_el.findall(_tag("font")))
        self._fonts_el.set("count", str(cnt))
        return cnt - 1

    # --- borders ---

    def _find_or_add_thin_border(self) -> int:
        sides = ("left", "right", "top", "bottom")
        for i, b in enumerate(self._borders_el.findall(_tag("border"))):
            if all(
                (s := b.find(_tag(side))) is not None and s.get("style") == "thin"
                for side in sides
            ):
                return i

        el = etree.SubElement(self._borders_el, _tag("border"))
        for side in sides:
            etree.SubElement(el, _tag(side)).set("style", "thin")
        etree.SubElement(el, _tag("diagonal"))
        cnt = len(self._borders_el.findall(_tag("border")))
        self._borders_el.set("count", str(cnt))
        return cnt - 1

    # --- numFmts ---

    def _resolve_numfmt(self, code: str) -> int:
        if code in _BUILTIN_NUMFMT:
            return _BUILTIN_NUMFMT[code]

        for nf in self._numfmts_el.findall(_tag("numFmt")):
            if nf.get("formatCode") == code:
                return int(nf.get("numFmtId", "164"))

        ids = [int(nf.get("numFmtId", "0"))
               for nf in self._numfmts_el.findall(_tag("numFmt"))]
        next_id = max(ids + [163]) + 1

        nf = etree.SubElement(self._numfmts_el, _tag("numFmt"))
        nf.set("numFmtId", str(next_id))
        nf.set("formatCode", code)
        self._numfmts_el.set(
            "count", str(len(self._numfmts_el.findall(_tag("numFmt")))),
        )
        return next_id

    # --- remap table ---

    _FMT_TYPES = ("header", "data", "date", "number_int", "number_dec")

    def _build_remap(self):
        next_idx = self._existing_count

        for old_idx, old_xf in enumerate(self._existing_xfs):
            old_fill = old_xf.get("fillId", "0")
            old_nf = old_xf.get("numFmtId", "0")
            old_align = old_xf.find(_tag("alignment"))
            old_prot = old_xf.find(_tag("protection"))

            for fmt in self._FMT_TYPES:
                font = self._bold_font_id if fmt == "header" else self._normal_font_id

                if fmt in ("header", "data"):
                    nf_id = old_nf  # preserve original numFmt
                elif fmt == "date":
                    nf_id = str(self._date_nf)
                elif fmt == "number_int":
                    nf_id = str(self._int_nf)
                else:
                    nf_id = str(self._dec_nf)

                xf = etree.SubElement(self._cellxfs_el, _tag("xf"))
                xf.set("numFmtId", str(nf_id))
                xf.set("fontId", str(font))
                xf.set("fillId", old_fill)
                xf.set("borderId", str(self._border_id))
                xf.set("xfId", "0")
                xf.set("applyFont", "1")
                xf.set("applyBorder", "1")

                if fmt in ("date", "number_int", "number_dec"):
                    xf.set("applyNumberFormat", "1")
                if old_fill != "0":
                    xf.set("applyFill", "1")
                if old_align is not None:
                    xf.append(copy.deepcopy(old_align))
                    xf.set("applyAlignment", "1")
                if old_prot is not None:
                    xf.append(copy.deepcopy(old_prot))
                    xf.set("applyProtection", "1")

                self.remap[(old_idx, fmt)] = next_idx
                next_idx += 1

        self._cellxfs_el.set("count", str(next_idx))

    def get_style(self, old_s: int, fmt_type: str) -> str:
        """Look up new xf index, with fallback to xf-0 variant."""
        s = self.remap.get((old_s, fmt_type))
        if s is None:
            s = self.remap.get((0, fmt_type), 0)
        return str(s)

    def build_fast_lookup(self) -> Dict[str, Dict[int, str]]:
        """Pre-compute {fmt_type: {old_s: new_s_str}} for hot-loop use."""
        lookup: Dict[str, Dict[int, str]] = {}
        for fmt in self._FMT_TYPES:
            d: Dict[int, str] = {}
            fallback = str(self.remap.get((0, fmt), 0))
            for i in range(self._existing_count):
                d[i] = str(self.remap.get((i, fmt), fallback))
            d[-1] = fallback  # sentinel for out-of-range old_s
            lookup[fmt] = d
        return lookup

    def serialize(self) -> bytes:
        return etree.tostring(
            self._tree, xml_declaration=True, encoding="UTF-8", standalone=True,
        )


# ===================================================================
# Sheet-level formatting
# ===================================================================


def _format_sheet(
    sheet_xml: bytes,
    style_ctx: StyleContext,
    config: SheetConfig,
    shared_strings: List[str],
    last_col: int,
    freeze_pane: bool,
    progress_callback=None,
) -> bytes:
    """Apply formatting to one sheet's XML.  Returns modified XML bytes.

    Optimised: in-place cell modification (no tree rebuild), inlined
    column-index parsing, pre-computed style dicts, width sampling.
    """
    tree = etree.fromstring(sheet_xml)
    sheet_data = tree.find(_tag("sheetData"))
    if sheet_data is None:
        return sheet_xml

    header_row = config.header_row

    # Pre-compute style lookups
    fast = style_ctx.build_fast_lookup()
    sty_header = fast["header"]
    sty_data = fast["data"]
    sty_date = fast["date"]
    sty_int = fast["number_int"]
    sty_dec = fast["number_dec"]
    existing_count = style_ctx._existing_count

    # Per-column style dict (eliminates per-cell type checks)
    date_cols: Set[int] = set()
    decimal_cols: Set[int] = set()  # number cols needing per-cell decimal check
    col_style: Dict[int, Dict[int, str]] = {}
    for ci in config.all_columns:
        if ci.user_format_type == "date":
            col_style[ci.index] = sty_date
            date_cols.add(ci.index)
        elif ci.user_format_type == "number":
            if ci.has_decimals:
                # Column has mixed int/decimal values — check per cell
                col_style[ci.index] = sty_int  # default to integer
                decimal_cols.add(ci.index)
            else:
                col_style[ci.index] = sty_int
        # text/other columns: use sty_data (default below)

    # Cache qualified tag strings
    C_TAG = _tag("c")
    V_TAG = _tag("v")
    F_TAG = _tag("f")
    IS_TAG = _tag("is")
    T_TAG = _tag("t")
    ROW_TAG = _tag("row")

    # Default styles for empty cells
    empty_header_s = sty_header.get(0, sty_header[-1])
    empty_data_s = sty_data.get(0, sty_data[-1])

    # Find last data row for progress reporting
    last_data_row = 0
    for row_el in sheet_data:
        if row_el.tag == ROW_TAG:
            rn = int(row_el.get("r", "0"))
            if rn > last_data_row:
                last_data_row = rn

    if last_data_row == 0:
        return sheet_xml

    total_rows = max(last_data_row - header_row + 1, 1)

    # Column widths — sampled from header + first N data rows
    WIDTH_SAMPLE_ROWS = 200
    col_widths: Dict[int, int] = {}
    min_w = EXCEL_MIN_COL_WIDTH
    width_rows_left = WIDTH_SAMPLE_ROWS

    processed = 0

    # ---- MAIN LOOP: iterate rows in-place (no tree rebuild) ----
    for row_el in sheet_data:
        if row_el.tag != ROW_TAG:
            continue

        row_num = int(row_el.get("r", "0"))
        if row_num < header_row:
            continue

        is_header = row_num == header_row
        sample_width = is_header or width_rows_left > 0
        cols_seen: Set[int] = set()

        for cell_el in row_el:
            if cell_el.tag != C_TAG:
                continue

            # Inline column-index extraction
            ref = cell_el.get("r", "")
            col = 0
            for ch in ref:
                if 'A' <= ch <= 'Z':
                    col = col * 26 + (ord(ch) - 64)
                elif 'a' <= ch <= 'z':
                    col = col * 26 + (ord(ch) - 96)
                else:
                    break
            cols_seen.add(col)

            old_s = int(cell_el.get("s", "0"))
            if old_s >= existing_count:
                old_s = 0

            # Style assignment
            if is_header:
                cell_el.set("s", sty_header.get(old_s, sty_header[-1]))
            else:
                if col in decimal_cols:
                    # Per-cell: integer cells get #,##0, decimal cells get #,##0.00
                    v_el = cell_el.find(V_TAG)
                    has_dec = False
                    if v_el is not None and v_el.text:
                        try:
                            _v = float(v_el.text)
                            has_dec = abs(_v - int(_v)) > 0.001
                        except ValueError:
                            pass
                    sd = sty_dec if has_dec else sty_int
                    cell_el.set("s", sd.get(old_s, sd[-1]))
                else:
                    sd = col_style.get(col, sty_data)
                    cell_el.set("s", sd.get(old_s, sd[-1]))
                    if col in date_cols:
                        _convert_date_cell(cell_el, shared_strings)

            # Width sampling
            if sample_width and col <= last_col:
                v_el = cell_el.find(V_TAG)
                txt = None
                if v_el is not None and v_el.text:
                    t_attr = cell_el.get("t", "")
                    if t_attr == "s":
                        try:
                            si = int(v_el.text)
                            if 0 <= si < len(shared_strings):
                                txt = shared_strings[si]
                        except (ValueError, IndexError):
                            pass
                    else:
                        txt = v_el.text
                elif cell_el.get("t") == "inlineStr":
                    is_el = cell_el.find(IS_TAG)
                    if is_el is not None:
                        t_el = is_el.find(T_TAG)
                        if t_el is not None:
                            txt = t_el.text
                if txt:
                    vlen = len(txt) + 2
                    if is_header:
                        vlen += 1
                    if vlen > col_widths.get(col, min_w):
                        col_widths[col] = vlen

        # Fill missing cells for consistent borders across the entire range
        needs_fill = False
        for ci in range(1, last_col + 1):
            if ci not in cols_seen:
                needs_fill = True
                break

        if needs_fill:
            es = empty_header_s if is_header else empty_data_s
            for ci in range(1, last_col + 1):
                if ci not in cols_seen:
                    c = etree.SubElement(row_el, C_TAG)
                    c.set("r", _cell_ref(ci, row_num))
                    c.set("s", es)
            # CRITICAL: cells must be sorted by column ref (OOXML spec)
            _sort_row_cells(row_el, C_TAG)

        if not is_header:
            width_rows_left -= 1
        processed += 1
        if processed % 1000 == 0 and progress_callback:
            progress_callback(processed / total_rows)

    # ---- END MAIN LOOP ----

    # Create entirely missing rows for complete border coverage
    existing_row_nums: Set[int] = set()
    for r in sheet_data:
        if r.tag == ROW_TAG:
            existing_row_nums.add(int(r.get("r", "0")))

    added_rows = False
    for rn in range(header_row, last_data_row + 1):
        if rn not in existing_row_nums:
            added_rows = True
            row_el = etree.SubElement(sheet_data, ROW_TAG)
            row_el.set("r", str(rn))
            es = empty_header_s if rn == header_row else empty_data_s
            for ci in range(1, last_col + 1):
                c = etree.SubElement(row_el, C_TAG)
                c.set("r", _cell_ref(ci, rn))
                c.set("s", es)

    # Sort rows in sheetData by row number (new rows were appended at end)
    if added_rows:
        rows_list = list(sheet_data)
        sheet_data[:] = []
        rows_list.sort(
            key=lambda r: int(r.get("r", "0")) if r.tag == ROW_TAG else 0,
        )
        for r in rows_list:
            sheet_data.append(r)

    _set_column_widths(tree, col_widths)

    if freeze_pane:
        _set_freeze_pane(tree, header_row)
    else:
        _remove_freeze_pane(tree)

    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------------------
# Column widths
# ---------------------------------------------------------------------------


def _set_column_widths(tree, col_widths: Dict[int, int]):
    if not col_widths:
        return

    max_w = EXCEL_MAX_COL_WIDTH
    sheet_data = tree.find(_tag("sheetData"))

    # Remove existing <cols> — we'll recreate with measured widths
    old_cols = tree.find(_tag("cols"))
    if old_cols is not None:
        tree.remove(old_cols)

    cols_el = etree.Element(_tag("cols"))
    for ci in sorted(col_widths):
        w = min(col_widths[ci], max_w)
        col_el = etree.SubElement(cols_el, _tag("col"))
        col_el.set("min", str(ci))
        col_el.set("max", str(ci))
        col_el.set("width", str(w))
        col_el.set("customWidth", "1")

    # Insert before <sheetData>
    sd_idx = list(tree).index(sheet_data)
    tree.insert(sd_idx, cols_el)


# ---------------------------------------------------------------------------
# Freeze pane
# ---------------------------------------------------------------------------


def _set_freeze_pane(tree, header_row: int):
    views = tree.find(_tag("sheetViews"))
    if views is None:
        views = etree.SubElement(tree, _tag("sheetViews"))
    view = views.find(_tag("sheetView"))
    if view is None:
        view = etree.SubElement(views, _tag("sheetView"))
        view.set("workbookViewId", "0")

    # Remove stale scroll position from the original file — if kept,
    # Excel may open the sheet scrolled far from the header.
    if "topLeftCell" in view.attrib:
        del view.attrib["topLeftCell"]

    # Remove existing pane
    pane = view.find(_tag("pane"))
    if pane is not None:
        view.remove(pane)

    # Remove ALL existing <selection> elements (there can be one per pane)
    for sel in view.findall(_tag("selection")):
        view.remove(sel)

    # Add freeze pane below the header row
    pane = etree.Element(_tag("pane"))
    pane.set("ySplit", str(header_row))
    pane.set("topLeftCell", f"A{header_row + 1}")
    pane.set("activePane", "bottomLeft")
    pane.set("state", "frozen")
    view.insert(0, pane)

    # Add clean selection for the scrollable pane
    sel = etree.SubElement(view, _tag("selection"))
    sel.set("pane", "bottomLeft")
    sel.set("activeCell", f"A{header_row + 1}")
    sel.set("sqref", f"A{header_row + 1}")


def _remove_freeze_pane(tree):
    views = tree.find(_tag("sheetViews"))
    if views is None:
        return
    view = views.find(_tag("sheetView"))
    if view is None:
        return
    pane = view.find(_tag("pane"))
    if pane is not None:
        view.remove(pane)


# ===================================================================
# Public entry point — format an entire workbook
# ===================================================================


def format_workbook(
    config: FileConfig,
    output_path: str,
    progress_callback: Optional[Callable[[str, float, str], None]] = None,
) -> None:
    """Format a workbook via direct XML manipulation.

    Args:
        config:  Fully-configured FileConfig (analysed + user-adjusted).
        output_path:  Full path for the formatted output file.
        progress_callback:  fn(file_name, progress_0_to_1, status_text).

    Raises on failure (caller should catch and fall back to openpyxl).
    """
    file_name = config.file_name

    def _report(pct: float, text: str):
        if progress_callback:
            progress_callback(file_name, pct, text)

    _report(0.0, "Loading workbook...")

    # 1. Read ZIP contents
    zip_contents: Dict[str, bytes] = {}
    with zipfile.ZipFile(config.file_path, "r") as zin:
        for name in zin.namelist():
            zip_contents[name] = zin.read(name)

    _report(0.05, "Preparing styles...")

    # 2. Shared strings
    shared_strings = _parse_shared_strings(
        zip_contents.get("xl/sharedStrings.xml"),
    )

    # 3. Sheet-name → ZIP-path mapping
    sheet_paths = _resolve_sheet_paths(zip_contents)

    # 4. Style context
    style_ctx = StyleContext(
        zip_contents["xl/styles.xml"],
        config.date_format,
        config.separator_style,
    )

    # 5. Format each selected sheet
    selected = [
        (name, sc)
        for name, sc in config.sheet_configs.items()
        if sc.selected
    ]
    total = len(selected) or 1

    for idx, (sheet_name, sc) in enumerate(selected):
        sheet_base = 0.05 + idx / total * 0.80
        sheet_span = 0.80 / total

        _report(sheet_base, f"Formatting: {sheet_name}")

        zip_path = sheet_paths.get(sheet_name)
        if not zip_path or zip_path not in zip_contents:
            continue

        last_col = max((ci.index for ci in sc.all_columns), default=1)

        def _sheet_progress(
            row_pct, _b=sheet_base, _s=sheet_span,
        ):
            _report(_b + row_pct * _s, f"Formatting: {sheet_name}")

        modified = _format_sheet(
            zip_contents[zip_path],
            style_ctx, sc, shared_strings,
            last_col, config.freeze_pane,
            _sheet_progress,
        )
        zip_contents[zip_path] = modified

    # 6. Write modified styles.xml
    zip_contents["xl/styles.xml"] = style_ctx.serialize()

    _report(0.85, "Saving...")

    # 7. Write output ZIP
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in zip_contents.items():
            zout.writestr(name, data)

    _report(1.0, "Done")
