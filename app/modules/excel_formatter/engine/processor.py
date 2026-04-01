"""Orchestrator: load workbook -> apply formatting -> save to output folder.

Uses fast XML-based engine (xml_formatter) by default for ~8-15x speedup.
Falls back to openpyxl cell-by-cell formatting if XML processing fails.
"""

import os
import traceback
from pathlib import Path
from typing import Callable, Optional

from app.modules.excel_formatter.models.file_config import FileConfig

# Check for lxml availability at import time
try:
    from app.modules.excel_formatter.engine.xml_formatter import format_workbook as _xml_format
    _HAS_XML_ENGINE = True
except ImportError:
    _HAS_XML_ENGINE = False


def process_file(
    config: FileConfig,
    output_folder: str,
    progress_callback: Optional[Callable[[str, float, str], None]] = None,
) -> bool:
    """Format a single file and save it to the output folder.

    Tries the fast XML engine first; falls back to openpyxl on failure.

    Args:
        config: Fully configured FileConfig (analyzed + user-adjusted).
        output_folder: Directory to write the formatted file.
        progress_callback: fn(file_name, progress_0_to_1, status_text)

    Returns:
        True on success, False on failure (error stored in config.error_message).
    """
    file_name = config.file_name
    # Preserve subfolder structure when files came from a nested folder
    out_dir = os.path.join(output_folder, config.relative_dir) if config.relative_dir else output_folder
    out_path = os.path.join(out_dir, file_name)

    def _report(pct: float, text: str):
        if progress_callback:
            progress_callback(file_name, pct, text)

    # Ensure output folder (including subfolder) exists
    Path(out_dir).mkdir(parents=True, exist_ok=True)

    # --- Fast path: XML engine ---
    if _HAS_XML_ENGINE:
        try:
            _xml_format(config, out_path, progress_callback)
            config.status = "Done"
            config.progress = 1.0
            return True
        except Exception as exc:
            # Log and fall through to openpyxl fallback
            traceback.print_exc()
            _report(0.0, "Retrying with fallback engine...")

    # --- Fallback: openpyxl cell-by-cell ---
    return _process_file_openpyxl(config, out_path, _report)


def _process_file_openpyxl(
    config: FileConfig,
    out_path: str,
    report: Callable[[float, str], None],
) -> bool:
    """Original openpyxl-based formatting (fallback for edge-case files)."""
    from openpyxl import load_workbook

    from app.modules.excel_formatter.engine.formatter import format_sheet

    try:
        report(0.0, "Loading workbook...")
        wb = load_workbook(config.file_path, data_only=False)

        selected_sheets = [
            (name, sc) for name, sc in config.sheet_configs.items() if sc.selected
        ]
        total = len(selected_sheets) or 1

        for idx, (sheet_name, sc) in enumerate(selected_sheets):
            sheet_base = idx / total
            sheet_span = 0.85 / total

            report(sheet_base, f"Formatting: {sheet_name}")

            def sheet_progress(row_pct, _base=sheet_base, _span=sheet_span):
                report(_base + row_pct * _span, f"Formatting: {sheet_name}")

            ws = wb[sheet_name]
            format_sheet(
                ws, sc, config.date_format, config.freeze_pane,
                separator_style=config.separator_style,
                progress_callback=sheet_progress,
            )

        report(0.85, "Saving...")
        wb.save(out_path)
        wb.close()

        config.status = "Done"
        config.progress = 1.0
        report(1.0, "Done")
        return True

    except Exception as exc:
        config.status = "Error"
        config.error_message = str(exc)
        config.progress = 0.0
        report(0.0, f"Error: {exc}")
        traceback.print_exc()
        return False
