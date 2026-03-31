"""Orchestrator: load workbook -> apply formatting -> save to output folder."""

import os
import traceback
from pathlib import Path
from typing import Callable, Optional

from openpyxl import load_workbook

from app.modules.excel_formatter.engine.formatter import format_sheet
from app.modules.excel_formatter.models.file_config import FileConfig


def process_file(
    config: FileConfig,
    output_folder: str,
    progress_callback: Optional[Callable[[str, float, str], None]] = None,
) -> bool:
    """Format a single file and save it to the output folder.

    Args:
        config: Fully configured FileConfig (analyzed + user-adjusted).
        output_folder: Directory to write the formatted file.
        progress_callback: fn(file_name, progress_0_to_1, status_text)

    Returns:
        True on success, False on failure (error stored in config.error_message).
    """
    file_path = config.file_path
    file_name = config.file_name

    def _report(pct: float, text: str):
        if progress_callback:
            progress_callback(file_name, pct, text)

    try:
        _report(0.0, "Loading workbook...")
        wb = load_workbook(file_path, data_only=False)

        selected_sheets = [
            (name, sc) for name, sc in config.sheet_configs.items() if sc.selected
        ]
        total = len(selected_sheets) or 1

        for idx, (sheet_name, sc) in enumerate(selected_sheets):
            sheet_base = idx / total
            sheet_span = 0.85 / total  # Leave 15% for saving

            _report(sheet_base, f"Formatting: {sheet_name}")

            # Per-row progress within this sheet
            def sheet_progress(row_pct, _base=sheet_base, _span=sheet_span):
                _report(_base + row_pct * _span, f"Formatting: {sheet_name}")

            ws = wb[sheet_name]
            format_sheet(
                ws, sc, config.date_format, config.freeze_pane,
                separator_style=config.separator_style,
                progress_callback=sheet_progress,
            )

        # Ensure output folder exists
        Path(output_folder).mkdir(parents=True, exist_ok=True)
        out_path = os.path.join(output_folder, file_name)

        _report(0.85, "Saving...")
        wb.save(out_path)
        wb.close()

        config.status = "Done"
        config.progress = 1.0
        _report(1.0, "Done")
        return True

    except Exception as exc:
        config.status = "Error"
        config.error_message = str(exc)
        config.progress = 0.0
        _report(0.0, f"Error: {exc}")
        traceback.print_exc()
        return False
