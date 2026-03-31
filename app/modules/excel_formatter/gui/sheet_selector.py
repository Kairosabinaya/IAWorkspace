"""Widget for selecting which sheets to format."""

import customtkinter as ctk

from app.core import theme
from app.modules.excel_formatter.models.file_config import SheetConfig


class SheetSelector(ctk.CTkFrame):
    """Checklist of sheet names with select-all toggle."""

    def __init__(self, master, sheet_configs: dict[str, SheetConfig], **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._vars: dict[str, ctk.BooleanVar] = {}

        label = ctk.CTkLabel(
            self, text="Sheets:",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL, "bold"),
            text_color=theme.TEXT_PRIMARY, anchor="w",
        )
        label.pack(fill="x", pady=(0, 4))

        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x")

        for name, sc in sheet_configs.items():
            var = ctk.BooleanVar(value=sc.selected)
            cb = ctk.CTkCheckBox(
                row, text=name, variable=var,
                font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
                text_color=theme.TEXT_PRIMARY,
                fg_color=theme.ACCENT_BLUE, hover_color=theme.HOVER_BLUE,
                border_color=theme.BORDER_GRAY,
            )
            cb.pack(side="left", padx=(0, theme.PADDING_LARGE))
            self._vars[name] = var

    def get_selection(self) -> dict[str, bool]:
        """Return {sheet_name: selected}."""
        return {name: var.get() for name, var in self._vars.items()}
