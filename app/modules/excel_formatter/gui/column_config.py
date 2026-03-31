"""Per-sheet column configuration panel with checkboxes and live preview."""

from datetime import datetime, timedelta
from typing import Optional

import customtkinter as ctk

from app.core import theme
from app.modules.excel_formatter.models.column_info import ColumnInfo
from app.modules.excel_formatter.models.file_config import SheetConfig

# Python strftime equivalents for preview rendering
_FORMAT_STRFTIME = {
    "DD-MMM-YY": "%d-%b-%y",
    "DD/MM/YYYY": "%d/%m/%Y",
    "DD-MM-YYYY": "%d-%m-%Y",
    "YYYY-MM-DD": "%Y-%m-%d",
    "DD MMMM YYYY": "%d %B %Y",
    "MMM DD, YYYY": "%b %d, %Y",
}

# Common input date parse formats
_PARSE_FMTS = [
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


class SheetColumnPanel(ctk.CTkFrame):
    """Per-sheet container showing ALL columns with checkboxes for formatting."""

    def __init__(
        self,
        master,
        sheet_config: SheetConfig,
        date_format: str,
        separator_style: str = ",",
        **kwargs,
    ):
        super().__init__(
            master, fg_color=theme.WHITE, corner_radius=theme.CORNER_RADIUS,
            border_width=1, border_color=theme.BORDER_GRAY, **kwargs,
        )
        self._sheet_config = sheet_config
        self._date_format = date_format
        self._separator_style = separator_style  # "," or "."
        self._separator_vars: dict[int, ctk.BooleanVar] = {}
        self._date_vars: dict[int, ctk.BooleanVar] = {}
        self._separator_cbs: dict[int, ctk.CTkCheckBox] = {}
        self._date_cbs: dict[int, ctk.CTkCheckBox] = {}
        self._preview_labels: dict[int, ctk.CTkLabel] = {}

        self._build_ui()

    def _build_ui(self):
        # --- Header row: sheet checkbox + header row info + override ---
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=theme.PADDING_NORMAL, pady=(theme.PADDING_NORMAL, 0))

        self._sheet_var = ctk.BooleanVar(value=self._sheet_config.selected)
        ctk.CTkCheckBox(
            header_frame,
            text=self._sheet_config.name,
            variable=self._sheet_var,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL, "bold"),
            text_color=theme.TEXT_PRIMARY,
            fg_color=theme.ACCENT_BLUE, hover_color=theme.HOVER_BLUE,
            border_color=theme.BORDER_GRAY,
            command=self._on_sheet_toggled,
        ).pack(side="left")

        ctk.CTkLabel(
            header_frame,
            text=f"  Header Row: Auto-detected Row {self._sheet_config.header_row}",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=theme.TEXT_MUTED,
        ).pack(side="left", padx=(theme.PADDING_LARGE, 0))

        ctk.CTkLabel(
            header_frame, text="Override:",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=theme.TEXT_SECONDARY,
        ).pack(side="left", padx=(theme.PADDING_NORMAL, 0))

        self._header_entry = ctk.CTkEntry(
            header_frame, width=50,
            placeholder_text=str(self._sheet_config.header_row),
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            fg_color=theme.LIGHT_GRAY, border_color=theme.BORDER_GRAY,
            text_color=theme.TEXT_PRIMARY,
        )
        self._header_entry.pack(side="left", padx=(4, 0))

        # --- Separator ---
        ctk.CTkFrame(self, fg_color=theme.BORDER_GRAY, height=1).pack(
            fill="x", padx=theme.PADDING_NORMAL, pady=(theme.PADDING_NORMAL, 4),
        )

        # --- Column header row ---
        col_header = ctk.CTkFrame(self, fg_color="transparent")
        col_header.pack(fill="x", padx=theme.PADDING_NORMAL, pady=(0, 2))

        ctk.CTkLabel(col_header, text="Column", width=55, anchor="w",
                     font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
                     text_color=theme.TEXT_MUTED).pack(side="left")
        ctk.CTkLabel(col_header, text="Header", width=140, anchor="w",
                     font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
                     text_color=theme.TEXT_MUTED).pack(side="left", padx=(4, 0))
        ctk.CTkLabel(col_header, text="1000 Sep.", width=65, anchor="center",
                     font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
                     text_color=theme.TEXT_MUTED).pack(side="left", padx=(4, 0))
        ctk.CTkLabel(col_header, text="Date Fmt", width=65, anchor="center",
                     font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
                     text_color=theme.TEXT_MUTED).pack(side="left", padx=(4, 0))
        ctk.CTkLabel(col_header, text="Sample Data", width=170, anchor="w",
                     font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
                     text_color=theme.TEXT_MUTED).pack(side="left", padx=(8, 0))
        ctk.CTkLabel(col_header, text="Preview", anchor="w",
                     font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
                     text_color=theme.TEXT_MUTED).pack(side="left", padx=(4, 0), fill="x", expand=True)

        # --- Column rows ---
        for ci in self._sheet_config.all_columns:
            self._build_column_row(ci)

        ctk.CTkFrame(self, fg_color="transparent", height=theme.PADDING_NORMAL).pack()

        # Apply initial sheet enabled/disabled state
        if not self._sheet_config.selected:
            self._set_columns_enabled(False)

    def _build_column_row(self, ci: ColumnInfo):
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", padx=theme.PADDING_NORMAL, pady=1)

        # Column letter
        ctk.CTkLabel(
            row, text=f"Col {ci.letter}", width=55, anchor="w",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            text_color=theme.TEXT_SECONDARY,
        ).pack(side="left")

        # Header name
        display_name = ci.header_name
        if len(display_name) > 20:
            display_name = display_name[:18] + ".."
        ctk.CTkLabel(
            row, text=f'"{display_name}"', width=140, anchor="w",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            text_color=theme.TEXT_PRIMARY,
        ).pack(side="left", padx=(4, 0))

        # Separator checkbox — uses its own callback
        sep_var = ctk.BooleanVar(value=(ci.user_format_type == "number"))
        sep_cb = ctk.CTkCheckBox(
            row, text="", variable=sep_var, width=65,
            fg_color=theme.ACCENT_BLUE, hover_color=theme.HOVER_BLUE,
            border_color=theme.BORDER_GRAY, checkbox_width=18, checkbox_height=18,
            command=lambda idx=ci.index: self._on_separator_toggled(idx),
        )
        sep_cb.pack(side="left", padx=(4, 0))
        self._separator_vars[ci.index] = sep_var
        self._separator_cbs[ci.index] = sep_cb

        # Date checkbox — uses its own callback
        date_var = ctk.BooleanVar(value=(ci.user_format_type == "date"))
        date_cb = ctk.CTkCheckBox(
            row, text="", variable=date_var, width=65,
            fg_color=theme.ACCENT_BLUE, hover_color=theme.HOVER_BLUE,
            border_color=theme.BORDER_GRAY, checkbox_width=18, checkbox_height=18,
            command=lambda idx=ci.index: self._on_date_toggled(idx),
        )
        date_cb.pack(side="left", padx=(4, 0))
        self._date_vars[ci.index] = date_var
        self._date_cbs[ci.index] = date_cb

        # Sample values
        samples = ", ".join(ci.sample_values[:3]) if ci.sample_values else "-"
        if len(samples) > 26:
            samples = samples[:24] + ".."
        ctk.CTkLabel(
            row, text=samples, width=170, anchor="w",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=theme.TEXT_MUTED,
        ).pack(side="left", padx=(8, 0))

        # Preview label (dynamic)
        preview_text, is_warning = self._generate_preview(ci, ci.user_format_type)
        preview_lbl = ctk.CTkLabel(
            row, text=preview_text, anchor="w",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=theme.TEXT_MUTED if is_warning else theme.ACCENT_BLUE,
        )
        preview_lbl.pack(side="left", padx=(4, 0), fill="x", expand=True)
        self._preview_labels[ci.index] = preview_lbl

    # ------------------------------------------------------------------
    # Sheet toggle — enable/disable all column checkboxes
    # ------------------------------------------------------------------

    def _on_sheet_toggled(self):
        self._set_columns_enabled(self._sheet_var.get())

    def _set_columns_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        fg = theme.ACCENT_BLUE if enabled else theme.DISABLED_TEXT
        border = theme.BORDER_GRAY if enabled else theme.DISABLED_TEXT
        self._header_entry.configure(state=state)
        for cb in self._separator_cbs.values():
            cb.configure(state=state, fg_color=fg, border_color=border)
        for cb in self._date_cbs.values():
            cb.configure(state=state, fg_color=fg, border_color=border)
        for lbl in self._preview_labels.values():
            lbl.configure(
                text_color=theme.ACCENT_BLUE if enabled else theme.DISABLED_TEXT
            )

    # ------------------------------------------------------------------
    # Checkbox callbacks — separate methods for proper mutual exclusivity
    # ------------------------------------------------------------------

    def _on_separator_toggled(self, col_index: int):
        """Separator checkbox was toggled. If checked, uncheck date."""
        if self._separator_vars[col_index].get():
            self._date_vars[col_index].set(False)
        self._update_preview(col_index)

    def _on_date_toggled(self, col_index: int):
        """Date checkbox was toggled. If checked, uncheck separator."""
        if self._date_vars[col_index].get():
            self._separator_vars[col_index].set(False)
        self._update_preview(col_index)

    def _update_preview(self, col_index: int):
        """Recalculate preview text and color after a checkbox change."""
        sep_on = self._separator_vars[col_index].get()
        date_on = self._date_vars[col_index].get()
        fmt_type = "number" if sep_on else ("date" if date_on else None)

        ci = next((c for c in self._sheet_config.all_columns if c.index == col_index), None)
        if ci is None:
            return
        preview, is_warning = self._generate_preview(ci, fmt_type)
        self._preview_labels[col_index].configure(
            text=preview,
            text_color=theme.TEXT_MUTED if is_warning else theme.ACCENT_BLUE,
        )

    # ------------------------------------------------------------------
    # Preview generation
    # ------------------------------------------------------------------

    def _generate_preview(self, ci: ColumnInfo, format_type: Optional[str]) -> tuple[str, bool]:
        """Return (preview_text, is_warning). is_warning=True for unsupported data."""
        if format_type is None or not ci.sample_values:
            return ("", False)

        if format_type == "number":
            result = self._preview_number(ci.sample_values)
            if not result:
                return ("(not numeric)", True)
            return (result, False)

        if format_type == "date":
            result = self._preview_date(ci.sample_values)
            if "(date)" in result:
                return (result, True)
            return (result, False)

        return ("", False)

    def _preview_number(self, samples: list[str]) -> str:
        """Format sample values with thousand separator. Returns '' if none parseable."""
        parts = []
        sep = self._separator_style
        for s in samples[:2]:
            try:
                raw = str(s).replace(",", "")
                num = float(raw)
                frac = abs(num - int(num))
                if sep == ".":
                    # Dot as thousand separator, comma as decimal
                    if frac > 0.001:
                        parts.append(f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    else:
                        parts.append(f"{num:,.0f}".replace(",", "."))
                else:
                    # Comma as thousand separator, dot as decimal (default)
                    if frac > 0.001:
                        parts.append(f"{num:,.2f}")
                    else:
                        parts.append(f"{num:,.0f}")
            except (ValueError, TypeError):
                continue
        if not parts:
            return ""
        return "->" + ", ".join(parts)

    def _preview_date(self, samples: list[str]) -> str:
        target_fmt = _FORMAT_STRFTIME.get(self._date_format, "%d-%b-%y")

        for s in samples[:1]:
            s = str(s).strip()
            # Try Excel serial date (numeric)
            try:
                serial = float(s)
                if 1 <= serial <= 60000:
                    base = datetime(1899, 12, 30)
                    dt = base + timedelta(days=int(serial))
                    return f"->{dt.strftime(target_fmt)}"
            except (ValueError, TypeError):
                pass

            # Try string date parsing
            for pfmt in _PARSE_FMTS:
                try:
                    dt = datetime.strptime(s, pfmt)
                    return f"->{dt.strftime(target_fmt)}"
                except ValueError:
                    continue

        return "->(date)"

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_selections(self) -> tuple:
        """Return (sheet_selected, header_row_override_or_None, {col_index: user_format_type})."""
        sheet_selected = self._sheet_var.get()

        header_override = None
        txt = self._header_entry.get().strip()
        if txt:
            try:
                val = int(txt)
                if val >= 1:
                    header_override = val
            except ValueError:
                pass

        col_formats = {}
        for col_idx in self._separator_vars:
            sep_on = self._separator_vars[col_idx].get()
            date_on = self._date_vars[col_idx].get()
            if sep_on:
                col_formats[col_idx] = "number"
            elif date_on:
                col_formats[col_idx] = "date"
            else:
                col_formats[col_idx] = None

        return sheet_selected, header_override, col_formats

    def validate_header_row(self) -> tuple[bool, str]:
        """Return (is_valid, error_message)."""
        txt = self._header_entry.get().strip()
        if not txt:
            return True, ""
        try:
            val = int(txt)
            if val < 1:
                return False, f"Header row for '{self._sheet_config.name}' must be a positive integer."
            return True, ""
        except ValueError:
            return False, f"Header row for '{self._sheet_config.name}' must be a number."

    def update_date_format(self, date_format: str):
        """Refresh all date previews when the date format changes."""
        self._date_format = date_format
        for ci in self._sheet_config.all_columns:
            if self._date_vars.get(ci.index) and self._date_vars[ci.index].get():
                preview, is_warn = self._generate_preview(ci, "date")
                self._preview_labels[ci.index].configure(
                    text=preview,
                    text_color=theme.TEXT_MUTED if is_warn else theme.ACCENT_BLUE,
                )

    def update_separator_style(self, separator_style: str):
        """Refresh all number previews when the separator style changes."""
        self._separator_style = separator_style
        for ci in self._sheet_config.all_columns:
            if self._separator_vars.get(ci.index) and self._separator_vars[ci.index].get():
                preview, is_warn = self._generate_preview(ci, "number")
                self._preview_labels[ci.index].configure(
                    text=preview,
                    text_color=theme.TEXT_MUTED if is_warn else theme.ACCENT_BLUE,
                )
