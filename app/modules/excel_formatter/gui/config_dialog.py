"""Per-file configuration dialog with per-sheet column panels."""

import customtkinter as ctk
from tkinter import messagebox

from app.core import theme
from app.modules.excel_formatter.gui.column_config import SheetColumnPanel
from app.modules.excel_formatter.models.file_config import FileConfig
from app.utils.constants import DATE_FORMATS, DEFAULT_DATE_FORMAT_INDEX, SEPARATOR_OPTIONS

# Example date for preview
_EXAMPLE_PREVIEWS = {
    "DD-MMM-YY": "01-Aug-24",
    "DD/MM/YYYY": "01/08/2024",
    "DD-MM-YYYY": "01-08-2024",
    "YYYY-MM-DD": "2024-08-01",
    "DD MMMM YYYY": "01 August 2024",
    "MMM DD, YYYY": "Aug 01, 2024",
}


class ConfigDialog(ctk.CTkToplevel):
    """Modal dialog for detailed per-file configuration."""

    def __init__(self, parent, config: FileConfig, date_format: str = "DD-MMM-YY"):
        super().__init__(parent)
        self._config = config
        self._date_format = date_format
        self._sheet_panels: dict[str, SheetColumnPanel] = {}

        self.title(f"Configure: {config.file_name}")
        self.geometry("880x660")
        self.minsize(750, 400)
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()

        self.configure(fg_color=theme.LIGHT_GRAY)

        self._build_ui()

        # Centre on parent
        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        py = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")

    def _build_ui(self):
        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=theme.PADDING_LARGE, pady=theme.PADDING_LARGE)

        # --- Format settings card ---
        fmt_card = ctk.CTkFrame(
            scroll, fg_color=theme.WHITE, corner_radius=theme.CORNER_RADIUS,
            border_width=1, border_color=theme.BORDER_GRAY,
        )
        fmt_card.pack(fill="x", pady=(0, theme.PADDING_NORMAL))

        # Row 1: Date format
        fmt_row1 = ctk.CTkFrame(fmt_card, fg_color="transparent")
        fmt_row1.pack(fill="x", padx=theme.PADDING_LARGE, pady=(theme.PADDING_NORMAL, 4))

        ctk.CTkLabel(
            fmt_row1, text="Date Format:", anchor="w", width=110,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL, "bold"),
            text_color=theme.TEXT_PRIMARY,
        ).pack(side="left")

        self._date_var = ctk.StringVar(value=self._date_format)
        ctk.CTkOptionMenu(
            fmt_row1, values=[d[0] for d in DATE_FORMATS],
            variable=self._date_var, width=170,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            fg_color=theme.LIGHT_GRAY, button_color=theme.SLATE_GRAY,
            button_hover_color=theme.HOVER_BLUE,
            text_color=theme.TEXT_PRIMARY,
            command=self._on_date_format_changed,
        ).pack(side="left", padx=(4, 0))

        example = _EXAMPLE_PREVIEWS.get(self._date_format, "01-Aug-24")
        self._date_example_label = ctk.CTkLabel(
            fmt_row1, text=f"Example: {example}",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            text_color=theme.ACCENT_BLUE, anchor="w",
        )
        self._date_example_label.pack(side="left", padx=(theme.PADDING_LARGE, 0))

        # Row 2: Number separator format
        fmt_row2 = ctk.CTkFrame(fmt_card, fg_color="transparent")
        fmt_row2.pack(fill="x", padx=theme.PADDING_LARGE, pady=(0, theme.PADDING_NORMAL))

        ctk.CTkLabel(
            fmt_row2, text="Number Format:", anchor="w", width=110,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL, "bold"),
            text_color=theme.TEXT_PRIMARY,
        ).pack(side="left")

        # Map current separator_style to display label
        current_sep = self._config.separator_style
        sep_labels = [s[0] for s in SEPARATOR_OPTIONS]
        current_label = next((s[0] for s in SEPARATOR_OPTIONS if s[1] == current_sep), sep_labels[0])

        self._sep_var = ctk.StringVar(value=current_label)
        ctk.CTkOptionMenu(
            fmt_row2, values=sep_labels,
            variable=self._sep_var, width=170,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            fg_color=theme.LIGHT_GRAY, button_color=theme.SLATE_GRAY,
            button_hover_color=theme.HOVER_BLUE,
            text_color=theme.TEXT_PRIMARY,
            command=self._on_separator_changed,
        ).pack(side="left", padx=(4, 0))

        sep_example = "1,500,000.50" if current_sep == "," else "1.500.000,50"
        self._sep_example_label = ctk.CTkLabel(
            fmt_row2, text=f"Example: {sep_example}",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            text_color=theme.ACCENT_BLUE, anchor="w",
        )
        self._sep_example_label.pack(side="left", padx=(theme.PADDING_LARGE, 0))

        # --- One SheetColumnPanel per sheet ---
        for sname, sc in self._config.sheet_configs.items():
            panel = SheetColumnPanel(
                scroll,
                sheet_config=sc,
                date_format=self._date_format,
                separator_style=current_sep,
            )
            panel.pack(fill="x", pady=(0, theme.PADDING_NORMAL))
            self._sheet_panels[sname] = panel

        # --- Buttons ---
        btn_row = ctk.CTkFrame(self, fg_color="transparent", height=50)
        btn_row.pack(fill="x", padx=theme.PADDING_LARGE, pady=(0, theme.PADDING_LARGE))

        ctk.CTkButton(
            btn_row, text="Apply",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL, "bold"),
            fg_color=theme.ACCENT_BLUE, hover_color=theme.HOVER_BLUE,
            text_color=theme.WHITE, height=theme.BUTTON_HEIGHT,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            command=self._apply,
        ).pack(side="left", padx=(0, theme.PADDING_NORMAL))

        ctk.CTkButton(
            btn_row, text="Cancel",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            fg_color=theme.LIGHT_GRAY, hover_color=theme.BORDER_GRAY,
            text_color=theme.TEXT_PRIMARY, height=theme.BUTTON_HEIGHT,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            border_width=1, border_color=theme.BORDER_GRAY,
            command=self.destroy,
        ).pack(side="left")

    # ------------------------------------------------------------------
    # Format change handlers
    # ------------------------------------------------------------------

    def _on_date_format_changed(self, choice: str):
        self._date_format = choice
        example = _EXAMPLE_PREVIEWS.get(choice, choice)
        self._date_example_label.configure(text=f"Example: {example}")
        for panel in self._sheet_panels.values():
            panel.update_date_format(choice)

    def _on_separator_changed(self, choice: str):
        sep_style = next((s[1] for s in SEPARATOR_OPTIONS if s[0] == choice), ",")
        sep_example = "1,500,000.50" if sep_style == "," else "1.500.000,50"
        self._sep_example_label.configure(text=f"Example: {sep_example}")
        for panel in self._sheet_panels.values():
            panel.update_separator_style(sep_style)

    # ------------------------------------------------------------------
    # Apply
    # ------------------------------------------------------------------

    def _apply(self):
        """Write user selections back into the FileConfig."""
        # Validate header rows
        for sname, panel in self._sheet_panels.items():
            valid, err = panel.validate_header_row()
            if not valid:
                messagebox.showwarning("Invalid Input", err)
                return

        # Store date format
        selected_label = self._date_var.get()
        for display, fmt in DATE_FORMATS:
            if display == selected_label:
                self._config.date_format = fmt
                break

        # Store separator style
        sep_label = self._sep_var.get()
        self._config.separator_style = next(
            (s[1] for s in SEPARATOR_OPTIONS if s[0] == sep_label), ","
        )

        # Apply column selections per sheet
        for sname, panel in self._sheet_panels.items():
            sc = self._config.sheet_configs[sname]
            sheet_selected, header_override, col_formats = panel.get_selections()

            sc.selected = sheet_selected

            if header_override is not None:
                sc.header_row = header_override
                sc.header_auto_detected = False

            for ci in sc.all_columns:
                if ci.index in col_formats:
                    ci.user_format_type = col_formats[ci.index]
                    ci.user_selected = col_formats[ci.index] is not None

        self.destroy()
