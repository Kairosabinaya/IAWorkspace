"""Panel showing the list of dropped/added Excel files."""

import customtkinter as ctk

from app.core import theme
from app.modules.excel_formatter.models.file_config import FileConfig


class FileListPanel(ctk.CTkFrame):
    """Scrollable list of files with status, size, sheet count, and action buttons."""

    def __init__(self, master, on_configure_click, on_remove_click,
                 on_format_click=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._on_configure = on_configure_click
        self._on_remove = on_remove_click
        self._on_format = on_format_click
        self._file_rows: dict[str, ctk.CTkFrame] = {}

        # Header
        ctk.CTkLabel(
            self, text="Files", font=(theme.FONT_FAMILY, theme.FONT_SIZE_LARGE, "bold"),
            text_color=theme.TEXT_PRIMARY, anchor="w",
        ).pack(fill="x", padx=theme.PADDING_NORMAL, pady=(theme.PADDING_NORMAL, 4))

        # Scrollable container
        self._scroll = ctk.CTkScrollableFrame(
            self, fg_color=theme.WHITE, corner_radius=theme.CORNER_RADIUS,
            border_width=1, border_color=theme.BORDER_GRAY,
        )
        self._scroll.pack(fill="both", expand=True, padx=theme.PADDING_NORMAL,
                          pady=(0, theme.PADDING_NORMAL))

        # Empty state label
        self._empty_label = ctk.CTkLabel(
            self._scroll,
            text="No files added yet.\nDrag & drop Excel files above, or click Browse.",
            text_color=theme.TEXT_MUTED, font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
        )
        self._empty_label.pack(pady=theme.PADDING_XLARGE)

    def add_file(self, config: FileConfig):
        """Add a file row to the list."""
        if config.file_path in self._file_rows:
            return
        self._empty_label.pack_forget()

        row = ctk.CTkFrame(self._scroll, fg_color=theme.LIGHT_GRAY, corner_radius=6)
        row.pack(fill="x", pady=2, padx=2)

        # File info
        info_frame = ctk.CTkFrame(row, fg_color="transparent")
        info_frame.pack(side="left", fill="x", expand=True,
                        padx=theme.PADDING_NORMAL, pady=6)

        ctk.CTkLabel(
            info_frame, text=config.file_name,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL, "bold"),
            text_color=theme.TEXT_PRIMARY, anchor="w",
        ).pack(fill="x")

        sheet_count = len(config.sheet_configs)
        detail_text = f"{sheet_count} sheet{'s' if sheet_count != 1 else ''}  |  {config.file_size}"
        detail_lbl = ctk.CTkLabel(
            info_frame, text=detail_text,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=theme.TEXT_SECONDARY, anchor="w",
        )
        detail_lbl.pack(fill="x")

        col_summary_lbl = ctk.CTkLabel(
            info_frame, text="",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=theme.ACCENT_BLUE, anchor="w",
        )
        col_summary_lbl.pack(fill="x")

        # Status label
        status_lbl = ctk.CTkLabel(
            row, text=config.status, width=90,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=self._status_color(config.status),
        )
        status_lbl.pack(side="left", padx=4)

        # Per-file Format button
        fmt_btn = ctk.CTkButton(
            row, text="Format", width=65, height=28,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            fg_color=theme.ACCENT_BLUE, hover_color=theme.HOVER_BLUE,
            text_color=theme.WHITE,
            corner_radius=4,
            command=lambda fp=config.file_path: (
                self._on_format(fp) if self._on_format else None
            ),
        )
        fmt_btn.pack(side="left", padx=2)

        # Configure button
        cfg_btn = ctk.CTkButton(
            row, text="Configure", width=75, height=28,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            fg_color=theme.LIGHT_GRAY, hover_color=theme.BORDER_GRAY,
            text_color=theme.TEXT_SECONDARY,
            border_width=1, border_color=theme.BORDER_GRAY,
            command=lambda fp=config.file_path: self._on_configure(fp),
        )
        cfg_btn.pack(side="left", padx=2)

        # Remove button
        remove_btn = ctk.CTkButton(
            row, text="X", width=32, height=28,
            font=(theme.FONT_FAMILY, 16),
            fg_color="transparent", hover_color=theme.ERROR_RED,
            text_color=theme.TEXT_SECONDARY,
            command=lambda fp=config.file_path: self._on_remove(fp),
        )
        remove_btn.pack(side="left", padx=(0, 6))

        # Store references
        row._status_label = status_lbl
        row._detail_label = detail_lbl
        row._col_summary_label = col_summary_lbl
        row._format_btn = fmt_btn
        row._config_btn = cfg_btn
        row._remove_btn = remove_btn
        self._file_rows[config.file_path] = row

    def update_file_status(self, file_path: str, status: str):
        row = self._file_rows.get(file_path)
        if row:
            row._status_label.configure(
                text=status, text_color=self._status_color(status),
            )

    def update_file_details(self, config: FileConfig):
        row = self._file_rows.get(config.file_path)
        if not row:
            return
        sheet_count = len(config.sheet_configs)
        detail_text = f"{sheet_count} sheet{'s' if sheet_count != 1 else ''}  |  {config.file_size}"
        row._detail_label.configure(text=detail_text)

        summary = self._build_column_summary(config)
        if summary:
            row._col_summary_label.configure(text=summary)

    def set_buttons_enabled(self, enabled: bool):
        """Enable or disable Format/Configure buttons on all rows."""
        state = "normal" if enabled else "disabled"
        for row in self._file_rows.values():
            row._format_btn.configure(state=state)
            row._config_btn.configure(state=state)

    def set_file_buttons_state(self, file_path: str, state: str):
        """Set button states for a specific file row.

        state: "ready" | "queued" | "processing" | "done" | "error" | "analyzing"
        """
        row = self._file_rows.get(file_path)
        if not row:
            return

        if state in ("ready", "done", "error"):
            row._format_btn.configure(state="normal", text="Format")
            row._config_btn.configure(state="normal")
            row._remove_btn.configure(state="normal")
        elif state == "queued":
            row._format_btn.configure(state="disabled", text="Queued")
            row._config_btn.configure(state="disabled")
            row._remove_btn.configure(state="normal")  # Can cancel from queue
        elif state == "processing":
            row._format_btn.configure(state="disabled", text="Formatting...")
            row._config_btn.configure(state="disabled")
            row._remove_btn.configure(state="disabled")
        elif state == "analyzing":
            row._format_btn.configure(state="disabled", text="Format")
            row._config_btn.configure(state="disabled")
            row._remove_btn.configure(state="normal")

    @staticmethod
    def _build_column_summary(config: FileConfig) -> str:
        date_letters: set[str] = set()
        number_letters: set[str] = set()
        for sc in config.sheet_configs.values():
            for ci in sc.all_columns:
                if ci.user_format_type == "date":
                    date_letters.add(ci.letter)
                elif ci.user_format_type == "number":
                    number_letters.add(ci.letter)
        parts = []
        if date_letters:
            parts.append(f"Dates: {', '.join(sorted(date_letters))}")
        if number_letters:
            parts.append(f"Numbers: {', '.join(sorted(number_letters))}")
        return "  |  ".join(parts) if parts else ""

    def remove_file(self, file_path: str):
        row = self._file_rows.pop(file_path, None)
        if row:
            row.destroy()
        if not self._file_rows:
            self._empty_label.pack(pady=theme.PADDING_XLARGE)

    def clear_all(self):
        for row in self._file_rows.values():
            row.destroy()
        self._file_rows.clear()
        self._empty_label.pack(pady=theme.PADDING_XLARGE)

    @staticmethod
    def _status_color(status: str) -> str:
        s = status.lower()
        if "done" in s:
            return theme.SUCCESS_GREEN
        elif "error" in s:
            return theme.ERROR_RED
        elif "processing" in s or "analyzing" in s or "formatting" in s:
            return theme.ACCENT_BLUE
        return theme.TEXT_SECONDARY
