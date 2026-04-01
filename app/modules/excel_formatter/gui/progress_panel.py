"""Progress panel showing per-file and overall progress bars."""

import customtkinter as ctk

from app.core import theme


class ProgressPanel(ctk.CTkFrame):
    """Displays progress bars for each file and an overall bar.

    Supports both batch initialisation via ``show()`` and incremental
    additions via ``add_file()`` so that new files can be enqueued while
    formatting is already running.
    """

    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._bars: dict[str, dict] = {}
        self._overall_bar = None
        self._overall_label = None
        self._overall_pct = None
        self._container = None
        self._file_container = None
        self._visible = False

    # ------------------------------------------------------------------
    # Initialisation helpers
    # ------------------------------------------------------------------

    def _ensure_visible(self):
        """Lazily create the panel skeleton if it isn't shown yet."""
        if self._visible:
            return

        self._visible = True

        header = ctk.CTkLabel(
            self, text="Progress",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_LARGE, "bold"),
            text_color=theme.TEXT_PRIMARY, anchor="w",
        )
        header.pack(fill="x", padx=theme.PADDING_NORMAL,
                    pady=(theme.PADDING_NORMAL, 4))

        self._container = ctk.CTkFrame(
            self, fg_color=theme.WHITE, corner_radius=theme.CORNER_RADIUS,
            border_width=1, border_color=theme.BORDER_GRAY,
        )
        self._container.pack(fill="x", padx=theme.PADDING_NORMAL, pady=(0, 4))

        # Inner frame that holds individual file rows
        self._file_container = ctk.CTkFrame(self._container,
                                            fg_color="transparent")
        self._file_container.pack(fill="x")

        # Separator + overall row
        sep = ctk.CTkFrame(self._container, fg_color=theme.BORDER_GRAY,
                           height=1)
        sep.pack(fill="x", padx=theme.PADDING_NORMAL, pady=4)

        overall_row = ctk.CTkFrame(self._container, fg_color="transparent")
        overall_row.pack(fill="x", padx=theme.PADDING_NORMAL,
                         pady=(0, theme.PADDING_NORMAL))

        self._overall_label = ctk.CTkLabel(
            overall_row, text="Overall:", width=200, anchor="w",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL, "bold"),
            text_color=theme.TEXT_PRIMARY,
        )
        self._overall_label.pack(side="left")

        self._overall_bar = ctk.CTkProgressBar(
            overall_row, width=300, height=14,
            progress_color=theme.ACCENT_BLUE,
            fg_color=theme.BORDER_GRAY, corner_radius=4,
        )
        self._overall_bar.pack(side="left", fill="x", expand=True, padx=8)
        self._overall_bar.set(0)

        self._overall_pct = ctk.CTkLabel(
            overall_row, text="0%", width=80, anchor="e",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL, "bold"),
            text_color=theme.TEXT_PRIMARY,
        )
        self._overall_pct.pack(side="left")

    @staticmethod
    def _truncate(text: str, max_chars: int = 32) -> str:
        """Shorten text to *max_chars* with trailing '...' if needed."""
        if len(text) <= max_chars:
            return text
        return text[: max_chars - 3] + "..."

    def _create_file_row(self, file_name: str):
        """Add a single file progress row to the panel."""
        row = ctk.CTkFrame(self._file_container, fg_color="transparent")
        row.pack(fill="x", padx=theme.PADDING_NORMAL, pady=3)

        lbl = ctk.CTkLabel(
            row, text=self._truncate(file_name), width=200, anchor="w",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=theme.TEXT_PRIMARY,
        )
        lbl.pack(side="left")

        bar = ctk.CTkProgressBar(
            row, width=300, height=14,
            progress_color=theme.ACCENT_BLUE,
            fg_color=theme.BORDER_GRAY, corner_radius=4,
        )
        bar.pack(side="left", fill="x", expand=True, padx=8)
        bar.set(0)

        status = ctk.CTkLabel(
            row, text="Waiting", width=80, anchor="e",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=theme.TEXT_MUTED,
        )
        status.pack(side="left")

        self._bars[file_name] = {
            "bar": bar, "label": lbl, "status": status, "row": row,
        }

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def show(self, file_names: list[str]):
        """Initialise progress bars for a batch of files."""
        self._clear()
        self._ensure_visible()
        for fn in file_names:
            if fn not in self._bars:
                self._create_file_row(fn)

    def add_file(self, file_name: str):
        """Incrementally add one file without rebuilding the panel."""
        self._ensure_visible()
        if file_name not in self._bars:
            self._create_file_row(file_name)

    def remove_file(self, file_name: str):
        """Remove a single file's progress row (e.g. cancelled)."""
        entry = self._bars.pop(file_name, None)
        if entry and "row" in entry:
            entry["row"].destroy()
        self._refresh_overall()

    def update_file(self, file_name: str, progress: float, status_text: str):
        """Update a single file's progress bar and status text."""
        entry = self._bars.get(file_name)
        if not entry:
            return
        entry["bar"].set(max(0, min(1, progress)))

        color = theme.TEXT_MUTED
        if "done" in status_text.lower():
            color = theme.SUCCESS_GREEN
        elif "error" in status_text.lower():
            color = theme.ERROR_RED
        elif progress > 0:
            color = theme.ACCENT_BLUE
            status_text = f"{int(progress * 100)}%"

        entry["status"].configure(text=status_text, text_color=color)
        self._refresh_overall()

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _refresh_overall(self):
        """Recalculate and display overall progress."""
        if not self._bars:
            return
        total = sum(b["bar"].get() for b in self._bars.values())
        pct = total / len(self._bars)
        if self._overall_bar:
            self._overall_bar.set(pct)
        if self._overall_pct:
            self._overall_pct.configure(text=f"{int(pct * 100)}%")

    def _clear(self):
        """Remove all child widgets."""
        for w in self.winfo_children():
            w.destroy()
        self._bars.clear()
        self._overall_bar = None
        self._overall_label = None
        self._overall_pct = None
        self._container = None
        self._file_container = None
        self._visible = False

    def hide(self):
        self._clear()
