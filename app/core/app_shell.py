"""Main application shell: sidebar navigation + content area."""

import customtkinter as ctk

from app.core import theme
from app.core.about_window import AboutWindow
from app.core.base_module import BaseModule


class AppShell(ctk.CTk):
    """Top-level window with a sidebar and swappable content panel."""

    def __init__(self, modules: list[BaseModule]):
        super().__init__()
        self._modules = modules
        self._current_view = None
        self._sidebar_buttons: list[ctk.CTkButton] = []

        # Window setup
        self.title(f"{theme.APP_NAME} - {theme.APP_ORG}")
        self.geometry(f"{theme.WINDOW_MIN_WIDTH}x{theme.WINDOW_MIN_HEIGHT}")
        self.minsize(theme.WINDOW_MIN_WIDTH, theme.WINDOW_MIN_HEIGHT)
        self.configure(fg_color=theme.WHITE)

        # Try to set icon
        try:
            import os, sys
            base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(base, "..", "..", "assets", "icon.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass

        self._build_layout()

        # Select the first ready module by default
        for i, mod in enumerate(modules):
            if mod.is_ready():
                self._select_module(i)
                break

    def _build_layout(self):
        # ---- Sidebar ----
        self._sidebar = ctk.CTkFrame(
            self, fg_color=theme.SIDEBAR_BG,
            width=theme.SIDEBAR_WIDTH, corner_radius=0,
        )
        self._sidebar.pack(side="left", fill="y")
        self._sidebar.pack_propagate(False)

        # App title in sidebar
        title_frame = ctk.CTkFrame(self._sidebar, fg_color="transparent")
        title_frame.pack(fill="x", padx=theme.PADDING_LARGE, pady=(theme.PADDING_XLARGE, theme.PADDING_LARGE))

        ctk.CTkLabel(
            title_frame, text=theme.APP_SHORT_NAME,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_HEADING, "bold"),
            text_color=theme.WHITE, anchor="w",
        ).pack(fill="x")

        # Separator
        ctk.CTkFrame(self._sidebar, fg_color=theme.SIDEBAR_HOVER, height=1).pack(
            fill="x", padx=theme.PADDING_NORMAL, pady=(0, theme.PADDING_NORMAL),
        )

        # Module buttons
        for i, mod in enumerate(self._modules):
            enabled = mod.is_ready()
            btn = ctk.CTkButton(
                self._sidebar,
                text=f"  {mod.get_name()}",
                font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
                fg_color="transparent",
                hover_color=theme.SIDEBAR_HOVER,
                text_color=theme.WHITE if enabled else theme.TEXT_MUTED,
                anchor="w", height=40,
                corner_radius=theme.BUTTON_CORNER_RADIUS,
                state="normal" if enabled else "disabled",
                command=lambda idx=i: self._select_module(idx),
            )
            btn.pack(fill="x", padx=theme.PADDING_NORMAL, pady=2)
            self._sidebar_buttons.append(btn)

            if not enabled:
                ctk.CTkLabel(
                    self._sidebar,
                    text="      Coming Soon",
                    font=(theme.FONT_FAMILY, 10),
                    text_color=theme.TEXT_MUTED, anchor="w",
                ).pack(fill="x", padx=theme.PADDING_NORMAL)

        # Spacer
        ctk.CTkFrame(self._sidebar, fg_color="transparent").pack(fill="both", expand=True)

        # Bottom: separator + About + version
        ctk.CTkFrame(self._sidebar, fg_color=theme.SIDEBAR_HOVER, height=1).pack(
            fill="x", padx=theme.PADDING_NORMAL, pady=(0, theme.PADDING_NORMAL),
        )

        about_btn = ctk.CTkButton(
            self._sidebar,
            text="  About",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            fg_color="transparent", hover_color=theme.SIDEBAR_HOVER,
            text_color=theme.WHITE, anchor="w", height=36,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            command=self._show_about,
        )
        about_btn.pack(fill="x", padx=theme.PADDING_NORMAL, pady=2)

        version_lbl = ctk.CTkLabel(
            self._sidebar, text=f"v{theme.APP_VERSION}",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=theme.TEXT_MUTED, anchor="w",
        )
        version_lbl.pack(fill="x", padx=theme.PADDING_XLARGE, pady=(0, 2))

        credit_lbl = ctk.CTkLabel(
            self._sidebar, text=f"Developed by\n{theme.APP_DEVELOPER}",
            font=(theme.FONT_FAMILY, 10),
            text_color=theme.TEXT_MUTED, anchor="w", justify="left",
        )
        credit_lbl.pack(fill="x", padx=theme.PADDING_XLARGE, pady=(0, theme.PADDING_LARGE))

        # ---- Content area ----
        self._content = ctk.CTkFrame(self, fg_color=theme.WHITE, corner_radius=0)
        self._content.pack(side="left", fill="both", expand=True)

        # ---- Status bar ----
        self._status_bar = ctk.CTkFrame(
            self, fg_color=theme.LIGHT_GRAY, height=28, corner_radius=0,
        )
        self._status_bar.pack(side="bottom", fill="x")
        self._status_bar.pack_propagate(False)

        self._status_label = ctk.CTkLabel(
            self._status_bar, text="Ready",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=theme.TEXT_MUTED, anchor="w",
        )
        self._status_label.pack(side="left", padx=theme.PADDING_NORMAL)

    # ------------------------------------------------------------------
    # Module switching
    # ------------------------------------------------------------------

    def _select_module(self, index: int):
        mod = self._modules[index]
        if not mod.is_ready():
            return

        # Highlight active button
        for i, btn in enumerate(self._sidebar_buttons):
            if i == index:
                btn.configure(fg_color=theme.SIDEBAR_ACTIVE)
            else:
                btn.configure(fg_color="transparent")

        # Swap content
        if self._current_view:
            self._current_view.destroy()

        view = mod.create_view(self._content)
        view.pack(fill="both", expand=True)
        self._current_view = view

        self._status_label.configure(text=mod.get_name())

    def _show_about(self):
        AboutWindow(self, self._modules)

    def set_status(self, text: str):
        self._status_label.configure(text=text)
