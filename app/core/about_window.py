"""About dialog — auto-generated from registered modules."""

import customtkinter as ctk

from app.core import theme


class AboutWindow(ctk.CTkToplevel):
    """Shows app info, list of available tools, and developer credit."""

    def __init__(self, parent, modules: list):
        super().__init__(parent)
        self.title(f"About {theme.APP_SHORT_NAME}")
        self.geometry("520x500")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.configure(fg_color=theme.WHITE)

        # Centre on parent
        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        py = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")

        self._build(modules)

    def _build(self, modules):
        pad = theme.PADDING_XLARGE

        # Title
        ctk.CTkLabel(
            self, text=theme.APP_NAME,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_TITLE, "bold"),
            text_color=theme.NAVY_BLUE,
        ).pack(pady=(pad, 0))

        ctk.CTkLabel(
            self, text=theme.APP_ORG,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_LARGE),
            text_color=theme.TEXT_SECONDARY,
        ).pack()

        ctk.CTkLabel(
            self, text=f"Version {theme.APP_VERSION}",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            text_color=theme.TEXT_MUTED,
        ).pack(pady=(2, 0))

        # Separator
        ctk.CTkFrame(self, fg_color=theme.BORDER_GRAY, height=1).pack(
            fill="x", padx=pad, pady=theme.PADDING_LARGE,
        )

        # Available tools header
        ctk.CTkLabel(
            self, text="Available Tools:",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_LARGE, "bold"),
            text_color=theme.TEXT_PRIMARY, anchor="w",
        ).pack(fill="x", padx=pad)

        # Module list (auto-generated)
        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=pad, pady=(4, 0))

        for mod in modules:
            ready = mod.is_ready()
            icon = mod.get_icon()
            name = mod.get_name()
            suffix = "" if ready else " (Coming Soon)"
            color = theme.TEXT_PRIMARY if ready else theme.TEXT_MUTED

            ctk.CTkLabel(
                scroll, text=f"{name}{suffix}",
                font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL, "bold"),
                text_color=color, anchor="w",
            ).pack(fill="x", pady=(theme.PADDING_NORMAL, 0))

            desc = mod.get_description()
            ctk.CTkLabel(
                scroll, text=desc,
                font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
                text_color=theme.TEXT_SECONDARY if ready else theme.DISABLED_TEXT,
                anchor="w", justify="left",
            ).pack(fill="x", padx=(24, 0))

        # Separator
        ctk.CTkFrame(self, fg_color=theme.BORDER_GRAY, height=1).pack(
            fill="x", padx=pad, pady=theme.PADDING_LARGE,
        )

        # Credit + security
        ctk.CTkLabel(
            self, text=f"Developed by: {theme.APP_DEVELOPER}",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            text_color=theme.TEXT_PRIMARY,
        ).pack()

        ctk.CTkLabel(
            self, text=theme.APP_SECURITY_MSG,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=theme.TEXT_MUTED,
        ).pack(pady=(2, 0))

        # Close button
        ctk.CTkButton(
            self, text="Close", width=100,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            fg_color=theme.LIGHT_GRAY, hover_color=theme.BORDER_GRAY,
            text_color=theme.TEXT_PRIMARY, height=theme.BUTTON_HEIGHT,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            border_width=1, border_color=theme.BORDER_GRAY,
            command=self.destroy,
        ).pack(pady=theme.PADDING_LARGE)
