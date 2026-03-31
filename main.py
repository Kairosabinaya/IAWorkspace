"""Entry point for Internal Audit Workspace."""

import sys
import os

# Ensure the project root is on sys.path so absolute imports work
# when running as `python main.py` from within ia-workspace/.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import customtkinter as ctk

from app.core.app_shell import AppShell
from app.core.module_registry import MODULES


def main():
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    app = AppShell(MODULES)
    app.mainloop()


if __name__ == "__main__":
    main()
