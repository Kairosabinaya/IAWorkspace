"""Abstract base class for all workspace modules."""

from abc import ABC, abstractmethod


class BaseModule(ABC):
    """Every feature/tool in the workspace must extend this class."""

    @abstractmethod
    def get_name(self) -> str:
        """Module name for sidebar display, e.g. 'Excel Formatter'."""

    @abstractmethod
    def get_icon(self) -> str:
        """Short icon text for sidebar, or empty string."""

    @abstractmethod
    def get_description(self) -> str:
        """Short description for About window (bullet-point style)."""

    @abstractmethod
    def create_view(self, parent_frame) -> object:
        """Create and return the main widget/frame for this module."""

    def is_ready(self) -> bool:
        """Return False to show module as 'Coming Soon' in sidebar."""
        return True
