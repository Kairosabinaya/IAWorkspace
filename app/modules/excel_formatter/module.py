"""Excel Formatter module — BaseModule implementation."""

from app.core.base_module import BaseModule
from app.modules.excel_formatter.gui.formatter_view import FormatterView


class ExcelFormatterModule(BaseModule):

    def get_name(self) -> str:
        return "Excel Formatter"

    def get_icon(self) -> str:
        return ""

    def get_description(self) -> str:
        return (
            "Automatically format Excel files with consistent styling:\n"
            "- Standardize fonts, borders, and headers\n"
            "- Smart detection of date and numeric columns\n"
            "- Thousand separator for amounts (not IDs)\n"
            "- Auto-fit column widths and freeze panes\n"
            "- Batch process multiple files at once"
        )

    def create_view(self, parent_frame) -> FormatterView:
        return FormatterView(parent_frame)
