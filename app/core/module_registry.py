"""Registry of all available workspace modules.

To add a new module:
1. Create a subfolder in app/modules/your_module/
2. Implement a class extending BaseModule
3. Add one import + one line to MODULES below
"""

from app.modules.excel_formatter.module import ExcelFormatterModule
# from app.modules.pdf_manager.module import PdfManagerModule  # uncomment when ready

MODULES = [
    ExcelFormatterModule(),
    # PdfManagerModule(),  # uncomment when ready
]
