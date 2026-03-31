"""Global constants and per-module defaults."""

# Excel Formatter defaults
EXCEL_FONT_NAME = "Arial"
EXCEL_FONT_SIZE = 9
EXCEL_MIN_COL_WIDTH = 8
EXCEL_MAX_COL_WIDTH = 50
EXCEL_HEADER_SCAN_ROWS = 20
EXCEL_SAMPLE_ROWS = 100
EXCEL_DATE_THRESHOLD = 0.6  # >60% date-like values to recommend as date column

# Date format options — (display label, openpyxl number_format)
DATE_FORMATS = [
    ("DD-MMM-YY", "DD-MMM-YY"),
    ("DD/MM/YYYY", "DD/MM/YYYY"),
    ("DD-MM-YYYY", "DD-MM-YYYY"),
    ("YYYY-MM-DD", "YYYY-MM-DD"),
    ("DD MMMM YYYY", "DD MMMM YYYY"),
    ("MMM DD, YYYY", "MMM DD, YYYY"),
]
DEFAULT_DATE_FORMAT_INDEX = 0

# Number formats — keyed by separator style
NUMBER_FORMAT_INTEGER = "#,##0"
NUMBER_FORMAT_DECIMAL = "#,##0.00"
NUMBER_FORMAT_DOT_INTEGER = "#.##0"
NUMBER_FORMAT_DOT_DECIMAL = "#.##0,00"

SEPARATOR_OPTIONS = [
    ("1,000.00", ","),
    ("1.000,00", "."),
]

# Output
DEFAULT_OUTPUT_FOLDER = "Formatted"

# Supported file extensions
SUPPORTED_EXTENSIONS = {".xlsx"}

# Header detection heuristic keywords
DATE_COLUMN_KEYWORDS = {
    "date", "tanggal", "tgl", "period", "bulan", "tahun",
    "due", "effective", "posting", "entry", "expired", "maturity",
}

ID_COLUMN_KEYWORDS = {
    "id", "no", "no.", "number", "num", "kode", "code", "ref",
    "reference", "journal", "voucher", "doc", "document", "nomor",
    "index", "key", "seq",
}

AMOUNT_COLUMN_KEYWORDS = {
    "amount", "total", "balance", "qty", "quantity", "harga", "price",
    "nilai", "value", "saldo", "debit", "credit", "kredit", "cost",
    "revenue", "sales", "budget", "actual", "variance", "fee", "tax",
    "pajak", "gross", "net", "sum", "subtotal", "payment", "bayar",
    "tagihan", "piutang", "hutang", "receivable", "payable", "unit",
    "volume",
}
