from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parents[1]

EXCEL_FILE = SCRIPT_DIR / "Nhap_du_lieu.xlsm"
TEMPLATE_DIR = SCRIPT_DIR / "Templates"
OUTPUT_DIR = SCRIPT_DIR / "Output"

OUTPUT_DIR.mkdir(exist_ok=True)