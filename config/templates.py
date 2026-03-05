import pandas as pd
from config.paths import EXCEL_FILE

def load_cfg_templates():
    df = pd.read_excel(EXCEL_FILE, sheet_name="CFG_TEMPLATES")
    df.columns = [c.strip() for c in df.columns]

    cfg = {}
    for _, row in df.iterrows():
        code = str(row["TemplateCode"]).strip()
        if not code:
            continue

        enabled = str(row.get("Enabled", "")).strip().upper() in ("1", "TRUE", "YES", "Y")

        cfg[code] = {
            "excel_sheet": row["ExcelSheet"],
            "docx_file": row["DocxFile"],
            "file_prefix": row["FilePrefix"],
            "description": row.get("Description", ""),
            "enabled": enabled,
        }

    return cfg
