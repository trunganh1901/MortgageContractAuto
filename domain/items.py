import pandas as pd
from config.paths import EXCEL_FILE
from utils.numbers import to_number, format_vn

def load_items():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="Items")
    except Exception:
        return []

    cols = {c: str(c).strip().lower().replace(" ", "_") for c in df.columns}
    df = df.rename(columns=cols)

    items = []
    for _, r in df.iterrows():
        so = to_number(r.get("so_luong", 0))
        dg = to_number(r.get("don_gia", 0))
        tt = to_number(r.get("thanh_tien")) or so * dg

        items.append({
            "ten_hang": str(r.get("ten_hang", "")),
            "dvt": str(r.get("dvt", "")),
            "so_luong_num": so,
            "don_gia_num": dg,
            "thanh_tien_num": tt,
            "so_luong": format_vn(so, force_decimals=True),
            "don_gia": format_vn(dg),
            "thanh_tien": format_vn(tt),
        })

    return items
