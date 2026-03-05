import pandas as pd
from datetime import datetime
from utils.vn_number_words import vnd_to_words
from config.paths import EXCEL_FILE
from utils.numbers import round_half_up, to_number, format_vn
from domain.vat import calculate_vat, DEFAULT_VAT_RATE

def build_context(sheet_name):
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None, usecols=[0, 3])
    df.columns = ["key", "value"]

    ctx = {}
    for k, v in zip(df.key, df.value):
        if pd.isna(k):
            continue
        ctx[str(k).strip()] = "" if pd.isna(v) else str(v)

    now = datetime.now()
    ctx.setdefault("DAY", now.day)
    ctx.setdefault("MONTH", now.month)
    ctx.setdefault("YEAR", now.year)

    return ctx

def enrich_totals(ctx, items):
    grand = sum(i["thanh_tien_num"] for i in items)
    rate = to_number(ctx.get("VAT_RATE", DEFAULT_VAT_RATE)) or DEFAULT_VAT_RATE

    vat, total = calculate_vat(grand, rate)

    ctx.update({
        "items": items,
        "grand_total": round_half_up(grand),
        "grand_total_formatted": format_vn(grand),
        "vat_amount_formatted": format_vn(vat),
        "grand_total_vat_formatted": format_vn(total),
        "grand_total_text": vnd_to_words(round_half_up(grand)),
        "grand_total_vat_text": vnd_to_words(total),
    })
