from domain.items import load_items
from domain.context import build_context, enrich_totals
from infra.rendering import render

def run(wb, template_code, cfg):
    override = wb.sheets["UI_DASHBOARD"].range("B8").value
    items = load_items()

    keys = cfg.keys() if template_code == "ALL" else [template_code]

    last_folder = None
    for k in keys:
        if not cfg[k]["enabled"]:
            continue

        sheet = override or cfg[k]["excel_sheet"]
        ctx = build_context(sheet)

        seq = ctx.get("STT_HD", "00")
        seq = str(int(float(seq))).zfill(2) if str(seq).strip() else "00"

        enrich_totals(ctx, items)

        name = ctx.get("TEN_KH") or ctx.get("KH_ABB") or "contract"
        last_folder = render(cfg[k], ctx, seq, name)

    return str(last_folder)
