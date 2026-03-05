import xlwings as xw
from pathlib import Path
import re
import unicodedata
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
from docxtpl import DocxTemplate
from utils.vn_number_words import vnd_to_words

# ---------- CONFIG ----------

SCRIPT_DIR = Path(__file__).resolve().parent
EXCEL_FILE = SCRIPT_DIR / "Nhap_du_lieu.xlsm"
TEMPLATE_DIR = SCRIPT_DIR / "Templates"
OUTPUT_DIR = SCRIPT_DIR / "Output"
OUTPUT_DIR.mkdir(exist_ok=True)

# ---------- LOAD TEMPLATE CONFIG ----------

def load_cfg_templates():
    df = pd.read_excel(EXCEL_FILE, sheet_name="CFG_TEMPLATES")
    df.columns = [c.strip() for c in df.columns]

    cfg = {}
    for _, row in df.iterrows():
        code = str(row["TemplateCode"]).strip()
        if not code:
            continue

        enabled_raw = str(row.get("Enabled", "")).strip().upper()
        is_enabled = enabled_raw in ("1", "TRUE", "YES", "Y")

        cfg[code] = {
            "excel_sheet": row["ExcelSheet"],
            "docx_file": row["DocxFile"],
            "file_prefix": row["FilePrefix"],
            "description": row.get("Description", ""),
            "enabled": is_enabled,
        }
    return cfg

# ---------- HELPERS ----------

def round_half_up(n):
    """Vietnamese rounding: ROUND_HALF_UP, not bankers rounding to even"""
    return int(Decimal(n).quantize(0, rounding=ROUND_HALF_UP))


def smart_cast(n):
    """Return int when whole number, else float"""
    try:
        f = float(n)
    except Exception:
        return n
    if f.is_integer():
        return int(round_half_up(f))
    return f


def make_safe_filename(text):
    if not text:
        return "contract"
    text = unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode()
    text = re.sub(r"[^A-Za-z0-9_\- ]+", "_", text)
    return text.strip().replace(" ", "_") or "contract"


def to_number(x):
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        try:
            if pd.isna(x):
                return 0.0
        except Exception:
            pass
        return float(x)

    s = str(x).strip()
    if s == "":
        return 0.0

    s = s.replace("\u00A0", "").replace(" ", "")
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        if s.count(".") > 1:
            s = s.replace(".", "")
        if "," in s:
            s = s.replace(",", ".")

    s = re.sub(r"[^0-9.\-]", "", s)
    if s in ("", ".", "-", "-."):
        return 0.0

    try:
        return float(s)
    except Exception:
        return 0.0


def format_vn(n, force_decimals=False, *, decimals=None):
    if n is None:
        return ""
    try:
        num = float(n)
    except Exception:
        return str(n)

    # --- determine decimals ---
    if force_decimals:
        dec = 2 if decimals is None else int(decimals)
    else:
        if num.is_integer():
            dec = 0
        else:
            dec = 2 if decimals is None else int(decimals)

    # --- rounding (ROUND_HALF_UP) ---
    if dec == 0:
        return f"{round_half_up(num):,}".replace(",", ".")

    q = Decimal(num).quantize(
        Decimal("1." + "0" * dec),
        rounding=ROUND_HALF_UP
    )
    s = f"{q:,.{dec}f}"

    # --- Vietnamese format ---
    s = s.replace(",", "_").replace(".", ",").replace("_", ".")

    return s

# ---------- VAT CALCULATION ----------
DEFAULT_VAT_RATE = 0.08  # 8% as default
def calculate_vat(grand_total, vat_rate):
    """
    VAT rules (LOCKED):
    - VAT rate comes from context (already resolved)
    - VAT is calculated on grand total
    - VAT amount is rounded HALF UP
    - Total with VAT is rounded HALF UP
    """
    vat_amount = round_half_up(grand_total * vat_rate)
    total_with_vat = round_half_up(grand_total + vat_amount)
    return vat_amount, total_with_vat

# ---------- CONTEXT BUILDING ----------

def build_context_from_sheet(sheet_name):
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None, usecols=[0, 3])
    df.columns = ["key", "value"]

    context = {}
    for k, v in zip(df["key"], df["value"]):
        if pd.isna(k):
            continue
        key = str(k).strip()
        if key:
            context[key] = "" if pd.isna(v) else str(v)
    return context

# ---------- ITEMS ----------

def load_items():
    try:
        items_df = pd.read_excel(EXCEL_FILE, sheet_name="Items")
    except Exception:
        return []

    items = []
    if items_df is not None:
        cols_norm = {c: str(c).strip().lower().replace(" ", "_").replace("-", "_") for c in items_df.columns}
        items_df = items_df.rename(columns=cols_norm)

        for _, row in items_df.iterrows():
            ten = (row.get("ten_hang", "") or row.get("tên_hang", "") or row.get("tên hàng", "") or "")
            dvt = row.get("dvt", "") or row.get("đvt", "") or ""
            so_raw = row.get("so_luong", row.get("số_lượng", 0))
            dg_raw = row.get("don_gia", row.get("đơn_gia", 0))
            tt_raw = row.get("thanh_tien", row.get("thành_tiền", None))

            so = to_number(so_raw)
            dg = to_number(dg_raw)
            tt = to_number(tt_raw) if tt_raw not in (None, "", " ") else so * dg

            items.append({
                # numeric (math-safe)
                "so_luong_num": so,
                "don_gia_num": dg,
                "thanh_tien_num": tt,

                # raw (smart_cast preserved)
                "so_luong_raw": smart_cast(so),
                "don_gia_raw": smart_cast(dg),
                "thanh_tien_raw": smart_cast(tt),

                # formatted (Word)
                "so_luong": format_vn(so, force_decimals=True),
                "don_gia": format_vn(dg),
                "thanh_tien": format_vn(tt),

                # text
                "ten_hang": str(ten),
                "dvt": str(dvt),
            })

    return items

# ---------- TOTALS & VAT ----------

def enrich_totals(context, items):
    grand_total = sum(i["thanh_tien_num"] for i in items)

    vat_rate = to_number(context.get("VAT_RATE", DEFAULT_VAT_RATE))
    if vat_rate <= 0:
        vat_rate = DEFAULT_VAT_RATE

    vat_amount, grand_total_vat = calculate_vat(grand_total, vat_rate)

    context.update({
        "items": items,
        "grand_total": round_half_up(grand_total),
        "grand_total_formatted": format_vn(grand_total),
        "vat_amount_raw": vat_amount,
        "vat_amount_formatted": format_vn(vat_amount),
        "grand_total_vat_raw": grand_total_vat,
        "grand_total_vat_formatted": format_vn(grand_total_vat),
        "grand_total_text": vnd_to_words(round_half_up(grand_total)),
        "grand_total_vat_text": vnd_to_words(grand_total_vat),
    })

# ---------- TEMPLATE RENDERING ----------

def render_template(template_code, context, seq_num, name_hint, cfg):
    tpl_cfg = cfg[template_code]
    template_path = TEMPLATE_DIR / tpl_cfg["docx_file"]

    customer_folder = OUTPUT_DIR / make_safe_filename(name_hint)
    customer_folder.mkdir(exist_ok=True)

    output_file = customer_folder / (
        f"{seq_num}_{tpl_cfg['file_prefix']}_{make_safe_filename(name_hint)}.docx"
    )

    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_file)

    print("Created:", output_file)
    return customer_folder

# ---------- MAIN WORKFLOW ----------

def run_workflow(wb, template_code, cfg):
    try:
        source_sheet_override = wb.sheets["UI_DASHBOARD"].range("B8").value
    except Exception:
        source_sheet_override = None

    items = load_items()
    now = datetime.now()

    if template_code == "ALL":
        template_keys = [
            k for k, v in cfg.items()
            if v.get("enabled", True)
        ]
    else:
        template_keys = [template_code]

    if template_code == "ALL" and not template_keys:
        raise ValueError("No templates are enabled in CFG_TEMPLATES.")

    created_folder = None

    for key in template_keys:
        sheet_name = source_sheet_override or cfg[key]["excel_sheet"]

        context = build_context_from_sheet(sheet_name)

        seq_raw = context.get("STT_HD", "")
        try:
            seq_num = str(int(float(seq_raw))).zfill(2)
        except Exception:
            seq_num = "00"

        context.setdefault("DAY", now.day)
        context.setdefault("MONTH", now.month)
        context.setdefault("YEAR", now.year)

        enrich_totals(context, items)

        name_hint = context.get("TEN_KH") or context.get("KH_ABB") or "contract"
        created_folder = render_template(key, context, seq_num, name_hint, cfg)

    return str(created_folder)

# ---------- ENTRY POINT ----------

def main():
    wb = xw.Book.caller()
    cfg = load_cfg_templates()

    template_code = wb.sheets["UI_DASHBOARD"].range("B2").value

    if template_code != "ALL" and template_code not in cfg:
        raise ValueError(
            f"Unknown template code: {template_code}. "
            f"Valid codes: {list(cfg.keys()) + ['ALL']}"
        )

    folder = run_workflow(wb, template_code, cfg)
    wb.sheets["UI_DASHBOARD"].range("B7").value = folder


if __name__ == "__main__":
    xw.Book(EXCEL_FILE).set_mock_caller()
    main()