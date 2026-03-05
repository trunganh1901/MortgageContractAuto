import re
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP

def round_half_up(n):
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

def to_number(x):
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return 0.0 if pd.isna(x) else float(x)

    s = str(x).strip()
    if not s:
        return 0.0

    s = s.replace("\u00A0", "").replace(" ", "")
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        if s.count(".") > 1:
            s = s.replace(".", "")
        s = s.replace(",", ".")

    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s)
    except Exception:
        return 0.0

def format_vn(n, force_decimals=False, *, decimals=None):
    if n is None:
        return ""
    num = float(n)

    if force_decimals:
        dec = 2 if decimals is None else int(decimals)
    else:
        dec = 0 if num.is_integer() else (2 if decimals is None else int(decimals))

    if dec == 0:
        return f"{round_half_up(num):,}".replace(",", ".")

    q = Decimal(num).quantize(Decimal("1." + "0" * dec), rounding=ROUND_HALF_UP)
    s = f"{q:,.{dec}f}"
    return s.replace(",", "_").replace(".", ",").replace("_", ".")
