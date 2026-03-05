""" # vn-number-words

Vietnamese number-to-words conversion for Python.

## Features
- Correct Vietnamese grammar (linh, mốt, lăm)
- Decimal truncation (no rounding)
- Optional decimal reading
- VND money conversion
- No dependencies

"""

__version__ = "1.0.0"

from typing import List

# -------------------- CONSTANTS --------------------

DIGITS = (
    "không", "một", "hai", "ba", "bốn",
    "năm", "sáu", "bảy", "tám", "chín"
)

UNITS = ("", "nghìn", "triệu", "tỷ", "nghìn tỷ")

# -------------------- CORE HELPERS --------------------

def _read_unit(unit: int, ten: int) -> str:
    if unit == 1 and ten > 1:
        return "mốt"
    if unit == 5 and ten >= 1:
        return "lăm"
    return DIGITS[unit]


def _read_hundreds(number: int, full: bool) -> List[str]:
    hundred = number // 100
    ten = (number % 100) // 10
    unit = number % 10

    words: List[str] = []

    if hundred:
        words.extend([DIGITS[hundred], "trăm"])
    elif full and (ten or unit):
        words.extend(["không", "trăm"])

    if ten == 0 and unit:
        if hundred or full:
            words.append("linh")
        words.append(DIGITS[unit])
    elif ten == 1:
        words.append("mười")
        if unit:
            words.append(_read_unit(unit, ten))
    elif ten > 1:
        words.extend([DIGITS[ten], "mươi"])
        if unit:
            words.append(_read_unit(unit, ten))

    return words


def _split_thousands(number: int) -> List[int]:
    parts: List[int] = []
    while number:
        parts.append(number % 1000)
        number //= 1000
    return parts


def _read_integer(number: int, use_commas: bool = False) -> str:
    if number == 0:
        return "không"

    parts = _split_thousands(number)
    blocks: List[str] = []

    highest_idx = max(i for i, p in enumerate(parts) if p != 0)

    for idx, part in enumerate(parts):
        if part == 0:
            continue

        full = idx < highest_idx
        words = _read_hundreds(part, full)
        unit = UNITS[idx]

        text = " ".join(words + ([unit] if unit else []))
        blocks.insert(0, text)

    sep = ", " if use_commas else " "
    return sep.join(blocks)


def _read_decimal(decimal: str) -> str:
    return " ".join(DIGITS[int(d)] for d in decimal)


def _truncate_decimal(value: float, decimal_places: int) -> str:
    text = str(value)

    if "." not in text:
        return "0" * decimal_places

    decimal = text.split(".", 1)[1]
    decimal = decimal[:decimal_places]

    return decimal.ljust(decimal_places, "0")

# -------------------- PUBLIC API --------------------

def number_to_words(
    value: int | float,
    *,
    decimal_places: int = 0,
    use_commas: bool = False
) -> str:
    if not isinstance(value, (int, float)):
        raise TypeError("value must be int or float")

    if decimal_places < 0:
        raise ValueError("decimal_places must be >= 0")

    negative = value < 0
    value = abs(value)

    integer_part = int(value)
    result = _read_integer(integer_part, use_commas)

    if decimal_places > 0:
        decimal_str = _truncate_decimal(value, decimal_places)
        decimal_words = _read_decimal(decimal_str)
        result = f"{result} phẩy {decimal_words}"

    if negative:
        result = f"âm {result}"

    return result.capitalize()


def vnd_to_words(
    amount: int,
    *,
    use_commas: bool = False,
    append_chan: bool = True
) -> str:
    if not isinstance(amount, int):
        raise TypeError("amount must be int (VND has no decimals)")

    words = _read_integer(abs(amount), use_commas)

    if append_chan:
        words += " đồng chẵn"
    else:
        words += " đồng"

    if amount < 0:
        words = f"âm {words}"

    return words.capitalize()


__all__ = ["number_to_words", "vnd_to_words"]