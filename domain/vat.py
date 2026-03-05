from utils.numbers import round_half_up

DEFAULT_VAT_RATE = 0.08

def calculate_vat(grand_total, vat_rate):
    vat_amount = round_half_up(grand_total * vat_rate)
    total_with_vat = round_half_up(grand_total + vat_amount)
    return vat_amount, total_with_vat

