import re
import pandas as pd

def normalize_phone(raw):
    """
    From original main.py: normalize phone numbers to +62 / +... format.
    Preserves same behavior as original function.
    """
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    s = str(raw).strip()
    if s.startswith('+'):
        digits = re.sub(r'\D', '', s[1:])
        return '+' + digits if digits else None
    digits = re.sub(r'\D', '', s)
    if not digits:
        return None
    if digits.startswith('0'):
        return '+62' + digits[1:]
    if digits.startswith('62'):
        return '+' + digits
    if digits.startswith('8'):
        return '+62' + digits
    return '+' + digits
