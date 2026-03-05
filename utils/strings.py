import re
import unicodedata

def make_safe_filename(text):
    if not text:
        return "contract"
    text = unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode()
    text = re.sub(r"[^A-Za-z0-9_\- ]+", "_", text)
    return text.strip().replace(" ", "_") or "contract"