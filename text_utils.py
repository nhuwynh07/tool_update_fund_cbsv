"""
text_utils.py – Các hàm tiện ích xử lý văn bản và số tiền.
"""
import re
import unicodedata

import pandas as pd


def remove_accents(value: str) -> str:
    """Bỏ dấu tiếng Việt, chuyển đ/Đ → d/D."""
    value = "" if value is None else str(value)
    value = unicodedata.normalize("NFD", value)
    value = "".join(ch for ch in value if unicodedata.category(ch) != "Mn")
    return value.replace("đ", "d").replace("Đ", "D")


def norm_text(value) -> str:
    """Chuẩn hoá chuỗi: bỏ dấu → lower → chỉ giữ a-z0-9 → cắt khoảng trắng."""
    value = remove_accents(value).lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def parse_money(value) -> int:
    """Chuyển giá trị ô Excel sang số nguyên (đồng). Trả về 0 nếu rỗng/lỗi."""
    if value is None or value == "":
        return 0
    if isinstance(value, (int, float)):
        return int(round(value))
    value = str(value).strip()
    if value.startswith("="):
        return 0
    value = re.sub(r"[^0-9-]", "", value)
    return int(value) if value not in ("", "-") else 0


def parse_date(value):
    """Chuyển giá trị ô Excel sang pandas Timestamp. Trả về NaT nếu lỗi."""
    if value is None or value == "":
        return pd.NaT
    return pd.to_datetime(value, dayfirst=True, errors="coerce")
