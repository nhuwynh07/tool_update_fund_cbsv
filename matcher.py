"""
matcher.py – Khớp tên thành viên, phân loại giao dịch, xác định tháng nộp quỹ.
"""
import re

try:
    from rapidfuzz import fuzz
except ImportError:
    fuzz = None
    from difflib import SequenceMatcher

from text_utils import norm_text
from config import (
    STOP_WORDS,
    PHAT_RULES, QUY_KEYWORDS, QUY_AMOUNT_HINTS, PHAT_AMOUNT_HINTS,
    ALLOW_PHAT_AMOUNT_HINT,
    MIN_MATCH_SCORE, AMBIGUOUS_SCORE_GAP,
    DEFAULT_COLLECTION_QUARTER, QUY_MONTH_COLS,
)


# ── Trích xuất cụm tên người từ nội dung giao dịch ──────────────────────────

def compact_person_phrase(text: str) -> str:
    """Lọc stop-words, số, từ 1 ký tự; giữ tối đa 6 token cuối."""
    tokens = []
    for tok in norm_text(text).split():
        if tok in STOP_WORDS or any(ch.isdigit() for ch in tok) or len(tok) <= 1:
            continue
        tokens.append(tok)
    return " ".join(tokens[-6:])


def candidate_texts(search_text: str) -> list:
    """
    Sinh danh sách các cụm văn bản ứng viên từ nội dung tìm kiếm,
    để thử khớp với tên thành viên.
    """
    t = norm_text(search_text)
    candidates = []

    base = compact_person_phrase(t)
    if len(base.split()) >= 2:
        candidates.append(base)

    patterns = [
        r"ct\s*tu\s+\d[\d\s]*\s+([a-z\s]{4,100}?)\s+t\s*oi\b",
        r"(?:^|\bqr\b|[-])\s*([a-z\s]{4,90}?)\s+chuyen\b",
        r"([a-z\s]{4,90}?)\s+transfer\b",
        r"([a-z\s]{4,90}?)\s+nop\b",
        r"[-]\s*([a-z\s]{4,90}?)$",
    ]
    for pat in patterns:
        for m in re.finditer(pat, t):
            cand = compact_person_phrase(m.group(1))
            if len(cand.split()) >= 2:
                candidates.append(cand)

    for phrase in re.split(r"\bchuyen\b|\bnop\b|\btransfer\b", t):
        cand = compact_person_phrase(phrase)
        if len(cand.split()) >= 2:
            candidates.append(cand)

    candidates.append(t)
    return list(dict.fromkeys(candidates))


# ── Tính điểm khớp tên ──────────────────────────────────────────────────────

def fuzzy_score(a: str, b: str) -> float:
    """Điểm khớp mờ giữa hai chuỗi đã chuẩn hoá (0–100)."""
    if fuzz is not None:
        return max(float(fuzz.token_set_ratio(a, b)), float(fuzz.partial_ratio(a, b)))
    return SequenceMatcher(None, a, b).ratio() * 100.0


def score_name(name_norm: str, candidate_norm: str) -> float:
    """Tính điểm khớp giữa tên thành viên và một cụm ứng viên từ sao kê."""
    name_norm      = norm_text(name_norm)
    candidate_norm = norm_text(candidate_norm)
    if not name_norm or not candidate_norm:
        return 0.0
    if name_norm in candidate_norm:
        return 100.0
    name_tokens = set(name_norm.split())
    cand_tokens = [
        tok for tok in candidate_norm.split()
        if tok not in STOP_WORDS and not any(ch.isdigit() for ch in tok)
    ]
    if len(cand_tokens) < 2:
        return 0.0
    cand_set = set(cand_tokens)
    if cand_set.issubset(name_tokens):
        return min(99.0, 88.0 + len(cand_set) * 4.0)
    overlap = 100.0 * len(name_tokens & cand_set) / max(len(name_tokens), 1)
    return max(overlap, fuzzy_score(name_norm, " ".join(cand_tokens)))


def match_member(tx: dict, members: list) -> tuple:
    """
    Tìm thành viên khớp nhất với giao dịch.
    Trả về (member | None, best_score, ghi_chu).
    """
    candidates = candidate_texts(tx["search_text"])
    ranked = []
    for member in members:
        score = max(score_name(member["name_norm"], cand) for cand in candidates)
        ranked.append((score, member))
    ranked.sort(key=lambda x: x[0], reverse=True)

    best_score, best_member = ranked[0]
    second_score = ranked[1][0] if len(ranked) > 1 else 0.0

    if best_score < MIN_MATCH_SCORE:
        return None, best_score, "Tên không khớp đủ ngưỡng"
    if second_score >= MIN_MATCH_SCORE and best_score - second_score <= AMBIGUOUS_SCORE_GAP:
        return None, best_score, (
            f"Tên khớp mơ hồ: {best_member['name']} và {ranked[1][1]['name']}"
        )
    return best_member, best_score, "OK"


# ── Phân loại giao dịch ─────────────────────────────────────────────────────

def classify_transaction(tx: dict) -> tuple:
    """
    Phân loại giao dịch thành: 'ignore' | 'phat' | 'quy' | 'unknown'.
    Trả về (tx_type, target_col | None, ghi_chu).

    Lưu ý: giao dịch 20.000đ không rõ tháng/quý → phạt KQHT (col 7),
    tránh ghi sai vào tháng giao dịch thực tế (VD: nộp trễ tháng 4).
    """
    details_norm = norm_text(tx["details"])

    # Lãi tài khoản → bỏ qua
    if "tra lai so du" in details_norm or "lai so du" in details_norm:
        return "ignore", None, "Bỏ qua lãi tài khoản"

    # Khớp từ khóa phạt
    for rule in PHAT_RULES:
        if any(norm_text(k) in details_norm for k in rule["keywords"]):
            return "phat", rule["col"], rule["label"]

    # Từ khóa quỹ/đảng phí
    if any(norm_text(k) in details_norm for k in QUY_KEYWORDS):
        return "quy", None, "Từ khóa quỹ/đảng phí"

    # Gợi ý theo số tiền
    if tx["amount"] in QUY_AMOUNT_HINTS:
        period = QUY_AMOUNT_HINTS[tx["amount"]]
        if period == 1:
            # 20k không rõ tháng/quý → phạt KQHT, không điền vào quỹ tránh sai tháng
            has_month   = bool(re.search(r"\bthang\s*(1[0-2]|[1-9])\b", details_norm))
            has_quarter = bool(re.search(r"\b(q\s*[1-4]|quy\s*[1-4i]{1,3})\b", details_norm))
            if not has_month and not has_quarter:
                return "phat", 7, "20k không rõ tháng/quý → phạt KQHT"
        return "quy", None, "Suy luận quỹ theo số tiền"

    if ALLOW_PHAT_AMOUNT_HINT and tx["amount"] in PHAT_AMOUNT_HINTS:
        return "phat", PHAT_AMOUNT_HINTS[tx["amount"]], "Suy luận phạt nợ môn theo số tiền"

    return "unknown", None, "Không nhận diện được loại thu"


# ── Xác định tháng nộp quỹ ──────────────────────────────────────────────────

def months_of_quarter(q: int) -> list:
    """Trả về [tháng_đầu, tháng_giữa, tháng_cuối] của quý q."""
    start = (q - 1) * 3 + 1
    return [start, start + 1, start + 2]


def get_quy_months(tx: dict) -> list:
    """
    Xác định danh sách các tháng cần cập nhật quỹ cho giao dịch tx.
    Ưu tiên: tháng tường minh > quý tường minh > cả năm > gợi ý số tiền > tháng GD.
    """
    t = norm_text(tx["details"])

    # Tháng ghi rõ trong nội dung
    explicit_month = re.search(r"\bthang\s*(1[0-2]|[1-9])\b", t)
    if explicit_month:
        return [int(explicit_month.group(1))]

    # Quý ghi rõ trong nội dung
    quarter_patterns = {
        1: [r"\bq\s*1\b", r"\bquy\s*1\b", r"\bquy\s*i\b",   r"\bquy1\b"],
        2: [r"\bq\s*2\b", r"\bquy\s*2\b", r"\bquy\s*ii\b",  r"\bquy2\b"],
        3: [r"\bq\s*3\b", r"\bquy\s*3\b", r"\bquy\s*iii\b", r"\bquy3\b"],
        4: [r"\bq\s*4\b", r"\bquy\s*4\b", r"\bquy\s*iv\b",  r"\bquy4\b"],
    }
    for q, patterns in quarter_patterns.items():
        if any(re.search(p, t) for p in patterns):
            return months_of_quarter(q)

    # Cả năm
    if "ca nam" in t or tx["amount"] >= 240_000:
        return list(range(1, 13))

    # Gợi ý theo số tiền (chỉ period 3 và 12; period 1 đã xử lý ở classify)
    if tx["amount"] in QUY_AMOUNT_HINTS:
        period = QUY_AMOUNT_HINTS[tx["amount"]]
        if period == 12:
            return list(range(1, 13))
        if period == 3:
            return months_of_quarter(DEFAULT_COLLECTION_QUARTER)

    # Mặc định: tháng của ngày giao dịch
    return [int(tx["date"].month)]
