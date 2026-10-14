"""
config.py – Tất cả hằng số và cấu hình cho hệ thống đối soát chi bộ.
Chỉnh sửa file này để thay đổi đường dẫn, cột, ngưỡng khớp, v.v.

Mức đóng quỹ và mức phạt có thể chỉnh trên giao diện (tab "Quản lý mức thu"),
được lưu vào settings.json và ghi đè các giá trị mặc định bên dưới.
"""
import copy
import json
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
SETTINGS_PATH = BASE_DIR / "settings.json"

# ── Đường dẫn file ──────────────────────────────────────────────────────────
SAO_KE_PATH   = BASE_DIR / "SaoKeTK_04102026_12102026_1.xlsx"
TONG_HOP_PATH = BASE_DIR / "2026.10.10CBSV1 BC Dang phi.xlsx"
OUTPUT_PATH   = BASE_DIR / "2026.10.10CBSV1 BC Dang phi_UPDATED.xlsx"

# ── Tên sheet ────────────────────────────────────────────────────────────────
SHEET_QUY  = "BCQD 2026"
SHEET_PHAT = "BC che tai CB"
LOG_SHEET  = "DoiSoat_Log"

# ── Cấu trúc bảng tổng hợp ──────────────────────────────────────────────────
DATA_START_ROW = 8          # Dòng bắt đầu có dữ liệu thành viên
NAME_COLS      = (2, 3)     # Cột họ, tên
TT_COL         = 1          # Cột số thứ tự

# ── Cột tháng trong sheet Quỹ ───────────────────────────────────────────────
QUY_MONTH_COLS = {
    1: 5,  2: 6,  3: 7,
    4: 9,  5: 10, 6: 11,
    7: 13, 8: 14, 9: 15,
    10: 17, 11: 18, 12: 19,
}
QUY_QUARTER_TOTAL_COLS  = {1: 8, 2: 12, 3: 16, 4: 20}
QUY_TOTAL_COL            = 21
QUY_STANDARD_MONTH_AMOUNT = 20_000
# Thu đảng phí: GD đúng số tiền này được phân loại 'dangphi' (chỉ hiển thị, không ghi báo cáo). 0 = tắt
DANG_PHI_AMOUNT = 5_000
DEFAULT_COLLECTION_QUARTER = 1   # Quý đang thu (1 = Q1)

# ── Quy tắc phạt ────────────────────────────────────────────────────────────
# amount       : mức phạt cho 1 lần vi phạm (0 = chưa đặt)
# max_multiple : nhận diện theo số tiền khi GD = amount × 1..max_multiple
#                (0 = không nhận diện theo số tiền, chỉ theo từ khóa)
PHAT_RULES = [
    {"col": 5,  "label": "Không viết BB họp", "amount": 0, "max_multiple": 0,
     "keywords": ["khong viet bb hop", "khong viet bien ban", "bb hop"]},
    {"col": 6,  "label": "Không mang sổ họp", "amount": 0, "max_multiple": 0,
     "keywords": ["khong mang so hop", "quen so hop", "so hop"]},
    {"col": 7,  "label": "Không hoàn thành nhiệm vụ học tập", "amount": 20_000, "max_multiple": 1,
     "keywords": ["khong hoan thanh nhiem vu hoc tap", "kqht", "khong dat duoc kha", "hoc tap"]},
    {"col": 8,  "label": "Bị nợ môn", "amount": 50_000, "max_multiple": 3,
     "keywords": ["no mon", "rot mon", "thi lai"]},
    {"col": 9,  "label": "Vắng họp", "amount": 0, "max_multiple": 0,
     "keywords": ["vang hop", "vang mat hop", "khong bao truoc"]},
    {"col": 10, "label": "Không đọc thông báo", "amount": 0, "max_multiple": 0,
     "keywords": ["khong doc thong bao", "khong thuc hien thong bao", "thong bao chi bo"]},
]
PHAT_TOTAL_COL       = 11
ALLOW_PHAT_AMOUNT_HINT = True
# Số tiền → cột phạt (tự sinh từ PHAT_RULES, xem apply_settings)
PHAT_AMOUNT_HINTS    = {50_000: 8, 100_000: 8, 150_000: 8}
# GD bằng đúng mức quỹ 1 tháng nhưng không ghi rõ tháng/quý → ghi vào cột phạt này
# (None = coi là quỹ của tháng giao dịch)
QUY_UNSPECIFIED_PHAT_COL = 7

# ── Từ khóa & gợi ý số tiền Quỹ/Đảng phí ───────────────────────────────────
QUY_KEYWORDS = [
    "tien quy", "nop quy", "quy dang", "dang phi",
    "dang phi va quy", "quy chi bo", "chuyen quy", "quy quy",
]
# Các mức thêm ngoài 1/3/12 tháng × mức chuẩn: số tiền → số tháng (1, 3 hoặc 12)
QUY_EXTRA_AMOUNT_HINTS = {69_000: 3}
# Giá trị → số tháng tương ứng (tự sinh, xem apply_settings)
QUY_AMOUNT_HINTS = {20_000: 1, 60_000: 3, 69_000: 3, 240_000: 12}

# ── Ngưỡng khớp tên ─────────────────────────────────────────────────────────
MIN_MATCH_SCORE    = 86
AMBIGUOUS_SCORE_GAP = 4

# ── Stop words lọc khi trích xuất tên người ─────────────────────────────────
STOP_WORDS = set("""
ngan hang tmcp ngoai thuong cong thuong dau trien nong nghiep ptnt
techcombank vietcombank bidv tpb agribank mb timo ntmk qr
service jsc chuyen tien khoan nhanh qua zalo momo transfer nop dang quy phi
ct tu toi tai bank bnk ft mbvcb
account remitter debit credit so but toan dien giai no mon khong bb hop so hop
""".split())


# ── Cài đặt mức thu chỉnh trên giao diện (settings.json) ────────────────────
_DEFAULT_SETTINGS = {
    "quy_month_amount":         QUY_STANDARD_MONTH_AMOUNT,
    "dang_phi_amount":          DANG_PHI_AMOUNT,
    "quy_extra_amount_hints":   [{"amount": a, "months": m} for a, m in QUY_EXTRA_AMOUNT_HINTS.items()],
    "phat_rules":               copy.deepcopy(PHAT_RULES),
    "quy_unspecified_phat_col": QUY_UNSPECIFIED_PHAT_COL,
}


def default_settings() -> dict:
    """Bản sao cài đặt mặc định (theo quy định cũ)."""
    return copy.deepcopy(_DEFAULT_SETTINGS)


def load_settings() -> dict:
    """Đọc settings.json, ghép với mặc định (thiếu khóa nào lấy mặc định khóa đó)."""
    settings = default_settings()
    if SETTINGS_PATH.exists():
        with open(SETTINGS_PATH, encoding="utf-8") as f:
            saved = json.load(f)
        settings.update({k: v for k, v in saved.items() if k in settings})
        # Giữ cột/từ khóa của quy tắc phạt theo cột, chỉ ghi đè trường đã lưu
        saved_rules = {r["col"]: r for r in saved.get("phat_rules", [])}
        settings["phat_rules"] = [
            {**rule, **saved_rules.get(rule["col"], {})} for rule in _DEFAULT_SETTINGS["phat_rules"]
        ]
    return settings


def save_settings(settings: dict) -> None:
    """Lưu cài đặt ra settings.json và áp dụng ngay."""
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
    apply_settings(settings)


def build_quy_amount_hints(month_amount: int, extra: list) -> dict:
    """Số tiền → số tháng: 1/3/12 tháng × mức chuẩn, cộng các mức thêm."""
    hints = {int(e["amount"]): int(e["months"]) for e in extra if int(e["amount"]) > 0}
    hints.update({month_amount: 1, month_amount * 3: 3, month_amount * 12: 12})
    return hints


def build_phat_amount_hints(rules: list) -> dict:
    """Số tiền → cột phạt, sinh từ mức phạt × 1..max_multiple của từng quy tắc."""
    hints = {}
    for rule in rules:
        amount = int(rule.get("amount") or 0)
        for k in range(1, int(rule.get("max_multiple") or 0) + 1):
            if amount > 0:
                hints.setdefault(amount * k, rule["col"])
    return hints


def apply_settings(settings: dict) -> None:
    """Cập nhật các hằng số module theo cài đặt (các module đọc qua config.X)."""
    global QUY_STANDARD_MONTH_AMOUNT, QUY_AMOUNT_HINTS, DANG_PHI_AMOUNT
    global PHAT_RULES, PHAT_AMOUNT_HINTS, QUY_UNSPECIFIED_PHAT_COL
    QUY_STANDARD_MONTH_AMOUNT = int(settings["quy_month_amount"])
    DANG_PHI_AMOUNT = int(settings["dang_phi_amount"])
    QUY_AMOUNT_HINTS = build_quy_amount_hints(
        QUY_STANDARD_MONTH_AMOUNT, settings["quy_extra_amount_hints"]
    )
    PHAT_RULES = copy.deepcopy(settings["phat_rules"])
    PHAT_AMOUNT_HINTS = build_phat_amount_hints(PHAT_RULES)
    QUY_UNSPECIFIED_PHAT_COL = settings["quy_unspecified_phat_col"]


apply_settings(load_settings())
