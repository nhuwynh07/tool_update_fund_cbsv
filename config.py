"""
config.py – Tất cả hằng số và cấu hình cho hệ thống đối soát chi bộ.
Chỉnh sửa file này để thay đổi đường dẫn, cột, ngưỡng khớp, v.v.
"""
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

# ── Đường dẫn file ──────────────────────────────────────────────────────────
SAO_KE_PATH   = BASE_DIR / "SaoKeTK_02022026_02052026.xlsx"
TONG_HOP_PATH = BASE_DIR / "2026.01.CBSV1 BC Dang phi.xlsx"
OUTPUT_PATH   = BASE_DIR / "2026.01.CBSV1 BC Dang phi_UPDATED.xlsx"

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
DEFAULT_COLLECTION_QUARTER = 1   # Quý đang thu (1 = Q1)

# ── Quy tắc phạt ────────────────────────────────────────────────────────────
PHAT_RULES = [
    {"col": 5,  "label": "Không viết BB họp",
     "keywords": ["khong viet bb hop", "khong viet bien ban", "bb hop"]},
    {"col": 6,  "label": "Không mang sổ họp",
     "keywords": ["khong mang so hop", "quen so hop", "so hop"]},
    {"col": 7,  "label": "Không hoàn thành nhiệm vụ học tập",
     "keywords": ["khong hoan thanh nhiem vu hoc tap", "kqht", "khong dat duoc kha", "hoc tap"]},
    {"col": 8,  "label": "Bị nợ môn",
     "keywords": ["no mon", "rot mon", "thi lai"]},
    {"col": 9,  "label": "Vắng họp",
     "keywords": ["vang hop", "vang mat hop", "khong bao truoc"]},
    {"col": 10, "label": "Không đọc thông báo",
     "keywords": ["khong doc thong bao", "khong thuc hien thong bao", "thong bao chi bo"]},
]
PHAT_TOTAL_COL       = 11
ALLOW_PHAT_AMOUNT_HINT = True
PHAT_AMOUNT_HINTS    = {50_000: 8, 100_000: 8, 150_000: 8}

# ── Từ khóa & gợi ý số tiền Quỹ/Đảng phí ───────────────────────────────────
QUY_KEYWORDS = [
    "tien quy", "nop quy", "quy dang", "dang phi",
    "dang phi va quy", "quy chi bo", "chuyen quy", "quy quy",
]
# Giá trị → số tháng tương ứng
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
