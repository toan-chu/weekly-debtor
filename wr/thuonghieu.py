# -*- coding: utf-8 -*-
"""Màu và phông của thương hiệu Trustana, dùng chung cho mọi file Excel tool sinh ra."""

TIM = "4D148C"      # cột số liệu công nợ
CAM = "FF6200"      # cột định danh khách hàng
TRANG = "FFFFFF"
XAM_NHAT = "F2F2F2"
VANG_CANH_BAO = "FFF2CC"
DO_CANH_BAO = "FCE4E4"

# Chữ trên báo cáo — báo cáo có lúc gửi đối tác nước ngoài nên để tiếng Anh.
# Sửa chữ ở đây là đủ, tool ghi lại các dòng này mỗi lần chạy.
CONG_TY = "TRUSTANA VIETNAM COMPANY LIMITED"
TEN_BAO_CAO = "Overdue Debtors Reports"
TEN_BANG = "Accounts receivables"
SHEET_BANG = "Receivables"
SHEET_REVIEW = "Review"
# v0.4 — các sheet còn lại của full credit report (dashboard WR-05 đọc theo đúng tên này)
SHEET_SUMMARY = "Summary"
SHEET_RISK = "AR Risk"
SHEET_PAYABLE = "Payable"
SHEET_CASH = "Cash Flow"
SHEET_METHOD = "Methodology"
SHEET_INVOICES = "Invoices"
SHEET_DANH_GIA = "Đánh giá"    # WR-15: bảng nâng hạng từng khách như file v3 của Head FIN
SHEET_THAM_SO = "Tham số"      # WR-15: bản chụp tham số tuần đó (sửa ở file credit term)
SHEET_CONTRACT = "Contracts"   # WR-08: cảnh báo hợp đồng / credit term (Client Master)
TEN_BAO_CAO_TRA = "Creditors Report"
TEN_BANG_TRA = "Accounts payable"

PHONG_TIEU_DE = "Roboto"   # heading
PHONG_NOI_DUNG = "Calibri" # body

# Cột định danh (nền cam) và cột tiền (nền tím) của sheet Receivable
COT = [
    ("No.", "ten", 6),
    ("Code", "ten", 16),
    ("Customer's Name", "ten", 46),
    ("Risk level", "ten", 13),                     # WR-13 (FIN 05/11): ô Mức độ cạnh tên khách — công thức từ AR Risk
    ("Credit Term", "ten", 14),
    ("Total", "tien", 16),
    ("Current", "tien", 15),
    ("1 - 30", "tien", 15),
    ("31 - 60", "tien", 15),
    ("61 - 90", "tien", 15),
    ("91 - 120", "tien", 15),
    ("120+", "tien", 15),
    ("Salesman", "ten", 24),
    ("Job Number", "ten", 30),
    ("Reason last week (read only)", "ten", 40),   # WR-12 (user 29/10): tuần trước bên TRÁI, tuần này bên phải
    ("Reason for late payment", "ten", 60),
    # WR-13 (FIN 05/11): Sales đánh giá mỗi tuần, tuần mới giữ giá trị tuần trước
    ("Adverse info? (Sales)", "ten", 11),
    ("Customer responds? (Sales)", "ten", 11),
    ("Broken promises / often late? (Sales)", "ten", 13),
]


def c(ten: str) -> int:
    """Số thứ tự cột (bắt đầu 1) của cột tên `ten` ở sheet Receivables — đừng viết số cột cứng."""
    return [t for t, _, _ in COT].index(ten) + 1


C_TEN, C_MUC, C_TERM, C_TOTAL = c("Customer's Name"), c("Risk level"), c("Credit Term"), c("Total")
C_NHOM0, C_NHOM5 = c("Current"), c("120+")
C_SALES, C_JOB = c("Salesman"), c("Job Number")
COT_LY_DO_TRUOC_SO = c("Reason last week (read only)")
COT_LY_DO_NAY = c("Reason for late payment")
COT_DANH_GIA = [c("Adverse info? (Sales)"), c("Customer responds? (Sales)"), c("Broken promises / often late? (Sales)")]
# giá trị mặc định khi chưa ai chọn (như file FIN): không có thông tin bất lợi · khách có phản hồi · không thất hứa
DANH_GIA_MAC_DINH = ("Không", "Có", "Không")
# WR-15 (Head FIN, file W43 tiêu chí nâng hạng v3, 06/11): 5 mức màu 0–4, CHỈ tô ô Risk level (user: tô cả hàng đau mắt)
MUC_FIN = ("Trắng", "Vàng", "Cam", "Đỏ", "Tím")
MAU_MUC_FIN = {"Trắng": ("FFFFFF", "000000"), "Vàng": ("FFEB9C", "000000"), "Cam": ("F8CBAD", "000000"),
               "Đỏ": ("FF7C80", "000000"), "Tím": ("7030A0", "FFFFFF")}

# v0.5 (WR-06): cột thêm khi có file Bán hàng — dư nợ Misa so với tổng job chưa thanh toán
COT_DOI_CHIEU = ("Đối chiếu job (FIN)", 46)

# WR-11 PA1: lý do Sales ghi ở báo cáo tuần trước — chỉ để đọc (nền xám). Cột "Reason for late payment"
# để trống mỗi tuần mới cho Sales ghi lý do tuần này.
COT_LY_DO_TRUOC = ("Reason last week (read only)", 40)

# Dòng 1 công ty · 2 tên báo cáo · 3 ngày chốt · 4 tên bảng · 5 tên cột
DONG_TIEU_DE = 5      # dòng chứa tên cột
DONG_DAU_DU_LIEU = 6  # dòng dữ liệu đầu tiên


# --- Tên Salesman: CEO chốt 08/10 dùng TÊN ĐẦY ĐỦ như trên SMS -------------
# Báo cáo cũ của kế toán trưởng dùng tên ngắn; tool đổi sang tên đầy đủ theo bảng này.
# Đối chiếu từ W40: cùng khách, bản KTT ghi tên ngắn, SMS ghi tên đầy đủ.
# Có nhân viên mới thì thêm một dòng ở đây.
TEN_SALES_DAY_DU = {
    "hiếu": "Trần Văn Hiếu",
    "hiếu 1": "Phạm Trần Hiếu",
    "hoài": "Phạm Thị Thương Hoài",
    "trân": "Võ Lê Huyền Trân",
    "hằng": "Lê Thị Hằng",
    "hanna": "Ngô Thúy Hằng",
    "thư": "Kim Thanh Thư",
    "tân": "Nguyễn Duy Tân",
}

# Bản cũ (tới v0.6.1) tự viết lý do có nhãn này khi Sales chưa ghi. Từ WR-11 không viết nữa (user 27/10:
# "xoá toàn bộ chữ Tự động"); vẫn giữ nhãn để đọc báo cáo cũ thì coi câu đó như trống.
NHAN_TU_DONG = "[TỰ ĐỘNG]"


def ten_sales(ten: str) -> str:
    """'Admin,Admin' → 'Admin'; 'Hiếu' → 'Trần Văn Hiếu'; tên đầy đủ giữ nguyên."""
    import unicodedata
    phan = []
    for x in str(ten or "").split(","):
        x = unicodedata.normalize("NFC", x).strip()
        if not x or x in ("+",):
            continue
        x = TEN_SALES_DAY_DU.get(x.lower(), x)
        if x not in phan:
            phan.append(x)
    return ", ".join(phan)



def nen_cf(mau: str):
    """Nền cho định dạng có điều kiện (Conditional Formatting). Excel đọc màu nền của CF ở bgColor, không phải fgColor
    như ô thường — chỉ ghi fgColor thì Excel không tô gì (lỗi W43 chạy lại 29/10). Ghi cả hai cho Excel lẫn LibreOffice."""
    from openpyxl.styles import PatternFill
    return PatternFill(fill_type="solid", fgColor=mau, bgColor=mau)
