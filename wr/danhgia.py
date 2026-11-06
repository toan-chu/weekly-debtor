# -*- coding: utf-8 -*-
"""WR-13 — đánh giá khách hàng phải thu theo tiêu chí FIN ban hành (file 'Đánh giá khách hàng phải thu và tiêu chí',
Head FIN + FIN chốt 05/11/2026).

- 3 mức theo SỐ NGÀY QUÁ HẠN LỚN NHẤT của khách: Bình thường (chưa quá hạn hoặc 1–30) · Cảnh báo (31–60) · Rủi ro cao (>60).
  Tính cả hoá đơn chỉ còn vài nghìn đồng (FIN 05/11).
- Nâng mức theo 3 ô Sales đánh giá mỗi tuần: thông tin bất lợi → Rủi ro cao; đang quá hạn mà thất hứa / trễ nhiều lần →
  Rủi ro cao; đang quá hạn mà không phản hồi → Cảnh báo.
- 3 tiêu chí số liệu (TC1 nợ quá hạn ≥ X, TC2 dư nợ > hạn mức, TC3 khách mới / chưa có HĐ nợ ≥ Y) CHƯA áp: FIN chọn chờ có
  lãi gộp và Client Master (WR-14).
- Nhóm khách Mới / Thường xuyên / Lâu năm (tự xếp) × mức → Phương án xử lý (bảng 'mức độ' của FIN) và Hình thức thanh
  toán đề xuất. Điểm lịch sử thanh toán A–D chỉ để tham khảo.

WR-15 (Head FIN, file 'W43 tiêu chí nâng hạng rủi ro v3', user 06/11): MÀU / MỨC CUỐI 0–4 thay 3 mức ở trên:
- Mức gốc theo nhóm tuổi CŨ NHẤT còn tiền: 0 Trắng chưa quá hạn · 1 Vàng 1–30 · 2 Cam 31–60 · 3 Đỏ 61–90 · 4 Tím ≥ 91.
- +1 mức (không cộng dồn) khi có ít nhất một: T1 quá hạn ≥ X · T2 dư nợ > hạn mức · T3 ≥ N hoá đơn trả trễ > 15 ngày trong
  6 tháng · T4 chưa có hợp đồng / chưa có credit term / khách mới nợ ≥ Y — trừ khách ✓ "Không nâng hạng màu".
- Lên thẳng ít nhất Đỏ: lý do có mất liên lạc / hứa không trả…, ô Sales bất lợi = Có, hoặc đang quá hạn + thất hứa = Có.
  Đang quá hạn + không phản hồi → ít nhất Cam.
- 3 mức FIN ở trên vẫn dùng cho Phương án xử lý / Hình thức thanh toán: Trắng, Vàng → 1 · Cam → 2 · Đỏ, Tím → 3.
Tham số (lãi gộp, tỷ lệ, Y, N…) FIN sửa mỗi tháng ở sheet "Tham số" của file credit term (khachhang.doc_tham_so).
Cùng luật với công thức ở sheet AR Risk (ghifull.cong_thuc_fin) và dashboard (app.js) — sửa một chỗ thì sửa cả ba."""
from __future__ import annotations

import datetime as dt
import math
from typing import Dict, List, Optional, Sequence, Tuple

MUC = ("Bình thường", "Cảnh báo", "Rủi ro cao")      # mã 1, 2, 3
NGUONG_BINH_THUONG = 30                               # quá hạn tối đa còn Bình thường (ngày)
NGUONG_CANH_BAO = 60                                  # quá hạn tối đa còn Cảnh báo (ngày)
THANG_KHACH_MOI = 6                                   # hoá đơn đầu tiên trong N tháng gần nhất = khách Mới
NHOM = ("Mới", "Thường xuyên", "Lâu năm")

# Nguyên văn sheet 'mức độ' của FIN (cột Phương án xử lý)
PHUONG_AN = {
    ("Mới", 1): "Thanh toán trước khi phát sinh lô hàng; nếu đã phát sinh công nợ thì thu hồi gấp. Không phân biệt giá trị lớn hay nhỏ.",
    ("Mới", 2): "Liên hệ xác nhận nguyên nhân chậm trả + yêu cầu lịch thanh toán cụ thể. Hợp đồng còn hiệu lực, đủ hồ sơ; nếu sắp hết hạn hoặc thiếu chứng từ thì bổ sung gấp. Rà soát background (thông tin minh bạch, hoạt động ổn định).",
    ("Mới", 3): "Xiết chặt theo dõi công nợ rủi ro. Yêu cầu thanh toán trước 50% hoặc 100% trước khi tiếp tục cung cấp dịch vụ mới. Nếu có nguy cơ không thu hồi được thì ngừng hẳn dịch vụ mới.",
    ("Thường xuyên", 1): "KH 7 ngày hoặc 15/20 ngày: nếu chưa đến hạn thì theo dõi và đôn đốc thu hồi.",
    ("Thường xuyên", 2): "Liên hệ xác nhận nguyên nhân + yêu cầu lịch thanh toán cụ thể. Nếu KH vẫn hoạt động nhưng hay chậm, liên hệ có phản hồi, có khả năng trả: chốt ngày thanh toán cụ thể, theo dõi đến khi tiền về tài khoản.",
    ("Thường xuyên", 3): "(1) Xác minh lại tình trạng doanh nghiệp, đánh giá khả năng thu hồi, báo cáo gấp cấp trên. (2) Yêu cầu cam kết thanh toán bằng văn bản; dừng cung cấp dịch vụ. (3) Rà soát hợp đồng, bổ sung hồ sơ, xác định căn cứ và quyền xử lý công nợ. (4) Chuyển hồ sơ pháp lý/kiện tụng nếu cần. (5) Lập dự phòng nợ khó đòi.",
    ("Lâu năm", 1): "KH có hợp đồng 30 ngày: chưa tới hạn thì theo dõi thường xuyên, nhất là công nợ phát sinh giá trị lớn.",
    ("Lâu năm", 2): "Liên hệ xác nhận nguyên nhân + yêu cầu lịch thanh toán cụ thể. Nếu KH vẫn hoạt động nhưng hay chậm: chốt ngày thanh toán cụ thể, theo dõi đến khi tiền về tài khoản.",
    ("Lâu năm", 3): "(1) Xác minh lại tình trạng doanh nghiệp, đánh giá khả năng thu hồi, báo cáo gấp cấp trên. (2) Yêu cầu cam kết thanh toán bằng văn bản; dừng cung cấp dịch vụ. (3) Rà soát hợp đồng, bổ sung hồ sơ, xác định căn cứ và quyền xử lý công nợ. (4) Chuyển hồ sơ pháp lý/kiện tụng nếu cần. (5) Lập dự phòng nợ khó đòi.",
}
KHONG_PHAN_HOI = ("KHÔNG PHẢN HỒI / DẤU HIỆU BẤT THƯỜNG: gửi email nhắc nợ chính thức; đồng thời báo cáo công nợ có nguy cơ "
                  "rủi ro; xác minh, tìm cách gặp trực tiếp KH, cam kết lịch thanh toán.")
HINH_THUC = {3: ("Thanh toán trước 50% hoặc 100%", "Dừng dịch vụ mới; chỉ thanh toán trước"),
             2: "Cấp tín dụng hạn chế, theo dõi sát", 1: "Tín dụng theo hạn mức"}

# Bảng chấm điểm lịch sử thanh toán (sheet 'hệ số hạn mức' của FIN) — tham khảo, không ảnh hưởng mức
DIEM_NGAY = ((0, 40), (NGUONG_BINH_THUONG, 30), (NGUONG_CANH_BAO, 15))      # ≤ ngưỡng → điểm; còn lại 0
DIEM_THU_HOI = ((0.95, 30), (0.8, 20), (0.5, 10))                          # ≥ ngưỡng → điểm
DIEM_QUA_HAN = ((0, 20), (0.3, 15), (0.6, 8))                              # ≤ ngưỡng → điểm
DIEM_TRA_DU = ((0.9, 10), (0.7, 6), (0.5, 3))                              # ≥ ngưỡng → điểm
HANG = (("A", 80), ("B", 60), ("C", 40), ("D", 0))


def _cong_thang(d: dt.date, n: int) -> dt.date:
    import calendar
    m = d.month - 1 + n
    y, m = d.year + m // 12, m % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def nhom_khach(ngay_dau: Optional[dt.date], ngay_chot: dt.date) -> str:
    """Như cột 'Nhóm KH áp dụng' của FIN: chưa có hoá đơn → Mới; hoá đơn đầu trước 01/01 năm nay → Lâu năm; hoá đơn đầu
    trong 6 tháng gần nhất → Mới; còn lại → Thường xuyên."""
    if ngay_dau is None:
        return "Mới"
    if ngay_dau < dt.date(ngay_chot.year, 1, 1):
        return "Lâu năm"
    if ngay_dau > _cong_thang(ngay_chot, -THANG_KHACH_MOI):
        return "Mới"
    return "Thường xuyên"


def muc_theo_ngay(ngay_qh: int) -> int:
    return 3 if ngay_qh > NGUONG_CANH_BAO else 2 if ngay_qh > NGUONG_BINH_THUONG else 1


def muc_cuoi(ngay_qh: int, danh_gia: Sequence[str]) -> int:
    """danh_gia = [thông tin bất lợi, khách có phản hồi, thất hứa / trễ nhiều lần] ("Có"/"Không")."""
    bat_loi, phan_hoi, that_hua = (list(danh_gia) + ["", "", ""])[:3]
    m = muc_theo_ngay(ngay_qh)
    if bat_loi == "Có" or (ngay_qh > 0 and that_hua == "Có"):
        m = max(m, 3)
    if ngay_qh > 0 and phan_hoi == "Không":
        m = max(m, 2)
    return m


def can_cu(ngay_qh: int, danh_gia: Sequence[str]) -> str:
    bat_loi, phan_hoi, that_hua = (list(danh_gia) + ["", "", ""])[:3]
    ra = [f"Quá hạn {ngay_qh} ngày." if ngay_qh > 0 else "Chưa quá hạn."]
    if bat_loi == "Có":
        ra.append("Thông tin bất lợi.")
    if that_hua == "Có" and ngay_qh > 0:
        ra.append("Trễ hạn nhiều lần/thất hứa.")
    if phan_hoi == "Không" and ngay_qh > 0:
        ra.append("Không phản hồi.")
    return " ".join(ra)


def hinh_thuc(muc: int, nhom: str) -> str:
    if muc == 3:
        return HINH_THUC[3][0] if nhom == "Mới" else HINH_THUC[3][1]
    return HINH_THUC[muc]


def phuong_an(muc: int, nhom: str, ngay_qh: int, danh_gia: Sequence[str]) -> str:
    phan_hoi = (list(danh_gia) + ["", "", ""])[1]
    t = PHUONG_AN[(nhom, muc)]
    if phan_hoi == "Không" and ngay_qh > 0:
        t += "\n" + KHONG_PHAN_HOI
    return t


def _bac(v: float, bang, lon_hon: bool) -> int:
    for nguong, diem in bang:
        if (v >= nguong) if lon_hon else (v <= nguong):
            return diem
    return 0


def cham_diem(ngay_qh: int, ty_le_thu_hoi: float, ty_le_qua_han: float, ty_le_tra_du: float) -> Tuple[int, str]:
    """Điểm lịch sử thanh toán (0–100) và hạng A–D theo bảng của FIN (tham khảo)."""
    d = (_bac(ngay_qh, DIEM_NGAY, False) + _bac(ty_le_thu_hoi, DIEM_THU_HOI, True)
         + _bac(ty_le_qua_han, DIEM_QUA_HAN, False) + _bac(ty_le_tra_du, DIEM_TRA_DU, True))
    return d, next(h for h, toi_thieu in HANG if d >= toi_thieu)


# --- WR-15: 5 mức màu của Head FIN -------------------------------------------------------------------------------
MUC5 = ("Trắng", "Vàng", "Cam", "Đỏ", "Tím")
# Nâng thẳng lên Đỏ theo lý do. Không dùng "không trả" trơn — bắt nhầm câu như "khách không trả kịp vì bank" (W43).
TU_KHOA_NANG_THANG = ("mất liên lạc", "không liên lạc được", "hứa không trả", "từ chối thanh toán",
                      "từ chối trả", "không chịu thanh toán", "không chịu trả", "bỏ trốn")
THAM_SO_MAC_DINH: Dict[str, object] = {
    "quy": "",                 # VD Q3/2026 — chỉ để ghi chú
    "lai_gop": [None, None, None],   # lãi gộp 3 tháng của quý gần nhất (VND)
    "ty_le_lai_gop": 0.25,     # X1 = 25% × lãi gộp bình quân / tháng
    "top_qua_han": 0.15,       # X2 = nợ quá hạn của khách cuối trong top 15% (làm tròn lên)
    "y": 50_000_000,           # T4: khách mới nợ ≥ Y (Head FIN: giả định tạm). Trống → Y = X / 2
    "he_so_an_toan": 1.3,      # T2: hạn mức tạm tính = billing bình quân × hệ số term × 1,3
    "so_hd_tre": 2,            # T3: từ N hoá đơn trả trễ
    "ngay_tre": 15,            # trễ hơn hạn quá 15 ngày
    "so_thang": 6,             # xét 6 tháng gần nhất (T3 và billing bình quân)
}
# Bảng hệ số điều khoản (tháng) — sheet Tham_số của Head FIN; FIN thêm dòng cho kiểu term mới
HE_SO_TERM_MAC_DINH: List[Tuple[str, float, str]] = [
    ("30 days", 1.0, "30 ngày từ ngày hoá đơn = 1"),
    ("45 days", 1.5, "45 ngày = 1,5"),
    ("60 days", 2.0, "60 ngày = 2"),
    ("20 days", 20 / 30, "Số ngày / 30"),
    ("15 days", 0.5, "Số ngày / 30"),
    ("14 days", 14 / 30, "Số ngày / 30"),
    ("7 days", 7 / 30, "Số ngày / 30"),
    ("thanh toán ngay", 0.0, "0 tháng → hạn mức tạm tính = 0, mọi dư nợ đều vượt"),
    ("15 ngày kể từ ngày gửi bảng công nợ", 1.0, "Gửi bảng công nợ cuối tháng: 15 ngày + 15 ngày bình quân"),
    ("14 ngày của tháng tiếp theo", 29 / 30, "Ngày 14 tháng sau ≈ 14 ngày từ cuối tháng + 15 ngày bình quân"),
    ("30 ngày kể từ ngày nhận đủ hồ sơ thanh toán", 1.0, "Tính như 30 ngày từ hoá đơn"),
]


def muc_goc5(nhom: Sequence[float], nguong: float = 0.5) -> int:
    """Nhóm tuổi CŨ NHẤT còn tiền: 0 chưa quá hạn · 1 (1–30) · 2 (31–60) · 3 (61–90) · 4 (91–120, 120+)."""
    if nhom[4] + nhom[5] > nguong:
        return 4
    for i in (3, 2, 1):
        if nhom[i] > nguong:
            return i
    return 0


def nguong_x(qua_han: Sequence[float], ts: Dict) -> Tuple[float, float, float]:
    """(X, X1, X2). X1 = tỷ lệ × lãi gộp bình quân các tháng đã nhập (0 nếu chưa nhập). X2 = nợ quá hạn của khách xếp
    cuối trong top 15% khách đang quá hạn (số khách làm tròn lên). X = MIN(X1, X2); chưa có lãi gộp thì X = X2."""
    lg = [float(v) for v in ts.get("lai_gop") or [] if isinstance(v, (int, float)) and v > 0]
    x1 = (sum(lg) / len(lg)) * float(ts.get("ty_le_lai_gop") or 0) if lg else 0.0
    ds = sorted((q for q in qua_han if q > 0.5), reverse=True)
    k = math.ceil(round(len(ds) * float(ts.get("top_qua_han") or 0), 9)) if ds else 0
    x2 = ds[k - 1] if k else 0.0
    return (min(x1, x2) if x1 > 0 and x2 > 0 else (x1 or x2)), round(x1, 2), round(x2, 2)


def he_so_term(term: str, bang: Sequence[Tuple[str, float, str]], han_cua=None) -> Optional[float]:
    """Số tháng của điều khoản thanh toán: tra bảng (đúng chữ, không phân biệt hoa thường); không có trong bảng thì
    tự tính = số ngày từ hoá đơn ngày 15 tới hạn / 30. Chưa có term → None."""
    t = " ".join(str(term or "").split()).lower()
    if not t:
        return None
    for ten, hs, *_ in bang:
        if " ".join(str(ten).split()).lower() == t:
            return float(hs)
    if han_cua is None:
        return None
    goc = dt.date(2026, 1, 15)
    return round((han_cua(goc, term) - goc).days / 30, 4)


def nang_thang(ly_do: str, danh_gia: Sequence[str], dang_qua_han: bool) -> bool:
    bat_loi, _phan_hoi, that_hua = (list(danh_gia) + ["", "", ""])[:3]
    c = (ly_do or "").lower()
    return bat_loi == "Có" or (dang_qua_han and that_hua == "Có") or any(k in c for k in TU_KHOA_NANG_THANG)


def muc_cuoi5(goc: int, nang: bool, thang: bool, dang_qua_han: bool, danh_gia: Sequence[str]) -> int:
    phan_hoi = (list(danh_gia) + ["", "", ""])[1]
    m = max(goc, 3) if thang else goc + (1 if nang else 0)
    if dang_qua_han and phan_hoi == "Không":
        m = max(m, 2)
    return min(4, m)


def muc3_tu_muc5(m5: int) -> int:
    """5 mức màu → 3 mức FIN cho Phương án xử lý / Hình thức thanh toán."""
    return 3 if m5 >= 3 else 2 if m5 == 2 else 1
