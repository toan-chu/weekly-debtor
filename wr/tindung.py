# -*- coding: utf-8 -*-
"""v0.4 — full credit report (WR-04).

Một lần chạy tính đủ các phần mà file Excel tuần cần:
  ① Receivables  tiền theo Tổng hợp 131, tuổi theo hoá đơn trong Chi tiết 131
  ② AR Risk      luật rủi ro của chị kế toán trưởng (sheet "data W39 so W40" + "Tham chieu")
  ③ Payable      Tổng hợp 331
  ④ Cash Flow    phiếu thu trong kỳ tách theo TK đối ứng + tiền đã trả vendor
  ⑤ Summary      các con số tuần này, lưu vào sổ ghi để làm lịch sử nhiều tuần
Luật chốt: handoff/20261008_WR-04_full-credit-report.md, handoff/MAP.md.
"""
from __future__ import annotations

import calendar
import datetime as dt
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from . import chitiet as ctm
from . import ghep
from . import soghi as sg
from . import thuonghieu as th_
from .misa import PhaiTra, TongHop, chuan_hoa_ten
from .tinhtoan import DongBaoCao, _cac_job, _chenh_ty_gia

# --- tham số của chị KTT (sheet "Tham chieu") ------------------------------
SO_DAU_HIEU_NO_XAU = 2
NGUONG_CEO = 100_000_000
NGUONG_DU_LE = 1_000              # VND: dư nợ nhỏ hơn thế là số lẻ còn sót, không đưa vào báo cáo
TU_KHOA_LOAI_TRU = ("cấn trừ", "bù trừ", "thanh toán hộ", "supplier của mình")
TU_KHOA_PHAP_LY = ("đang kiện", "mất liên lạc", "khó đòi", "công nợ xấu", "phá sản")

N0, NX, N4, N3, N2, N1 = ("0·Chưa đến hạn", "X·Không phải rủi ro tín dụng",
                          "N4·Nợ xấu–sự kiện pháp lý", "N3·Nợ xấu", "N2·Cần theo dõi",
                          "N1·Quá hạn bình thường")


@dataclass
class DongRuiRo:
    ma: str
    ten: str
    credit_term: str
    dau_ky: float          # Misa: dư đầu kỳ
    tang: float            # Misa: phát sinh Nợ (nợ mới)
    giam: float            # Misa: phát sinh Có (đã thu, cấn trừ, tỷ giá)
    cuoi_ky: float         # = Total của Receivables
    nhom: List[float]
    salesman: str
    job: str
    ly_do: str
    tt_total: float        # tuần trước (U)
    tt_qua_han: float      # (V)
    tt_61: float           # (W)
    co_tuan_truoc: bool
    nhap_bo_sung: Optional[float]   # dư đầu kỳ Misa − dư cuối báo cáo tuần trước
    s: List[int] = field(default_factory=lambda: [0] * 6)
    so_dau_hieu: int = 0
    ma_dau_hieu: str = ""
    nhom_rui_ro: str = ""
    xu_huong: str = ""
    viec_can_lam: str = ""
    ceo: bool = False
    chia_sales: str = ""   # WR-08 (CX): khách nhiều Sales — job nào của Sales nào
    ngay_hen: Optional[dt.date] = None   # ngày Sales hẹn khách trả, đọc từ lý do
    s7: int = 0                          # S7 lỡ hẹn: đã qua ngày hẹn mà quá hạn không giảm
    # WR-11: lý do tuần trước (PA1) · file credit term
    ly_do_truoc: str = ""
    han_muc: Optional[float] = None      # file credit term: Hạn mức nợ (VND) — TC2 của FIN chờ WR-14
    co_hd: Optional[bool] = None         # file credit term: Đã có hợp đồng (None = file chưa có cột này)
    link_hd: str = ""                    # WR-12: link file hợp đồng (PDF) — ô "Mở hợp đồng" ở AR Risk
    # WR-13: đánh giá theo tiêu chí FIN (wr/danhgia.py)
    danh_gia: List[str] = field(default_factory=lambda: list(th_.DANH_GIA_MAC_DINH))  # bất lợi, phản hồi, thất hứa
    ngay_qh: int = 0                     # số ngày quá hạn lớn nhất
    ngay_hd_dau: Optional[dt.date] = None
    nhom_kh: str = ""                    # Mới / Thường xuyên / Lâu năm
    muc_ngay: int = 1                    # mức theo ngày quá hạn (1–3)
    muc_fin: int = 1                     # mức cuối sau nâng (1–3)
    can_cu: str = ""
    hinh_thuc: str = ""
    phuong_an: str = ""
    cham: Dict[str, float] = field(default_factory=dict)   # so_hd, ty_le_thu_hoi, ty_le_qua_han, ty_le_tra_du
    diem: int = 0
    hang: str = ""
    # WR-15: 5 mức màu của Head FIN (danhgia.muc_cuoi5)
    khong_nang: bool = False             # file credit term: ✓ Không nâng hạng màu
    so_hd_tre: int = 0                   # T3: hoá đơn trả trễ trong N tháng (tool đếm)
    billing: float = 0.0                 # billing bình quân / tháng trong N tháng
    he_so: Optional[float] = None        # hệ số điều khoản (tháng)
    han_muc_ap: Optional[float] = None   # hạn mức áp dụng: hợp đồng, không có thì tạm tính
    nguon_han_muc: str = ""              # Hợp đồng / Tạm tính / Thiếu dữ liệu
    muc_goc: int = 0                     # 0–4 theo tuổi nợ
    tieu_chi: List[int] = field(default_factory=lambda: [0, 0, 0, 0])   # T1..T4
    nang: int = 0                        # +1 mức
    nang_thang: int = 0                  # lên thẳng ít nhất Đỏ
    muc5: int = 0                        # mức cuối 0–4
    theo_doi: str = ""                   # khách trong hạn có dấu hiệu T1–T4: "↑T2 T3" (không đổi màu)
    can_cu_goc: str = ""                 # phần căn cứ tool tính (tuổi nợ, T1–T4) — Excel nối thêm phần ô Sales

    @property
    def ly_do_hieu_luc(self) -> str:
        """Lý do tuần này; Sales chưa ghi thì dùng lý do tuần trước để xét nhóm (pháp lý, cấn trừ, ngày hẹn)."""
        return self.ly_do if self.ly_do.strip() else self.ly_do_truoc

    @property
    def ten_muc(self) -> str:
        return th_.MUC_FIN[self.muc5] + (f" {self.theo_doi}" if self.muc5 == 0 and self.theo_doi else "")

    @property
    def qua_han(self) -> float:
        return round(sum(self.nhom[1:]), 2)


@dataclass
class DongTra:
    ma: str
    ten: str
    dau_ky: float
    no_moi: float
    da_tra: float
    cuoi_ky: float
    trang_thai: str
    nhap_bo_sung: Optional[float] = None


@dataclass
class DongThu:
    ma: str
    ten: str
    theo_nhom: Dict[str, float]

    @property
    def tong(self) -> float:
        return round(sum(self.theo_nhom.values()), 2)


@dataclass
class KetQuaFull:
    dong: List[DongBaoCao]
    rui_ro: List[DongRuiRo] = field(default_factory=list)
    phai_tra: Optional[List[DongTra]] = None
    thu: List[DongThu] = field(default_factory=list)
    tra_vendor: List[DongTra] = field(default_factory=list)
    chi_so: Dict[str, float] = field(default_factory=dict)
    chi_so_truoc: Dict[str, float] = field(default_factory=dict)
    lich_su: List[dict] = field(default_factory=list)
    da_bo: List[tuple] = field(default_factory=list)
    lech_chi_tiet: List[tuple] = field(default_factory=list)    # tên, Tổng hợp, Chi tiết
    khach_tra_xong: List[tuple] = field(default_factory=list)   # tên, dư tuần trước
    vendor_tra_truoc: List[tuple] = field(default_factory=list) # mã, tên, dư Nợ
    sms_cat: List[tuple] = field(default_factory=list)
    master: Dict[str, dict] = field(default_factory=dict)       # WR-11: file credit term đã đọc
    them_khach: str = ""                                        # WR-12: kết quả tự thêm khách mới vào file credit term
    x: float = 0.0                                              # WR-15: ngưỡng T1 = MIN(X1, X2)
    x1: float = 0.0
    x2: float = 0.0
    y: float = 0.0                                              # ngưỡng T4 khách mới
    tham_so: Dict[str, object] = field(default_factory=dict)   # sheet Tham số của file credit term
    co_tham_so: bool = False                                    # file credit term đã có sheet Tham số
    sms_khong_misa: List[tuple] = field(default_factory=list)
    sms_chua_ghep: List[tuple] = field(default_factory=list)
    # v0.5 (WR-06): đối chiếu job chưa thanh toán trên Misa
    job_misa: bool = False                                      # có file Bán hàng hay không
    het_no_con_hd: List[tuple] = field(default_factory=list)    # tên, số HĐ, tiền: Misa hết nợ, HĐ vẫn "chưa TT"
    ten_khong_ghep: List[tuple] = field(default_factory=list)   # tên, số HĐ, tiền: không tìm ra khách Misa


def _co(chu: str, tu_khoa) -> bool:
    c = (chu or "").lower()
    return any(t in c for t in tu_khoa)


def _ghep_sms(th: TongHop, dong_sms, bang):
    """Gán từng dòng SMS cho một mã khách Misa (bảng ghép > trùng mã > trùng tên)."""
    con_no = {k.ma for k in th.khach_con_no()}
    theo_ma = {k.ma.upper(): k.ma for k in th.dong.values()}
    theo_ten = {k.ten_chuan: k.ma for k in th.dong.values()}
    theo_khach: Dict[str, list] = {}
    khong_misa, chua_ghep = [], {}
    for x in dong_sms or []:
        ma, cach = ghep.ma_misa_cua(x, bang, theo_ma, theo_ten)
        if ma is not None:
            ma = theo_ma.get(ma.upper(), ma)
        if ma is None and cach == "chưa ghép":
            chua_ghep.setdefault(x.ma_sms, []).append(x)
        elif ma is None or ma not in con_no:
            ly_do = ("FIN khai không có trên Misa" if ma is None else
                     f"Misa không còn nợ ({ma}) — nhiều khả năng đã thu, SMS chưa gạch paid")
            khong_misa.append((x.ma_sms, x.ten_sms, x.job, x.so_tien, ly_do))
        else:
            theo_khach.setdefault(ma, []).append(x)
    gop: Dict[tuple, list] = {}
    for ma, ten, job, tien, ly_do in khong_misa:
        g = gop.setdefault((ma, job), [ma, ten, job, 0.0, ly_do])
        g[3] += tien
    chua = [(m, ds[0].ten_sms, sum(x.so_tien for x in ds), len(ds)) for m, ds in chua_ghep.items()]
    return theo_khach, [tuple(g) for g in gop.values()], chua


# --- S7 lỡ hẹn (user 13/10, chọn phương án A) ------------------------------------------
_RE_NGAY_THANG = re.compile(r"ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})(?:\s*(?:năm|/)\s*(\d{4}))?")
_RE_DMY = re.compile(r"(?<![\d.,/])(\d{1,2})[/.-](\d{1,2})(?:[/.-](\d{4}|\d{2}))?(?![.,]?\d)")
_RE_THANG = re.compile(r"(đầu|giữa|cuối)?\s*(?:tháng|\bt)\s*(\d{1,2})(?:\s*(?:năm|/)\s*(\d{4}))?(?!\d)")


def ngay_hen(ly_do: str, ngay_chot: dt.date) -> Optional[dt.date]:
    """Ngày hẹn trả muộn nhất ghi trong lý do: '15/9', 'ngày 15 tháng 9', 'tháng 9' (= cuối tháng),
    'đầu tháng 9' (= 10/9), 'giữa tháng 9' (= 20/9), 'T9'. Không ghi năm → năm của ngày chốt,
    lùi một năm nếu ra ngày xa hơn 6 tháng sau ngày chốt. Không thấy → None."""
    t = str(ly_do or "").lower()
    if not t.strip() or t.startswith(th_.NHAN_TU_DONG.lower()):
        return None
    ra: List[dt.date] = []

    def them(d, m, y, cuoi=False):
        try:
            d, m = int(d), int(m)
            if not 1 <= m <= 12:
                return
            nam = int(y) + (2000 if y and int(y) < 100 else 0) if y else ngay_chot.year
            if y is None and dt.date(nam, m, 1) > ngay_chot + dt.timedelta(days=183):
                nam -= 1                         # 'tháng 12' ghi hồi tháng 1 → tháng 12 năm trước
            elif y is None and dt.date(nam, m, 28) < ngay_chot - dt.timedelta(days=183):
                nam += 1                         # 'tháng 2' ghi hồi tháng 10 → tháng 2 năm sau
            ra.append(dt.date(nam, m, calendar.monthrange(nam, m)[1] if cuoi else d))
        except ValueError:
            pass

    for m in _RE_NGAY_THANG.finditer(t):
        them(m.group(1), m.group(2), m.group(3))
    t = _RE_NGAY_THANG.sub(" ", t)
    for m in _RE_DMY.finditer(t):
        them(m.group(1), m.group(2), m.group(3))
    t = _RE_DMY.sub(" ", t)
    for m in _RE_THANG.finditer(t):
        moc = {"đầu": 10, "giữa": 20}.get(m.group(1) or "")
        them(moc or 1, m.group(2), m.group(3), cuoi=not moc)
    return max(ra) if ra else None


def _xu_huong(r: DongRuiRo) -> str:
    """Câu tự sinh, chép đúng công thức cột 'Trend (Auto)' của chị KTT."""
    s1, s2, s3, s4, s5, s6 = r.s
    q, v, p = r.qua_han, r.tt_qua_han, r.nhom[5]
    c = []
    if q < 0.5:
        c.append("Chưa đến hạn thanh toán.")
    if not r.co_tuan_truoc or r.tt_total < 0.5:
        c.append("Khách mới xuất hiện trong bảng tuần này.")
    if s1:
        c.append("Khách phát sinh lô mới nhưng nợ cũ không giảm.")
    if s2:
        c.append("Quá hạn đứng yên so với tuần trước, chưa thu được đồng nào.")
    if r.s7 and r.ngay_hen:
        c.append(f"Đã quá ngày hẹn trả ({r.ngay_hen:%d/%m/%Y}) mà quá hạn chưa giảm.")
    if r.tt_total > 0.5 and 0.5 < q < v - 0.5:
        c.append(f"Đã thu bớt {v - q:,.0f} so với tuần trước.")
    if s3:
        c.append("Nợ đang già đi, tiền chuyển sang nhóm tuổi cao hơn.")
    if s6:
        c.append("Toàn bộ dư nợ đã quá hạn.")
    if p > 0.5:
        c.append(f"Có {p:,.0f} nằm ở nhóm trên 120 ngày.")
    if s5:
        c.append("Chưa có điều khoản thanh toán nên chưa xác định chắc chắn được quá hạn.")
    if s4:
        c.append("CHƯA CÓ GIẢI TRÌNH TỪ SALES.")
    return " ".join(c)


def _chia_tuoi_dong_con(d: DongBaoCao) -> None:
    """Chia sáu nhóm tuổi của khách cho các dòng con (user 08/10: số tiền job phải nằm đúng
    cột 1-30, 31-60...). Không có ngày hoá đơn cho từng job, nên phân bổ theo thứ tự: nợ cũ không
    có job nhận nhóm già nhất trước, rồi tới job cũ nhất. Cộng các dòng con theo từng cột luôn
    bằng đúng dòng khách."""
    con_lai = list(d.nhom)
    for g in d.dong_con:                  # dòng FIN khai theo hoá đơn: nhóm tuổi đã biết chắc
        if g.get("nhom_co_dinh"):
            g["nhom"] = g["nhom_co_dinh"]
            con_lai = [round(a - b, 2) for a, b in zip(con_lai, g["nhom"])]
    con = [g for g in d.dong_con if not g.get("nhom_co_dinh")]
    if d.can_doi > 0:
        con = [{"_can_doi": True, "tien": d.can_doi}] + sorted(con, key=lambda g: (g.get("ngay") or dt.date.min))
    else:
        con = sorted(con, key=lambda g: (g.get("ngay") or dt.date.min))
    for g in con:
        tien, nhom = g["tien"], [0.0] * 6
        for i in range(5, -1, -1):
            lay = min(max(con_lai[i], 0), tien)
            if lay > 0:
                nhom[i] = round(lay, 2)
                con_lai[i] = round(con_lai[i] - lay, 2)
                tien = round(tien - lay, 2)
        if tien > 0.5:                     # job SMS nhiều hơn phần còn lại (tỷ giá): để ở Current
            nhom[0] = round(nhom[0] + tien, 2)
        if g.get("_can_doi"):
            d.can_doi_nhom = nhom
        else:
            g["nhom"] = nhom


# --- v0.5 (WR-06): đối chiếu dư nợ Misa với tổng job chưa thanh toán ---------------------------
DUNG_SAI_KHOP = 1_000          # VND: làm tròn
TY_LE_LECH_NHO = 0.02          # lệch ≤ 2% dư nợ: nghi tỷ giá
MAU_DOI_CHIEU = {"khop": "Khớp", "do": "Đỏ", "vang": "Vàng", "cam": "Cam", "tim": "Tím"}


def _dong_con_misa(d: DongBaoCao, hds, term: str, tai_ngay: dt.date, sales_job: Dict[str, str],
                   sales_cu: str) -> None:
    """Mỗi job của hoá đơn chưa thanh toán (luật FIN: chỉ 'Đã thanh toán' mới bỏ) thành một dòng
    con. Biết ngày hoá đơn nên biết đúng nhóm tuổi của từng job."""
    for h in sorted(hds, key=lambda h: h.ngay):
        han = sg.han_tu_term(h.ngay, term)
        i = sg.nhom_tu_so_ngay((tai_ngay - han).days)
        for job, tien in h.jobs.items():
            nhom = [0.0] * 6
            nhom[i] = tien
            d.dong_con.append({"job": job or f"HĐ {h.so_hd} (chưa có job)",
                               "sales": sales_job.get(job.upper(), "") if job else "",
                               "ngay": h.ngay, "nguon_ngay": "HĐ", "tien": tien, "so_hd": h.so_hd,
                               "trang_thai": (h.trang_thai if h.chua_tt else
                                              f"{h.trang_thai} ngày {h.thu_ngay:%d/%m} (sau kỳ)"),
                               "nhom_co_dinh": nhom, "misa": True})
    if d.dong_con:
        d.job = ", ".join(_cac_job(", ".join(g["job"] for g in d.dong_con if not g["job"].startswith("HĐ "))))
        tien_sale: Dict[str, float] = {}
        for g in d.dong_con:
            if g["sales"]:
                tien_sale[g["sales"]] = tien_sale.get(g["sales"], 0) + g["tien"]
        if tien_sale:
            d.salesman = ", ".join(sorted(tien_sale, key=lambda t: -tien_sale[t]))
        elif sales_cu:
            d.salesman = sales_cu


def _dong_con_theo_so(d: DongBaoCao, job_misa, ma: str, khai_no_cu: list, sales_job: Dict[str, str],
                      sales_cu: str, tai_ngay: dt.date) -> None:
    """WR-08 (FIN sửa tay 9 khách ở W41): tiền mỗi dòng job = số CÒN NỢ của hoá đơn theo sổ Chi tiết 131
    (phiếu thu ghi số HĐ trừ đúng HĐ đó, còn lại trừ HĐ cũ trước) — không lấy nguyên tiền hoá đơn.
    Job của hoá đơn: Bán hàng + Sổ chi tiết (năm nay), bảng job nợ cũ (năm trước). Cộng các dòng = dư nợ."""
    theo_hd = {}
    for h in job_misa.theo_ma.get(ma, []):
        n = ctm._so_hd(h.so_hd)
        if n is not None:
            theo_hd[(n, h.ngay.year)] = h
    no_cu = {}
    for so_hd, job, sales in khai_no_cu:
        no_cu.setdefault(so_hd, (job, sales))
    for ngay, so_hd, han, tien in sorted(d.hoa_don, key=lambda x: (x[0] or dt.date.max)):
        if tien <= 0.5:
            continue
        nhom = [0.0] * 6
        i = sg.nhom_tu_so_ngay((tai_ngay - han).days)
        n = ctm._so_hd(str(so_hd)) if ngay else None
        h = theo_hd.get((n, ngay.year)) if n is not None else None
        g0 = {"ngay": ngay, "nguon_ngay": "HĐ", "so_hd": so_hd if ngay else "", "misa": True,
              "han": han}                       # WR-13 (FIN 05/11): hạn của job = hạn hoá đơn (ngày HĐ + credit term)
        if ngay is None:
            nhom[i] = tien
            d.dong_con.append({**g0, "job": "Chưa rõ hoá đơn (Tổng hợp 131 cao hơn Chi tiết)", "sales": "",
                               "tien": tien, "trang_thai": "", "nhom_co_dinh": nhom})
            continue
        if h is not None:
            tong = sum(h.jobs.values()) or 1.0
            phan = [(j, round(tien * t / tong, 2)) for j, t in h.jobs.items()]
            lech = round(tien - sum(p for _, p in phan), 2)
            if lech and phan:
                phan[-1] = (phan[-1][0], round(phan[-1][1] + lech, 2))
            tt = h.trang_thai if h.chua_tt else f"Misa ghi '{h.trang_thai}' nhưng sổ còn nợ"
            for job, p in phan:
                nh = [0.0] * 6
                nh[i] = p
                d.dong_con.append({**g0, "job": job or f"HĐ {so_hd} (chưa có job)",
                                   "sales": sales_job.get(job.upper(), "") if job else "",
                                   "tien": p, "trang_thai": tt, "nhom_co_dinh": nh})
            continue
        nhom[i] = tien
        if n is None:                     # số chứng từ không phải số hoá đơn (NVK…): bút toán điều chỉnh
            d.dong_con.append({**g0, "job": f"Điều chỉnh / chênh lệch tỷ giá ({so_hd})", "sales": "",
                               "tien": tien, "trang_thai": "Bút toán điều chỉnh trên sổ", "nhom_co_dinh": nhom})
            continue
        if ngay < job_misa.tu_ngay:
            job, sales = no_cu.get(n, ("", ""))
            d.dong_con.append({**g0, "job": job or f"HĐ {so_hd} (nợ cũ, chưa có job)", "sales": th_.ten_sales(sales),
                               "tien": tien, "trang_thai": "Nợ cũ – bảng FIN" if job else "Nợ cũ – FIN điền job",
                               "tham_chieu": bool(job), "nhom_co_dinh": nhom})
        else:
            d.dong_con.append({**g0, "job": f"HĐ {so_hd} (không thấy trên Bán hàng)", "sales": "",
                               "tien": tien, "trang_thai": "", "nhom_co_dinh": nhom})
    if d.dong_con:
        d.job = ", ".join(_cac_job(", ".join(g["job"] for g in d.dong_con if not g["job"].startswith(("HĐ ", "Chưa rõ", "Điều chỉnh")))))
        tien_sale: Dict[str, float] = {}
        for g in d.dong_con:
            if g["sales"]:
                tien_sale[g["sales"]] = tien_sale.get(g["sales"], 0) + g["tien"]
        if tien_sale:
            d.salesman = ", ".join(sorted(tien_sale, key=lambda t: -tien_sale[t]))
        elif sales_cu:
            d.salesman = sales_cu


def _tr(v: float) -> str:
    return f"{v / 1e9:,.2f} tỷ".replace(".00 tỷ", " tỷ") if abs(v) >= 1e9 else f"{v / 1e6:,.0f}tr" if abs(v) >= 1e6 \
        else f"{v:,.0f}"


def chia_theo_sales(dong_con: List[dict]) -> str:
    """CX 12/10: khách có từ 2 Sales trở lên thì ghi rõ job nào của Sales nào, để sếp đọc hiểu.
    'Trần Văn Hiếu (45tr): EXSANA…, EXHANA… | Lê Thị Hằng (12tr): TRAANA…'. Một Sales thì để trống."""
    nhom: Dict[str, list] = {}
    for g in dong_con:
        x = nhom.setdefault(g.get("sales") or "Không thấy job trên SMS", [0.0, []])
        x[0] += g["tien"]
        if g["job"] not in x[1]:
            x[1].append(g["job"])
    co_ten = [s for s in nhom if s != "Không thấy job trên SMS"]
    if len(co_ten) < 2:
        return ""
    return " | ".join(f"{s} ({_tr(v)}): {', '.join(j)}" for s, (v, j) in sorted(nhom.items(), key=lambda x: -x[1][0]))


def doi_chieu(T: float, dong_con: List[dict]) -> dict:
    """FIN 09/10: cộng các job chưa thanh toán, so với dư nợ của khách; lệch thì báo màu để FIN
    xem lại (quên đánh giá chênh lệch tỷ giá, thanh toán một phần mà Misa không tự trừ...)."""
    S = round(sum(g["tien"] for g in dong_con), 2)
    L = round(T - S, 2)
    mot_phan = sum(1 for g in dong_con if "một phần" in (g.get("trang_thai") or "").lower())
    if abs(L) <= DUNG_SAI_KHOP:
        mau, goi_y = "khop", ""
    elif not dong_con:
        mau, goi_y = "cam", ("Misa không ghi hoá đơn nào chưa thanh toán: nợ trước năm nay hoặc Misa ghi "
                             "'Đã thanh toán' mà sổ còn nợ. Dòng job đã tính theo sổ.")
    elif abs(L) <= TY_LE_LECH_NHO * max(T, S):
        mau, goi_y = "vang", "Lệch nhỏ: nghi quên đánh giá chênh lệch tỷ giá, hoặc làm tròn."
    elif L < 0:
        mau = "do"
        goi_y = ("Misa ghi chưa thanh toán NHIỀU hơn dư nợ: đã thu nhưng chưa đối trừ vào hoá đơn, "
                 "hoặc thanh toán một phần Misa chưa tự trừ → đối trừ trên Misa. Dòng job đã tính theo sổ.")
        if mot_phan:
            goi_y += f" Có {mot_phan} dòng 'Thanh toán một phần'."
    else:
        mau, goi_y = "cam", ("Dư nợ LỚN hơn số Misa ghi chưa thanh toán: nợ trước năm nay, hoặc hoá đơn Misa "
                             "ghi 'Đã thanh toán' mà sổ còn nợ. Dòng job đã tính theo sổ.")
    return {"mau": mau, "lech": L, "tong_job": S, "goi_y": goi_y}


def danh_gia_rui_ro(r: DongRuiRo, ngay_chot: Optional[dt.date] = None) -> DongRuiRo:
    J, Q = r.cuoi_ky, r.qua_han
    U, V, W = r.tt_total, r.tt_qua_han, r.tt_61
    nop = sum(r.nhom[3:])
    r.s = [
        int(U > 0.5 and J > U + 0.5 and Q >= V - 0.5),            # S1 nợ tăng, quá hạn không giảm
        int(Q > 0.5 and abs(Q - V) < 1),                          # S2 quá hạn đứng yên
        int(nop > W + 0.5),                                       # S3 nợ trên 60 ngày tăng
        int(not r.ly_do.strip()),                                 # S4 chưa có giải trình
        int(not r.credit_term.strip()),                           # S5 chưa có term
        int(J > 0.5 and abs(Q - J) < 1),                          # S6 toàn bộ đã quá hạn
    ]
    r.ngay_hen = ngay_hen(r.ly_do_hieu_luc, ngay_chot) if ngay_chot else None
    r.s7 = int(bool(r.ngay_hen and r.ngay_hen < ngay_chot and Q > 0.5 and Q >= V - 0.5))
    r.so_dau_hieu = sum(r.s[:3]) + r.s7
    r.ma_dau_hieu = " ".join([f"S{i + 1}" for i, v in enumerate(r.s) if v] + (["S7"] if r.s7 else []))
    if Q < 0.5:
        r.nhom_rui_ro = N0
    elif _co(r.ly_do_hieu_luc, TU_KHOA_LOAI_TRU):
        r.nhom_rui_ro = NX
    elif _co(r.ly_do_hieu_luc, TU_KHOA_PHAP_LY):
        r.nhom_rui_ro = N4
    elif r.nhom[5] > 0.5 or r.so_dau_hieu >= SO_DAU_HIEU_NO_XAU:
        r.nhom_rui_ro = N3
    elif r.so_dau_hieu >= 1:
        r.nhom_rui_ro = N2
    else:
        r.nhom_rui_ro = N1
    if r.nhom_rui_ro == N4:
        r.viec_can_lam = ("Chuyển hồ sơ thu hồi. Rà soát chứng từ gốc để xét trích dự phòng "
                          "theo TT48/2019.")
    elif r.nhom_rui_ro == N3:
        r.viec_can_lam = ("Dừng cấp dịch vụ mới. Chốt công văn đòi nợ có xác nhận gửi. Kế toán "
                          "trưởng ký xác nhận phân nhóm." if r.nhom[5] > 0.5 else
                          "Dừng cấp dịch vụ mới. Yêu cầu cam kết NGÀY cụ thể bằng văn bản trong 7 ngày.")
    r.ceo = Q >= NGUONG_CEO and not r.ly_do.strip()
    r.xu_huong = _xu_huong(r)
    return r


def danh_gia_fin(r: DongRuiRo, d: DongBaoCao, kct, tai_ngay: dt.date, da_tt=None) -> DongRuiRo:
    """WR-13: mức rủi ro theo tiêu chí FIN, nhóm khách, phương án xử lý, điểm A–D (wr/danhgia.py)."""
    from . import danhgia as dg
    han = [(tai_ngay - h).days for _n, _s, h, t in (d.hoa_don or []) if t > 0.5 and h]
    if han:
        r.ngay_qh = max(0, max(han))
    else:                                   # khó đòi / ngoại lệ không có hoá đơn theo sổ: lấy cận trên của nhóm tuổi
        r.ngay_qh = (0, 30, 60, 90, 120, 121)[d.nhom_gia_nhat]
    r.ngay_hd_dau = kct.ngay_hd_dau() if kct is not None else None
    r.nhom_kh = dg.nhom_khach(r.ngay_hd_dau, tai_ngay)
    r.muc_ngay = dg.muc_theo_ngay(r.ngay_qh)
    tk = kct.thong_ke_cham_diem(tai_ngay, da_tt) if kct is not None else {"so_hd": 0, "so_tra_du": 0,
                                                                           "gia_tri": 0.0, "da_thu": 0.0}
    r.cham = {"so_hd": tk["so_hd"],
              "ty_le_thu_hoi": 1.0 if tk["gia_tri"] <= 0 else round(tk["da_thu"] / tk["gia_tri"], 4),
              "ty_le_qua_han": 0.0 if r.cuoi_ky <= 0 else round(r.qua_han / r.cuoi_ky, 4),
              "ty_le_tra_du": 0.0 if not tk["so_hd"] else round(tk["so_tra_du"] / tk["so_hd"], 4)}
    r.diem, r.hang = dg.cham_diem(r.ngay_qh, r.cham["ty_le_thu_hoi"], r.cham["ty_le_qua_han"], r.cham["ty_le_tra_du"])
    return r


def danh_gia_mau5(r: DongRuiRo, x: float, y: float, so_hd_tre: int) -> DongRuiRo:
    """WR-15 (Head FIN): mức cuối 0–4 = mức gốc theo tuổi nợ, +1 nếu có T1–T4 (trừ khách 'Không nâng hạng màu'), lên thẳng
    ít nhất Đỏ khi mất liên lạc / bất lợi / thất hứa. Cùng luật với ghifull.cong_thuc_fin và dashboard."""
    from . import danhgia as dg
    J, Q = r.cuoi_ky, r.qua_han
    r.muc_goc = dg.muc_goc5(r.nhom)
    r.tieu_chi = [
        int(Q > 0.5 and Q >= x),                                                         # T1
        int(r.han_muc_ap is not None and J > r.han_muc_ap),                              # T2
        int(r.so_hd_tre >= so_hd_tre),                                                   # T3
        # WR-15b (Head FIN 12/11): T4 chỉ còn khách mới nợ ≥ Y. Chưa có term / chưa ✓ hợp đồng là việc cần bổ sung
        # (sheet Contracts), không nâng màu.
        int(J > 0.5 and r.tt_total < 0.5 and J >= y),                                    # T4
    ]
    co_dau_hieu = not r.khong_nang and any(r.tieu_chi)
    # WR-15b (Head FIN 12/11): khách đang trong hạn KHÔNG +1 màu — giữ Trắng, chỉ gắn "↑T…" để theo dõi
    r.nang = int(co_dau_hieu and Q > 0.5)
    r.theo_doi = ("↑" + " ".join(f"T{i}" for i, v in enumerate(r.tieu_chi, start=1) if v)) \
        if co_dau_hieu and Q <= 0.5 else ""
    r.nang_thang = int(dg.nang_thang(r.ly_do_hieu_luc, r.danh_gia, Q > 0.5))
    r.muc5 = dg.muc_cuoi5(r.muc_goc, bool(r.nang), bool(r.nang_thang), Q > 0.5, r.danh_gia)
    r.muc_fin = dg.muc3_tu_muc5(r.muc5)
    tc = ("T1 quá hạn ≥ X", "T2 dư nợ > hạn mức", f"T3 {r.so_hd_tre} hoá đơn trả trễ",
          "T4 khách mới nợ ≥ Y")
    phan = [f"Tuổi nợ: {dg.MUC5[r.muc_goc]}."]
    if any(r.tieu_chi):
        phan.append("; ".join(t for t, v in zip(tc, r.tieu_chi) if v) + (
            " (không nâng: khách ✓)." if r.khong_nang else " → +1." if r.nang else " (đang trong hạn: theo dõi, không đổi màu)."))
    r.can_cu_goc = " ".join(phan)
    if r.nang_thang:
        phan.append("Lên thẳng Đỏ (mất liên lạc / bất lợi / thất hứa).")
    if Q > 0.5 and (list(r.danh_gia) + ["", ""])[1] == "Không":
        phan.append("Không phản hồi.")
    r.can_cu = " ".join(phan)
    r.hinh_thuc = dg.hinh_thuc(r.muc_fin, r.nhom_kh)
    r.phuong_an = dg.phuong_an(r.muc_fin, r.nhom_kh, r.ngay_qh, r.danh_gia)
    return r


def tinh_full(th: TongHop, ct: ctm.ChiTiet, so: sg.SoGhi, dong_sms=None, bang=None,
              pt: Optional[PhaiTra] = None, ngay_chot: Optional[dt.date] = None,
              job_no_cu: Optional[Dict[str, list]] = None, job_misa=None,
              credit_terms: Optional[Dict[str, str]] = None, ma_no_cu=None,
              master: Optional[Dict[str, dict]] = None, tham_so=None) -> KetQuaFull:
    from . import danhgia as dg
    ts, bang_term = tham_so or (dict(dg.THAM_SO_MAC_DINH), list(dg.HE_SO_TERM_MAC_DINH))
    # Ngày chốt (tên trang sổ ghi, "At dd Mon yyyy") = hôm sau ngày cuối kỳ, như file Tuổi nợ cũ.
    # Tuổi nợ tính tại ngày cuối kỳ: hoá đơn đến hạn đúng ngày đó vẫn là "Current".
    ngay_chot = ngay_chot or th.den_ngay + dt.timedelta(days=1)
    tai_ngay = th.den_ngay
    truoc = so.trang_truoc(ngay_chot) or {}
    truoc_theo_ma = {v.ma.upper(): v for v in truoc.values() if v.ma}
    # WR-11 PA1: cột lý do tuần này bắt đầu TRỐNG mỗi tuần mới; chỉ khi chạy lại cùng tuần (sổ ghi đã có trang
    # của chính ngày chốt này, VD Sales đã ghi vào báo cáo tuần này) thì giữ lý do tuần này
    nay = so.trang.get(ngay_chot.isoformat()) or {}
    nay_theo_ma = {v.ma.upper(): v for v in nay.values() if v.ma}
    theo_khach, khong_misa, chua_ghep = _ghep_sms(th, dong_sms, bang)
    kq = KetQuaFull(dong=[], sms_khong_misa=khong_misa, sms_chua_ghep=chua_ghep)
    sales_job = {}
    for x in dong_sms or []:
        for j in _cac_job(x.job or ""):
            if x.salesman and th_.ten_sales(x.salesman):
                sales_job.setdefault(j.upper(), th_.ten_sales(x.salesman))
    if job_misa is not None:
        kq.job_misa = True
        for ten, n, tien in job_misa.khong_ghep_tai(tai_ngay):
            kct_ = ct.khach.get(chuan_hoa_ten(ten))
            if kct_ is not None and abs(kct_.so_du(tai_ngay)) < 1:
                kq.het_no_con_hd.append((ten, n, tien))
            else:
                kq.ten_khong_ghep.append((ten, n, tien))
    kho_doi = bang.kho_doi if bang else {}

    for k in th.khach_con_no():
        if k.ten_chuan in so.bo_qua:
            kq.da_bo.append((k.ten, k.du_no, so.bo_qua[k.ten_chuan].get("ly_do", "")))
            continue
        T = round(k.du_no, 2)
        if T < NGUONG_DU_LE:              # WR-08: FIN xoá tay Phan Thuý Ngân 84đ ở W41
            kq.da_bo.append((k.ten, T, "Dư nợ lẻ dưới 1.000đ — không đưa vào bảng"))
            continue
        tr = truoc.get(k.ten_chuan) or truoc_theo_ma.get(k.ma.upper())
        da_tt = set()
        if job_misa is not None:
            for h in job_misa.theo_ma.get(k.ma, []):
                if h.ngay <= tai_ngay and not h.chua_tt_tai(tai_ngay) and str(h.so_hd).strip().isdigit():
                    da_tt.add((int(h.so_hd), h.ngay.year))
        # WR-08: bảng credit term FIN giữ ở sample/ thắng; rồi tới lệnh khai tay; rồi báo cáo tuần trước
        term = ((credit_terms or {}).get(k.ma.upper())
                or so.dieu_chinh.get(k.ten_chuan, {}).get("credit_term")
                or (tr.credit_term if tr else "") or "")
        term_ngay = sg.so_ngay_term(term)
        hn = nay.get(k.ten_chuan) or nay_theo_ma.get(k.ma.upper())
        d = DongBaoCao(ma=k.ma, ten=k.ten, ten_chuan=k.ten_chuan, credit_term=term, total=T,
                       nhom=[0.0] * 6, salesman=tr.salesman if tr else "",
                       ghi_chu=tr.ghi_chu if tr else "", ly_do=hn.ly_do if hn else "",
                       ly_do_truoc=tr.ly_do if tr else "")
        # WR-13 (FIN 05/11): 3 ô Sales đánh giá — chạy lại cùng tuần giữ ô tuần này; tuần mới mang giá trị tuần trước
        # FIN 12/11: bỏ 3 ô Sales khỏi báo cáo → luôn giá trị mặc định (giá trị cũ trong sổ ghi không còn tác dụng)
        d.danh_gia = list(th_.DANH_GIA_MAC_DINH)
        ngoai = so.ngoai_le.get(k.ten_chuan)
        kct = ct.khach.get(k.ten_chuan)
        if ngoai and "nhom" in ngoai:
            d.nhom[sg.TEN_NHOM.index(ngoai["nhom"])] = T
            d.lo = sg.lo_tu_nhom(d.nhom, tai_ngay)
            d.nguon, d.phan_nguon = "ngoai-le", {"ngoai-le": T}
            d.cho_xem_lai.append(f"Ngoại lệ do kế toán khai: xếp cả {T:,.0f} vào nhóm {ngoai['nhom']}")
        elif k.ma in kho_doi or "khó đòi" in (d.ly_do or d.ly_do_truoc).lower():
            d.nhom[5] = T
            d.lo = sg.lo_tu_nhom(d.nhom, tai_ngay)
            d.nguon, d.phan_nguon = "kho-doi", {"kho-doi": T}
            d.ly_do = d.ly_do or kho_doi.get(k.ma) or d.ly_do_truoc or "Khó đòi"
            # W42 (FIN giữ dòng "HĐ 2519 nợ cũ" dưới Minh Khôi): khách khó đòi mà FIN đã khai hoá đơn ở bảng
            # job nợ cũ thì vẫn liệt kê hoá đơn theo sổ, tất cả ở 120+; khó đòi chưa khai thì giữ một dòng
            if kct and k.ma.upper() in (ma_no_cu or set()):
                con_kd, _ = kct.hoa_don_con_no(tai_ngay)
                if con_kd and abs(sum(t for *_x, t in con_kd) - T) < 1:
                    d.hoa_don = [(n, s, sg.han_tu_term(n, term), t) for n, s, t in con_kd]
        else:
            # FIN 17/10: tiền thu không ghi số HĐ trừ trước vào hoá đơn Misa đã ghi "Đã thanh toán"
            con, _ = kct.hoa_don_con_no(tai_ngay, da_tt) if kct else ([], 0.0)
            # mỗi hoá đơn: [ngày HĐ, số HĐ, hạn, còn nợ] — giữ lại cho sheet Invoices
            hd = [[n, so_hd, sg.han_tu_term(n, term), t] for n, so_hd, t in con]
            S = round(sum(x[3] for x in hd), 2)
            phan = {"hoa-don": min(S, T)} if S else {}
            if kct is None:
                d.cho_xem_lai.append("Không thấy khách này trong file Chi tiết 131 (tên khác nhau?) — "
                                     "tạm tính là nợ mới")
            if S > T + 1:
                du = round(S - T, 2)        # Chi tiết ghi nợ nhiều hơn Tổng hợp: trừ tiếp vào HĐ cũ
                for x in hd:
                    tru = min(x[3], du)
                    x[3], du = round(x[3] - tru, 2), round(du - tru, 2)
                    if du <= 0:
                        break
                hd = [x for x in hd if x[3] > 0.5]
            if abs(S - T) >= 1:
                kq.lech_chi_tiet.append((k.ten, T, S))
                d.cho_xem_lai.append(f"Chi tiết 131 ghi {S:,.0f}, Tổng hợp 131 ghi {T:,.0f}: tiền theo "
                                     "Tổng hợp. Thường do hai file xuất khác thời điểm")
            lo = [sg.Lo(x[3], x[2].isoformat()) for x in hd]
            if S < T - 1:
                thieu = round(T - S, 2)
                han = tai_ngay + dt.timedelta(days=term_ngay)
                lo.append(sg.Lo(thieu, han.isoformat(), True))
                hd.append([None, "(Tổng hợp > Chi tiết)", han, thieu])
                phan["uoc-tinh"] = thieu
            d.hoa_don = [tuple(x) for x in hd]
            d.lo = lo
            d.nhom = sg.nhom_tu_lo(lo, tai_ngay)
            d.phan_nguon = phan
            d.nguon = "uoc-tinh" if "uoc-tinh" in phan else "hoa-don"

        # job và salesman từ SMS — mỗi job còn nợ thành một dòng con dưới khách (CEO 08/10:
        # sếp cần biết ai đang làm job của khách nợ để hỏi thẳng)
        ds = theo_khach.get(k.ma, [])
        if job_misa is not None:
            # v0.5 (WR-06): job + tiền lấy từ Misa (Bán hàng + Sổ chi tiết), SMS chỉ cho tên sales
            misa_chua = job_misa.chua_tt_cua(k.ma, tai_ngay)
            if d.hoa_don:
                _dong_con_theo_so(d, job_misa, k.ma, (job_no_cu or {}).get(k.ma.upper(), []), sales_job,
                                  tr.salesman if tr else "", tai_ngay)
                if d.nguon == "kho-doi":
                    for g in d.dong_con:
                        g["nhom_co_dinh"] = [0.0] * 5 + [g["tien"]]
            else:                             # khó đòi / ngoại lệ: không có hoá đơn theo sổ
                _dong_con_misa(d, misa_chua, term, tai_ngay, sales_job, tr.salesman if tr else "")
            for h in job_misa.theo_ma.get(k.ma, []):
                n = ctm._so_hd(h.so_hd)
                if n is not None:
                    d.hd_misa[(n, h.ngay.year)] = (h.trang_thai, ", ".join(j for j in h.jobs if j))
        elif ds:
            jobs = [[x, x.so_tien] for x in sorted(ds, key=lambda x: (x.ngay or dt.date.min))]
            S = sum(x.so_tien for x in ds)
            if S > T + _chenh_ty_gia(T) + 1 and d.nguon not in {"kho-doi", "ngoai-le"}:
                can_bo = S - T              # SMS ghi nhiều hơn Misa: job cũ nhất đã thu, chưa gạch paid
                for j in jobs:
                    if can_bo <= 0.5:
                        break
                    bo = min(j[1], can_bo)
                    can_bo -= bo
                    j[1] -= bo
                    kq.sms_cat.append((k.ten, j[0].job, bo, j[0].ngay))
            gop: Dict[str, dict] = {}
            for x, con in jobs:
                if con <= 0.5:
                    continue
                g = gop.setdefault(x.job, {"job": x.job, "sales": th_.ten_sales(x.salesman),
                                           "ngay": x.ngay, "nguon_ngay": x.nguon_ngay, "tien": 0.0})
                g["tien"] = round(g["tien"] + con, 2)
            d.dong_con = sorted(gop.values(), key=lambda g: (g["ngay"] or dt.date.min))
            d.job = ", ".join(_cac_job(", ".join(g["job"] for g in d.dong_con)))
            tien_sale: Dict[str, float] = {}
            for g in d.dong_con:
                if g["sales"]:
                    tien_sale[g["sales"]] = tien_sale.get(g["sales"], 0) + g["tien"]
            if tien_sale:
                d.salesman = ", ".join(sorted(tien_sale, key=lambda t: -tien_sale[t]))
        elif tr:
            d.job = ", ".join(_cac_job(tr.job))
        # nợ cũ không có job trên SMS: FIN khai một lần ở sample/ (hoá đơn nào thuộc job nào, Sales nào)
        con_lai = round(T - sum(g["tien"] for g in d.dong_con), 2)
        if job_misa is not None:
            con_lai = float("inf")            # nợ cũ trước năm nay: FIN khai theo hoá đơn, không giới hạn
        if job_no_cu and con_lai >= 1 and d.hoa_don and job_misa is None:
            mo = {}
            for ngay, so_hd, han, tien in d.hoa_don:
                if job_misa is not None and (ngay is None or ngay >= job_misa.tu_ngay):
                    continue                  # hoá đơn năm nay đã có job trên Misa
                n = ctm._so_hd(str(so_hd))
                if n is not None:
                    mo.setdefault(n, []).append((ngay, han, tien))
            for so_hd, job, sales in job_no_cu.get(k.ma.upper(), []):
                if con_lai < 1:
                    break
                for ngay, han, tien in mo.pop(so_hd, []):
                    lay = round(min(tien, con_lai), 2)
                    nhom = [0.0] * 6      # biết đúng hoá đơn nên biết đúng nhóm tuổi
                    nhom[sg.nhom_tu_so_ngay((tai_ngay - han).days)] = lay
                    d.dong_con.append({"job": job or f"HĐ {so_hd}", "sales": th_.ten_sales(sales), "ngay": ngay,
                                       "nguon_ngay": "HĐ", "tien": lay, "tham_chieu": True, "nhom_co_dinh": nhom})
                    con_lai = round(con_lai - lay, 2)
            if d.dong_con:
                d.job = ", ".join(_cac_job(", ".join(g["job"] for g in d.dong_con)))
                if not d.salesman:
                    d.salesman = ", ".join(dict.fromkeys(g["sales"] for g in d.dong_con if g["sales"]))
        if job_misa is not None:
            # đối chiếu: số Misa GHI chưa thanh toán so với dư nợ — chỉ để FIN dọn Misa; dòng job đã theo sổ
            d.doi_chieu = doi_chieu(T, [{"tien": h.tien, "trang_thai": h.trang_thai} for h in misa_chua])
            du_ct = kct.so_du(tai_ngay) if kct else None
            if (d.doi_chieu["mau"] != "khop" and du_ct is not None and abs(du_ct - T) > DUNG_SAI_KHOP
                    and abs(du_ct - d.doi_chieu["tong_job"]) <= DUNG_SAI_KHOP):
                # FIN 09/10 (GOLDEN GLOBE): job khớp Chi tiết 131, chỉ Tổng hợp 131 lệch → file xuất khác lúc
                d.doi_chieu["goi_y"] = (f"Tổng job khớp Chi tiết 131 ({du_ct:,.0f}); chỉ file Tổng hợp 131 ghi "
                                        f"{T:,.0f}: hai file xuất khác lúc, Misa đã sửa hoá đơn ở giữa. "
                                        "Xuất lại cả 6 file cùng lúc là khớp.")
            if d.nguon == "kho-doi" and d.doi_chieu["mau"] == "cam":
                d.doi_chieu.update(mau="kho_doi", goi_y="Nợ khó đòi FIN đã khai ở sample/: không cần job.")
        elif d.dong_con and abs(con_lai) >= 1:
            d.can_doi = con_lai
        _chia_tuoi_dong_con(d)
        if not term:
            d.thieu_thong_tin.append("Chưa có Credit Term, tạm tính 30 ngày")
        d.salesman = th_.ten_sales(d.salesman)
        if not d.salesman:
            d.thieu_thong_tin.append("Chưa có Salesman")
        lech = round(d.total - sum(d.nhom), 2)
        if abs(lech) >= 1:
            d.nhom[0] = round(d.nhom[0] + lech, 2)
            d.cho_xem_lai.append(f"Tổng các nhóm lệch Total {lech:,.0f}, đã dồn vào Current")
        kq.dong.append(d)

        r = DongRuiRo(
            ma=k.ma, ten=k.ten, credit_term=term, dau_ky=round(k.dau_ky, 2), tang=round(k.ps_no, 2),
            giam=round(k.ps_co, 2), cuoi_ky=T, nhom=list(d.nhom), salesman=d.salesman, job=d.job,
            ly_do=d.ly_do, tt_total=round(tr.tong, 2) if tr else 0.0,
            tt_qua_han=round(sum(tr.nhom[1:]), 2) if tr else 0.0,
            tt_61=round(sum(tr.nhom[3:]), 2) if tr else 0.0, co_tuan_truoc=tr is not None,
            nhap_bo_sung=round(k.dau_ky - tr.tong, 2) if tr else None, ly_do_truoc=d.ly_do_truoc,
        )
        if master:
            m = master.get(k.ma.upper())
            r.han_muc = (m or {}).get("han_muc")
            r.khong_nang = bool((m or {}).get("khong_nang"))
            r.link_hd = (m or {}).get("link_hd") or ""
            co_cot = any(v.get("co_hd") is not None for v in master.values())
            r.co_hd = bool((m or {}).get("co_hd")) if co_cot else None
        r.chia_sales = chia_theo_sales(d.dong_con)
        r.danh_gia = list(d.danh_gia)
        # WR-15: T2 hạn mức (hợp đồng, không có thì tạm tính) và T3 hoá đơn trả trễ
        so_thang = int(ts.get("so_thang") or 6)
        if kct is not None:
            r.billing = kct.billing_binh_quan(tai_ngay, so_thang)
            r.so_hd_tre = len(kct.hoa_don_tra_tre(tai_ngay, lambda n, t=term: sg.han_tu_term(n, t), da_tt,
                                                  so_ngay=int(ts.get("ngay_tre") or 15), so_thang=so_thang))
        r.he_so = dg.he_so_term(term, bang_term, sg.han_tu_term)
        if r.han_muc:
            r.han_muc_ap, r.nguon_han_muc = float(r.han_muc), "Hợp đồng"
        elif r.he_so is not None and r.billing > 0:
            r.han_muc_ap = round(r.billing * r.he_so * float(ts.get("he_so_an_toan") or 1), 2)
            r.nguon_han_muc = "Tạm tính"
        else:
            r.nguon_han_muc = "Thiếu dữ liệu"
        kq.rui_ro.append(danh_gia_fin(danh_gia_rui_ro(r, ngay_chot), d, kct, tai_ngay, da_tt))

    # WR-15: X cần nợ quá hạn của mọi khách (top 15%) → tính mức màu sau vòng lặp
    kq.x, kq.x1, kq.x2 = dg.nguong_x([r.qua_han for r in kq.rui_ro], ts)
    kq.y = float(ts["y"]) if isinstance(ts.get("y"), (int, float)) else round(kq.x / 2, 2)
    kq.tham_so = ts
    for r in kq.rui_ro:
        danh_gia_mau5(r, kq.x, kq.y, int(ts.get("so_hd_tre") or 2))

    if job_misa is not None:
        # khách có trên Tổng hợp nhưng không còn dư nợ, mà hoá đơn vẫn ghi chưa thanh toán
        da_co = {d.ma for d in kq.dong} | {k.ma for k in th.khach_con_no()}
        ten_ma = {k.ma: k.ten for k in th.dong.values()}
        for ma in job_misa.theo_ma:
            if ma in da_co:
                continue
            hds = job_misa.chua_tt_cua(ma, tai_ngay)
            if hds:
                kq.het_no_con_hd.append((ten_ma.get(ma, ma), len(hds), round(sum(h.tien for h in hds), 2)))
        kq.het_no_con_hd.sort(key=lambda x: -x[2])

    # gộp SMS cần gạch paid theo job
    gop: Dict[tuple, list] = {}
    for khach_, job, tien, ngay in kq.sms_cat:
        g = gop.setdefault((khach_, job), [khach_, job, 0.0, ngay])
        g[2] += tien
    kq.sms_cat = [tuple(g) for g in gop.values()]
    kq.dong.sort(key=lambda d: (-d.nhom_gia_nhat, -d.total))
    thu_tu = {N4: 0, N3: 1, N2: 2, N1: 3, NX: 4, N0: 5}
    kq.rui_ro.sort(key=lambda r: (thu_tu.get(r.nhom_rui_ro, 9), -r.qua_han, -r.cuoi_ky))

    # khách tuần trước còn nợ, tuần này hết
    con = {k.ten_chuan for k in th.khach_con_no()}
    for tc, hs in truoc.items():
        if tc not in con and hs.tong > 0.5:
            kq.khach_tra_xong.append((hs.ten, hs.tong))

    # --- tiền thu trong kỳ, tách theo TK đối ứng
    for tc, kh in ct.khach.items():
        theo = kh.thu_trong_ky(th.tu_ngay, th.den_ngay)
        if theo:
            ma = th.dong[tc].ma if tc in th.dong else ""
            kq.thu.append(DongThu(ma=ma, ten=kh.ten, theo_nhom=theo))
    kq.thu.sort(key=lambda x: -x.tong)

    # --- phải trả
    if pt is not None:
        truoc_tra = so.phai_tra_truoc(ngay_chot)
        kq.phai_tra = []
        for v in pt.dong:
            dau = round(v.dk_co - v.dk_no, 2)
            cuoi = round(v.du_co, 2)
            if v.du_no > 0.5:
                kq.vendor_tra_truoc.append((v.ma, v.ten, v.du_no))
            if cuoi <= 0.5 and not (dau > 0.5 and v.ps_no > 0.5):
                if v.ps_no > 0.5:
                    kq.tra_vendor.append(DongTra(v.ma, v.ten, dau, v.ps_co, v.ps_no, cuoi, ""))
                continue
            if cuoi <= 0.5:
                tt = "Paid off this week"
            elif dau <= 0.5:
                tt = "New this week"
            elif v.ps_no > 0.5:
                tt = "Partly paid"
            else:
                tt = "Outstanding"
            dong_tra = DongTra(v.ma, v.ten, dau, round(v.ps_co, 2), round(v.ps_no, 2), cuoi, tt,
                               round(dau - truoc_tra.get(v.ma, 0.0), 2) if truoc_tra is not None else None)
            kq.phai_tra.append(dong_tra)
            if v.ps_no > 0.5:
                kq.tra_vendor.append(dong_tra)
        kq.phai_tra.sort(key=lambda x: (x.trang_thai == "Paid off this week", -x.cuoi_ky))
        kq.tra_vendor.sort(key=lambda x: -x.da_tra)

    kq.chi_so = chi_so_tuan(kq, th, pt, ngay_chot)
    kq.chi_so_truoc = _chi_so_truoc(so, ngay_chot, truoc)
    return kq


def chi_so_tuan(kq: KetQuaFull, th: TongHop, pt: Optional[PhaiTra], ngay_chot: dt.date) -> dict:
    nhom = [round(sum(d.nhom[i] for d in kq.dong), 2) for i in range(6)]
    thu: Dict[str, float] = {}
    for x in kq.thu:
        for k, v in x.theo_nhom.items():
            thu[k] = round(thu.get(k, 0) + v, 2)
    xau = [r for r in kq.rui_ro if r.nhom_rui_ro in (N3, N4)]
    cs = {
        "ngay_chot": ngay_chot.isoformat(),
        "tuan": f"W{th.den_ngay.isocalendar()[1]:02d}",
        "tu_ngay": th.tu_ngay.isoformat(), "den_ngay": th.den_ngay.isoformat(),
        "tong_phai_thu": round(sum(d.total for d in kq.dong), 2),
        "chua_den_han": nhom[0], "qua_han": round(sum(nhom[1:]), 2),
        "qua_han_1_30": nhom[1], "qua_han_31_60": nhom[2], "qua_han_61_90": nhom[3],
        "qua_han_91_120": nhom[4], "qua_han_120": nhom[5],
        "so_khach": len(kq.dong),
        "no_moi": round(sum(k.ps_no for k in th.dong.values()), 2),
        "da_thu": round(sum(k.ps_co for k in th.dong.values()), 2),
        "da_thu_tien": round(thu.get("ngan-hang", 0) + thu.get("tien-mat", 0), 2),
        "thu_ngan_hang": thu.get("ngan-hang", 0.0), "thu_tien_mat": thu.get("tien-mat", 0.0),
        "thu_can_tru": thu.get("can-tru", 0.0), "thu_ty_gia_phi": thu.get("ty-gia-phi", 0.0),
        "thu_khac": thu.get("khac", 0.0),
        "no_xau": round(sum(r.qua_han for r in xau), 2), "so_khach_no_xau": len(xau),
        "no_xau_120": round(sum(r.nhom[5] for r in xau), 2),
        "no_xau_phap_ly": round(sum(r.qua_han for r in xau if r.nhom_rui_ro == N4), 2),
        "ceo_can_xem": sum(1 for r in kq.rui_ro if r.ceo),
    }
    if pt is not None:
        cs.update({
            "tong_phai_tra": round(pt.tong_con_no, 2), "da_tra_vendor": round(pt.tong_da_tra, 2),
            "no_vendor_moi": round(pt.tong_no_moi, 2),
            "so_vendor": sum(1 for v in pt.dong if v.du_co > 0.5),
            "dong_tien_rong": round(cs["da_thu_tien"] - pt.tong_da_tra, 2),
        })
    return cs


def _chi_so_truoc(so: sg.SoGhi, ngay_chot: dt.date, truoc: dict) -> dict:
    cu = [n for n in so.lich_su_tuan if dt.date.fromisoformat(n) < ngay_chot]
    if cu:
        return dict(so.lich_su_tuan[max(cu)])
    if not truoc:
        return {}
    # chưa có lịch sử (tuần đầu dùng v0.4): dựng lại phần phải thu từ trang sổ ghi tuần trước
    nhom = [round(sum(h.nhom[i] for h in truoc.values()), 2) for i in range(6)]
    return {"tong_phai_thu": round(sum(h.tong for h in truoc.values()), 2), "chua_den_han": nhom[0],
            "qua_han": round(sum(nhom[1:]), 2), "qua_han_1_30": nhom[1], "qua_han_31_60": nhom[2],
            "qua_han_61_90": nhom[3], "qua_han_91_120": nhom[4], "qua_han_120": nhom[5],
            "so_khach": len(truoc)}
