# -*- coding: utf-8 -*-
"""Luật điền sheet Receivable. Bản đặc tả: handoff/docs/SPEC-logic-bao-cao-tuan.md

Tóm tắt:
  - Danh sách khách và cột Total  <- file Tổng hợp (đây là số tiền đúng)
  - Sáu nhóm tuổi                 <- file Tuổi nợ nếu tin được, nếu không thì sổ ghi
  - Các cột do người điền         <- mang từ trang trước của sổ ghi sang
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from . import soghi as sg
from . import thuonghieu as th
from .misa import DongTuoiNo, TongHop, TuoiNo

LECH_CHO_PHEP = 2.0  # đồng


@dataclass
class DongBaoCao:
    ma: str
    ten: str
    ten_chuan: str
    credit_term: str
    total: float
    nhom: List[float]
    salesman: str = ""
    ghi_chu: str = ""
    ly_do: str = ""
    ly_do_truoc: str = ""           # WR-11 PA1: lý do Sales ghi ở báo cáo tuần trước (chỉ để đọc)
    danh_gia: List[str] = field(default_factory=lambda: list(th.DANH_GIA_MAC_DINH))  # WR-13: 3 ô Sales đánh giá
    nguon: str = ""                 # misa | so-ghi | khoi-dong-lanh | ngoai-le
    cho_xem_lai: List[str] = field(default_factory=list)      # vấn đề về SỐ
    thieu_thong_tin: List[str] = field(default_factory=list)  # thiếu cột người điền
    lo: List[sg.Lo] = field(default_factory=list)
    job: str = ""
    phan_nguon: Dict[str, float] = field(default_factory=dict)  # sms | so-ghi | kho-doi | uoc-tinh
    ly_do_tu_dong: str = ""         # câu tool tự viết khi Sales chưa ghi lý do (v0.4)
    dong_con: List[dict] = field(default_factory=list)   # job SMS còn nợ: job, sales, ngay, tien
    can_doi: float = 0.0            # Misa − tổng job SMS: nợ cũ / chưa có job (+) hoặc SMS ghi dư (−)
    can_doi_nhom: List[float] = field(default_factory=list)   # sáu nhóm tuổi của dòng cân đối
    hoa_don: List[tuple] = field(default_factory=list)   # (ngày HĐ, số HĐ, hạn, còn nợ) — sheet Invoices
    hd_misa: Dict = field(default_factory=dict)     # v0.5: (số HĐ int, năm) -> (trạng thái Misa, "job, job")
    doi_chieu: Dict = field(default_factory=dict)   # v0.5: {mau, lech, tong_job, goi_y} — dư nợ vs tổng job Misa

    @property
    def nhom_gia_nhat(self) -> int:
        for i in range(5, -1, -1):
            if self.nhom[i] > 0.5:
                return i
        return 0


def _tin_duoc_file_tuoi_no(a: Optional[DongTuoiNo], total: float) -> bool:
    """File Tuổi nợ chỉ dùng được khi khớp số tiền với file Tổng hợp VÀ mọi hoá
    đơn đều có ghi hạn thanh toán."""
    return bool(a) and abs(a.tong - total) < LECH_CHO_PHEP and a.khong_han < 1


def tinh_mot_khach(
    ten_chuan: str,
    ma: str,
    ten: str,
    total: float,
    a: Optional[DongTuoiNo],
    truoc: Optional[sg.HoSoKhach],
    ngay_chot: dt.date,
    ngoai_le: Dict[str, dict],
    dieu_chinh: Optional[Dict[str, dict]] = None,
) -> DongBaoCao:
    dieu_chinh = dieu_chinh or {}
    term = (dieu_chinh.get(ten_chuan, {}).get("credit_term")
            or (truoc.credit_term if truoc else "") or "")
    d = DongBaoCao(
        ma=ma, ten=ten, ten_chuan=ten_chuan, credit_term=term, total=round(total, 2),
        nhom=[0.0] * 6,
        salesman=(truoc.salesman if truoc else ""),
        ghi_chu=(truoc.ghi_chu if truoc else ""),
        ly_do=(truoc.ly_do if truoc else ""),
    )
    term_ngay = sg.so_ngay_term(term)

    ngoai = ngoai_le.get(ten_chuan)
    if ngoai and "nhom" in ngoai:
        i = sg.TEN_NHOM.index(ngoai["nhom"])
        d.nhom = [0.0] * 6
        d.nhom[i] = round(total, 2)
        d.lo = sg.lo_tu_nhom(d.nhom, ngay_chot)
        d.nguon = "ngoai-le"
        d.cho_xem_lai.append(f"Ngoại lệ do kế toán khai: xếp cả {total:,.0f} vào nhóm {ngoai['nhom']}"
                             + (f" — {ngoai.get('ly_do')}" if ngoai.get("ly_do") else ""))
        return d

    if _tin_duoc_file_tuoi_no(a, total):
        d.nhom = [round(a.truoc_han, 2)] + [round(x, 2) for x in a.qua_han]
        d.lo = sg.lo_tu_nhom(d.nhom, ngay_chot)
        d.nguon = "misa"
    else:
        if truoc and truoc.lo:
            lo = sg.dieu_chinh_lo(truoc.lo, total, ngay_chot, term_ngay)
            d.nguon = "so-ghi"
        else:
            lo = [sg.Lo(round(total, 2), (ngay_chot + dt.timedelta(days=term_ngay)).isoformat(), True)]
            d.nguon = "khoi-dong-lanh"
            d.cho_xem_lai.append(
                "Khách chưa có trong sổ ghi: tạm xếp toàn bộ vào Current, tuổi nợ là ước tính"
            )
        d.lo = lo
        d.nhom = sg.nhom_tu_lo(lo, ngay_chot)
        if a is None:
            d.cho_xem_lai.append("Không có trong file Tuổi nợ, chia theo sổ ghi")
        elif abs(a.tong - total) >= LECH_CHO_PHEP:
            d.cho_xem_lai.append(
                f"Hai file Misa lệch nhau: Tổng hợp {total:,.0f} — Tuổi nợ {a.tong:,.0f}. "
                "Chia theo sổ ghi, tiền lấy theo Tổng hợp"
            )
        elif a.khong_han >= 1:
            d.cho_xem_lai.append(
                f"Hoá đơn không ghi hạn thanh toán: {a.khong_han:,.0f}. Tuổi nợ tính theo sổ ghi"
            )

    if not term:
        d.thieu_thong_tin.append("Chưa có Credit Term, tạm tính 30 ngày")
    if not d.salesman:
        d.thieu_thong_tin.append("Chưa có Salesman")
    lech = round(d.total - sum(d.nhom), 2)
    if abs(lech) >= 1:
        d.nhom[0] = round(d.nhom[0] + lech, 2)   # tổng các nhóm luôn phải bằng Total
        d.cho_xem_lai.append(f"Tổng các nhóm lệch Total {lech:,.0f}, đã dồn phần lệch vào Current")
    return d


def tinh_bang(
    th: TongHop,
    tn: TuoiNo,
    so: sg.SoGhi,
    ngay_chot: Optional[dt.date] = None,
) -> List[DongBaoCao]:
    ngay_chot = ngay_chot or tn.den_ngay
    truoc = so.trang_truoc(ngay_chot) or {}
    ra: List[DongBaoCao] = []
    da_bo: List[tuple] = []
    for k in th.khach_con_no():
        if k.ten_chuan in so.bo_qua:
            da_bo.append((k.ten, k.du_no, so.bo_qua[k.ten_chuan].get("ly_do", "")))
            continue
        ra.append(
            tinh_mot_khach(
                ten_chuan=k.ten_chuan, ma=k.ma, ten=k.ten, total=k.du_no,
                a=tn.dong.get(k.ten_chuan), truoc=truoc.get(k.ten_chuan),
                ngay_chot=ngay_chot, ngoai_le=so.ngoai_le, dieu_chinh=so.dieu_chinh,
            )
        )
    # nợ lâu nhất lên trên; cùng nhóm thì số tiền lớn lên trên
    ra.sort(key=lambda d: (-d.nhom_gia_nhat, -d.total))
    tinh_bang.da_bo = da_bo          # để báo cáo liệt kê ở sheet Cần xem lại
    return ra


def thanh_ho_so(d: DongBaoCao) -> sg.HoSoKhach:
    return sg.HoSoKhach(
        ma=d.ma, ten=d.ten, credit_term=d.credit_term, salesman=d.salesman,
        ghi_chu=d.ghi_chu, ly_do=d.ly_do, tong=d.total, nhom=d.nhom, lo=d.lo, nguon=d.nguon,
        job=d.job, danh_gia=list(d.danh_gia),
    )


# --- v0.3: ghép thêm nguồn SMS (WR-03) ------------------------------------
#
# Misa quyết TIỀN. SMS quyết NGÀY của từng job. Sổ ghi (mà nguồn là bản final
# kế toán trưởng đã duyệt) quyết tuổi của phần nợ cũ không có trong SMS.
# Luật: handoff/MAP.md mục "Luồng chạy", handoff/docs/NOTE-ke-toan-tra-loi-20261006.md

@dataclass
class KetQuaSMS:
    dong: List[DongBaoCao]
    da_bo: List[tuple] = field(default_factory=list)
    sms_cat: List[tuple] = field(default_factory=list)          # khách, job, số tiền bỏ, ngày
    sms_khong_misa: List[tuple] = field(default_factory=list)   # mã SMS, tên SMS, job, số tiền, lý do
    sms_chua_ghep: List[tuple] = field(default_factory=list)    # mã SMS, tên SMS, số tiền, số dòng
    ngay_som_nhat: Optional[dt.date] = None


def _cac_job(*chuoi: str) -> List[str]:
    ra: List[str] = []
    for c in chuoi:
        for j in re.split(r"[,;\n]", re.sub(r"(?i)^\s*job\s*:\s*", "", c or "")):
            j = j.strip()
            if j and j not in ra:
                ra.append(j)
    return ra


def _chenh_ty_gia(total: float) -> float:
    """Ngưỡng coi chênh SMS–Misa là do tỷ giá: 1% số dư, tối đa 500.000đ."""
    return min(500_000.0, abs(total) * 0.01)


def tinh_bang_sms(th: TongHop, tn: TuoiNo, so: sg.SoGhi, dong_sms, bang=None,
                  ngay_chot: Optional[dt.date] = None) -> KetQuaSMS:
    from . import ghep
    ngay_chot = ngay_chot or tn.den_ngay
    truoc = so.trang_truoc(ngay_chot) or {}
    truoc_theo_ma = {v.ma.upper(): v for v in truoc.values() if v.ma}
    khach = th.khach_con_no()
    con_no = {k.ma: k for k in khach}
    theo_ma = {k.ma.upper(): k.ma for k in th.dong.values()}
    theo_ten = {k.ten_chuan: k.ma for k in th.dong.values()}
    co_ngay = [x.ngay for x in dong_sms if x.ngay]
    kq = KetQuaSMS(dong=[], ngay_som_nhat=min(co_ngay) if co_ngay else None)

    theo_khach: Dict[str, list] = {}
    chua_ghep: Dict[str, list] = {}
    for x in dong_sms:
        ma, cach = ghep.ma_misa_cua(x, bang, theo_ma, theo_ten)
        if ma is not None:
            ma = theo_ma.get(ma.upper(), ma)
        if ma is None and cach == "chưa ghép":
            chua_ghep.setdefault(x.ma_sms, []).append(x)
        elif ma is None or ma not in con_no:
            ly_do = ("FIN khai không có trên Misa" if ma is None else
                     f"Misa không còn nợ ({ma}) — nhiều khả năng đã thu, SMS chưa gạch paid")
            kq.sms_khong_misa.append((x.ma_sms, x.ten_sms, x.job, x.so_tien, ly_do))
        else:
            theo_khach.setdefault(ma, []).append(x)
    for ma_sms, ds in chua_ghep.items():
        kq.sms_chua_ghep.append((ma_sms, ds[0].ten_sms, sum(x.so_tien for x in ds), len(ds)))

    kho_doi = bang.kho_doi if bang else {}
    for k in khach:
        if k.ten_chuan in so.bo_qua:
            kq.da_bo.append((k.ten, k.du_no, so.bo_qua[k.ten_chuan].get("ly_do", "")))
            continue
        T = round(k.du_no, 2)
        tr = truoc.get(k.ten_chuan) or truoc_theo_ma.get(k.ma.upper())
        term = (so.dieu_chinh.get(k.ten_chuan, {}).get("credit_term")
                or (tr.credit_term if tr else "") or "")
        term_ngay = sg.so_ngay_term(term)
        d = DongBaoCao(ma=k.ma, ten=k.ten, ten_chuan=k.ten_chuan, credit_term=term, total=T,
                       nhom=[0.0] * 6, salesman=tr.salesman if tr else "",
                       ghi_chu=tr.ghi_chu if tr else "", ly_do=tr.ly_do if tr else "")
        ngoai = so.ngoai_le.get(k.ten_chuan)
        if ngoai and "nhom" in ngoai:
            d.nhom[sg.TEN_NHOM.index(ngoai["nhom"])] = T
            d.lo = sg.lo_tu_nhom(d.nhom, ngay_chot)
            d.nguon, d.phan_nguon = "ngoai-le", {"ngoai-le": T}
            d.cho_xem_lai.append(f"Ngoại lệ do kế toán khai: xếp cả {T:,.0f} vào nhóm {ngoai['nhom']}"
                                 + (f" — {ngoai.get('ly_do')}" if ngoai.get("ly_do") else ""))
            d.job = ", ".join(_cac_job(tr.job if tr else ""))
        elif k.ma in kho_doi:
            d.nhom[5] = T
            d.lo = sg.lo_tu_nhom(d.nhom, ngay_chot)
            d.nguon, d.phan_nguon = "kho-doi", {"kho-doi": T}
            d.ly_do = d.ly_do or kho_doi[k.ma]
            d.job = ", ".join(_cac_job(tr.job if tr else ""))
        else:
            ds = theo_khach.get(k.ma, [])
            cap = []
            for x in ds:
                goc = x.ngay or kq.ngay_som_nhat or ngay_chot
                cap.append([x, sg.Lo(round(x.so_tien, 2), (goc + dt.timedelta(days=term_ngay)).isoformat())])
            if any(x.ngay is None for x in ds):
                d.cho_xem_lai.append("Có job trên SMS không ghi ETD/ETA, tạm lấy ngày sớm nhất của file SMS")
            S = round(sum(l.so_tien for _, l in cap), 2)
            if cap and abs(S - T) <= _chenh_ty_gia(T):
                # SMS quy VND theo tỷ giá lúc làm job, Misa theo tỷ giá lúc hạch toán:
                # vênh vài chục nghìn là tỷ giá, không phải nợ thiếu/thừa — dồn vào job mới nhất
                moi = max(cap, key=lambda c: c[1].han)[1]
                moi.so_tien = round(moi.so_tien + (T - S), 2)
                S = T
            if S > T + 1:
                can_bo = round(S - T, 2)
                for x, l in sorted(cap, key=lambda c: c[1].han):
                    if can_bo <= 0:
                        break
                    bo = min(l.so_tien, can_bo)
                    l.so_tien = round(l.so_tien - bo, 2)
                    can_bo = round(can_bo - bo, 2)
                    kq.sms_cat.append((k.ten, x.job, bo, x.ngay))
                d.cho_xem_lai.append(f"SMS ghi nợ {S:,.0f}, Misa {T:,.0f}: đã bỏ {S - T:,.0f} ở các job "
                                     "cũ nhất — xem mục SMS cần gạch paid")
            lo = [l for _, l in cap if l.so_tien > 0.5]
            s_con = round(sum(l.so_tien for l in lo), 2)
            phan = {"sms": s_con} if s_con else {}
            R = round(T - s_con, 2)
            if R >= 1:
                con = R
                for l in sorted(tr.lo if tr else [], key=lambda l: l.han):   # cũ nhất trước
                    if con < 0.5:
                        break
                    lay = min(l.so_tien, con)
                    lo.append(sg.Lo(round(lay, 2), l.han, l.uoc_tinh))
                    con = round(con - lay, 2)
                if R - con >= 0.5:
                    phan["so-ghi"] = round(R - con, 2)
                if con >= 1:
                    goc = kq.ngay_som_nhat or ngay_chot
                    lo.append(sg.Lo(round(con, 2), (goc + dt.timedelta(days=term_ngay)).isoformat(), True))
                    phan["uoc-tinh"] = round(con, 2)
                    d.cho_xem_lai.append(
                        f"{con:,.0f} không có trong SMS và chưa có trong bản final/sổ ghi tuần trước: "
                        f"tạm tính tuổi tối thiểu (từ {goc:%d/%m/%Y}). Kế toán xếp lại nhóm trên bản final")
            d.lo = lo
            d.nhom = sg.nhom_tu_lo(lo, ngay_chot)
            d.phan_nguon = phan
            d.nguon = ("uoc-tinh" if "uoc-tinh" in phan else
                       "sms" if set(phan) == {"sms"} else
                       "so-ghi" if set(phan) == {"so-ghi"} else "sms+so-ghi")
            job_sms = [x.job for x, l in cap if l.so_tien > 0.5]
            d.job = ", ".join(_cac_job(", ".join(job_sms),
                                        tr.job if (tr and "so-ghi" in phan) else ""))
            if not d.salesman and ds:
                sale = [x.salesman for x in ds if x.salesman]
                d.salesman = max(set(sale), key=sale.count) if sale else ""
        if not term:
            d.thieu_thong_tin.append("Chưa có Credit Term, tạm tính 30 ngày")
        if not d.salesman:
            d.thieu_thong_tin.append("Chưa có Salesman")
        lech = round(d.total - sum(d.nhom), 2)
        if abs(lech) >= 1:
            d.nhom[0] = round(d.nhom[0] + lech, 2)
            d.cho_xem_lai.append(f"Tổng các nhóm lệch Total {lech:,.0f}, đã dồn phần lệch vào Current")
        kq.dong.append(d)
    kq.dong.sort(key=lambda d: (-d.nhom_gia_nhat, -d.total))
    # SMS ghi nhiều dòng phí cho một job: gộp theo job cho FIN dễ gạch
    gop: Dict[tuple, list] = {}
    for khach_, job, tien, ngay in kq.sms_cat:
        g = gop.setdefault((khach_, job), [khach_, job, 0.0, ngay])
        g[2] += tien
    kq.sms_cat = [tuple(g) for g in gop.values()]
    gop2: Dict[tuple, list] = {}
    for ma, ten, job, tien, ly_do in kq.sms_khong_misa:
        g = gop2.setdefault((ma, job), [ma, ten, job, 0.0, ly_do])
        g[3] += tien
    kq.sms_khong_misa = [tuple(g) for g in gop2.values()]
    return kq
