# -*- coding: utf-8 -*-
"""Sổ ghi — trí nhớ của tool, thay cho trí nhớ của kế toán.

Mỗi lần chạy ghi một TRANG theo ngày chốt số. Mỗi trang lưu, cho từng khách:
số dư, cách chia 6 nhóm tuổi, các LÔ nợ kèm hạn thanh toán, và các cột do người
điền (Credit Term, Salesman, Ghi chú, Lý do chưa thu hồi).

Nhờ lô nợ có hạn, tuổi nợ tính theo SỐ NGÀY THẬT đã trôi qua, không theo số lần
tool chạy. Nghỉ mấy tuần rồi chạy lại vẫn ra tuổi đúng cho khách cũ.
"""
from __future__ import annotations

import datetime as dt
import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

PHIEN_BAN = 1

# Ranh giới 6 nhóm tuổi, tính theo số ngày quá hạn
TEN_NHOM = ["Current", "1 - 30", "31 - 60", "61 - 90", "91 - 120", "120+"]
# Hạn giả định khi nạp trạng thái ban đầu từ một bảng đã chia sẵn nhóm tuổi:
# lấy giữa khoảng của nhóm đó
GIUA_NHOM = [-15, 15, 45, 75, 105, 150]


def nhom_tu_so_ngay(qua_han: int) -> int:
    if qua_han <= 0:
        return 0
    if qua_han <= 30:
        return 1
    if qua_han <= 60:
        return 2
    if qua_han <= 90:
        return 3
    if qua_han <= 120:
        return 4
    return 5


@dataclass
class Lo:
    """Một khoản nợ kèm hạn thanh toán của chính nó."""
    so_tien: float
    han: str              # ISO date
    uoc_tinh: bool = False

    @property
    def ngay_han(self) -> dt.date:
        return dt.date.fromisoformat(self.han)


@dataclass
class HoSoKhach:
    ma: str = ""
    ten: str = ""
    credit_term: str = ""
    salesman: str = ""
    ghi_chu: str = ""
    ly_do: str = ""
    tong: float = 0.0
    nhom: List[float] = field(default_factory=lambda: [0.0] * 6)
    lo: List[Lo] = field(default_factory=list)
    nguon: str = ""
    job: str = ""
    danh_gia: List[str] = field(default_factory=list)   # WR-13: [bất lợi, có phản hồi, thất hứa] — "Có"/"Không"

    @classmethod
    def tu_dict(cls, d: dict) -> "HoSoKhach":
        d = dict(d)
        d["lo"] = [Lo(**x) for x in d.get("lo", [])]
        d.setdefault("nhom", [0.0] * 6)
        biet = set(cls.__dataclass_fields__)
        return cls(**{k: v for k, v in d.items() if k in biet})

    def thanh_dict(self) -> dict:
        d = asdict(self)
        return d


class SoGhi:
    def __init__(self, duong_dan: Path):
        self.duong_dan = Path(duong_dan)
        self.trang: Dict[str, Dict[str, HoSoKhach]] = {}
        self.ngoai_le: Dict[str, dict] = {}
        # Kế toán khai đè một số thông tin không có trong file Misa,
        # ví dụ credit term mới ký hợp đồng. Khai một lần, dùng mãi.
        self.dieu_chinh: Dict[str, dict] = {}
        # Khách kế toán chủ động loại khỏi bảng dù Misa còn ghi nợ,
        # ví dụ nợ ảo do lỗi xuất hoá đơn. Vẫn hiện ở sheet Cần xem lại.
        self.bo_qua: Dict[str, dict] = {}
        # Bản final người duyệt đã nạp: tên file -> dấu thời gian sửa, để không nạp lại
        self.da_nap: Dict[str, str] = {}
        # v0.4: các con số của sheet Summary từng tuần (ngày chốt ISO -> {chỉ số: giá trị})
        self.lich_su_tuan: Dict[str, dict] = {}
        # v0.4: số còn phải trả từng vendor theo ngày chốt (ISO -> {mã vendor: số dư Có})
        self.phai_tra: Dict[str, Dict[str, float]] = {}
        # trang của ngày chốt nào lấy từ file nào, có phải bản người làm (không có tab Review)
        self.nguon_trang: Dict[str, dict] = {}
        # WR-08: báo cáo do tool ghi: tên file -> giờ sửa lúc tool ghi. Giờ sửa đổi = người đã sửa trên file
        self.file_tool: Dict[str, float] = {}
        if self.duong_dan.exists():
            self._doc()

    # --- đọc ghi ---------------------------------------------------------
    def _doc(self) -> None:
        raw = json.loads(self.duong_dan.read_text(encoding="utf-8"))
        self.ngoai_le = raw.get("ngoai_le", {})
        self.dieu_chinh = raw.get("dieu_chinh", {})
        self.bo_qua = raw.get("bo_qua", {})
        self.da_nap = raw.get("da_nap", {})
        self.lich_su_tuan = raw.get("lich_su_tuan", {})
        self.phai_tra = raw.get("phai_tra", {})
        self.nguon_trang = raw.get("nguon_trang", {})
        self.file_tool = raw.get("file_tool", {})
        for ngay, khach in raw.get("trang", {}).items():
            self.trang[ngay] = {k: HoSoKhach.tu_dict(v) for k, v in khach.items()}

    def ghi(self) -> None:
        self.duong_dan.parent.mkdir(parents=True, exist_ok=True)
        raw = {
            "phien_ban": PHIEN_BAN,
            "cap_nhat": dt.datetime.now().isoformat(timespec="seconds"),
            "ngoai_le": self.ngoai_le,
            "dieu_chinh": self.dieu_chinh,
            "bo_qua": self.bo_qua,
            "da_nap": self.da_nap,
            "lich_su_tuan": dict(sorted(self.lich_su_tuan.items())),
            "phai_tra": dict(sorted(self.phai_tra.items())),
            "nguon_trang": self.nguon_trang,
            "file_tool": self.file_tool,
            "trang": {
                ngay: {k: v.thanh_dict() for k, v in khach.items()}
                for ngay, khach in sorted(self.trang.items())
            },
        }
        tam = self.duong_dan.with_suffix(".tmp")
        tam.write_text(json.dumps(raw, ensure_ascii=False, indent=1), encoding="utf-8")
        tam.replace(self.duong_dan)

    # --- truy cập --------------------------------------------------------
    def ghi_trang(self, ngay_chot: dt.date, khach: Dict[str, HoSoKhach]) -> None:
        """Trang của một ngày chốt đã có thì ghi đè đúng trang đó."""
        self.trang[ngay_chot.isoformat()] = khach

    def trang_truoc(self, ngay_chot: dt.date) -> Optional[Dict[str, HoSoKhach]]:
        cu = [n for n in self.trang if dt.date.fromisoformat(n) < ngay_chot]
        if not cu:
            return None
        return self.trang[max(cu)]

    def phai_tra_truoc(self, ngay_chot: dt.date) -> Optional[Dict[str, float]]:
        cu = [n for n in self.phai_tra if dt.date.fromisoformat(n) < ngay_chot]
        return self.phai_tra[max(cu)] if cu else None

    def ngay_trang_truoc(self, ngay_chot: dt.date) -> Optional[dt.date]:
        cu = [n for n in self.trang if dt.date.fromisoformat(n) < ngay_chot]
        return dt.date.fromisoformat(max(cu)) if cu else None


# --- thao tác trên lô nợ --------------------------------------------------

def lo_tu_nhom(nhom: List[float], ngay_chot: dt.date, uoc_tinh: bool = True) -> List[Lo]:
    """Dựng lô nợ từ một bảng đã chia sẵn nhóm tuổi (nạp ban đầu, hoặc lấy từ
    file Tuổi nợ của Misa)."""
    ra: List[Lo] = []
    for i, tien in enumerate(nhom):
        if abs(tien) < 1:
            continue
        han = ngay_chot - dt.timedelta(days=GIUA_NHOM[i])
        ra.append(Lo(so_tien=round(tien, 2), han=han.isoformat(), uoc_tinh=uoc_tinh))
    return ra


def dieu_chinh_lo(lo: List[Lo], tong_moi: float, ngay_chot: dt.date, term_ngay: int) -> List[Lo]:
    """Nợ tăng thêm là hoá đơn mới — hạn = ngày chốt + credit term.
    Nợ giảm đi là khách đã trả — trừ vào lô cũ nhất trước."""
    lo = [Lo(x.so_tien, x.han, x.uoc_tinh) for x in lo]
    hien_tai = sum(x.so_tien for x in lo)
    chenh = round(tong_moi - hien_tai, 2)
    if chenh > 0.5:
        han = ngay_chot + dt.timedelta(days=term_ngay)
        lo.append(Lo(so_tien=chenh, han=han.isoformat(), uoc_tinh=False))
    elif chenh < -0.5:
        con_tru = -chenh
        lo.sort(key=lambda x: x.han)          # cũ nhất trước
        for x in lo:
            if con_tru <= 0:
                break
            tru = min(x.so_tien, con_tru)
            x.so_tien = round(x.so_tien - tru, 2)
            con_tru = round(con_tru - tru, 2)
        lo = [x for x in lo if x.so_tien > 0.5]
    return lo


def nhom_tu_lo(lo: List[Lo], ngay_chot: dt.date) -> List[float]:
    ra = [0.0] * 6
    for x in lo:
        qua_han = (ngay_chot - x.ngay_han).days
        ra[nhom_tu_so_ngay(qua_han)] += x.so_tien
    return [round(v, 2) for v in ra]


def so_ngay_term(credit_term: str, mac_dinh: int = 30) -> int:
    """'30 days' -> 30 | 'thanh toán ngay' -> 0 | trống -> mặc định."""
    import re as _re
    t = str(credit_term or "").strip().lower()
    if not t:
        return mac_dinh
    if any(x in t for x in ("thanh toán ngay", "trả ngay", "trả trước")) or t in {"cod", "prepay", "prepaid"}:
        return 0
    m = _re.search(r"(\d+)", t)
    return int(m.group(1)) if m else mac_dinh


def han_tu_term(ngay: dt.date, credit_term: str, mac_dinh: int = 30) -> dt.date:
    """Ngày đến hạn của một hoá đơn. '30 days' → +30 ngày. '14 ngày của tháng tiếp theo'
    → ngày 14 của tháng sau tháng hoá đơn. Không ghi gì → +mặc định."""
    import re as _re
    t = str(credit_term or "").strip().lower()
    if _re.search(r"tháng (tiếp theo|sau|kế tiếp)|next month|following month", t):
        m = _re.search(r"(\d+)", t)
        ngay_thang = min(int(m.group(1)), 28) if m else 1
        nam, thang = (ngay.year + 1, 1) if ngay.month == 12 else (ngay.year, ngay.month + 1)
        return dt.date(nam, thang, ngay_thang)
    return ngay + dt.timedelta(days=so_ngay_term(credit_term, mac_dinh))
