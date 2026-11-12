# -*- coding: utf-8 -*-
"""v0.4 — ghi các sheet của full credit report: Summary, AR Risk, Payable, Cash Flow,
Methodology. Sheet Receivables và Review vẫn do baocao.ghi_bao_cao ghi.

Ghi SỐ, không ghi công thức: dashboard (WR-05) đọc file bằng trình duyệt nên không
tính lại được công thức. Cách tính viết bằng lời ở sheet Methodology.
"""
from __future__ import annotations

import datetime as dt
from typing import List, Optional, Sequence

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: F401
from openpyxl.utils import get_column_letter

from . import chitiet as ctm
from . import thuonghieu as th
from .tindung import KetQuaFull

TIEN = "#,##0;[Red]-#,##0"
PHAN_TRAM = "0.0%"
THANG_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_VIEN = Side(style="thin", color="D9D9D9")
VIEN = Border(left=_VIEN, right=_VIEN, top=_VIEN, bottom=_VIEN)


def _ngay_en(d: dt.date) -> str:
    return f"At {d.day:02d} {THANG_EN[d.month - 1]} {d.year}"


def _dau_trang(ws, ten_bao_cao: str, ngay_chot: dt.date, ten_bang: str, ky: str = "") -> None:
    ws["A1"], ws["A2"], ws["A3"], ws["A4"] = th.CONG_TY, ten_bao_cao, _ngay_en(ngay_chot), ten_bang
    if ky:
        ws["A3"] = f"{_ngay_en(ngay_chot)} · Reporting period {ky}"
    ws["A1"].font = Font(name=th.PHONG_TIEU_DE, size=12, bold=True, color=th.TIM)
    ws["A2"].font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True)
    ws["A3"].font = Font(name=th.PHONG_NOI_DUNG, size=10, italic=True)
    ws["A4"].font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True, color=th.CAM)


def bang(ws, hang: int, cot: Sequence[tuple], dong: List[list], tong: bool = True,
         mau_dong: Optional[List[Optional[str]]] = None) -> int:
    """cot: [(tên, loại 'ten'|'tien'|'so'|'pt', độ rộng)]. Trả về dòng kế tiếp còn trống."""
    for i, (ten, loai, rong) in enumerate(cot, start=1):
        o = ws.cell(row=hang, column=i, value=ten)
        o.fill = PatternFill("solid", fgColor=th.TIM if loai in ("tien", "pt") else th.CAM)
        o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
        o.alignment = Alignment(wrap_text=True, vertical="center",
                                horizontal="center" if loai != "ten" else "left")
        if rong:
            ws.column_dimensions[get_column_letter(i)].width = rong
    dau = hang + 1
    hang += 1
    for j, gia_tri in enumerate(dong):
        for i, v in enumerate(gia_tri, start=1):
            loai = cot[i - 1][1]
            if loai in ("tien", "so") and isinstance(v, (int, float)) and abs(v) < 0.4:
                v = None
            o = ws.cell(row=hang, column=i, value=v)
            o.font = Font(name=th.PHONG_NOI_DUNG, size=10)
            o.border = VIEN
            o.alignment = Alignment(vertical="top")
            if loai == "tien":
                o.number_format = TIEN
            elif loai == "pt":
                o.number_format = PHAN_TRAM
            else:
                o.alignment = Alignment(vertical="top", wrap_text=loai == "ten" and cot[i - 1][2] >= 30)
            if mau_dong and mau_dong[j]:
                o.fill = PatternFill("solid", fgColor=mau_dong[j])
        hang += 1
    if tong and dong:
        for i, (ten, loai, _) in enumerate(cot, start=1):
            o = ws.cell(row=hang, column=i)
            o.fill = PatternFill("solid", fgColor=th.TIM)
            o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
            if i == 2 or (i == 1 and cot[1][1] != "ten"):
                o.value = "TOTAL"
            if loai == "tien":
                o.value = round(sum(r[i - 1] for r in dong if isinstance(r[i - 1], (int, float))), 2)
                o.number_format = TIEN
        hang += 1
    ws.freeze_panes = ws.cell(row=dau, column=4 if len(cot) > 6 else 1)
    return hang


# --- WR-12: link hợp đồng (PDF) bấm mở trong trình duyệt -----------------------------------------------------

NHAN_LINK = "Mở hợp đồng"


def la_link(u) -> bool:
    return str(u or "").strip().lower().startswith(("http://", "https://"))


def _o_link(o, url) -> None:
    """Ô 'Mở hợp đồng' có hyperlink tới file hợp đồng; link không phải http(s) thì để trống (không mở file lạ)."""
    if not la_link(url):
        o.value = None
        return
    o.value = NHAN_LINK
    o.hyperlink = str(url).strip()
    o.font = Font(name=th.PHONG_NOI_DUNG, size=10, color="0563C1", underline="single")


# --- Summary ---------------------------------------------------------------

CHI_SO = [
    ("tong_phai_thu", "Total receivables", "tien"),
    ("chua_den_han", "  Not yet due", "tien"),
    ("qua_han", "  Total overdue", "tien"),
    ("qua_han_1_30", "    Overdue 1 - 30 days", "tien"),
    ("qua_han_31_60", "    Overdue 31 - 60 days", "tien"),
    ("qua_han_61_90", "    Overdue 61 - 90 days", "tien"),
    ("qua_han_91_120", "    Overdue 91 - 120 days", "tien"),
    ("qua_han_120", "    Overdue over 120 days", "tien"),
    ("so_khach", "Customers with balance", "so"),
    ("no_moi", "New invoices this week", "tien"),
    ("da_thu", "Collected this week (all credits)", "tien"),
    ("thu_ngan_hang", "  of which: bank", "tien"),
    ("thu_tien_mat", "  of which: cash", "tien"),
    ("thu_can_tru", "  of which: offset with payables", "tien"),
    ("thu_ty_gia_phi", "  of which: FX & bank fees", "tien"),
    ("thu_khac", "  of which: other", "tien"),
    ("no_xau", "Bad debt (overdue of N3 + N4)", "tien"),
    ("so_khach_no_xau", "  Bad-debt customers", "so"),
    ("no_xau_phap_ly", "  of which legal event (N4)", "tien"),
    ("no_xau_120", "  of which over 120 days", "tien"),
    ("ceo_can_xem", "Customers CEO must review (overdue ≥ 100M, no reason)", "so"),
    ("tong_phai_tra", "Total payables", "tien"),
    ("no_vendor_moi", "New vendor bills this week", "tien"),
    ("da_tra_vendor", "Paid to vendors this week", "tien"),
    ("so_vendor", "Vendors with balance", "so"),
    ("dong_tien_rong", "Net cash (bank + cash collected − paid to vendors)", "tien"),
    # WR-08: hợp đồng / credit term (sheet Contracts)
    ("hd_het_han", "Contracts expired, customer still owes", "so"),
    ("hd_sap_het_30", "Contracts expiring within 30 days", "so"),
    ("khach_chua_term", "Customers owing without a credit term in the credit term file", "so"),
    ("term_doi", "Credit terms changed since last week", "so"),
    ("khach_chua_han_hd", "Customers owing without contract expiry date", "so"),
    ("khach_chua_hd", "Customers owing without a signed contract (credit term file ✓; blank = column missing)", "so"),
    # WR-11
    ("khach_chua_master", "New customers owing, not in the credit term file before this run (end of Contracts)", "so"),
    ("khach_chua_han_muc", "Customers owing without a credit limit in the credit term file", "so"),
    # WR-15: 5 mức màu Head FIN (AR Risk, cột Risk level)
    ("muc_4", "Risk level · Tím (customers)", "so"),
    ("muc_3", "Risk level · Đỏ (customers)", "so"),
    ("muc_2", "Risk level · Cam (customers)", "so"),
    ("muc_1", "Risk level · Vàng (customers)", "so"),
    ("muc_0", "Risk level · Trắng (customers)", "so"),
    ("trang_theo_doi", "  of which Trắng ↑ (within term, watch T1–T3)", "so"),
    ("no_rui_ro_cao", "  Balance of Đỏ + Tím customers", "tien"),
    ("nguong_x", "  X (T1 threshold: overdue ≥)", "tien"),
    ("hang_d", "Payment-history grade D (customers, reference)", "so"),
    ("sales_chua_giai_trinh", "Salesmen with overdue customers and no reason this week (table at the end)", "so"),
]

LICH_SU = [
    ("tuan", "Week", "ten", 8), ("den_ngay", "Period end", "ten", 12),
    ("tong_phai_thu", "Total AR", "tien", 16), ("chua_den_han", "Not yet due", "tien", 15),
    ("qua_han", "Overdue", "tien", 15), ("qua_han_61_90", "61 - 90", "tien", 14),
    ("qua_han_91_120", "91 - 120", "tien", 14), ("qua_han_120", "120+", "tien", 14),
    ("da_thu", "Collected", "tien", 15), ("da_thu_tien", "Collected (bank + cash)", "tien", 15),
    ("no_moi", "New invoices", "tien", 15), ("no_xau", "Bad debt", "tien", 15),
    ("tong_phai_tra", "Total AP", "tien", 15), ("da_tra_vendor", "Paid to vendors", "tien", 15),
    ("no_vendor_moi", "New vendor bills", "tien", 15), ("dong_tien_rong", "Net cash", "tien", 15),
]


def _thay_doi(cs: dict, tr: dict, khoa: str) -> str:
    nay, cu = cs.get(khoa), tr.get(khoa)
    if not isinstance(cu, (int, float)) or not cu or nay is None:
        return "no last-week figure"
    return f"{'▲' if nay >= cu else '▼'} {abs(nay - cu):,.0f} vs last week ({(nay - cu) / cu:+.1%})"


def _o_so(ws, hang: int, a: str, b: str, nhan: str, gia_tri, phu: str, mau: str) -> None:
    """Một ô số lớn kiểu thẻ: nhãn nhỏ · số to · dòng phụ."""
    for r, v, f in ((hang, nhan, Font(name=th.PHONG_NOI_DUNG, size=9, bold=True, color=mau)),
                    (hang + 1, gia_tri, Font(name=th.PHONG_TIEU_DE, size=16, bold=True, color=mau)),
                    (hang + 2, phu, Font(name=th.PHONG_NOI_DUNG, size=8, color="7F7F7F"))):
        if a != b:
            ws.merge_cells(f"{a}{r}:{b}{r}")
        o = ws[f"{a}{r}"]
        o.value = v
        o.font = f
        o.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        if isinstance(v, (int, float)):
            o.number_format = "#,##0"
        for c in range(ord(a), ord(b) + 1):
            ws[f"{chr(c)}{r}"].fill = PatternFill("solid", fgColor="F7F5FB")
    ws[f"{a}{hang}"].border = Border(top=Side(style="medium", color=mau))
    ws.row_dimensions[hang + 1].height = 24


def ghi_summary(wb, kq: KetQuaFull, ngay_chot: dt.date, lich_su: List[dict]) -> None:
    ws = wb.create_sheet(th.SHEET_SUMMARY)
    cs, tr = kq.chi_so, kq.chi_so_truoc
    ky = f"{dt.date.fromisoformat(cs['tu_ngay']):%d/%m/%Y} - {dt.date.fromisoformat(cs['den_ngay']):%d/%m/%Y}"
    _dau_trang(ws, "Weekly Credit Summary", ngay_chot, "Key figures", ky)
    tong = cs.get("tong_phai_thu") or 0
    o_so = [   # 3 hàng thẻ: phải thu · phải trả · dòng tiền. Cột F hẹp nên chỉ để số đếm.
        ("RECEIVABLES · Total", cs.get("tong_phai_thu"), _thay_doi(cs, tr, "tong_phai_thu"), th.TIM),
        ("Overdue", cs.get("qua_han"), f"{(cs.get('qua_han') or 0) / tong:.0%} of receivables" if tong else "",
         th.CAM),
        ("Bad debt (N3 + N4)", cs.get("no_xau"), f"{cs.get('so_khach_no_xau', 0)} customers", "C00000"),
        ("CEO must review", cs.get("ceo_can_xem"), "overdue ≥ 100M, no reason", "C00000"),
        ("PAYABLES · Total", cs.get("tong_phai_tra"), _thay_doi(cs, tr, "tong_phai_tra"), th.TIM),
        ("Paid to vendors this week", cs.get("da_tra_vendor"), "Misa TK331 debit", th.CAM),
        ("New vendor bills this week", cs.get("no_vendor_moi"), "Misa TK331 credit", th.CAM),
        ("Vendors", cs.get("so_vendor"), "with balance", th.TIM),
        ("CASH · Collected this week", cs.get("da_thu"), "all credits to TK131", "2E7D32"),
        ("Collected to bank + cash", cs.get("da_thu_tien"), "real money in", "2E7D32"),
        ("Net cash this week", cs.get("dong_tien_rong"), "bank + cash collected − paid to vendors", "2E7D32"),
        ("Customers", cs.get("so_khach"), "with balance", th.TIM),
    ]
    vi_tri = [("A", "A"), ("B", "C"), ("D", "E"), ("F", "F")]
    for i, (nhan, gia_tri, phu, mau) in enumerate(o_so):
        r = 6 + 4 * (i // 4)
        a, b = vi_tri[i % 4]
        _o_so(ws, r, a, b, nhan, gia_tri, phu, mau)
    ws.cell(row=18, column=1, value="This week vs last week").font = Font(
        name=th.PHONG_TIEU_DE, size=11, bold=True, color=th.CAM)
    dong = []
    for khoa, nhan, loai in CHI_SO:
        if khoa not in cs:
            continue
        nay, cu = cs.get(khoa), tr.get(khoa)
        chenh = (nay - cu) if isinstance(cu, (int, float)) else None
        ty_le = (chenh / cu) if chenh is not None and cu else None
        dong.append([nhan, cu, nay, chenh, ty_le, khoa])
    hang = bang(ws, 19, [("Metric", "ten", 52), ("Last week", "tien", 18), ("This week", "tien", 18),
                        ("Change", "tien", 18), ("Change %", "pt", 11), ("key", "ten", 18)],
                dong, tong=False)
    for r in range(20, hang):
        ws.cell(row=r, column=6).font = Font(name=th.PHONG_NOI_DUNG, size=8, color="A6A6A6")
    hang += 1
    ws.cell(row=hang, column=1, value="Weekly history").font = Font(
        name=th.PHONG_TIEU_DE, size=11, bold=True, color=th.CAM)
    hang += 1
    cot = [(nhan, loai, rong) for _, nhan, loai, rong in LICH_SU]
    dl = [[x.get(k) for k, *_ in LICH_SU] for x in lich_su]
    bang(ws, hang, cot, dl, tong=False)
    # bảng lịch sử dùng chung cột với bảng trên: đặt lại độ rộng cho bảng trên dễ đọc
    hang = _bang_sales_chua_giai_trinh(ws, kq, hang + len(dl) + 3)
    for chu, rong in zip("ABCDEF", (52, 18, 18, 18, 15, 18)):
        ws.column_dimensions[chu].width = rong
    ws.freeze_panes = None
    ws.sheet_view.showGridLines = False


def _bang_sales_chua_giai_trinh(ws, kq: KetQuaFull, hang: int) -> int:
    """WR-11 (ảnh 2: 'thông báo sale nào chưa điền lý do'): mỗi Sales — số khách đang quá hạn mà cột lý do tuần
    này còn trống, tổng quá hạn, tên khách. Sales ghi lý do ở Receivables, cột Reason for late payment."""
    gom: dict = {}
    for r in kq.rui_ro:
        if r.qua_han > 0.5 and r.s[3]:
            g = gom.setdefault(r.salesman or "(chưa có Sales)", [0, 0.0, []])
            g[0] += 1
            g[1] += r.qua_han
            g[2].append(r.ma or r.ten)
    ws.cell(row=hang, column=1, value="Reasons missing this week — by salesman").font = Font(
        name=th.PHONG_TIEU_DE, size=11, bold=True, color=th.CAM)
    if not gom:
        ws.cell(row=hang + 1, column=1, value="Mọi khách quá hạn đã có lý do tuần này.").font = Font(
            name=th.PHONG_NOI_DUNG, size=10, italic=True)
        return hang + 2
    dong = [[s, n, v, ", ".join(ds)] for s, (n, v, ds) in sorted(gom.items(), key=lambda x: -x[1][1])]
    return bang(ws, hang + 1, [("Salesman", "ten", 52), ("Customers", "so", 18), ("Overdue", "tien", 18),
                               ("Customer codes", "ten", 18)], dong, tong=False)


# --- AR Risk ---------------------------------------------------------------

COT_RISK = [
    ("No.", "so", 6), ("Code", "ten", 14), ("Customer's Name", "ten", 40), ("Credit Term", "ten", 14),
    ("Salesman", "ten", 12), ("Jobs by Salesman", "ten", 40), ("Opening (Misa)", "tien", 15), ("Increase", "tien", 15),
    ("Decrease", "tien", 15), ("Ending = Total", "tien", 15), ("Backdated vs last report", "tien", 15),
    ("Not yet due", "tien", 14), ("1 - 30", "tien", 14), ("31 - 60", "tien", 14), ("61 - 90", "tien", 14),
    ("91 - 120", "tien", 14), ("120+", "tien", 14), ("Total overdue", "tien", 15),
    ("Last week total", "tien", 15), ("Last week overdue", "tien", 15), ("Last week 61+", "tien", 15),
    ("S1", "so", 5), ("S2", "so", 5), ("S3", "so", 5), ("S4", "so", 5), ("S5", "so", 5), ("S6", "so", 5),
    ("Promise date", "ten", 11), ("S7", "so", 5), ("Signals (S1-S3, S7)", "so", 10), ("Risk Code", "ten", 14), ("Risk Tier", "ten", 26),
    ("CEO review", "ten", 9), ("Trend (auto)", "ten", 50),
    ("Reason last week", "ten", 36),       # WR-11 PA1: lý do Sales ghi tuần trước (chỉ để đọc) — bên trái tuần này
    ("Reason for late payment", "ten", 36),
    ("Required action", "ten", 40), ("Contract (PDF)", "ten", 14), ("Job Number", "ten", 30),
    # WR-13 (FIN 05/11): đánh giá theo tiêu chí FIN — số do tool tính (3 ô Sales đã bỏ 12/11)
    ("Max days overdue", "so", 9), ("First invoice", "ten", 11), ("Customer group", "ten", 13),
    # WR-15 (Head FIN v3, 06/11): 5 mức màu — số liệu T1–T4 do tool tính; lên thẳng Đỏ + mức cuối là công thức (đổi theo
    # lý do Sales ghi). Tham số X, Y, N ghi ở đầu sheet; FIN sửa ở sheet "Tham số" của file credit term.
    ("Billing avg / month", "tien", 15), ("Credit limit", "tien", 15), ("Credit limit source", "ten", 11),
    ("Late invoices", "so", 8), ("Do not upgrade", "ten", 9), ("Base level (0-4)", "so", 8),
    ("T1: Overdue ≥ X", "so", 7), ("T2: Total > credit limit", "so", 8), ("T3: ≥ N late invoices", "so", 8),
    ("T4: New customer ≥ Y", "so", 10), ("Upgrade +1", "so", 7), ("Straight to Red", "so", 7),
    ("Final level (0-4)", "so", 7), ("Watch (not overdue)", "ten", 10), ("Risk level", "ten", 13), ("Basis", "ten", 40),
    ("Suggested payment terms", "ten", 26), ("Action plan (FIN)", "ten", 60),
    # điểm lịch sử thanh toán A–D (tham khảo, không ảnh hưởng mức)
    ("Invoices scored", "so", 8), ("Collection rate", "pt", 9), ("Overdue / balance", "pt", 9),
    ("Paid-in-full rate", "pt", 9), ("Score", "so", 7), ("Grade", "ten", 7),
]
MAU_HANG = {"A": ("C6EFCE", "006100"), "B": ("DDEBF7", "1F4E78"), "C": ("FFEB9C", "9C5700"), "D": ("FFC7CE", "9C0006")}


def cong_thuc_risk(h: int, ngay_chot: Optional[dt.date] = None) -> dict:
    """WR-08: công thức của Head of FIN (sheet 'data W39 so W40' của chị) cho dòng h của AR Risk.
    Cột tiền là số; các cột tính ra (S1–S6, nhóm, xu hướng, việc cần làm) là công thức, lý do lấy thẳng
    từ Receivables — Sales sửa lý do ở Receivables thì nhóm rủi ro đổi ngay khi mở file."""
    from . import thuonghieu as th_
    from .tindung import NGUONG_CEO, SO_DAU_HIEU_NO_XAU, TU_KHOA_LOAI_TRU, TU_KHOA_PHAP_LY
    L = {ten: f"{get_column_letter(i)}{h}" for i, (ten, _, _) in enumerate(COT_RISK, start=1)}
    J, E = L["Ending = Total"], L["Credit Term"]
    nd, b1, b2, b3, b4, b5 = (L[x] for x in ("Not yet due", "1 - 30", "31 - 60", "61 - 90", "91 - 120", "120+"))
    Q, U, V, W = L["Total overdue"], L["Last week total"], L["Last week overdue"], L["Last week 61+"]
    S1, S2, S3, S4, S5, S6, S7 = (L[f"S{i}"] for i in range(1, 8))
    PD, SIG = L["Promise date"], L["Signals (S1-S3, S7)"]
    nc = ngay_chot or dt.date.today()
    chot = f"DATE({nc.year},{nc.month},{nc.day})"
    LD, TIER = L["Reason for late payment"], L["Risk Tier"]
    nhan = th_.NHAN_TU_DONG
    LW = L["Reason last week"]
    _ld = get_column_letter(th.COT_LY_DO_NAY)          # cột lý do tuần này ở Receivables (O từ WR-12)
    # lý do để xét nhóm: lý do tuần này; Sales chưa ghi (S4) thì lý do tuần trước (WR-11 PA1)
    goc = f'IF({S4}=1,{LW},{LD})'
    tim = lambda ds: "OR(" + ",".join(f'ISNUMBER(SEARCH("{t}",{goc}))' for t in ds) + ")"
    ct = {
        "Total overdue": f"=SUM({b1}:{b5})",
        "S1": f"=IF(AND({U}>0,{J}-{U}>0.5,{Q}-{V}>=-0.5),1,0)",
        "S2": f"=IF(AND({Q}>0.5,ABS({Q}-{V})<1),1,0)",
        "S3": f"=IF(({b3}+{b4}+{b5})-{W}>0.5,1,0)",
        "S4": f'=IF(OR(TRIM({LD})="",LEFT(TRIM({LD}),{len(nhan)})="{nhan}"),1,0)',
        "S5": f'=IF(TRIM({E})="",1,0)',
        "S6": f"=IF(AND({J}>0.5,ABS({Q}-{J})<1),1,0)",
        # S7 lỡ hẹn (user 13/10): Promise date do tool đọc từ lý do, FIN sửa tay được
        "S7": f'=IF(AND(ISNUMBER({PD}),{PD}<{chot},{Q}>0.5,{Q}-{V}>=-0.5),1,0)',
        "Signals (S1-S3, S7)": f"=SUM({S1}:{S3})+{S7}",
        "Risk Code": (f'=TRIM(IF({S1}=1,"S1 ","")&IF({S2}=1,"S2 ","")&IF({S3}=1,"S3 ","")&IF({S4}=1,"S4 ","")'
                      f'&IF({S5}=1,"S5 ","")&IF({S6}=1,"S6 ","")&IF({S7}=1,"S7",""))'),
        "Risk Tier": (f'=IF({Q}<0.5,"0·Chưa đến hạn",IF({tim(TU_KHOA_LOAI_TRU)},"X·Không phải rủi ro tín dụng",'
                      f'IF({tim(TU_KHOA_PHAP_LY)},"N4·Nợ xấu–sự kiện pháp lý",IF(OR({b5}>0.5,'
                      f'{SIG}>={SO_DAU_HIEU_NO_XAU}),"N3·Nợ xấu",IF({SIG}>=1,'
                      f'"N2·Cần theo dõi","N1·Quá hạn bình thường")))))'),
        "CEO review": f'=IF(AND({Q}>={NGUONG_CEO},{S4}=1),"Yes","")',
        "Trend (auto)": (f'=TRIM(IF({Q}<0.5,"Chưa đến hạn thanh toán. ","")&IF({U}<0.5,"Khách mới xuất hiện trong '
                         f'bảng tuần này. ","")&IF({S1}=1,"Khách phát sinh lô mới nhưng nợ cũ không giảm. ","")'
                         f'&IF({S2}=1,"Quá hạn đứng yên so với tuần trước, chưa thu được đồng nào. ","")'
                         f'&IF({S7}=1,"Đã quá ngày hẹn trả ("&TEXT({PD},"dd/mm/yyyy")&") mà quá hạn chưa giảm. ","")'
                         f'&IF(AND({U}>0.5,{Q}>0.5,{Q}<{V}-0.5),"Đã thu bớt "&TEXT({V}-{Q},"#,##0")&" so với tuần '
                         f'trước. ","")&IF({S3}=1,"Nợ đang già đi, tiền chuyển sang nhóm tuổi cao hơn. ","")'
                         f'&IF({S6}=1,"Toàn bộ dư nợ đã quá hạn. ","")&IF({b5}>0.5,"Có "&TEXT({b5},"#,##0")&" nằm ở '
                         f'nhóm trên 120 ngày. ","")&IF({S5}=1,"Chưa có điều khoản thanh toán nên chưa xác định chắc '
                         f'chắn được quá hạn. ","")&IF({S4}=1,"CHƯA CÓ GIẢI TRÌNH TỪ SALES. ",""))'),
        "Reason for late payment": (f"=IFERROR(INDEX('{th.SHEET_BANG}'!${_ld}:${_ld},MATCH($B{h},'{th.SHEET_BANG}'!$B:$B,0))"
                                    f'&"","")'),
        "Required action": (f'=IF(LEFT({TIER},2)="N4","Chuyển hồ sơ thu hồi. Rà soát chứng từ gốc để xét trích dự '
                            f'phòng theo TT48/2019.",IF(LEFT({TIER},2)="N3",IF({b5}>0.5,"Dừng cấp dịch vụ mới. Chốt '
                            f'công văn đòi nợ có xác nhận gửi. Kế toán trưởng ký xác nhận phân nhóm.","Dừng cấp dịch '
                            f'vụ mới. Yêu cầu cam kết NGÀY cụ thể bằng văn bản trong 7 ngày."),""))'),
    }
    ct.update(cong_thuc_fin(L, h))
    return ct


def cong_thuc_fin(L: dict, h: int) -> dict:
    """WR-15: cùng luật với wr/danhgia.py (muc_cuoi5, nang_thang, muc3_tu_muc5, hinh_thuc, phuong_an) và dashboard —
    sửa một chỗ thì sửa cả ba. Lên thẳng Đỏ theo lý do ở Receivables: Sales sửa lý do thì mức đổi ngay khi mở file.
    T1–T4, Upgrade +1, Base level là SỐ do tool tính. (FIN 12/11: bỏ 3 ô Sales đánh giá.)"""
    from . import danhgia as dg
    G = L["Customer group"]
    BASE, UP, STR, LV = L["Base level (0-4)"], L["Upgrade +1"], L["Straight to Red"], L["Final level (0-4)"]
    LD, LW = L["Reason for late payment"], L["Reason last week"]
    q = lambda t: '"' + t.replace('"', '""') + '"'
    m3 = f"IF({LV}>=3,3,IF({LV}=2,2,1))"
    pa = lambda nhom: "CHOOSE(" + m3 + "," + ",".join(q(dg.PHUONG_AN[(nhom, m)]) for m in (1, 2, 3)) + ")"
    ly_do = f'IF(TRIM({LD})="",{LW},{LD})'
    tu = ",".join(f'ISNUMBER(SEARCH("{k}",{ly_do}))' for k in dg.TU_KHOA_NANG_THANG)
    return {
        "Straight to Red": f"=IF(OR({tu}),1,0)",
        "Final level (0-4)": f"=MIN(4,IF({STR}=1,MAX({BASE},3),{BASE}+{UP}))",
        # WR-15b: khách trong hạn có dấu hiệu T1–T4 vẫn Trắng, chỉ gắn "↑T…" (cột Watch)
        "Risk level": ("=CHOOSE(" + LV + "+1," + ",".join(f'"{m}"' for m in dg.MUC5) + ")"
                       + f'&IF(AND({LV}=0,{L["Watch (not overdue)"]}<>"")," "&{L["Watch (not overdue)"]},"")'),
        "Suggested payment terms": (f'=IF({m3}=3,IF({G}="Mới",{q(dg.HINH_THUC[3][0])},{q(dg.HINH_THUC[3][1])}),'
                                    f'IF({m3}=2,{q(dg.HINH_THUC[2])},{q(dg.HINH_THUC[1])}))'),
        "Action plan (FIN)": f'=IF({G}="Mới",{pa("Mới")},IF({G}="Thường xuyên",{pa("Thường xuyên")},{pa("Lâu năm")}))',
    }


def _can_cu_cong_thuc(L: dict, tinh: str) -> str:
    """Basis: phần tool tính (tuổi nợ, T1–T4) + phần đổi theo lý do Sales ghi."""
    STR = L["Straight to Red"]
    tinh = tinh.replace('"', '""')
    return f'=TRIM("{tinh} "&IF({STR}=1,"Lên thẳng Đỏ (lý do: mất liên lạc / hứa không trả…).",""))'


def ghi_risk(wb, kq: KetQuaFull, ngay_chot: dt.date) -> None:
    from openpyxl.formatting.rule import FormulaRule
    ws = wb.create_sheet(th.SHEET_RISK)
    _dau_trang(ws, "Receivables Movement & Risk", ngay_chot, "Risk tier per customer")
    ws["F3"] = ("Các cột S1–S7, Signals, Risk Code/Tier, CEO, Trend, Reason, Required action là CÔNG THỨC "
                "(luật Head of FIN). Lý do sửa ở sheet Receivables. Cột Risk level: 5 màu Head FIN (chỉ tô ở sheet này).")
    ws["F3"].font = Font(name=th.PHONG_NOI_DUNG, size=9, italic=True, color="7F7F7F")
    dong = []
    for i, r in enumerate(kq.rui_ro, start=1):
        # giá trị python vẫn ghi trước (để bang() định dạng), rồi thay bằng công thức ở dưới
        dong.append([i, r.ma, r.ten, r.credit_term, r.salesman, r.chia_sales, r.dau_ky, r.tang, r.giam,
                     r.cuoi_ky, r.nhap_bo_sung, *r.nhom, r.qua_han, r.tt_total, r.tt_qua_han, r.tt_61, *r.s, r.ngay_hen, r.s7,
                     r.so_dau_hieu, r.ma_dau_hieu, r.nhom_rui_ro, "Yes" if r.ceo else "",
                     r.xu_huong, r.ly_do_truoc, r.ly_do, r.viec_can_lam, NHAN_LINK if la_link(r.link_hd) else "", r.job,
                     r.ngay_qh, r.ngay_hd_dau, r.nhom_kh, r.billing or None, r.han_muc_ap,
                     r.nguon_han_muc, r.so_hd_tre, "Có" if r.khong_nang else "", r.muc_goc, *r.tieu_chi, r.nang,
                     r.nang_thang, r.muc5, r.theo_doi, r.ten_muc, r.can_cu,
                     r.hinh_thuc, r.phuong_an, r.cham.get("so_hd"), r.cham.get("ty_le_thu_hoi"),
                     r.cham.get("ty_le_qua_han"), r.cham.get("ty_le_tra_du"), r.diem, r.hang])
    hang = bang(ws, 5, COT_RISK, dong)
    vi_tri = {ten: i for i, (ten, _, _) in enumerate(COT_RISK, start=1)}
    so_cot = {"Not yet due", "1 - 30", "31 - 60", "61 - 90", "91 - 120", "120+", "Last week total",
              "Last week overdue", "Last week 61+", "Max days overdue", "Late invoices", "Base level (0-4)",
              "T1: Overdue ≥ X", "T2: Total > credit limit", "T3: ≥ N late invoices", "T4: New customer ≥ Y",
              "Upgrade +1"}
    for h in range(6, 6 + len(dong)):
        for ten in so_cot:                     # ô số 0 phải là 0 (không để trống) để công thức so sánh đúng
            o = ws.cell(row=h, column=vi_tri[ten])
            if o.value is None:
                o.value = 0
        o = ws.cell(row=h, column=vi_tri["Promise date"])
        o.number_format = "dd/mm/yyyy"
        o.alignment = Alignment(horizontal="center", vertical="top")
        Lh = {ten: f"{get_column_letter(i)}{h}" for i, (ten, _, _) in enumerate(COT_RISK, start=1)}
        ct_h = cong_thuc_risk(h, ngay_chot)
        ct_h["Basis"] = _can_cu_cong_thuc(Lh, kq.rui_ro[h - 6].can_cu_goc)
        for ten, ct in ct_h.items():
            o = ws.cell(row=h, column=vi_tri[ten], value=ct)
            if ten.startswith("S") and len(ten) == 2 or ten.startswith("Signals"):
                o.alignment = Alignment(horizontal="center", vertical="top")
            elif ten in ("Trend (auto)", "Reason for late payment", "Required action", "Basis", "Action plan (FIN)",
                         "Suggested payment terms"):
                o.alignment = Alignment(vertical="top", wrap_text=True)
            else:
                o.alignment = Alignment(horizontal="center", vertical="top")
        _o_link(ws.cell(row=h, column=vi_tri["Contract (PDF)"]), kq.rui_ro[h - 6].link_hd)
        o = ws.cell(row=h, column=vi_tri["Reason last week"])
        o.font = Font(name=th.PHONG_NOI_DUNG, size=10, italic=True, color="595959")
        o.alignment = Alignment(vertical="top", wrap_text=True)
        o = ws.cell(row=h, column=vi_tri["First invoice"])
        o.number_format = "dd/mm/yyyy"
        o.alignment = Alignment(horizontal="center", vertical="top")
        for ten in ("Max days overdue", "Customer group", "Score", "Grade", "Invoices scored", "Credit limit source",
                    "Late invoices", "Do not upgrade", "Base level (0-4)", "T1: Overdue ≥ X", "T2: Total > credit limit",
                    "T3: ≥ N late invoices", "T4: New customer ≥ Y", "Upgrade +1"):
            ws.cell(row=h, column=vi_tri[ten]).alignment = Alignment(horizontal="center", vertical="top")
    if dong:
        # WR-15: 5 màu Head FIN, CHỈ tô ô Risk level (và Grade) — không tô cả hàng (user 06/11: nhìn đau mắt).
        # Risk Tier N0–N4 giữ nguyên, không tô.
        cuoi = 5 + len(dong)
        for ten_cot, bang_mau in (("Risk level", th.MAU_MUC_FIN), ("Grade", MAU_HANG)):
            cl = get_column_letter(vi_tri[ten_cot])
            for gia_tri, (nen, chu) in bang_mau.items():      # LEFT: "Trắng ↑T3" vẫn là Trắng
                ws.conditional_formatting.add(f"{cl}6:{cl}{cuoi}", FormulaRule(
                    formula=[f'LEFT(${cl}6,{len(gia_tri)})="{gia_tri}"'], fill=th.nen_cf(nen),
                    font=Font(color=chu, bold=True)))
    ws.auto_filter.ref = f"A5:{get_column_letter(len(COT_RISK))}{5 + len(dong)}"
    # WR-15: tham số tiêu chí nâng hạng (Head FIN) ghi ở đầu sheet, trên các cột T1–T4
    c0 = vi_tri["T1: Overdue ≥ X"]
    ts = kq.tham_so or {}
    lg = [v for v in ts.get("lai_gop") or [] if isinstance(v, (int, float)) and v > 0]
    for i, (nhan, v, ghi) in enumerate((
            ("X (T1): overdue ≥", kq.x, f"X = MIN(X1, X2). X1 = {ts.get('ty_le_lai_gop', 0):.0%} lãi gộp bình quân/tháng"
             + (f" {ts.get('quy') or ''} = {kq.x1:,.0f}" if lg else " — CHƯA NHẬP lãi gộp, X tạm = X2")
             + f" · X2 = nợ quá hạn khách cuối top {ts.get('top_qua_han', 0):.0%} = {kq.x2:,.0f}"),
            ("Y (T4): new customer ≥", kq.y, "Khách mới (tuần trước chưa có dư nợ) nợ từ Y"
             + ("" if isinstance(ts.get("y"), (int, float)) else " — Y = X / 2")),
            ("N (T3): late invoices ≥", ts.get("so_hd_tre", 2),
             f"Hoá đơn trả trễ hơn hạn quá {ts.get('ngay_tre', 15)} ngày trong {ts.get('so_thang', 6)} tháng. "
             "Sửa tham số: sheet 'Tham số' của file credit term")), start=1):
        ws.cell(row=i, column=c0 - 2, value=nhan).font = Font(name=th.PHONG_NOI_DUNG, size=9, bold=True)
        o = ws.cell(row=i, column=c0, value=v)
        o.number_format, o.font = TIEN, Font(name=th.PHONG_NOI_DUNG, size=9, bold=True, color=th.TIM)
        ws.cell(row=i, column=c0 + 1, value=ghi).font = Font(name=th.PHONG_NOI_DUNG, size=9, italic=True, color="7F7F7F")
    if kq.khach_tra_xong:
        hang += 1
        ws.cell(row=hang, column=2, value="Customers fully paid since last report").font = Font(
            name=th.PHONG_TIEU_DE, size=11, bold=True, color=th.CAM)
        bang(ws, hang + 1, [("No.", "so", 6), ("Customer's Name", "ten", 14), ("Last report balance", "tien", 40)],
             [[i, t, v] for i, (t, v) in enumerate(kq.khach_tra_xong, start=1)])
    ws.freeze_panes = "D6"


# --- WR-15: Đánh giá + Tham số (bố cục file v3 của Head FIN) --------------------------------------------------------

COT_DANH_GIA = [("No.", "so", 6), ("Code", "ten", 14), ("Customer's Name", "ten", 40), ("Credit Term", "ten", 18),
                ("Total Outstanding", "tien", 16), ("Overdue Amount", "tien", 16),
                ("Billing TB/tháng (6 tháng gần nhất)", "tien", 16), ("Credit limit trong hợp đồng", "tien", 16),
                ("Hệ số điều khoản (tháng)", "so", 10), ("Credit limit tạm tính", "tien", 16),
                ("Credit limit áp dụng", "tien", 16), ("Nguồn credit limit", "ten", 13), ("Xếp hạng Overdue", "so", 9),
                ("T1: Overdue ≥ X", "so", 8), ("T2: Total > Credit limit", "so", 8), ("Số HĐ trả trễ", "so", 8),
                ("T3: ≥ N HĐ trễ", "so", 8), ("T4: Khách mới ≥ Y", "so", 10),
                ("Không nâng hạng", "ten", 9), ("KẾT QUẢ", "ten", 20), ("Mức gốc", "ten", 9), ("Risk level", "ten", 11),
                ("Ghi chú", "ten", 50)]


def ghi_danh_gia(wb, kq: KetQuaFull, ngay_chot: dt.date) -> None:
    """Sheet 'Đánh giá' (từng khách: hạn mức, T1–T4, kết quả nâng hạng) và 'Tham số' (số tool đã dùng tuần này). Cùng bố
    cục file v3 của Head FIN, nhưng tool tự điền — không ai nhập tay ở đây. Tham số sửa ở file credit term."""
    from openpyxl.formatting.rule import FormulaRule
    from . import danhgia as dg
    ws = wb.create_sheet(th.SHEET_DANH_GIA)
    _dau_trang(ws, "Đánh giá nâng hạng rủi ro công nợ", ngay_chot, "Theo từng khách hàng (Head of FIN)")
    rr = kq.rui_ro
    so_nang = sum(1 for r in rr if r.nang)
    ws["F2"] = (f"X áp dụng (VND): {kq.x:,.0f}   |   Y: {kq.y:,.0f}   |   N: {(kq.tham_so or {}).get('so_hd_tre', 2)} HĐ"
                f"   |   Số khách nâng hạng: {so_nang}/{len(rr)}")
    ws["F2"].font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True, color=th.TIM)
    ws["F3"] = ("Tool tự tính mỗi tuần — không nhập tay ở đây. Hạn mức hợp đồng, Không nâng hạng: file credit term. Lãi "
                "gộp và tỷ lệ: sheet Tham số của file credit term. Risk level cuối (sau lý do Sales): sheet AR Risk.")
    ws["F3"].font = Font(name=th.PHONG_NOI_DUNG, size=9, italic=True, color="7F7F7F")
    qh = sorted({r.qua_han for r in rr if r.qua_han > 0.5}, reverse=True)
    hang_qh = {v: i for i, v in enumerate(qh, start=1)}
    dong = []
    for i, r in enumerate(rr, start=1):
        tam = (round(r.billing * r.he_so * float((kq.tham_so or {}).get("he_so_an_toan") or 1), 2)
               if r.he_so is not None and r.billing > 0 else None)
        ghi = []
        if r.nguon_han_muc == "Thiếu dữ liệu" and r.cuoi_ky > 0.5:
            ghi.append("T2 chưa đánh giá được: thiếu billing / term / credit limit.")
        dat = [t for t, v in zip(("T1", "T2", "T3", "T4"), r.tieu_chi) if v]
        if dat:
            ghi.append("Thỏa " + ", ".join(dat) + (" — khách ✓ không nâng hạng." if r.khong_nang else
                                                   " — đang trong hạn: giữ Trắng, theo dõi." if r.theo_doi else "."))
        if not (r.credit_term or "").strip() or r.co_hd is False:
            ghi.append("Bổ sung: " + " · ".join(x for x, c in (("credit term", not (r.credit_term or "").strip()),
                                                              ("✓ hợp đồng", r.co_hd is False)) if c)
                       + " (sheet Contracts) — không nâng màu.")
        dong.append([i, r.ma, r.ten, r.credit_term, r.cuoi_ky, r.qua_han, r.billing or None, r.han_muc or None, r.he_so,
                     tam, r.han_muc_ap, r.nguon_han_muc, hang_qh.get(r.qua_han), *r.tieu_chi[:2], r.so_hd_tre,
                     *r.tieu_chi[2:], "Có" if r.khong_nang else "",
                     "NÂNG HẠNG RỦI RO" if r.nang else ("THEO DÕI (trong hạn)" if r.theo_doi else ""),
                     dg.MUC5[r.muc_goc], r.ten_muc, " ".join(ghi)])
    bang(ws, 5, COT_DANH_GIA, dong)
    vt = {t: i for i, (t, _, _) in enumerate(COT_DANH_GIA, start=1)}
    for h in range(6, 6 + len(dong)):
        ws.cell(h, vt["Hệ số điều khoản (tháng)"]).number_format = "0.00"
        for t in ("T1: Overdue ≥ X", "T2: Total > Credit limit", "Số HĐ trả trễ", "T3: ≥ N HĐ trễ",
                  "T4: Khách mới ≥ Y", "Xếp hạng Overdue", "Không nâng hạng", "Mức gốc", "Risk level"):
            o = ws.cell(h, vt[t])
            if o.value is None and t.startswith("T"):
                o.value = 0
            o.alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(h, vt["Ghi chú"]).alignment = Alignment(vertical="top", wrap_text=True)
    if dong:
        cuoi = 5 + len(dong)
        for ten_cot in ("Mức gốc", "Risk level"):
            cl = get_column_letter(vt[ten_cot])
            for gia_tri, (nen, chu) in th.MAU_MUC_FIN.items():
                ws.conditional_formatting.add(f"{cl}6:{cl}{cuoi}", FormulaRule(
                    formula=[f'LEFT(${cl}6,{len(gia_tri)})="{gia_tri}"'], fill=th.nen_cf(nen),
                    font=Font(color=chu, bold=True)))
        cl = get_column_letter(vt["KẾT QUẢ"])
        ws.conditional_formatting.add(f"{cl}6:{cl}{cuoi}", FormulaRule(
            formula=[f'LEFT(${cl}6,4)="NÂNG"'], fill=th.nen_cf("FFC7CE"), font=Font(color="9C0006", bold=True)))
    ws.auto_filter.ref = f"A5:{get_column_letter(len(COT_DANH_GIA))}{5 + len(dong)}"
    ws.freeze_panes = "D6"

    # Tham số: bản chụp số tool dùng tuần này
    ts = {**dg.THAM_SO_MAC_DINH, **(kq.tham_so or {})}
    w2 = wb.create_sheet(th.SHEET_THAM_SO)
    _dau_trang(w2, "Tham số tiêu chí nâng hạng rủi ro", ngay_chot, "Số tool đã dùng tuần này")
    w2["A5"] = ("CHỈ ĐỂ XEM. Sửa ở file credit term (sample/SAMPLE_credit_term_khach.xlsx), sheet 'Tham số' — FIN cập nhật "
                "lãi gộp mỗi tháng." + ("" if kq.co_tham_so else " File credit term CHƯA có sheet Tham số: tuần này dùng "
                                        "số mặc định của Head FIN."))
    w2["A5"].font = Font(name=th.PHONG_NOI_DUNG, size=9, italic=True, color="C00000")
    lg = list(ts.get("lai_gop") or []) + [None, None, None]
    lg_ok = [v for v in lg if isinstance(v, (int, float)) and v > 0]
    rows = [("Quý gần nhất", ts.get("quy") or "", "@"),
            ("Lãi gộp tháng 1 của quý (VND)", lg[0], TIEN), ("Lãi gộp tháng 2 của quý (VND)", lg[1], TIEN),
            ("Lãi gộp tháng 3 của quý (VND)", lg[2], TIEN),
            ("Lãi gộp bình quân / tháng", (sum(lg_ok) / len(lg_ok)) if lg_ok else None, TIEN),
            ("Tỷ lệ áp dụng trên lãi gộp", ts.get("ty_le_lai_gop"), "0%"),
            ("X1 = tỷ lệ × lãi gộp bình quân", kq.x1 or None, TIEN),
            ("Tỷ lệ top khách đang quá hạn", ts.get("top_qua_han"), "0%"),
            ("Số khách đang quá hạn", len(qh), "0"),
            ("X2 = nợ quá hạn khách cuối top (làm tròn lên)", kq.x2, TIEN),
            ("X áp dụng = MIN(X1, X2)", kq.x, TIEN),
            ("Y (T4) — trống ở file credit term thì = X / 2", kq.y, TIEN),
            ("Hệ số an toàn hạn mức tạm tính (T2)", ts.get("he_so_an_toan"), "0.0"),
            ("N: số HĐ trả trễ tối thiểu (T3)", ts.get("so_hd_tre"), "0"),
            ("Trễ hơn hạn quá (ngày)", ts.get("ngay_tre"), "0"),
            ("Số tháng xét (T3, billing)", ts.get("so_thang"), "0")]
    bang(w2, 7, [("Tham số", "ten", 48), ("Giá trị", "tien", 20)], [[a, b] for a, b, _ in rows], tong=False)
    for i, (_a, v, dd) in enumerate(rows, start=8):
        w2.cell(i, 2).value = v                     # bang() làm tròn cột tiền → tỷ lệ 0,25 thành trống
        w2.cell(i, 2).number_format = dd
    w2.column_dimensions["A"].width = 48


# --- Payable ---------------------------------------------------------------

COT_CONTRACT = [("No.", "so", 6), ("Alert", "ten", 30), ("Code", "ten", 14), ("Customer's Name", "ten", 40),
                ("Salesman", "ten", 22), ("Balance", "tien", 16), ("Contract No.", "ten", 16),
                ("Credit Term", "ten", 24), ("Contract expiry", "ten", 13), ("Days left", "so", 9),
                ("Note", "ten", 60), ("Contract (PDF)", "ten", 14)]


def ghi_contracts(wb, canh_bao: List[dict], ngay_chot: dt.date, them_khach: str = "") -> None:
    """WR-08 (user 13/10): cảnh báo hợp đồng / credit term đọc từ Client Master. Chỉ cảnh báo —
    hợp đồng hết hạn vẫn tính tuổi nợ theo term cũ; FIN sửa Client Master."""
    ws = wb.create_sheet(th.SHEET_CONTRACT)
    _dau_trang(ws, "Contracts & Credit Terms", ngay_chot, "Alerts from the credit term file")
    ws["F3"] = ("Nguồn: file credit term sample/SAMPLE_credit_term_khach.xlsx (FIN sửa lần cuối). Hợp đồng hết hạn "
                "vẫn tính tuổi nợ theo term cũ.")
    ws["F3"].font = Font(name=th.PHONG_NOI_DUNG, size=9, italic=True, color="7F7F7F")
    dong = [[i, c["loai"], c["ma"], c["ten"], c["sales"], c["du_no"], c["so_hd"], c["term"], c["han_dung"],
             c["con_lai"], c["ghi_chu"], NHAN_LINK if la_link(c.get("link")) else ""]
            for i, c in enumerate(canh_bao, start=1)]
    mau = [{"do": th.DO_CANH_BAO, "vang": th.VANG_CANH_BAO}.get(c["muc"]) for c in canh_bao]
    if not dong:
        ws["A6"] = "Không có cảnh báo hợp đồng / credit term tuần này."
        ws["A6"].font = Font(name=th.PHONG_NOI_DUNG, size=10, italic=True)
        return
    hang = bang(ws, 5, COT_CONTRACT, dong, tong=False, mau_dong=mau)
    for h in range(6, 6 + len(dong)):
        ws.cell(row=h, column=9).number_format = "dd/mm/yyyy"
        ws.cell(row=h, column=9).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(row=h, column=11).alignment = Alignment(vertical="top", wrap_text=True)
        _o_link(ws.cell(row=h, column=12), canh_bao[h - 6].get("link"))
    ws.auto_filter.ref = f"A5:{get_column_letter(len(COT_CONTRACT))}{5 + len(dong)}"
    # WR-11 (FIN 27/10): khách đang nợ chưa có trong Client Master — 3 cột đầu xếp đúng thứ tự Client Master
    # (Mã kh · Khách hàng · Sales phụ trách) để FIN copy một lần, dán vào cột Mã kh ở dòng trống đầu tiên
    from .khachhang import LOAI_CHUA_MASTER
    moi = [c for c in canh_bao if c["loai"] == LOAI_CHUA_MASTER]
    if moi:
        hang += 1
        da = any(c.get("da_them") for c in moi)
        ws.cell(row=hang, column=3, value=(f"New customers — added to the credit term file ({len(moi)})" if da else
                                           f"New customers — NOT yet in the credit term file ({len(moi)})")).font = Font(
            name=th.PHONG_TIEU_DE, size=11, bold=True, color=th.CAM)
        # WR-12: tool tự thêm vào file credit term; không thêm được (file đang mở…) thì FIN vẫn dán được bảng này
        ws.cell(row=hang + 1, column=3, value=(them_khach or "") + ("" if da else
            " Muốn thêm ngay: chọn 3 cột Mã kh · Khách hàng · Sales phụ trách → Copy → dán vào ô Mã kh ở dòng trống "
            "đầu tiên của file credit term.")).font = Font(name=th.PHONG_NOI_DUNG, size=9, italic=True, color="7F7F7F")
        # ghi thẳng vào cột C–F (đúng dưới Code · Customer's Name · Salesman · Balance của bảng trên) để giữ độ
        # rộng cột; cột A để trống nên dashboard không đọc nhầm thành cảnh báo
        h0 = hang + 2
        for j, t in enumerate(("Mã kh", "Khách hàng", "Sales phụ trách", "Dư nợ (tham khảo)"), start=3):
            o = ws.cell(row=h0, column=j, value=t)
            o.fill = PatternFill("solid", fgColor=th.TIM if j == 6 else th.CAM)
            o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
        for i, c in enumerate(moi, start=h0 + 1):
            for j, v in enumerate((c["ma"], c["ten"], c["sales"], c["du_no"]), start=3):
                o = ws.cell(row=i, column=j, value=v)
                o.font = Font(name=th.PHONG_NOI_DUNG, size=10)
                o.border = VIEN
                if j == 6:
                    o.number_format = TIEN
    ws.freeze_panes = "E6"


def ghi_payable(wb, kq: KetQuaFull, ngay_chot: dt.date, ky: str) -> None:
    ws = wb.create_sheet(th.SHEET_PAYABLE)
    _dau_trang(ws, th.TEN_BAO_CAO_TRA, ngay_chot, th.TEN_BANG_TRA, ky)
    if kq.phai_tra is None:
        ws["A6"] = "Tuần này chưa có file Tổng hợp công nợ phải trả (TK331) trong input — sheet để trống."
        ws["A6"].font = Font(name=th.PHONG_NOI_DUNG, size=10, bold=True, color="C00000")
        return
    # FIN 09/10: "làm y chang mẫu của em... chỉ cần lấy số tổng, không phân tích gì cả" —
    # 5 cột như sheet Payable của FIN, chỉ vendor còn số dư, lớn trước. Tiền đã trả nằm ở Cash Flow.
    cot = [("No.", "so", 6), ("Code", "ten", 20), ("Vendor's Name", "ten", 54), ("Current", "tien", 20),
           ("Total", "tien", 20)]
    con = sorted((v for v in kq.phai_tra if v.cuoi_ky > 0.5), key=lambda v: -v.cuoi_ky)
    dong = [[i, v.ma, v.ten, v.cuoi_ky, v.cuoi_ky] for i, v in enumerate(con, start=1)]
    bang(ws, 5, cot, dong)
    ws.auto_filter.ref = f"A5:E{5 + len(dong)}"
    ws.freeze_panes = "D6"


# --- Cash Flow -------------------------------------------------------------

def ghi_cash(wb, kq: KetQuaFull, ngay_chot: dt.date, ky: str) -> None:
    ws = wb.create_sheet(th.SHEET_CASH)
    _dau_trang(ws, "Cash Movement", ngay_chot, "A. Collections from customers", ky)
    nhom = ["ngan-hang", "tien-mat", "can-tru", "ty-gia-phi", "khac"]
    cot = [("No.", "so", 6), ("Code", "ten", 16), ("Customer's Name", "ten", 46)] + \
          [(ctm.TEN_NHOM_THU[k], "tien", 16) for k in nhom] + [("Total", "tien", 16)]
    dong = [[i, x.ma, x.ten] + [x.theo_nhom.get(k, 0.0) for k in nhom] + [x.tong]
            for i, x in enumerate(kq.thu, start=1)]
    hang = bang(ws, 5, cot, dong)
    hang += 1
    ws.cell(row=hang, column=1, value="B. Payments to vendors (Misa TK331 debit — may include offsets)"
            ).font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True, color=th.CAM)
    hang += 1
    if kq.phai_tra is None:
        ws.cell(row=hang, column=1, value="Chưa có file Tổng hợp công nợ phải trả tuần này.")
        hang += 2
    else:
        dong_b = [[i, v.ma, v.ten, v.da_tra, v.cuoi_ky] for i, v in enumerate(kq.tra_vendor, start=1)]
        hang = bang(ws, hang, [("No.", "so", 6), ("Code", "ten", 16), ("Vendor's Name", "ten", 46),
                               ("Paid this week", "tien", 16), ("Closing", "tien", 16)], dong_b)
    hang += 1
    ws.cell(row=hang, column=1, value="C. Net").font = Font(
        name=th.PHONG_TIEU_DE, size=11, bold=True, color=th.CAM)
    cs = kq.chi_so
    tom = [["Cash in (bank + cash collected)", cs.get("da_thu_tien")],
           ["Cash out (paid to vendors)", cs.get("da_tra_vendor")],
           ["Net cash", cs.get("dong_tien_rong")]]
    bang(ws, hang + 1, [("Item", "ten", 46), ("Amount", "tien", 18)], tom, tong=False)


# --- Invoices -------------------------------------------------------------

def ghi_invoices(wb, kq: KetQuaFull, ngay_chot: dt.date) -> None:
    """Mỗi dòng một hoá đơn còn nợ. Lọc theo một khách, cộng cột Amount = Total ở Receivables."""
    from . import soghi as sg
    ws = wb.create_sheet(th.SHEET_INVOICES)
    den = dt.date.fromisoformat(kq.chi_so["den_ngay"])
    _dau_trang(ws, "Open invoices behind every receivable", ngay_chot,
               f"Filter one customer: the Amount column adds up to its Total in {th.SHEET_BANG}")
    cot = [("No.", "so", 6), ("Code", "ten", 14), ("Customer's Name", "ten", 40), ("Invoice No.", "ten", 20),
           ("Invoice Date", "ten", 12), ("Credit Term", "ten", 14), ("Due Date", "ten", 12),
           ("Days Overdue", "so", 9), ("Bucket", "ten", 12), ("Amount", "tien", 16), ("Salesman", "ten", 24),
           ("Note", "ten", 30)]
    if kq.job_misa:                       # v0.5 (WR-06): trạng thái + job theo Misa (Bán hàng, Sổ chi tiết)
        cot += [("Misa status", "ten", 18), ("Job (Misa)", "ten", 30)]
    dong, i = [], 0
    for d in kq.dong:
        if d.hoa_don:
            for ngay, so_hd, han, tien in d.hoa_don:
                qua = (den - han).days
                i += 1
                dong.append([i, d.ma, d.ten, so_hd, ngay.strftime("%d/%m/%Y") if ngay else "",
                             d.credit_term, han.strftime("%d/%m/%Y"), max(qua, 0),
                             sg.TEN_NHOM[sg.nhom_tu_so_ngay(qua)], tien, d.salesman,
                             "Tổng hợp 131 cao hơn Chi tiết 131 — tạm coi là nợ mới" if ngay is None else ""])
                if kq.job_misa:
                    n = ctm._so_hd(str(so_hd)) if ngay else None
                    tt, jobs = d.hd_misa.get((n, ngay.year), ("", "")) if n is not None else ("", "")
                    if not tt and ngay:
                        tt = "(trước năm nay)" if ngay.year < den.year else "(không thấy trên Bán hàng)"
                    dong[-1] += [tt, jobs]
        else:
            i += 1
            nhom = sg.TEN_NHOM[d.nhom_gia_nhat]
            dong.append([i, d.ma, d.ten, "(whole balance)", "", d.credit_term, "", None, nhom, d.total,
                         d.salesman, {"kho-doi": "Hard-to-collect: always 120+",
                                      "ngoai-le": "Exception declared by accounting"}.get(d.nguon, d.nguon)])
            if kq.job_misa:
                dong[-1] += ["", ""]
    bang(ws, 5, cot, dong)
    ws.auto_filter.ref = f"A5:{get_column_letter(len(cot))}{5 + len(dong)}"
    ws.freeze_panes = "D6"


# --- Methodology -----------------------------------------------------------

CACH_TINH = [
    ("All sheets", "Money", "Every amount is VND as recorded in Misa. Excel holds values, not formulas.",
     "—"),
    ("Receivables", "Total", "Ending debit balance of the customer.",
     "Misa — Tổng hợp công nợ phải thu (TK131), column Số dư cuối kỳ · Nợ"),
    ("Receivables", "Customer list", "Customers whose ending debit balance > 0. Prepaid customers "
     "(credit balance) are not listed.", "Misa — Tổng hợp công nợ phải thu"),
    ("Receivables", "Ageing buckets", "Each invoice: due date = invoice date + credit term ('14 ngày của "
     "tháng tiếp theo' = 14th of the next month). A receipt that names invoice numbers in its "
     "description ('theo HD 00001574, 00001419') pays those invoices (an FX-difference voucher on the same "
     "day goes with them); the rest of the money first pays invoices Misa (Bán hàng) marks 'Đã thanh toán' at "
     "the period end, then the OLDEST open invoice (FIN 17/10, WR-09). Days overdue = period end − due date; the due date itself is still "
     "Current, overdue starts the next day. Buckets: Current, 1-30, 31-60, "
     "61-90, 91-120, 120+.", "Misa — Chi tiết công nợ phải thu (TK131) from 01/01/2022: Ngày hoá đơn, "
     "Số hoá đơn, PS Nợ, PS Có"),
    ("Receivables", "Hard-to-collect", "Whole balance in 120+ when FIN marked the customer "
     "'không có khả năng thu hồi' or the reason says 'khó đòi'.", "sample/ ghép khách · last report"),
    ("Receivables", "Credit Term", "Carried from last week's report (as edited by accounting) when Client "
     "Master has none.", "Report of last week in the workspace root"),
    ("Receivables", "Salesman, Job Number", "Open jobs of the customer in SMS (A/R remain > 0), "
     "matched through the customer mapping file.", "SMS AR-AP · sample/ ghép khách"),
    ("AR Risk", "Opening / Increase / Decrease", "Opening and movements of the week from Misa. "
     "Opening + Increase − Decrease = Ending.", "Misa — Tổng hợp TK131: Số dư đầu kỳ, Phát sinh Nợ, "
     "Phát sinh Có"),
    ("AR Risk", "Backdated vs last report", "Misa opening − last report's Total. Not zero means "
     "invoices or receipts were posted later with an earlier date.", "Misa + last report"),
    ("AR Risk", "S1", "Balance grew vs last week AND overdue did not decrease (new lots, old debt "
     "not paid).", "this week vs last week"),
    ("AR Risk", "S2", "Overdue > 0 and exactly equal to last week (nothing collected).",
     "this week vs last week"),
    ("AR Risk", "S3", "Overdue over 60 days (61-90 + 91-120 + 120+) larger than last week "
     "(debt is ageing).", "this week vs last week"),
    ("AR Risk", "S4 / S5 / S6", "S4 no reason from sales · S5 no credit term · S6 whole balance "
     "overdue. Informational, not counted.", "Receivables"),
    ("AR Risk", "S7 · Promise date", "Missed promise: the reason names a payment date or month "
     "('15/9', 'tháng 9' = end of September, 'đầu tháng 9' = 10/9, 'giữa tháng 9' = 20/9) that is "
     "before the report date, and overdue did not go down vs last week. Counted like S1–S3 (Chairman "
     "13/10). Promise date is read by the tool from the reason; FIN may overwrite it.", "Receivables reason"),
    ("Receivables", "Reason for late payment · Reason last week", "WR-11 (option 1): 'Reason for late payment' "
     "starts EMPTY every new week — sales write this week's reason there. 'Reason last week (read only)' shows "
     "what was written in last week's report. The tool writes no reason of its own any more (no '[TỰ ĐỘNG]'). "
     "Re-running the same week keeps this week's reasons. Hard-to-collect customers keep FIN's declaration. "
     "Summary lists, per salesman, overdue customers with no reason this week.",
     "Last week's report · this week's report when re-run"),
    ("Receivables", "Salesman", "Full name as in SMS (CEO 08/10). Short names in old reports are "
     "converted (Hiếu → Trần Văn Hiếu, Hiếu 1 → Phạm Trần Hiếu …).", "SMS AR-AP · thuonghieu.py"),
    ("Receivables", "Rows under a customer (+ / 1 | 2)", "Click + (or the 2 button top-left) to open "
     "the jobs behind a customer: each SMS job still owed, its salesman and amount. The last line "
     "balances to the Total (old debt without an SMS job, or SMS over-stated). Jobs carry no invoice "
     "date, so the customer's ageing buckets are shared out: old debt without a job and the oldest "
     "jobs take the oldest buckets first; rows FIN mapped to a known invoice (sample/"
     "Tham_chieu_job_no_cu.xlsx) take that invoice's own bucket. Columns of the rows always add up "
     "to the customer row. (Used only when the Bán hàng file is missing.)", "SMS AR-AP"),
    ("Receivables", "Rows under a customer — from Misa (v0.6)", "One row per job of every invoice still "
     "OWED in the books: amount = what is left on that invoice after receipts (same rule as Ageing buckets) "
     "— not the full invoice amount (WR-08: FIN had to fix 9 "
     "customers by hand in W41). Rows always add up to the customer Total. Job = 'Mã đối tượng THCP' of the "
     "invoice (Sổ chi tiết TK131); an invoice with several jobs is split pro rata. Invoices before this year: "
     "job from sample/Tham_chieu_job_no_cu.xlsx. Each row sits in the bucket of its own invoice due date. "
     "Salesman of the job from SMS. Column 'Đối chiếu' shows the Misa payment status for FIN to clean up.",
     "Misa — Chi tiết TK131 · Bán hàng · Sổ chi tiết TK131 · SMS (salesman)"),
    ("Receivables", "Đối chiếu job (FIN)", "Customer Total − invoices Misa still MARKS unpaid (Bán hàng). "
     "Khớp: within 1,000. VÀNG: gap ≤ 2% (exchange-rate revaluation forgotten, rounding). ĐỎ: Misa marks "
     "more unpaid than the balance (paid but not matched to the invoice in Misa, or partial payment not "
     "deducted). CAM: balance exceeds what Misa marks unpaid (older debt, or invoice marked paid while "
     "the books still show it owed). The job rows are NOT affected — they follow the books. "
     "Review lists each colour, plus customers with no balance whose invoices still show unpaid (ĐỎ) "
     "and names on Bán hàng not found in Misa (TÍM). Rows are NOT forced to add up: the gap is what "
     "FIN checks.", "Misa — Tổng hợp 131 vs Bán hàng"),
    ("Invoices", "Open invoices", "Every invoice still owed after payments are applied (see Ageing "
     "buckets). Sum per customer = Total in Receivables.", "Misa — Chi tiết TK131"),
    ("AR Risk", "Formulas (WR-08)", "S1–S7, Signals, Risk Code, Risk Tier, CEO review, Trend, Required "
     "action are Excel formulas — the same as the Head of FIN's sheet. Reason for late payment is pulled "
     "from Receivables: write or change the reason there and the tier updates when the file is opened. "
     "Money columns are values. Empty reason (or '[TỰ ĐỘNG] …' in old reports) counts as no reason (S4); "
     "the tier keywords then read 'Reason last week'.",
     "Head of FIN — sheet 'data W39 so W40'"),
    ("Contracts", "Alerts (WR-08)", "Read from the credit term file (sample/SAMPLE_credit_term_khach.xlsx, FIN has the "
     "final edit): contract expired while the customer still owes (red) · expiring within 30 / 60 / 90 days "
     "(30 = yellow) · customer owing not in the credit term file yet (yellow; the tool adds them to the file itself, WR-12 — "
     "list at the end of the sheet) · in the credit term file but no credit term (yellow) · credit term changed "
     "since last week (yellow). Expiry in use = renewal date + 12 months, else expiry date. An expired "
     "contract keeps its old credit term for ageing (Head of FIN 13/10).", "Credit term file"),
    ("AR Risk", "Required action (N3)", "Shown as a sentence only. FIN checks and decides whether to stop "
     "service (Head of FIN 13/10).", "Head of FIN"),
    ("AR Risk", "Jobs by Salesman", "Only for customers with two or more salesmen: each salesman, the "
     "amount of their jobs and the job numbers (Customer Experience 12/10).", "Receivables job rows"),
    ("Receivables", "Credit Term", "Taken from sample/SAMPLE_credit_term_khach.xlsx (FIN keeps it); "
     "if blank there, last week's report. No term → 30 days and S5. Review lists customers without a term.",
     "sample/ · last week's report"),
    ("Receivables", "Customers left out", "Balance below 1,000 VND (rounding residue) — listed in Review.",
     "Misa — Tổng hợp TK131"),
    ("AR Risk", "Risk level (Head of FIN, WR-15)", "Base level by the OLDEST ageing bucket with money: 0 Trắng (not "
     "overdue) · 1 Vàng (1–30) · 2 Cam (31–60) · 3 Đỏ (61–90) · 4 Tím (91+). +1 level (not cumulative) if at least one: "
     "T1 overdue ≥ X · T2 balance > credit limit · T3 ≥ N invoices paid more than 15 days late in 6 months · T4 new "
     "customer (no balance last week) owing ≥ Y — except customers ticked 'Không nâng hạng màu' in the credit term file. "
     "Head of FIN 12/11: the +1 applies ONLY to customers with overdue money; a customer fully within term stays Trắng "
     "and gets a watch mark (e.g. 'Trắng ↑T3'). Missing credit term / no ✓ contract no longer raises the colour: it is "
     "listed in Contracts as data to complete. Straight to at least Đỏ: reason mentions lost contact / promises not to pay / refuses to pay. Max Tím. "
     "X = MIN(X1 = 25% × average monthly gross profit of the last quarter, X2 = overdue of the last customer in the top "
     "15% overdue customers, rounded up); no gross profit entered → X = X2. Credit limit = 'Hạn mức nợ' in the credit "
     "term file, otherwise average monthly billing (6 months) × term factor (months) × 1.3. Parameters: sheet 'Tham "
     "số' of the credit term file (FIN updates monthly). Colour on the Risk level cell of AR Risk only (FIN 12/11: not in Receivables). "
     "For the action plan / payment terms: Trắng, Vàng → 1 · Cam → 2 · Đỏ, Tím → 3.",
     "Chi tiết TK131 · credit term file · Receivables (sales)"),
    ("AR Risk", "Customer group", "Mới: first invoice within the last 6 months (or none) · Lâu năm: first invoice before "
     "1 January this year · Thường xuyên: the rest. First invoice = earliest invoice in Chi tiết TK131 (exported from "
     "2022).", "Chi tiết TK131"),
    ("AR Risk", "Action plan · Suggested payment terms", "FIN's table 'mức độ' by customer group × risk level; 'no "
     "response' adds the escalation sentence. Payment terms: 3 → new customer prepay 50–100%, others stop new service / "
     "prepay only · 2 → limited credit, close follow-up · 1 → credit within limit. Required action (N0–N4) is kept too "
     "(FIN 05/11).", "FIN — sheet 'mức độ'"),
    ("AR Risk", "Score · Grade (reference)", "FIN's payment-history score: longest days overdue (0 → 40, ≤30 → 30, "
     "≤60 → 15) + collection rate (≥95% → 30, ≥80% → 20, ≥50% → 10) + overdue / balance (0 → 20, ≤30% → 15, ≤60% → 8) "
     "+ invoices paid in full (≥90% → 10, ≥70% → 6, ≥50% → 3). A ≥ 80 · B ≥ 60 · C ≥ 40 · D. Invoices issued this year "
     "plus older invoices still open on 1 January. Collection rate = (invoice value + opening − still owed) / invoice "
     "value; credit notes are counted as collected (the ledger does not separate them).", "Chi tiết TK131"),
    ("Receivables", "Job due date (WR-13)", "Each job row shows the due date of its invoice = invoice date + credit "
     "term (FIN 05/11), the same date used for its ageing bucket.", "Chi tiết TK131 · credit term file"),
    ("AR Risk · Contracts", "Contract (PDF)", "'Mở hợp đồng' opens the contract file in the browser — the link "
     "Sales/FIN paste in the credit term file, column 'Link hợp đồng' (http/https only).", "Credit term file"),
    ("Receivables", "Locked past weeks (WR-12)", "Once next week's report exists, edits to an older week's report "
     "are no longer read back: 'Reason last week' always shows what was written while that week was open.",
     "_tool/so-ghi.json"),
    ("AR Risk", "Risk Tier", "0 no overdue → X reason mentions offset/payment on behalf → N4 reason "
     "mentions legal event (đang kiện, mất liên lạc, khó đòi, công nợ xấu, phá sản) → N3 any 120+ or "
     "≥ 2 signals (S1-S3 + S7) → N2 one signal → N1 otherwise.", "Rules of the chief accountant "
     "(sheet Tham chieu)"),
    ("AR Risk", "CEO review", "Overdue ≥ 100,000,000 and no reason yet.", "Tham chieu"),
    ("Payable", "Current / Total", "Ending credit balance of every vendor that still has one; no filter, "
     "no ageing (same layout as FIN's Payable sheet, FIN 09/10). Both columns hold the same amount.",
     "Misa — Tổng hợp công nợ phải trả (TK331), Số dư cuối kỳ · Có"),
    ("Cash Flow", "Payments to vendors", "Debit movements of TK331 in the week, per vendor.",
     "Misa — Tổng hợp TK331: Phát sinh Nợ"),
    ("Cash Flow", "Collections", "Credit lines of TK131 posted in the week, split by contra account: "
     "112 bank, 111 cash, 331 offset with payables, 515/635/642/811 FX & bank fees, other.",
     "Misa — Chi tiết TK131: TK đối ứng, PS Có"),
    ("Cash Flow", "Net cash", "(bank + cash collected) − paid to vendors. Ledger view, not a bank "
     "statement.", "Cash Flow A + B"),
    ("Summary", "Last week", "Figures stored by the tool when last week's report was produced.",
     "_tool/so-ghi.json"),
]


def ghi_methodology(wb, ngay_chot: dt.date, nguon: List[str]) -> None:
    ws = wb.create_sheet(th.SHEET_METHOD)
    _dau_trang(ws, "How every figure is calculated", ngay_chot, "Methodology")
    hang = bang(ws, 5, [("Sheet", "ten", 14), ("Figure", "ten", 28), ("How it is calculated", "ten", 90),
                        ("Source", "ten", 60)], [list(x) for x in CACH_TINH], tong=False)
    hang += 1
    ws.cell(row=hang, column=1, value="Source files of this report").font = Font(
        name=th.PHONG_TIEU_DE, size=11, bold=True, color=th.CAM)
    for i, n in enumerate(nguon, start=hang + 1):
        ws.cell(row=i, column=2, value=n).font = Font(name=th.PHONG_NOI_DUNG, size=10)


THU_TU_SHEET = [th.SHEET_SUMMARY, th.SHEET_BANG, th.SHEET_RISK, th.SHEET_DANH_GIA, th.SHEET_THAM_SO, th.SHEET_CONTRACT, th.SHEET_PAYABLE, th.SHEET_CASH,
                th.SHEET_INVOICES, th.SHEET_METHOD, th.SHEET_REVIEW]


def sap_xep(wb) -> None:
    thu_tu = [n for n in THU_TU_SHEET if n in wb.sheetnames] + \
             [n for n in wb.sheetnames if n not in THU_TU_SHEET]
    wb._sheets = [wb[n] for n in thu_tu]
    wb.active = 0
    for ws in wb.worksheets:      # in / xuất PDF: ngang trang, vừa một trang chiều rộng
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.print_options.horizontalCentered = True
        ws.page_margins.left = ws.page_margins.right = 0.4
