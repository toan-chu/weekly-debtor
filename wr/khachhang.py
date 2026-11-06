# -*- coding: utf-8 -*-
"""Client master list (credit term + hợp đồng) — WR-08.

File sống trên OneDrive, Sales và FIN cùng điền. Tool chỉ ĐỌC cột "Mã kh" và "Thời hạn"
(ghep.doc_credit_term tìm theo tên tiêu đề trong 6 dòng đầu), nên tiêu đề dòng 6 phải giữ đúng chữ.
Module này dựng file theo màu/phông của báo cáo: dùng cho file mẫu và cho lần tạo đầu tiên
(kiemtra.tao_file_credit_term)."""
from pathlib import Path
from typing import Dict, List, Optional, Sequence

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import thuonghieu as th
from . import ghep

# WR-12 (user 27/10): file Excel này gọi là "file credit term" (SAMPLE_credit_term_khach.xlsx). "Client Master" dành cho
# SharePoint List chứa hợp đồng PDF sau này. File cũ có sheet "Client Master" vẫn đọc được.
SHEET_MASTER = "Credit term"
SHEET_MASTER_CU = ("Client Master",)
TEN_FILE_MAC_DINH = "SAMPLE_credit_term_khach.xlsx"
SHEET_HUONG_DAN = "Hướng dẫn"
SHEET_THAM_SO = "Tham số"      # WR-15: tham số tiêu chí nâng hạng của Head FIN — FIN sửa tay mỗi tháng
DONG_NHOM = 5          # dải tên nhóm cột
DONG_TIEU_DE = 6       # tiêu đề cột (tool đọc dòng này)
DONG_DAU = 7
DONG_CUOI = 400        # công thức + danh sách thả xuống kéo sẵn tới đây cho dòng Sales thêm mới
NGAY = "dd/mm/yyyy"
DAU_TICH = "✓"

# (khoá, tiêu đề, nhóm, độ rộng, loại) — loại: ten | ngay | so | cong_thuc
COT = [
    ("stt",      "STT",                   "Khách hàng", 6,  "cong_thuc"),
    ("ma",       "Mã kh",                 "Khách hàng", 16, "ten"),
    ("ten",      "Khách hàng",            "Khách hàng", 44, "ten"),
    ("sales",    "Sales phụ trách",       "Khách hàng", 24, "ten"),
    ("pic",      "Người liên hệ (PIC)",   "Khách hàng", 20, "ten"),
    ("mst",      "Mã số thuế",            "Khách hàng", 14, "ten"),
    ("chu_tk",   "Chủ tài khoản",         "Tài khoản ngân hàng", 22, "ten"),
    ("so_tk",    "Số tài khoản",          "Tài khoản ngân hàng", 18, "ten"),
    ("ngan_hang","Ngân hàng",             "Tài khoản ngân hàng", 18, "ten"),
    ("chi_nhanh","Chi nhánh",             "Tài khoản ngân hàng", 16, "ten"),
    ("co_hd",    "Đã có hợp đồng",        "Hợp đồng", 10, "tick"),   # FIN 19/10 (Sales đề nghị): ✓ = đã ký
    ("so_hd",    "Số hợp đồng",           "Hợp đồng", 16, "ten"),
    ("link_hd",  "Link hợp đồng",         "Hợp đồng", 30, "ten"),    # WR-11: link file hợp đồng (SharePoint)
    ("loai_hd",  "Loại hợp đồng",         "Hợp đồng", 18, "ten"),
    ("term",     "Thời hạn",              "Hợp đồng", 30, "ten"),
    ("han_muc",  "Hạn mức nợ (VND)",      "Hợp đồng", 16, "so"),     # WR-11: credit term = số ngày + hạn mức
    ("khong_nang","Không nâng hạng màu",  "Hợp đồng", 11, "tick"),   # WR-15 (Head FIN): giữ hàng, đi hàng đều, hồ sơ tốt
    ("hieu_luc", "Ngày hiệu lực",         "Hợp đồng", 13, "ngay"),
    ("het_han",  "Ngày hết hạn",          "Hợp đồng", 13, "ngay"),
    ("gia_han",  "Ngày gia hạn",          "Gia hạn", 13, "ngay"),
    ("het_moi",  "Hết hạn sau gia hạn",   "Gia hạn", 13, "cong_thuc"),
    ("han_dung", "Hạn đang áp dụng",      "Tình trạng hợp đồng", 13, "cong_thuc"),
    ("con_lai",  "Còn lại (ngày)",        "Tình trạng hợp đồng", 10, "cong_thuc"),
    ("trang_thai","Trạng thái",           "Tình trạng hợp đồng", 22, "cong_thuc"),
    ("job_dau",  "Ngày job đầu tiên",     "Ghi chú", 13, "ngay"),
    ("ghi_chu",  "Ghi chú",               "Ghi chú", 44, "ten"),
]
VI_TRI = {k: get_column_letter(i) for i, (k, *_r) in enumerate(COT, start=1)}
LOAI_HD = ["Hợp đồng nguyên tắc", "Hợp đồng từng lô", "Báo giá / email", "Chưa có hợp đồng"]
TRANG_THAI = ["Đã hết hạn", "Sắp hết hạn ≤30 ngày", "Sắp hết hạn ≤60 ngày", "Sắp hết hạn ≤90 ngày",
              "Còn hiệu lực", "Chưa có hạn HĐ"]

_VIEN = Side(style="thin", color="D9D9D9")
VIEN = Border(left=_VIEN, right=_VIEN, top=_VIEN, bottom=_VIEN)
_TIM_NHAT = "E9E1F3"   # nền dải nhóm của cột hợp đồng
_CAM_NHAT = "FFE8D9"   # nền dải nhóm của cột định danh


def _mau_cot(nhom: str) -> str:
    return th.CAM if nhom in ("Khách hàng", "Tài khoản ngân hàng") else th.TIM


def _cong_thuc(k: str, r: int) -> str:
    c = VI_TRI
    if k == "stt":
        return f'=IF({c["ten"]}{r}="","",ROW()-{DONG_TIEU_DE})'
    if k == "het_moi":
        return f'=IF({c["gia_han"]}{r}="","",EDATE({c["gia_han"]}{r},12))'
    if k == "han_dung":
        # gia hạn chỉ có hiệu lực khi ngày gia hạn đã tới; chưa tới thì vẫn là ngày hết hạn hiện tại
        gh, hm, hh = (f'{c[x]}{r}' for x in ("gia_han", "het_moi", "het_han"))
        return (f'=IF(AND({hm}<>"",{gh}<=TODAY()),{hm},IF({hh}<>"",{hh},IF({hm}<>"",{hm},"")))')
    if k == "con_lai":
        return f'=IF({c["han_dung"]}{r}="","",{c["han_dung"]}{r}-TODAY())'
    if k == "trang_thai":
        h, s = f'{c["han_dung"]}{r}', f'{c["con_lai"]}{r}'
        return (f'=IF({c["ten"]}{r}="","",IF({h}="","{TRANG_THAI[5]}",IF({s}<0,"{TRANG_THAI[0]}",'
                f'IF({s}<=30,"{TRANG_THAI[1]}",IF({s}<=60,"{TRANG_THAI[2]}",IF({s}<=90,"{TRANG_THAI[3]}",'
                f'"{TRANG_THAI[4]}"))))))')
    raise KeyError(k)


def _dau_trang(ws, ghi_chu: str) -> None:
    ws["A1"], ws["A2"], ws["A3"], ws["A4"] = (th.CONG_TY, "Credit term khách hàng", ghi_chu,
                                               "Credit term & hợp đồng khách hàng")
    ws["A1"].font = Font(name=th.PHONG_TIEU_DE, size=12, bold=True, color=th.TIM)
    ws["A2"].font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True)
    ws["A3"].font = Font(name=th.PHONG_NOI_DUNG, size=10, italic=True)
    ws["A4"].font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True, color=th.CAM)


def _tieu_de(ws) -> None:
    dau = 1
    for i in range(1, len(COT) + 2):
        if i > len(COT) or COT[i - 1][2] != COT[dau - 1][2]:
            nhom = COT[dau - 1][2]
            o = ws.cell(row=DONG_NHOM, column=dau, value=nhom)
            if i - 1 > dau:
                ws.merge_cells(start_row=DONG_NHOM, start_column=dau, end_row=DONG_NHOM, end_column=i - 1)
            for j in range(dau, i):
                ws.cell(row=DONG_NHOM, column=j).fill = PatternFill(
                    "solid", fgColor=_CAM_NHAT if _mau_cot(nhom) == th.CAM else _TIM_NHAT)
            o.font = Font(name=th.PHONG_TIEU_DE, size=9, bold=True, color=_mau_cot(nhom))
            o.alignment = Alignment(horizontal="center", vertical="center")
            dau = i
    for i, (k, ten, nhom, rong, loai) in enumerate(COT, start=1):
        o = ws.cell(row=DONG_TIEU_DE, column=i, value=ten)
        o.fill = PatternFill("solid", fgColor=_mau_cot(nhom))
        o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
        o.alignment = Alignment(wrap_text=True, vertical="center",
                                horizontal="left" if loai == "ten" else "center")
        o.border = VIEN
        ws.column_dimensions[get_column_letter(i)].width = rong
    ws.row_dimensions[DONG_TIEU_DE].height = 32
    ghi = {"ma": "FIN điền: mã khách trên Misa. Tool ghép báo cáo theo cột này.",
           "term": "Tool đọc cột này mỗi tuần. Viết: 'thanh toán ngay', '30 days', "
                   "'14 ngày của tháng tiếp theo'… — xem sheet Hướng dẫn.",
           "sales": "Tên đầy đủ như trên SMS. Khách nhiều Sales: ghi cách nhau dấu phẩy.",
           "han_muc": "Hạn mức nợ (VND) theo hợp đồng. Dư nợ vượt hạn mức → màu lên 1 mức (T2). Trống thì tool tạm tính "
                      "= billing bình quân 6 tháng × hệ số term × 1,3 (sheet Tham số).",
           "khong_nang": "✓ = khách tốt (giữ được hàng, đi hàng đều, hồ sơ tốt): báo cáo không nâng màu theo T1–T4.",
           "link_hd": "Dán link file hợp đồng (SharePoint / OneDrive).",
           "het_moi": "Tự tính: ngày gia hạn + 12 tháng.",
           "trang_thai": "Tự tính từ Hạn đang áp dụng (≤30 / ≤60 / ≤90 ngày)."}
    for k, t in ghi.items():
        cm = Comment(t, "Trustana")
        cm.width, cm.height = 260, 70
        ws[f"{VI_TRI[k]}{DONG_TIEU_DE}"].comment = cm


def _o_du_lieu(o, loai: str, co_gia_tri: bool) -> None:
    o.font = Font(name=th.PHONG_NOI_DUNG, size=10, color="595959" if loai == "cong_thuc" else None)
    o.border = VIEN
    if loai == "cong_thuc":
        o.fill = PatternFill("solid", fgColor=th.XAM_NHAT)
    if loai == "tick":
        o.font = Font(name=th.PHONG_NOI_DUNG, size=12, bold=True, color="2E7D32")
        o.alignment = Alignment(horizontal="center", vertical="top")
    elif loai == "so":
        o.number_format = "#,##0"
        o.alignment = Alignment(horizontal="right", vertical="top")
    elif loai in ("ngay",) or o.column_letter in (VI_TRI["het_moi"], VI_TRI["han_dung"]):
        o.number_format = NGAY
        o.alignment = Alignment(horizontal="center", vertical="top")
    elif o.column_letter in (VI_TRI["stt"], VI_TRI["con_lai"]):
        o.number_format = "0"
        o.alignment = Alignment(horizontal="center", vertical="top")
    else:
        o.alignment = Alignment(vertical="top", wrap_text=False)


def ghi_master(dich: Path, dong: Sequence[Dict], sales: Sequence[str], ghi_chu: str = "",
               cot_them: Sequence[str] = (), tham_so=None) -> Path:
    """dong: [{khoá cột: giá trị}] — khoá như COT; cột công thức tự điền. sales: danh sách thả xuống.
    cot_them: tiêu đề các cột người dùng tự chèn ở file cũ — ghi thêm sau cột cuối, giá trị lấy từ
    dong[i]["+<tiêu đề>"] (tools/nang_cap_client_master.py), để nâng cấp mẫu không làm mất cột của ai."""
    cuoi = max(DONG_CUOI, DONG_DAU + len(dong) + 100)   # file dài thì kéo công thức + thả xuống theo
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_MASTER
    ws.sheet_view.showGridLines = False
    _dau_trang(ws, ghi_chu or "File chung trên OneDrive · Sales và FIN điền trực tiếp · "
                              "tool báo cáo tuần đọc file này, tự thêm khách mới đang nợ")
    _tieu_de(ws)
    for r in range(DONG_DAU, cuoi + 1):
        d = dong[r - DONG_DAU] if r - DONG_DAU < len(dong) else {}
        for i, (k, _t, _n, _w, loai) in enumerate(COT, start=1):
            o = ws.cell(row=r, column=i)
            if loai == "cong_thuc":
                o.value = _cong_thuc(k, r)
            else:
                v = d.get(k)
                o.value = v if v not in ("", None) else None
            _o_du_lieu(o, loai, o.value is not None)
        for j, t in enumerate(cot_them, start=len(COT) + 1):
            o = ws.cell(row=r, column=j, value=d.get("+" + t))
            _o_du_lieu(o, "ten", o.value is not None)
    for j, t in enumerate(cot_them, start=len(COT) + 1):
        o = ws.cell(row=DONG_TIEU_DE, column=j, value=t)
        o.fill = PatternFill("solid", fgColor=th.TIM)
        o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
        o.alignment = Alignment(wrap_text=True, vertical="center")
        o.border = VIEN
        ws.column_dimensions[get_column_letter(j)].width = 20

    c = VI_TRI
    # Thả xuống: gợi ý, không chặn (Sales vẫn gõ được kiểu khác)
    hd = SHEET_HUONG_DAN
    for cot, nguon in (("term", f"'{hd}'!$B$22:$B$29"), ("sales", f"'{hd}'!$B$33:$B$52"),
                       ("loai_hd", f"'{hd}'!$B$56:$B$59")):
        dv = DataValidation(type="list", formula1=nguon, allow_blank=True, showErrorMessage=False)
        dv.add(f"{c[cot]}{DONG_DAU}:{c[cot]}{cuoi}")
        ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1=f'"{DAU_TICH}"', allow_blank=True, showErrorMessage=False,
                        promptTitle="Đã có hợp đồng", prompt="Chọn ✓ nếu đã ký hợp đồng; chưa có thì để trống",
                        showInputMessage=True)
    dv.add(f"{c['co_hd']}{DONG_DAU}:{c['co_hd']}{cuoi}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1=f'"{DAU_TICH}"', allow_blank=True, showErrorMessage=False,
                        promptTitle="Không nâng hạng màu", showInputMessage=True,
                        prompt="Chọn ✓ nếu khách giữ được hàng, đi hàng đều, hồ sơ tốt: báo cáo không nâng màu "
                               "theo các tiêu chí T1–T4 (vẫn lên Đỏ nếu mất liên lạc / hứa không trả)")
    dv.add(f"{c['khong_nang']}{DONG_DAU}:{c['khong_nang']}{cuoi}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                        errorTitle="Hạn mức nợ", error="Nhập số tiền VND, VD 500000000")
    dv.add(f"{c['han_muc']}{DONG_DAU}:{c['han_muc']}{cuoi}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="date", operator="greaterThan", formula1="DATE(2015,1,1)", allow_blank=True,
                        errorTitle="Ngày", error="Nhập ngày dạng dd/mm/yyyy")
    for k in ("hieu_luc", "het_han", "gia_han", "job_dau"):
        dv.add(f"{c[k]}{DONG_DAU}:{c[k]}{cuoi}")
    ws.add_data_validation(dv)

    do = th.nen_cf(th.DO_CANH_BAO)          # CF: Excel đọc bgColor (xem thuonghieu.nen_cf)
    vang = th.nen_cf(th.VANG_CANH_BAO)
    tt = f"{c['trang_thai']}{DONG_DAU}"
    ws.conditional_formatting.add(f"{c['trang_thai']}{DONG_DAU}:{c['trang_thai']}{cuoi}",
        FormulaRule(formula=[f'{tt}="{TRANG_THAI[0]}"'], fill=do, font=Font(color="C00000", bold=True)))
    ws.conditional_formatting.add(f"{c['trang_thai']}{DONG_DAU}:{c['trang_thai']}{cuoi}",
                                  FormulaRule(formula=[f'{tt}="{TRANG_THAI[1]}"'], fill=vang))
    # Khách có tên mà chưa có term / chưa có mã → ô vàng cho người điền
    for k in ("term", "ma"):
        ws.conditional_formatting.add(
            f"{c[k]}{DONG_DAU}:{c[k]}{cuoi}",
            FormulaRule(formula=[f'AND(${c["ten"]}{DONG_DAU}<>"",{c[k]}{DONG_DAU}="")'], fill=vang))
    # Trùng mã kh → đỏ
    ws.conditional_formatting.add(
        f"{c['ma']}{DONG_DAU}:{c['ma']}{cuoi}",
        FormulaRule(formula=[f'AND({c["ma"]}{DONG_DAU}<>"",COUNTIF(${c["ma"]}${DONG_DAU}:${c["ma"]}${cuoi},'
                             f'{c["ma"]}{DONG_DAU})>1)'], fill=do))

    ws.freeze_panes = f"{c['sales']}{DONG_DAU}"
    ws.auto_filter.ref = f"A{DONG_TIEU_DE}:{get_column_letter(len(COT) + len(cot_them))}{cuoi}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{DONG_NHOM}:{DONG_TIEU_DE}"
    ws.print_area = f"A1:{get_column_letter(len(COT))}{DONG_DAU + max(len(dong), 1) + 4}"

    _huong_dan(wb.create_sheet(SHEET_HUONG_DAN), sales)
    _ghi_tham_so(wb.create_sheet(SHEET_THAM_SO), *(tham_so or (None, None)))
    wb.save(dich)
    return Path(dich)


def _huong_dan(ws, sales: Sequence[str]) -> None:
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 70
    ws["A1"] = th.CONG_TY
    ws["A1"].font = Font(name=th.PHONG_TIEU_DE, size=12, bold=True, color=th.TIM)
    ws["A2"] = "Hướng dẫn điền file credit term khách hàng"
    ws["A2"].font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True)

    def muc(r, ten, mau=th.CAM):
        o = ws.cell(row=r, column=1, value=ten)
        o.font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True, color=mau)

    def dong(r, a, b="", fill=None):
        for col, v in ((2, a), (3, b)):
            o = ws.cell(row=r, column=col, value=v)
            o.font = Font(name=th.PHONG_NOI_DUNG, size=10)
            o.alignment = Alignment(vertical="top", wrap_text=True)
            o.border = VIEN
            if fill and col == 2:
                o.fill = PatternFill("solid", fgColor=fill)

    muc(4, "1. Ai điền gì")
    dong(5, "Sales", "Thêm dòng khách mới ở cuối bảng; điền tên, PIC, MST, hợp đồng, thời hạn, ngày hết hạn.")
    dong(6, "FIN", "Điền cột Mã kh (mã khách trên Misa) và xác nhận Thời hạn đúng hợp đồng.")
    dong(7, "Tool báo cáo tuần", "Đọc mọi cột. Chỉ THÊM dòng cho khách mới đang nợ (Ghi chú 'Tool thêm …', ô vàng để "
                                 "FIN điền); không sửa ô đã có. Trước khi thêm luôn sao lưu file vào _tool/luu-tru.")
    dong(8, "Cột nền xám", "Tự tính (STT, hết hạn sau gia hạn, còn lại, trạng thái) — đừng gõ đè.")
    dong(9, "Tiêu đề dòng 6", "Giữ nguyên chữ. Chèn cột thêm được; đổi tên cột tiêu đề thì tool không đọc được.")
    dong(10, "Đã có hợp đồng · Không nâng hạng màu",
         "Chọn ✓ trong ô (bấm mũi tên). Hạn mức nợ: số tiền VND. Link hợp đồng: dán link file.")

    muc(11, "2. Màu trong bảng")
    dong(12, "Ô vàng ở Mã kh / Thời hạn", "Khách đã có tên nhưng còn thiếu — cần điền.", th.VANG_CANH_BAO)
    dong(13, "Mã kh nền đỏ", "Mã bị trùng ở hai dòng — giữ một dòng.", th.DO_CANH_BAO)
    dong(14, "Trạng thái nền đỏ", "Hợp đồng đã hết hạn — cần gia hạn.", th.DO_CANH_BAO)
    dong(15, "Trạng thái nền vàng", "Hết hạn trong 30 ngày tới.", th.VANG_CANH_BAO)
    dong(16, "Chưa có hạn HĐ", "Chưa nhập Ngày hết hạn (hoặc ngày gia hạn).")

    muc(18, "3. Cách ghi Thời hạn", th.TIM)
    ws.cell(row=19, column=2, value="Chọn trong danh sách hoặc gõ đúng một trong các kiểu dưới. "
                                    "Cột phải = tool báo cáo hiểu thế nào.").font = Font(
        name=th.PHONG_NOI_DUNG, size=10, italic=True)
    for col, t in ((2, "Thời hạn"), (3, "Tool hiểu là")):
        o = ws.cell(row=21, column=col, value=t)
        o.fill = PatternFill("solid", fgColor=th.TIM)
        o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
    goi_y = list(ghep.GOI_Y_TERM)
    for i in range(8):
        t = goi_y[i] if i < len(goi_y) else ""
        dong(22 + i, t, ghep.hieu_term(t) if t else "")

    muc(31, "4. Danh sách Sales (nguồn thả xuống — có người mới thì thêm vào đây)", th.CAM)
    o = ws.cell(row=32, column=2, value="Sales")
    o.fill = PatternFill("solid", fgColor=th.CAM)
    o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
    for i in range(20):
        dong(33 + i, sales[i] if i < len(sales) else "")

    muc(54, "5. Loại hợp đồng (nguồn thả xuống)", th.TIM)
    o = ws.cell(row=55, column=2, value="Loại hợp đồng")
    o.fill = PatternFill("solid", fgColor=th.TIM)
    o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
    for i, t in enumerate(LOAI_HD):
        dong(56 + i, t)


# --- Đọc lại Client Master + cảnh báo hợp đồng cho báo cáo tuần (WR-08, user 13/10) -------------
# Luật đã chốt: FIN là người sửa lần cuối file này; hợp đồng hết hạn chưa gia hạn vẫn tính tuổi nợ theo
# term cũ (tool không đổi term, chỉ cảnh báo); N3 chỉ hiện câu việc cần làm, FIN tự quyết ngưng dịch vụ.
NGUONG_SAP_HET = (30, 60, 90)
LOAI_CHUA_MASTER = "Chưa có trong file credit term"


def _hom_nay():
    import datetime as dt
    return dt.date.today()


def _cong_thang(d, n: int):
    import calendar
    import datetime as dt
    m = d.month - 1 + n
    y, m = d.year + m // 12, m % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def han_ap_dung(het, gh, ngay):
    """Hạn hợp đồng đang áp dụng tại `ngay`: đã tới ngày gia hạn → ngày gia hạn + 12 tháng; chưa tới → ngày hết hạn."""
    if gh and gh <= ngay:
        return _cong_thang(gh, 12)
    return het or (_cong_thang(gh, 12) if gh else None)


TICH_CO = {"✓", "✔", "☑", "x", "có", "co", "yes", "y", "1", "true", "rồi", "đã ký", "da ky", "ok"}


def la_tich(v) -> bool:
    """Ô tích: chỉ ✓ (và vài cách gõ quen: x, có, yes, 1) là CÓ. "Chưa", "Không", "-"… là KHÔNG — bản WR-09
    coi ô nào có chữ cũng là đã có hợp đồng (lỗi tự soi 27/10)."""
    if v is True:
        return True
    return ghep._o(v).strip().lower() in TICH_CO


def so_tien(v):
    """Ô số tiền: số, hoặc chữ '500.000.000' / '500,000,000' / '500tr' / '1,5 tỷ'. Không đọc được → None."""
    import re
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v) if v > 0 else None
    t = ghep._o(v).lower().replace(" ", "")
    if not t:
        return None
    nhan = 1
    for duoi, he_so in (("tỷ", 1e9), ("ty", 1e9), ("triệu", 1e6), ("tr", 1e6)):
        if t.endswith(duoi):
            t, nhan = t[: -len(duoi)], he_so
            break
    if nhan > 1:
        t = t.replace(",", ".")
    else:
        t = re.sub(r"[.,](?=\d{3}(\D|$))", "", t)
    try:
        x = float(t) * nhan
    except ValueError:
        return None
    return x if x > 0 else None


def doc_master(thu_muc, ngay=None) -> Dict[str, Dict]:
    """{MÃ KH (hoa): {ten, term, so_hd, het_han, gia_han, han_dung, sales}} từ mọi bảng term trong sample/.
    Nhận cả file Client Master lẫn file nháp của KTT (tìm cột theo tên). File sửa sau thắng."""
    ra: Dict[str, Dict] = {}
    for ws, (r0, i_ma, i_term) in ghep._cac_bang_term(thu_muc):
        hang = [list(r) for r in ws.iter_rows(min_row=r0, values_only=True)]
        if not hang:
            continue
        tieu = [ghep._o(v).lower() for v in hang[0]]

        def cot(*ten, thu=0):
            vt = [i for i, x in enumerate(tieu) if x in ten]
            return vt[thu] if len(vt) > thu else None
        i_ten = cot("khách hàng", "customer's name", "tên khách")
        i_hd = cot("số hợp đồng", "contract no.")
        i_het = cot("ngày hết hạn", "expiry date")
        i_gh = cot("ngày gia hạn", "renewal date")
        i_sales = cot("sales phụ trách", "salesman")
        i_co_hd = cot("đã có hợp đồng")
        i_han_muc = cot("hạn mức nợ (vnd)", "hạn mức nợ", "credit limit", "credit limit (vnd)")
        i_link = cot("link hợp đồng", "contract link")
        i_khong_nang = cot("không nâng hạng màu", "không nâng hạng")
        g = lambda r, i: r[i] if i is not None and i < len(r) else None
        for r in hang[1:]:
            ma = ghep._o(g(r, i_ma)).upper()
            if not ma or ma in ("MÃ KH", "CODE"):
                continue
            het, gh = ghep._ngay(g(r, i_het)), ghep._ngay(g(r, i_gh))
            ra[ma] = dict(ten=ghep._o(g(r, i_ten)), term=ghep._o(g(r, i_term)), so_hd=ghep._o(g(r, i_hd)),
                          het_han=het, gia_han=gh, han_dung=han_ap_dung(het, gh, ngay or _hom_nay()),
                          sales=ghep._o(g(r, i_sales)),
                          co_hd=la_tich(g(r, i_co_hd)) if i_co_hd is not None else None,
                          han_muc=so_tien(g(r, i_han_muc)), khong_nang=la_tich(g(r, i_khong_nang)),
                          link_hd=ghep._o(g(r, i_link)))
    return ra


def _chuan_term(t) -> str:
    return " ".join(str(t or "").lower().split())


def canh_bao_hop_dong(master: Dict[str, Dict], rui_ro, ngay_chot, term_truoc: Dict[str, str]) -> List[Dict]:
    """Danh sách cảnh báo cho sheet Contracts. rui_ro: DongRuiRo của tuần (khách còn nợ).
    muc: 'do' | 'vang' | '' — dùng tô màu."""
    ra: List[Dict] = []
    con_no = {r.ma.upper(): r for r in rui_ro}

    def them(loai, muc, ma, ten, sales, du_no, m, ghi_chu):
        han = (m or {}).get("han_dung")
        ra.append(dict(loai=loai, muc=muc, ma=ma, ten=ten, sales=sales, du_no=du_no, link=(m or {}).get("link_hd", ""),
                       so_hd=(m or {}).get("so_hd", ""), term=(m or {}).get("term", ""), han_dung=han,
                       con_lai=(han - ngay_chot).days if han else None, ghi_chu=ghi_chu))

    # 1–2. Hạn hợp đồng: đã hết hạn mà còn nợ / còn phát sinh; sắp hết hạn (mọi khách trong danh sách)
    for ma, m in master.items():
        han = m["han_dung"]
        if not han:
            continue
        r = con_no.get(ma)
        ten = (r.ten if r else "") or m["ten"]
        sales = (r.salesman if r else "") or m["sales"]
        du_no = r.cuoi_ky if r else 0.0
        con = (han - ngay_chot).days
        if con < 0 and r:
            ghi = f"Hết hạn {han:%d/%m/%Y}, vẫn còn nợ" + (f"; tuần này phát sinh {r.tang:,.0f}".replace(",", ".") if r.tang > 0.5 else "")
            them("Hợp đồng đã hết hạn", "do", ma, ten, sales, du_no, m, ghi + ". Tuổi nợ vẫn tính theo term cũ.")
        elif 0 <= con <= NGUONG_SAP_HET[-1]:
            moc = next(n for n in NGUONG_SAP_HET if con <= n)
            them(f"Sắp hết hạn ≤{moc} ngày", "vang" if moc == 30 else "", ma, ten, sales, du_no, m,
                 f"Hết hạn {han:%d/%m/%Y} — Sales làm gia hạn, FIN cập nhật Ngày gia hạn.")
    # 3–4. Khách đang nợ: chưa có term trong danh sách · term vừa bị đổi (thiếu số HĐ / ngày hết hạn: chỉ đếm)
    for ma, r in con_no.items():
        m = master.get(ma)
        if not m:
            # WR-11 (FIN 27/10): khách đang nợ chưa có trong Client Master — sheet Contracts có bảng riêng,
            # cột xếp đúng thứ tự Client Master để FIN dán vào rồi điền term + hạn mức
            them(LOAI_CHUA_MASTER, "vang", r.ma, r.ten, r.salesman, r.cuoi_ky, m,
                 f"Khách mới. Tool đang tính theo: "
                 f"{r.credit_term or '30 ngày (mặc định)'}.")
            continue
        if not m["term"]:
            them("Chưa có credit term", "vang", r.ma, r.ten, r.salesman, r.cuoi_ky, m,
                 f"Có trong file credit term nhưng chưa điền Thời hạn — tool đang tính theo: {r.credit_term or '30 ngày (mặc định)'}.")
            continue
        cu = term_truoc.get(ma)
        if cu and _chuan_term(cu) != _chuan_term(r.credit_term):
            them("Credit term vừa đổi", "vang", r.ma, r.ten, r.salesman, r.cuoi_ky, m,
                 f"Tuần trước: {cu} → tuần này: {r.credit_term}. FIN xác nhận đúng hợp đồng.")
    thu_tu = {"Hợp đồng đã hết hạn": 0, "Credit term vừa đổi": 1, LOAI_CHUA_MASTER: 2, "Chưa có credit term": 2}
    ra.sort(key=lambda x: (thu_tu.get(x["loai"], 3 if x["loai"].startswith("Sắp") else 4),
                           x["con_lai"] if x["con_lai"] is not None else 0, -x["du_no"]))
    return ra


# --- WR-12 (FIN 27/10): khách mới tự vào file credit term ---------------------------------------------------

def sheet_credit_term(wb):
    """Sheet chính của file credit term (tên mới hoặc tên cũ 'Client Master'). Không có → None."""
    for t in (SHEET_MASTER,) + SHEET_MASTER_CU:
        if t in wb.sheetnames:
            return wb[t]
    return None


def doc_file(src) -> tuple:
    """Đọc toàn bộ file credit term để ghi lại: (dòng [{khoá: giá trị}], cột người dùng tự chèn, danh sách Sales).
    Cột công thức bỏ qua (ghi lại tự tính). Ngày gõ dạng chữ đổi thành ngày thật. Cột lạ giữ dưới khoá '+<tiêu đề>'."""
    import datetime as dt
    wb = openpyxl.load_workbook(src, data_only=False)
    ws = sheet_credit_term(wb)
    if ws is None:
        raise ValueError(f"{Path(src).name}: không có sheet '{SHEET_MASTER}'")
    tieu = [str(ws.cell(DONG_TIEU_DE, c).value or "").strip() for c in range(1, ws.max_column + 1)]
    khoa = {t: k for k, t, *_ in COT}
    loai = {k: l for k, _, _, _, l in COT}

    def gia(v, l):
        if isinstance(v, str):
            v = v.replace("\xa0", " ").strip()
            if not v:
                return None
        if l == "ngay" and v is not None and not isinstance(v, (dt.date, dt.datetime)):
            return ghep._ngay(v) or v
        if isinstance(v, dt.datetime) and l == "ngay":
            return v.date()
        return v

    cot_them = [t for t in tieu if t and t not in khoa]
    dong = []
    for r in range(DONG_DAU, ws.max_row + 1):
        d = {}
        for c, t in enumerate(tieu, start=1):
            if not t:
                continue
            v = ws.cell(r, c).value
            k = khoa.get(t)
            if k is None:
                d["+" + t] = gia(v, "ten")
            elif loai[k] != "cong_thuc":
                d[k] = gia(v, loai[k])
        if d.get("ten") or d.get("ma"):
            dong.append(d)
    sales = []
    if SHEET_HUONG_DAN in wb.sheetnames:
        hd = wb[SHEET_HUONG_DAN]
        sales = [str(hd.cell(r, 2).value).strip() for r in range(33, 53)
                 if hd.cell(r, 2).value and str(hd.cell(r, 2).value).strip()]
    return dong, cot_them, sales


def tim_file_credit_term(thu_muc_mau) -> Optional[Path]:
    """File credit term để tool thêm khách: file chung đã khai (--client-master) nếu có; không thì file mới sửa nhất
    trong sample/ có sheet 'Credit term' (hoặc 'Client Master'). Không có file đúng mẫu → None."""
    if ghep.FILE_MASTER and Path(ghep.FILE_MASTER).is_file():
        return Path(ghep.FILE_MASTER)
    if not thu_muc_mau or not Path(thu_muc_mau).is_dir():
        return None
    for f in sorted(Path(thu_muc_mau).glob("*.xlsx"), key=lambda f: -f.stat().st_mtime):
        if f.name.startswith(("~$", "KIEM_TRA", "PHIEU")):
            continue
        try:
            wb = openpyxl.load_workbook(f, read_only=True)
            co = any(t in wb.sheetnames for t in (SHEET_MASTER,) + SHEET_MASTER_CU)
            wb.close()
        except Exception:
            continue
        if co:
            return f
    return None


def dang_mo(f: Path) -> bool:
    """File đang mở trong Excel trên máy này: có file khoá '~$<tên>' cạnh nó, hoặc Windows không cho mở để ghi."""
    if (f.parent / f"~${f.name}").exists():
        return True
    try:
        with open(f, "r+b"):
            pass
    except OSError:
        return True
    return False


# --- WR-15: sheet "Tham số" (Head FIN) ----------------------------------------------------------------------
THAM_SO_DONG = [   # (khoá, nhãn ở cột A — tool đọc theo nhãn, định dạng, ghi chú)
    ("quy", "Quý gần nhất", "@", "VD Q3/2026 — chỉ để ghi nhớ số lãi gộp bên dưới là của quý nào"),
    ("lai_gop_1", "Lãi gộp tháng 1 của quý (VND)", "#,##0", "Lấy từ báo cáo P&L. Đổi quý thì nhập lại cả 3 tháng"),
    ("lai_gop_2", "Lãi gộp tháng 2 của quý (VND)", "#,##0", ""),
    ("lai_gop_3", "Lãi gộp tháng 3 của quý (VND)", "#,##0", "Trống cả 3 → X tạm = X2 (báo cáo ghi rõ)"),
    ("ty_le_lai_gop", "X1: tỷ lệ trên lãi gộp bình quân/tháng", "0%", "X1 = tỷ lệ × lãi gộp bình quân / tháng"),
    ("top_qua_han", "X2: top khách đang quá hạn", "0%",
     "X2 = nợ quá hạn của khách xếp cuối trong top này (làm tròn lên). X = số nhỏ hơn giữa X1 và X2"),
    ("y", "Y: khách mới nợ từ (VND)", "#,##0", "T4. Để trống thì Y = X / 2"),
    ("he_so_an_toan", "Hệ số an toàn hạn mức tạm tính", "0.0",
     "T2: khách chưa có Hạn mức nợ → hạn mức tạm tính = billing bình quân 6 tháng × hệ số term × hệ số này"),
    ("so_hd_tre", "T3: số hoá đơn trả trễ tối thiểu", "0", ""),
    ("ngay_tre", "T3: trễ hơn hạn quá (ngày)", "0", ""),
    ("so_thang", "Số tháng xét (T3, billing bình quân)", "0", ""),
]
TIEU_DE_BANG_TERM = "Credit term (đúng chữ như cột Thời hạn)"


def _ghi_tham_so(ws, ts=None, bang=None) -> None:
    from . import danhgia as dg
    ts = {**dg.THAM_SO_MAC_DINH, **(ts or {})}
    lg = list(ts.get("lai_gop") or []) + [None, None, None]
    gia = {**ts, "lai_gop_1": lg[0], "lai_gop_2": lg[1], "lai_gop_3": lg[2]}
    bang = list(bang) if bang else list(dg.HE_SO_TERM_MAC_DINH)
    ws.sheet_view.showGridLines = False
    for col, w in (("A", 44), ("B", 20), ("C", 90)):
        ws.column_dimensions[col].width = w
    ws["A1"] = th.CONG_TY
    ws["A1"].font = Font(name=th.PHONG_TIEU_DE, size=12, bold=True, color=th.TIM)
    ws["A2"] = "Tham số tiêu chí nâng hạng rủi ro (Head FIN) — FIN sửa các ô vàng, mỗi tháng một lần"
    ws["A2"].font = Font(name=th.PHONG_TIEU_DE, size=11, bold=True)
    vang = PatternFill("solid", fgColor="FFF2CC")

    def tieu_de(r, *chu):
        for i, t in enumerate(chu, start=1):
            o = ws.cell(r, i, t)
            o.fill = PatternFill("solid", fgColor=th.TIM)
            o.font = Font(name=th.PHONG_TIEU_DE, size=10, bold=True, color=th.TRANG)
            o.border = VIEN
    tieu_de(4, "Tham số", "Giá trị", "Ghi chú")
    r = 5
    for k, nhan, dd, ghi in THAM_SO_DONG:
        ws.cell(r, 1, nhan).border = VIEN
        o = ws.cell(r, 2, gia.get(k))
        o.fill, o.border, o.number_format = vang, VIEN, dd
        ws.cell(r, 3, ghi).font = Font(name=th.PHONG_NOI_DUNG, size=9, italic=True, color="7F7F7F")
        r += 1
    r += 1
    tieu_de(r, TIEU_DE_BANG_TERM, "Hệ số (tháng)", "Cách tính")
    for i in range(max(len(bang) + 15, 30)):
        t, hs, ghi = (list(bang[i]) + [None, None, None])[:3] if i < len(bang) else (None, None, None)
        for j, v in enumerate((t, hs, ghi), start=1):
            o = ws.cell(r + 1 + i, j, v)
            o.border = VIEN
            if j < 3:
                o.fill = vang
        ws.cell(r + 1 + i, 2).number_format = "0.00"
    ws.cell(r + len(bang) + 17, 1, "Term không có trong bảng: tool tự tính = số ngày từ hoá đơn tới hạn / 30. "
                                   "Chưa có term: không tạm tính được hạn mức (T2 bỏ qua, T4 nâng).").font = \
        Font(name=th.PHONG_NOI_DUNG, size=9, italic=True, color="7F7F7F")


def doc_tham_so_file(f) -> tuple:
    """(tham số {khoá: giá trị}, bảng hệ số term [(term, hệ số, ghi chú)]) từ sheet 'Tham số'. Không có sheet → (None, None).
    Ô trống / gõ sai → giá trị mặc định của Head FIN (danhgia.THAM_SO_MAC_DINH); riêng Y trống → None (= X / 2)."""
    from . import danhgia as dg
    try:
        wb = openpyxl.load_workbook(f, data_only=True)
    except Exception:
        return None, None
    if SHEET_THAM_SO not in wb.sheetnames:
        return None, None
    ws = wb[SHEET_THAM_SO]
    nhan = {n: (k, dd) for k, n, dd, _g in THAM_SO_DONG}
    ts: Dict[str, object] = {}
    bang, trong_bang, trong = [], False, 0
    for r in range(1, ws.max_row + 1):
        a, b, c = (ws.cell(r, i).value for i in (1, 2, 3))
        a_ = str(a or "").strip()
        if trong_bang:
            if not a_:
                trong += 1
                if trong > 30:
                    break
                continue
            trong = 0
            hs = so_tien(b) if not isinstance(b, (int, float)) else b
            if isinstance(hs, (int, float)):
                bang.append((a_, float(hs), str(c or "")))
            continue
        if a_ == TIEU_DE_BANG_TERM:
            trong_bang = True
        elif a_ in nhan:
            k, dd = nhan[a_]
            if k == "quy":
                ts[k] = str(b or "").strip()
            else:
                v = b if isinstance(b, (int, float)) else so_tien(b)
                ts[k] = v if isinstance(v, (int, float)) else None
    goc = dict(dg.THAM_SO_MAC_DINH)
    ra = {k: (ts.get(k) if ts.get(k) is not None else goc[k]) for k in goc if k not in ("lai_gop",)}
    ra["y"] = ts.get("y")                     # trống → None → Y = X / 2
    ra["lai_gop"] = [ts.get(f"lai_gop_{i}") for i in (1, 2, 3)]
    return ra, bang


def doc_tham_so(thu_muc_mau) -> tuple:
    """Tham số của file credit term đang dùng; chưa có sheet / file → mặc định của Head FIN, bảng term mặc định."""
    from . import danhgia as dg
    f = tim_file_credit_term(thu_muc_mau)
    ts, bang = doc_tham_so_file(f) if f else (None, None)
    return (ts or dict(dg.THAM_SO_MAC_DINH)), (bang or list(dg.HE_SO_TERM_MAC_DINH)), (ts is not None)


def thieu_mau(f: Path) -> List[str]:
    """Những gì file credit term còn thiếu so với mẫu hiện tại: tên cột thiếu, và 'tên sheet cũ' nếu sheet còn tên
    'Client Master'. Rỗng = file đã đúng mẫu. File không đọc được → rỗng (không đụng vào)."""
    try:
        wb = openpyxl.load_workbook(f, read_only=True)
    except Exception:
        return []
    try:
        ws = sheet_credit_term(wb)
        if ws is None:
            return []
        tieu = {str(v or "").strip() for v in next(ws.iter_rows(min_row=DONG_TIEU_DE, max_row=DONG_TIEU_DE,
                                                                values_only=True), ())}
        ra = [t for _k, t, *_r in COT if t not in tieu]
        if SHEET_THAM_SO not in wb.sheetnames:
            ra.append(f"sheet '{SHEET_THAM_SO}'")
        if ws.title != SHEET_MASTER:
            ra.append(f"tên sheet '{SHEET_MASTER}'")
        return ra
    finally:
        wb.close()


def them_khach_moi(dich: Path, moi: Sequence[Dict], luu_tru: Path, ngay) -> tuple:
    """Thêm khách đang nợ chưa có vào cuối file credit term. CHỈ thêm dòng (Mã kh, Khách hàng, Sales phụ trách, Ghi chú
    'Tool thêm dd/mm/yyyy'); ô đã có giữ nguyên. Sao lưu file trước khi ghi; file đang mở thì không ghi, trả lý do.
    Trả (số khách đã thêm, câu báo)."""
    import datetime as dt
    import os
    import shutil
    dich = Path(dich)
    if dich.exists() and dang_mo(dich):
        if not moi:
            return 0, (f"Chưa nâng {dich.name} lên mẫu mới: file đang mở. Đóng file — lần chạy sau tool làm tiếp."
                       if thieu_mau(dich) else "")
        return 0, (f"Chưa thêm {len(moi)} khách mới vào {dich.name}: file đang mở. Đóng file — lần chạy sau tool "
                   "thêm tiếp (hoặc chạy lại tuần này).")
    dong, cot_them, sales = doc_file(dich) if dich.exists() else ([], [], [])
    tham_so = doc_tham_so_file(dich) if dich.exists() else (None, None)     # giữ nguyên số FIN đã nhập
    co = {str(d.get("ma") or "").strip().upper() for d in dong}
    them = []
    for m in moi:
        ma = str(m.get("ma") or "").strip()
        if not ma or ma.upper() in co:
            continue
        co.add(ma.upper())
        them.append(dict(ma=ma, ten=m.get("ten") or "", sales=m.get("sales") or "",
                         ghi_chu=f"Tool thêm {ngay:%d/%m/%Y} — FIN điền Thời hạn, Hạn mức nợ"))
    thieu = thieu_mau(dich) if dich.exists() else []
    if not them and not thieu:
        return 0, ""                       # không có khách mới, file đã đúng mẫu → không ghi gì
    if dich.exists():
        Path(luu_tru).mkdir(parents=True, exist_ok=True)
        shutil.copy2(dich, Path(luu_tru) / f"{dich.stem}_truoc-khi-tool-ghi_{dt.datetime.now():%Y%m%d-%H%M%S-%f}.xlsx")
    tam = dich.with_name(dich.stem + ".tool-dang-ghi")
    ghi_master(tam, dong + them, sales, cot_them=cot_them, tham_so=tham_so)
    os.replace(tam, dich)
    cau = []
    if thieu:
        cau.append(f"Đã nâng {dich.name} lên mẫu mới (thêm: {', '.join(thieu)}); dữ liệu cũ giữ nguyên, bản trước khi "
                   "nâng nằm trong _tool/luu-tru.")
    if them:
        cau.append(f"Đã thêm {len(them)} khách mới vào {dich.name}: {', '.join(x['ma'] for x in them)}. "
                   "FIN điền Thời hạn, Hạn mức nợ ở các ô vàng.")
    return len(them), " ".join(cau)
