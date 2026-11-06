# -*- coding: utf-8 -*-
"""WR-15 (Head FIN, file 'W43 tiêu chí nâng hạng rủi ro v3', user 06/11): 5 mức màu 0–4 (Trắng · Vàng · Cam · Đỏ · Tím),
+1 mức theo T1–T4, lên thẳng Đỏ, CHỈ tô ô Risk level. Tham số ở sheet 'Tham số' của file credit term. Dữ liệu GIẢ."""
import datetime as dt
import shutil
import subprocess
import sys
import time
from pathlib import Path

import openpyxl
import pytest

GOC = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(GOC))
sys.path.insert(0, str(GOC / "tests"))

from wr import danhgia as dg, khachhang as kh  # noqa: E402
from wr import soghi as sg  # noqa: E402
from wr import thuonghieu as th  # noqa: E402
import runner  # noqa: E402
from test_wr4 import A, B, _tha_4_file, ws  # noqa: E402,F401

D = dt.date
MAC = ("Không", "Có", "Không")


# --- luật ---------------------------------------------------------------------

def test_muc_goc_theo_nhom_tuoi_cu_nhat():
    assert [dg.muc_goc5(n) for n in ([9, 0, 0, 0, 0, 0], [0, 5, 0, 0, 0, 0], [0, 5, 5, 0, 0, 0],
                                      [0, 0, 0, 5, 0, 0], [0, 0, 0, 0, 5, 0], [0, 0, 0, 0, 0, 5])] == [0, 1, 2, 3, 4, 4]
    assert dg.muc_goc5([0, 5, 0, 0, 0, 0.3]) == 1, "dưới 0,5đ là số lẻ"


def test_nguong_x_min_x1_x2_va_top_15_lam_tron_len():
    ts = dict(dg.THAM_SO_MAC_DINH)
    qh = [100, 90, 80, 70, 60, 50, 40, 30, 0]          # 8 khách quá hạn → 15% = 1,2 → 2 khách → X2 = 90
    assert dg.nguong_x(qh, ts) == (90, 0.0, 90), "chưa nhập lãi gộp → X = X2"
    ts["lai_gop"] = [200, 400, None]                    # bình quân 300 × 25% = 75
    assert dg.nguong_x(qh, ts) == (75, 75, 90)
    ts["lai_gop"] = [4000, 4000, 4000]
    assert dg.nguong_x(qh, ts)[0] == 90


def test_he_so_term():
    b = dg.HE_SO_TERM_MAC_DINH
    assert dg.he_so_term("30 days", b) == 1 and dg.he_so_term(" 45  Days ", b) == 1.5
    assert dg.he_so_term("thanh toán ngay", b) == 0 and dg.he_so_term("", b) is None
    assert dg.he_so_term("21 days", b, sg.han_tu_term) == 0.7, "không có trong bảng → số ngày / 30"


def test_muc_cuoi_nang_va_nang_thang():
    assert dg.muc_cuoi5(1, True, False, True, MAC) == 2
    assert dg.muc_cuoi5(4, True, False, True, MAC) == 4, "tối đa Tím"
    assert dg.muc_cuoi5(0, False, True, False, MAC) == 3, "lên thẳng: ít nhất Đỏ"
    assert dg.muc_cuoi5(4, False, True, True, MAC) == 4
    assert dg.muc_cuoi5(0, False, False, True, ("Không", "Không", "Không")) == 2, "quá hạn + không phản hồi → Cam"
    assert dg.nang_thang("", ("Có", "Có", "Không"), False)
    assert dg.nang_thang("", ("Không", "Có", "Có"), True) and not dg.nang_thang("", ("Không", "Có", "Có"), False)
    assert dg.nang_thang("Mất liên lạc với kế toán", MAC, True)
    assert not dg.nang_thang("khách không trả kịp vì bank Ấn Độ", MAC, True), "'không trả' trơn bắt nhầm (W43)"
    assert [dg.muc3_tu_muc5(m) for m in range(5)] == [1, 1, 2, 3, 3]


# --- sheet Tham số trong file credit term ------------------------------------

def test_sheet_tham_so_doc_ghi_va_giu_so_fin_khi_tool_ghi_lai(tmp_path):
    f = kh.ghi_master(tmp_path / kh.TEN_FILE_MAC_DINH, [dict(ma="KA", ten=A, term="30 days")], [])
    ts, bang = kh.doc_tham_so_file(f)
    assert ts["ty_le_lai_gop"] == 0.25 and ts["y"] == 50_000_000 and ts["lai_gop"] == [None] * 3
    assert len(bang) == len(dg.HE_SO_TERM_MAC_DINH)
    wb = openpyxl.load_workbook(f)
    sh = wb[kh.SHEET_THAM_SO]
    hang = {sh.cell(r, 1).value: r for r in range(1, 40)}
    sh.cell(hang["Lãi gộp tháng 1 của quý (VND)"], 2, 1_000_000_000)
    sh.cell(hang["Y: khách mới nợ từ (VND)"], 2).value = None      # cell(..., None) không xoá ô
    r = next(r for r in range(1, 80) if sh.cell(r, 1).value is None and r > hang[kh.TIEU_DE_BANG_TERM])
    sh.cell(r, 1, "90 days")
    sh.cell(r, 2, 3)
    wb.save(f)
    time.sleep(0.05)
    kh.them_khach_moi(f, [dict(ma="KB", ten=B)], tmp_path / "lt", D(2026, 11, 6))      # tool ghi lại cả file
    ts, bang = kh.doc_tham_so_file(f)
    assert ts["lai_gop"][0] == 1_000_000_000 and ts["y"] is None, "số FIN nhập vẫn còn sau khi tool ghi lại"
    assert ("90 days", 3.0) in [(t, h) for t, h, _ in bang]


def test_file_cu_chua_co_sheet_tham_so_thi_tu_them(tmp_path):
    f = kh.ghi_master(tmp_path / kh.TEN_FILE_MAC_DINH, [dict(ma="KA", ten=A, term="30 days")], [])
    wb = openpyxl.load_workbook(f)
    del wb[kh.SHEET_THAM_SO]
    wb.save(f)
    assert kh.thieu_mau(f) == [f"sheet '{kh.SHEET_THAM_SO}'"]
    kh.them_khach_moi(f, [], tmp_path / "lt", D(2026, 11, 6))
    assert kh.thieu_mau(f) == [] and kh.doc_tham_so_file(f)[0] is not None


# --- báo cáo thật từ input -----------------------------------------------------

def _hang(sh, ma):
    return next(r for r in range(6, sh.max_row + 1) if sh.cell(r, 2).value == ma)


def _sm(wb):
    return {r[5]: (r[2] or 0) for r in wb[th.SHEET_SUMMARY].iter_rows(min_row=20, values_only=True) if r and r[5]}


def _risk(wb):
    k = wb[th.SHEET_RISK]
    cot = {k.cell(5, c).value: c for c in range(1, k.max_column + 1)}
    return {k.cell(r, 2).value: {t: k.cell(r, c).value for t, c in cot.items()}
            for r in range(6, k.max_row + 1) if isinstance(k.cell(r, 1).value, int) and k.cell(r, 4).value is not None
            or isinstance(k.cell(r, 1).value, int) and k.cell(r, cot["Base level (0-4)"]).value is not None}


def _sua(f, ma, cot, gia_tri):
    wb = openpyxl.load_workbook(f)
    sh = wb[th.SHEET_BANG]
    sh.cell(_hang(sh, ma), cot, value=gia_tri)
    wb.save(f)


def test_w40_chua_co_file_credit_term_moi_khach_dinh_t4(ws):
    _tha_4_file(ws, 40)
    assert runner.quet_thu_muc(ws) == 1
    wb = openpyxl.load_workbook(next(ws.ra.glob("2026_W40_*.xlsx")))
    rk = _risk(wb)
    assert rk["KA"]["Base level (0-4)"] == 2 and rk["KA"]["T1: Overdue ≥ X"] == 1, "KA quá hạn 100 = X2 (top 15%)"
    assert all(v["T4: No contract / term, or new ≥ Y"] == 1 for v in rk.values()), "chưa có credit term → T4"
    sm = _sm(wb)
    assert (sm["muc_0"], sm["muc_1"], sm["muc_2"], sm["muc_3"], sm["muc_4"]) == (0, 2, 1, 1, 0)
    assert sm["nguong_x"] == 100
    rec = wb[th.SHEET_BANG]
    cf = [(str(c.sqref), [r.formula[0] for r in c.rules]) for c in rec.conditional_formatting]
    assert any(s_.startswith("D6") and len(f_) == 5 for s_, f_ in cf), "5 màu, chỉ tô ô Risk level"
    assert not any(s_.startswith("A6") for s_, _ in cf), "không tô cả hàng"


def test_t2_han_muc_khong_nang_hang_va_lai_gop(ws):
    ws.ghep.mkdir(parents=True, exist_ok=True)
    f = kh.ghi_master(ws.ghep / kh.TEN_FILE_MAC_DINH, [
        dict(ma="KA", ten=A, term="30 days", co_hd="✓", han_muc=50),          # T2: dư 100 > 50
        dict(ma="KB", ten=B, term="30 days", co_hd="✓", khong_nang="✓"),      # khách tốt: không nâng
        dict(ma="KC", ten="CÔNG TY C", term="30 days", co_hd="✓"),
        dict(ma="KK", ten="CÔNG TY KHÓ ĐÒI", term="30 days", co_hd="✓")], [])
    wb = openpyxl.load_workbook(f)
    sh = wb[kh.SHEET_THAM_SO]
    hang = {sh.cell(r, 1).value: r for r in range(1, 40)}
    for i in (1, 2, 3):
        sh.cell(hang[f"Lãi gộp tháng {i} của quý (VND)"], 2, 40)              # X1 = 25% × 40 = 10
    wb.save(f)
    _tha_4_file(ws, 40)
    assert runner.quet_thu_muc(ws) == 1
    wb = openpyxl.load_workbook(next(ws.ra.glob("2026_W40_*.xlsx")))
    rk = _risk(wb)
    assert _sm(wb)["nguong_x"] == 10, "X = MIN(X1 10, X2 100)"
    a, b = rk["KA"], rk["KB"]
    assert a["Credit limit"] == 50 and a["Credit limit source"] == "Hợp đồng" and a["T2: Total > credit limit"] == 1
    assert b["T1: Overdue ≥ X"] == 1 and b["Do not upgrade"] == "Có" and b["Upgrade +1"] == 0
    c = rk["KC"]                  # 1 hoá đơn 80 trong 6 tháng → billing 13,33 × 1 × 1,3 = 17,33 < dư 100 → T2 (tạm tính)
    assert c["Credit limit source"] == "Tạm tính" and c["Credit limit"] == 17.33 and c["T2: Total > credit limit"] == 1
    assert c["T4: No contract / term, or new ≥ Y"] == 0, "đã ✓ hợp đồng, có term, dư 100 < Y"
    dgia = wb[th.SHEET_DANH_GIA]                      # sheet Đánh giá như file v3 của Head FIN
    cot = {dgia.cell(5, c).value: c for c in range(1, dgia.max_column + 1)}
    ka = _hang(dgia, "KA")
    assert dgia.cell(ka, cot["Credit limit áp dụng"]).value == 50 and dgia.cell(ka, cot["KẾT QUẢ"]).value == \
        "NÂNG HẠNG RỦI RO"
    assert dgia.cell(_hang(dgia, "KB"), cot["KẾT QUẢ"]).value in (None, ""), "khách ✓ không nâng hạng"
    tso = {r[0]: r[1] for r in wb[th.SHEET_THAM_SO].iter_rows(min_row=8, values_only=True) if r[0]}
    assert tso["X áp dụng = MIN(X1, X2)"] == 10 and tso["Lãi gộp bình quân / tháng"] == 40
    assert tso["Tỷ lệ áp dụng trên lãi gộp"] == 0.25 and tso["Tỷ lệ top khách đang quá hạn"] == 0.15
    sm = _sm(wb)
    assert (sm["muc_0"], sm["muc_1"], sm["muc_2"], sm["muc_3"]) == (0, 3, 0, 1), \
        "KB Vàng (không nâng) · KC, KK Trắng + T2 = Vàng · KA Cam + 1 = Đỏ"


def test_o_sales_giu_tuan_truoc_va_nang_thang_len_do(ws):
    _tha_4_file(ws, 40)
    assert runner.quet_thu_muc(ws) == 1
    w40 = next(ws.ra.glob("2026_W40_*.xlsx"))
    time.sleep(1.1)
    _sua(w40, "KC", th.COT_DANH_GIA[0], "Có")                # Sales: KC có thông tin bất lợi
    _tha_4_file(ws, 41)
    assert runner.quet_thu_muc(ws) == 1
    wb = openpyxl.load_workbook(next(ws.ra.glob("2026_W41_*.xlsx")))
    rec = wb[th.SHEET_BANG]
    assert rec.cell(_hang(rec, "KC"), th.COT_DANH_GIA[0]).value == "Có", "tuần mới giữ ô Sales của tuần trước"
    sm = _sm(wb)
    assert (sm["muc_3"], sm["muc_2"], sm["muc_1"]) == (2, 1, 1), "KC lên thẳng Đỏ vì thông tin bất lợi (KA Đỏ sẵn)"


@pytest.mark.skipif(not shutil.which("soffice"), reason="Cần LibreOffice để tính lại công thức")
def test_cong_thuc_excel_khop_python_va_doi_theo_o_sales(ws, tmp_path):
    _tha_4_file(ws, 40)
    assert runner.quet_thu_muc(ws) == 1
    f = next(ws.ra.glob("2026_W40_*.xlsx"))
    out = tmp_path / "lo"
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(out), str(f)],
                   check=True, capture_output=True, timeout=120)
    wb = openpyxl.load_workbook(out / f.name, data_only=True)
    rec = wb[th.SHEET_BANG]
    muc = {rec.cell(r, 2).value: rec.cell(r, th.C_MUC).value for r in range(6, rec.max_row)
           if rec.cell(r, 2).value and rec.cell(r, 2).value != "TOTAL"}
    assert muc == {"KA": "Đỏ", "KB": "Cam", "KC": "Vàng", "KK": "Vàng"}, muc
    loi = [o.coordinate for s_ in wb for r in s_.iter_rows() for o in r
           if isinstance(o.value, str) and o.value.startswith(("#", "Err:"))]
    assert not loi, f"công thức lỗi: {loi[:5]}"
    # Sales sửa tay trong Excel → mức đổi theo, không cần chạy tool
    time.sleep(1.1)
    _sua(f, "KC", th.COT_DANH_GIA[0], "Có")
    out2 = tmp_path / "lo2"
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(out2), str(f)],
                   check=True, capture_output=True, timeout=120)
    rec = openpyxl.load_workbook(out2 / f.name, data_only=True)[th.SHEET_BANG]
    assert rec.cell(_hang(rec, "KC"), th.C_MUC).value == "Đỏ"
