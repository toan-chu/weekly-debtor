# -*- coding: utf-8 -*-
"""WR-13 (FIN chốt 05/11): 3 mức rủi ro theo tiêu chí FIN (Bình thường / Cảnh báo / Rủi ro cao), ô "Risk level" cạnh
tên khách ở Receivables, 3 ô Sales đánh giá (giữ giá trị tuần trước), nhóm khách, phương án xử lý, điểm A–D,
hạn từng job trong nhãn dòng con. Dữ liệu GIẢ."""
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

from wr import chitiet as ctm, danhgia as dg  # noqa: E402
from wr import thuonghieu as th  # noqa: E402
import runner  # noqa: E402
from test_wr4 import _tha_4_file, ws  # noqa: E402,F401

D = dt.date
KHONG_SAO = ("Không", "Có", "Không")      # mặc định: không bất lợi, có phản hồi, không thất hứa


# --- luật FIN (wr/danhgia.py) -------------------------------------------------

def test_muc_theo_ngay_qua_han_lon_nhat():
    assert [dg.muc_theo_ngay(n) for n in (0, 1, 30, 31, 60, 61, 400)] == [1, 1, 1, 2, 2, 3, 3]


def test_ba_o_sales_nang_muc():
    assert dg.muc_cuoi(0, ("Có", "Có", "Không")) == 3, "thông tin bất lợi → Rủi ro cao kể cả chưa quá hạn"
    assert dg.muc_cuoi(5, ("Không", "Có", "Có")) == 3, "đang quá hạn + thất hứa → Rủi ro cao"
    assert dg.muc_cuoi(0, ("Không", "Có", "Có")) == 1, "chưa quá hạn thì thất hứa không nâng"
    assert dg.muc_cuoi(5, ("Không", "Không", "Không")) == 2, "quá hạn + không phản hồi → ít nhất Cảnh báo"
    assert dg.muc_cuoi(0, ("Không", "Không", "Không")) == 1
    assert dg.muc_cuoi(90, ("Không", "Không", "Không")) == 3, "không hạ mức"
    assert dg.muc_cuoi(45, KHONG_SAO) == 2


def test_nhom_khach_va_phuong_an():
    chot = D(2026, 10, 31)
    assert dg.nhom_khach(None, chot) == "Mới"
    assert dg.nhom_khach(D(2025, 12, 31), chot) == "Lâu năm"
    assert dg.nhom_khach(D(2026, 6, 1), chot) == "Mới", "hoá đơn đầu trong 6 tháng gần nhất"
    assert dg.nhom_khach(D(2026, 2, 1), chot) == "Thường xuyên"
    assert dg.hinh_thuc(3, "Mới") == dg.HINH_THUC[3][0] and dg.hinh_thuc(3, "Lâu năm") == dg.HINH_THUC[3][1]
    assert dg.phuong_an(2, "Mới", 40, ("Không", "Không", "Không")).endswith(dg.KHONG_PHAN_HOI)
    assert dg.KHONG_PHAN_HOI not in dg.phuong_an(2, "Mới", 40, KHONG_SAO)
    assert set(dg.PHUONG_AN) == {(n, m) for n in dg.NHOM for m in (1, 2, 3)}


def test_cham_diem_a_d():
    assert dg.cham_diem(0, 1.0, 0.0, 1.0) == (100, "A")
    assert dg.cham_diem(30, 0.8, 0.3, 0.7) == (30 + 20 + 15 + 6, "B")
    assert dg.cham_diem(61, 0.49, 0.61, 0.49) == (0, "D")
    assert dg.cham_diem(60, 0.5, 0.6, 0.5) == (36, "D"), "15 + 10 + 8 + 3 = 36 < 40"
    assert dg.cham_diem(60, 0.8, 0.6, 0.5) == (46, "C")
    assert dg.cham_diem(0, 0.8, 0.0, 0.0) == (80, "A"), "đúng ngưỡng 80 là A"


def test_thong_ke_cham_diem_tu_dau_nam_va_no_nam_truoc():
    def bt(ngay, so_hd, no=0.0, co=0.0, dg_=""):
        return ctm.ButToan(ngay_ht=ngay, ngay_hd=ngay if no else None, so_hd=so_hd, so_ct="X", dien_giai=dg_,
                           tk_doi_ung="1122", no=no, co=co)
    k = ctm.KhachChiTiet("K", "K", [
        bt(D(2025, 3, 1), "00000001", no=100),                      # năm trước, trả hết năm trước → không xét
        bt(D(2025, 4, 1), "", co=100, dg_="theo HD 00000001"),
        bt(D(2025, 12, 1), "00000002", no=200),                     # năm trước còn nợ ở 01/01 → xét
        bt(D(2026, 2, 1), "", co=200, dg_="theo HD 00000002"),
        bt(D(2026, 5, 1), "00000003", no=300),                      # trong năm, trả 100/300
        bt(D(2026, 6, 1), "", co=100, dg_="theo HD 00000003"),
    ])
    t = k.thong_ke_cham_diem(D(2026, 10, 30))
    assert t["so_hd"] == 2 and t["so_tra_du"] == 1
    assert t["gia_tri"] == 500 and t["da_thu"] == 300, "mẫu số gồm cả phần năm trước còn nợ → tỷ lệ ≤ 100%"


# --- báo cáo thật từ input ------------------------------------------------

def _hang(sh, ma):
    return next(r for r in range(6, sh.max_row + 1) if sh.cell(r, 2).value == ma)


def _sm(wb):
    return {r[5]: (r[2] or 0) for r in wb[th.SHEET_SUMMARY].iter_rows(min_row=20, values_only=True) if r and r[5]}


def _sua(f, ma, cot, gia_tri):
    wb = openpyxl.load_workbook(f)
    sh = wb[th.SHEET_BANG]
    sh.cell(_hang(sh, ma), cot, value=gia_tri)
    wb.save(f)


def test_o_muc_do_canh_ten_khach_va_3_o_sales(ws):
    _tha_4_file(ws, 40)
    assert runner.quet_thu_muc(ws) == 1
    wb = openpyxl.load_workbook(next(ws.ra.glob("2026_W40_*.xlsx")))
    rec = wb[th.SHEET_BANG]
    assert rec.cell(5, th.C_TEN).value == "Customer's Name" and rec.cell(5, th.C_MUC).value == "Risk level"
    ka = _hang(rec, "KA")
    assert str(rec.cell(ka, th.C_MUC).value).startswith("=IFERROR(INDEX('AR Risk'!"), "ô mức độ lấy từ AR Risk"
    assert [rec.cell(ka, c).value for c in th.COT_DANH_GIA] == list(th.DANH_GIA_MAC_DINH)
    cf = [(str(c.sqref), [r.formula[0] for r in c.rules]) for c in rec.conditional_formatting]
    assert any(s.startswith("D6") and len(f) == 3 for s, f in cf), "tô ô mức độ theo 3 màu Good/Neutral/Bad"
    dv = {str(d.sqref): d.formula1 for d in rec.data_validations.dataValidation}
    assert any(v == '"Có,Không"' and "Q6" in k for k, v in dv.items()), "3 ô Sales là danh sách Có/Không"
    # KA: HĐ 01/08, term mặc định 30 ngày → quá hạn 32 ngày tại 02/10 → Cảnh báo; B/C/K Bình thường
    risk = wb[th.SHEET_RISK]
    h = next(r for r in range(1, 8) if risk.cell(r, 1).value == "No.")
    cot = {risk.cell(h, c).value: c for c in range(1, risk.max_column + 1)}
    rk = next(r for r in range(h + 1, risk.max_row + 1) if risk.cell(r, 2).value == "KA")
    assert risk.cell(rk, cot["Max days overdue"]).value == 32
    sm = _sm(wb)
    assert (sm["muc_1"], sm["muc_2"], sm["muc_3"]) == (3, 1, 0)


def test_o_sales_giu_tuan_truoc_va_chay_lai_giu_tuan_nay(ws):
    _tha_4_file(ws, 40)
    assert runner.quet_thu_muc(ws) == 1
    w40 = next(ws.ra.glob("2026_W40_*.xlsx"))
    time.sleep(1.1)
    _sua(w40, "KB", th.COT_DANH_GIA[0], "Có")                # Sales: KB có thông tin bất lợi
    _tha_4_file(ws, 41)
    assert runner.quet_thu_muc(ws) == 1
    w41 = next(ws.ra.glob("2026_W41_*.xlsx"))
    wb = openpyxl.load_workbook(w41)
    rec = wb[th.SHEET_BANG]
    assert rec.cell(_hang(rec, "KB"), th.COT_DANH_GIA[0]).value == "Có", "tuần mới giữ ô Sales của tuần trước"
    assert _sm(wb)["muc_3"] == 1, "KB lên Rủi ro cao vì thông tin bất lợi"
    # Sales sửa lại ở báo cáo W41 rồi chạy lại cùng tuần → lấy ô của tuần này
    time.sleep(1.1)
    _sua(w41, "KB", th.COT_DANH_GIA[0], "Không")
    for f in ws.vao.glob("[[]DONE[]] *"):
        if "2." in f.name or "41" in f.name:
            f.rename(f.with_name(f.name.replace("[DONE] ", "")))
    assert runner.quet_thu_muc(ws) == 1
    lai = openpyxl.load_workbook(next(ws.ra.glob("2026_W41_*chay lai*.xlsx")))
    assert lai[th.SHEET_BANG].cell(_hang(lai[th.SHEET_BANG], "KB"), th.COT_DANH_GIA[0]).value == "Không"
    assert _sm(lai)["muc_3"] == 0


def test_dong_job_co_han_thanh_toan(ws, monkeypatch):
    from test_wr6 import BH, SCT, tao_ban_hang, tao_so_chi_tiet
    from wr import tindung
    monkeypatch.setattr(tindung, "DUNG_SAI_KHOP", 0.5)
    monkeypatch.setattr(runner, "CHO_SMS_GIAY", 7200)
    _tha_4_file(ws)
    tao_ban_hang(ws.vao / "bh.xlsx", BH)
    tao_so_chi_tiet(ws.vao / "sct.xlsx", SCT)
    assert runner.quet_thu_muc(ws) == 1
    rec = openpyxl.load_workbook(next(ws.ra.glob("2026_W40_*.xlsx")))[th.SHEET_BANG]
    nhan = rec.cell(_hang(rec, "KA") + 1, th.C_TEN).value
    assert "· hạn 31/08/2026" in nhan, f"hạn = ngày hoá đơn 01/08 + 30 ngày: {nhan}"


@pytest.mark.skipif(not shutil.which("soffice"), reason="Cần LibreOffice để tính lại công thức")
def test_cong_thuc_muc_do_trong_excel_khop_python(ws, tmp_path):
    _tha_4_file(ws, 40)
    assert runner.quet_thu_muc(ws) == 1
    f = next(ws.ra.glob("2026_W40_*.xlsx"))
    time.sleep(1.1)
    _sua(f, "KC", th.COT_DANH_GIA[0], "Có")          # sửa tay trong Excel → công thức phải đổi theo
    out = tmp_path / "lo"
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(out), str(f)],
                   check=True, capture_output=True, timeout=120)
    wb = openpyxl.load_workbook(out / f.name, data_only=True)
    rec = wb[th.SHEET_BANG]
    muc = {rec.cell(r, 2).value: rec.cell(r, th.C_MUC).value for r in range(6, rec.max_row)
           if rec.cell(r, 2).value and rec.cell(r, 2).value != "TOTAL"}
    assert muc == {"KA": "Cảnh báo", "KB": "Bình thường", "KC": "Rủi ro cao", "KK": "Bình thường"}, muc
    risk = wb[th.SHEET_RISK]
    loi = [o.coordinate for r in risk.iter_rows() for o in r if isinstance(o.value, str) and o.value.startswith("#")]
    assert not loi, f"công thức lỗi: {loi[:5]}"
