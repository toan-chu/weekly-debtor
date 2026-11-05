# -*- coding: utf-8 -*-
"""Tool báo cáo công nợ tuần — điểm chạy.

Lịch chạy nền gọi:
    python runner.py                 quét thư mục input, thấy đủ 4 file thì chạy:
                                     Tổng hợp 131 + Chi tiết 131 (từ 2022) + Tổng hợp 331 + SMS

Kiểm tra file ghép khách (nút 2_Kiem_tra_ghep_khach.bat):
    python runner.py --kiem-tra-ghep-khach

Chạy tay khi thử nghiệm:
    python runner.py --tong-hop A.xlsx --chi-tiet B.xlsx [--phai-tra C.xlsx --sms D.xlsx] --ra <thư mục>
    python runner.py --tong-hop A.xlsx --tuoi-no B.xlsx [--sms C.xlsx --ghep D.xlsx] --ra <thư mục>  (v0.3)
    python runner.py --nap-bao-cao "2026_W36_....xlsx" --ngay-chot 2026-09-05
    python runner.py --ngoai-le "TEN KHACH" --nhom "1 - 30" --vi-sao "lô tàu chìm"
    python runner.py --credit-term "TEN KHACH" --gia-tri "15 days"
    python runner.py --bo-qua "TEN KHACH" --vi-sao "no ao do loi xuat hoa don"
    python runner.py --nhan-lai "TEN KHACH"

Đặt hoặc đổi thư mục đích (nơi kế toán thả file và nhận kết quả):
    python runner.py --dat-thu-muc "<đường dẫn thư mục đích>"
    python runner.py --xem-cau-hinh
"""
from __future__ import annotations

import sys as _sys
_sys.dont_write_bytecode = True          # thay cho PYTHONDONTWRITEBYTECODE trong chay_nen.bat cũ

import argparse
import datetime as dt
import os
import sys
import time
import traceback
from pathlib import Path
from typing import List, Optional, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wr import banhang, baocao, caidat, chitiet, ghep, ghifull, misa, tindung
from wr import soghi as sg
from wr import tinhtoan
from wr.misa import LoiDocFile, chuan_hoa_ten

# Mặc định ghi cạnh mã nguồn; có thư mục đích rồi thì chuyển vào đó.
NHAT_KY = Path(__file__).resolve().parent / "log" / "runner.log"
# Kế toán thả file Misa trước, file SMS / phải trả sau: chờ tối đa chừng này rồi mới chạy thiếu
CHO_SMS_GIAY = 2 * 3600
# Các file mỗi tuần phải xuất cùng lúc; cách nhau hơn chừng này thì cảnh báo ở tab Review
LECH_XUAT_GIAY = 24 * 3600


def dat_nhat_ky(duong_dan: Path) -> None:
    global NHAT_KY
    NHAT_KY = Path(duong_dan)


def _in(dong: str) -> None:
    """Chạy nền bằng pythonw thì không có màn hình, và console Windows có thể
    không hiện được tiếng Việt. Cả hai đều không được phép làm chết tiến trình."""
    try:
        print(dong)
    except (AttributeError, ValueError, UnicodeEncodeError, OSError):
        try:
            sys.stdout.buffer.write((dong + "\n").encode("utf-8", "replace"))
        except Exception:
            pass


def _mo_ta(f: Path) -> str:
    t = dt.datetime.fromtimestamp(f.stat().st_mtime)
    return f"{f.name} (file sửa lần cuối {t:%d/%m/%Y %H:%M})"


def ghi_log(dong: str) -> None:
    try:
        NHAT_KY.parent.mkdir(parents=True, exist_ok=True)
        with NHAT_KY.open("a", encoding="utf-8") as f:
            f.write(f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}  {dong}\n")
    except OSError:
        pass          # không ghi được nhật ký cũng không được chặn việc chính
    _in(dong)


def _ghep_cap(files: List[Path]):
    """Ghép từng cặp file CÙNG KỲ, trả về danh sách cặp xếp theo thứ tự thời gian.

    Thả nhiều tuần vào một lúc cũng chạy được: mỗi tuần một cặp, chạy lần lượt
    từ tuần cũ nhất, để sổ ghi nối đúng thứ tự.
    """
    ds_th, ds_tn, la = [], [], []
    for f in files:
        loai = misa.nhan_dang(f)
        try:
            if loai == "tong-hop":
                ds_th.append((misa.doc_tong_hop(f).den_ngay, f))
            elif loai == "tuoi-no":
                ds_tn.append((misa.doc_tuoi_no(f).den_ngay, f))
            else:
                la.append((f, "không nhận ra là file Misa nào"))
        except Exception as loi:
            la.append((f, str(loi)))

    cap, con_tn = [], dict(ds_tn)
    for ngay_th, f_th in sorted(ds_th):
        mong_doi = ngay_th + dt.timedelta(days=1)
        f_tn = next((f for n, f in sorted(ds_tn) if n == mong_doi and f in con_tn.values()), None)
        if f_tn is None:
            la.append((f_th, f"chưa có file Tuổi nợ chốt ngày {mong_doi:%d/%m/%Y} đi kèm"))
            continue
        con_tn = {n: f for n, f in con_tn.items() if f != f_tn}
        cap.append((ngay_th, f_th, f_tn))
    for _, f in sorted(con_tn.items()):
        la.append((f, "chưa có file Tổng hợp cùng kỳ đi kèm"))
    return cap, la


def _doi_ten(f: Path, tien_to: str) -> None:
    """Đổi tên file đầu vào. File đang bị khoá (OneDrive, Excel) thì ghi nhật ký,
    không để tiến trình chạy nền chết im lặng."""
    try:
        f.rename(f.with_name(f"{tien_to} {f.name}"))
    except OSError as loi:
        ghi_log(f"  ! Không đổi tên được {f.name} thành {tien_to}: {loi}")


def nap_ban_tay(thu_muc: Path, so: sg.SoGhi) -> int:
    """Nạp các báo cáo tuần nằm ở thư mục gốc (bản tool sinh rồi kế toán sửa, hoặc
    bản final kế toán trưởng bỏ vào) vào sổ ghi, đúng ngày chốt ghi trong file.
    Cùng một ngày chốt có nhiều bản thì bản sửa sau cùng thắng. Bản của người
    thắng bản của máy. File đã nạp (cùng giờ sửa) thì bỏ qua."""
    n = 0
    if not thu_muc or not Path(thu_muc).is_dir():
        return 0
    for f in sorted(Path(thu_muc).glob("*.xlsx"), key=lambda f: f.stat().st_mtime):
        if f.name.startswith("~$"):
            continue
        dau = f"{f.stat().st_mtime:.0f}"
        if so.da_nap.get(f.name) == dau:
            continue
        try:
            ngay = baocao.ngay_chot_trong_file(f)
            if ngay is None:
                ghi_log(f"  ! {f.name}: không thấy ngày chốt (dòng 'Reporting period dd/mm/yyyy - "
                        "dd/mm/yyyy' hoặc 'At dd Mon yyyy'), chưa nạp")
                continue
            # Bản người: không có tab Review (KTT tự làm), tên có chữ "human", hoặc file tool ghi mà
            # sau đó có người sửa (giờ sửa khác lúc tool ghi) — WR-08: FIN sửa thẳng trên file của tool
            ghi_luc = so.file_tool.get(f.name)
            ban_nguoi = (not baocao.co_tab_review(f) or "human" in f.name.lower()
                         or (ghi_luc is not None and abs(f.stat().st_mtime - ghi_luc) > 2))
            # WR-12 (user 27/10: "tránh Sales sửa lại lời mình đã ghi"): tuần đã có báo cáo tuần sau là tuần ĐÃ CHỐT —
            # sửa file của tuần đó không còn tác dụng (lý do tuần trước giữ đúng lời đã ghi khi tuần đó còn mở)
            if ngay.isoformat() in so.trang and any(n > ngay.isoformat() for n in so.trang):
                so.da_nap[f.name] = dau
                ghi_log(f"  {f.name}: tuần {ngay:%d/%m/%Y} đã chốt (đã có báo cáo tuần sau) — không đọc lại chỉnh sửa")
                continue
            cu = so.nguon_trang.get(ngay.isoformat())
            if cu and cu.get("ban_nguoi") and not ban_nguoi and cu.get("file") != f.name:
                # WR-04: bản người làm (không có tab Review) luôn thắng bản tool cùng ngày chốt
                so.da_nap[f.name] = dau
                ghi_log(f"  {f.name}: cùng ngày chốt {ngay:%d/%m/%Y} với bản người làm "
                        f"{cu.get('file')} — giữ bản người làm")
                continue
            k = baocao.nap_vao_so_ghi(so, f, ngay)
        except Exception as loi:
            ghi_log(f"  ! Không đọc được bản final {f.name}: {loi}")
            continue
        so.da_nap[f.name] = dau
        so.nguon_trang[ngay.isoformat()] = {"file": f.name, "ban_nguoi": ban_nguoi}
        n += 1
        ghi_log(f"Đã đọc lại báo cáo {f.name}: {k} khách, chốt {ngay:%d/%m/%Y}"
                f"{' (bản người làm)' if ban_nguoi else ''} — tuần sau tính tiếp từ bản này")
    return n


def chay_mot_lan(
    duong_tong_hop: Path,
    duong_tuoi_no: Path,
    thu_muc_ra: Path,
    duong_so_ghi: Path,
    mau_chuan: Path,
    duong_sms: Optional[Path] = None,
    thu_muc_ghep: Optional[Path] = None,
    duong_ghep: Optional[Path] = None,
) -> Path:
    th = misa.doc_tong_hop(duong_tong_hop)
    tn = misa.doc_tuoi_no(duong_tuoi_no)
    misa.kiem_cung_ky(th, tn)

    so = sg.SoGhi(duong_so_ghi)
    ngay_chot = tn.den_ngay
    truoc = so.ngay_trang_truoc(ngay_chot)
    if truoc is None:
        ghi_log("  ! Sổ ghi chưa có trang nào trước ngày chốt — chạy chế độ khởi động lạnh")
    elif (ngay_chot - truoc).days > 10:
        ghi_log(f"  ! Trang gần nhất của sổ ghi là {truoc:%d/%m/%Y}, cách {(ngay_chot - truoc).days} ngày. "
                "Nợ mới phát sinh trong lúc nghỉ sẽ bị đoán trẻ hơn thực tế")

    them: dict = {}
    nguon_them: list = []
    if duong_sms:
        dong_sms = ghep.doc_sms(duong_sms)
        f_ghep = Path(duong_ghep) if duong_ghep else ghep.tim_file_ghep(thu_muc_ghep)
        bang = ghep.doc_bang_ghep(f_ghep) if f_ghep else None
        if bang is None:
            ghi_log("  ! Chưa có file ghép khách trong thư mục sample — chỉ ghép được khách trùng tên/mã")
        else:
            for ma_sms, cu, moi in bang.xung_dot:
                ghi_log(f"  ! Mã SMS {ma_sms} được ghép với cả {cu} và {moi}: theo phía Misa ({moi})")
        kq = tinhtoan.tinh_bang_sms(th, tn, so, dong_sms, bang, ngay_chot)
        dong, da_bo = kq.dong, kq.da_bo
        them = {"che_do": "sms", "sms_cat": kq.sms_cat, "sms_khong_misa": kq.sms_khong_misa,
                "sms_chua_ghep": kq.sms_chua_ghep}
        nguon_them = [("", _mo_ta(Path(duong_sms))),
                      ("", f"Bảng ghép khách: {f_ghep.name}" if f_ghep else "Bảng ghép khách: (chưa có)")]
    else:
        dong = tinhtoan.tinh_bang(th, tn, so, ngay_chot)
        da_bo = getattr(tinhtoan.tinh_bang, "da_bo", [])
    dich = thu_muc_ra / baocao.ten_file_bao_cao(th.tu_ngay, th.den_ngay)
    if dich.exists():
        # Báo cáo cũ của tuần này có thể đã được kế toán sửa tay: không ghi đè (RULES 5.1)
        dich = dich.with_name(f"{dich.stem} (chay lai {dt.datetime.now():%d-%m %Hh%M}){dich.suffix}")
    khach_lech = []
    for k, v in tn.dong.items():
        t = th.dong.get(k)
        so_tong_hop = t.du_no if t else 0.0
        if abs(v.tong - so_tong_hop) >= 2:
            khach_lech.append((v.ten, so_tong_hop, v.tong))
    khach_lech.sort(key=lambda x: -abs(x[2] - x[1]))

    baocao.ghi_bao_cao(
        dong, ngay_chot, mau_chuan, dich, tong_misa=th.tong_du_no,
        doi_chieu={
            # Chế độ SMS không dùng file Tuổi nợ để chia tuổi, nên không liệt kê độ vênh của nó
            "tong_tuoi_no": None if duong_sms else sum(v.tong for v in tn.dong.values()),
            "khach_lech": [] if duong_sms else khach_lech[:30],
            "da_bo": da_bo,
            **them,
            "nguon": [("Nguồn dữ liệu", _mo_ta(Path(duong_tong_hop))),
                      ("", _mo_ta(Path(duong_tuoi_no))), *nguon_them,
                      ("", "Số trong hai file này thay đổi theo thời điểm xuất. "
                           "Chỉ so báo cáo với bản làm tay khi cả hai dùng cùng một lần xuất Misa.")],
        },
    )

    so.file_tool[dich.name] = dich.stat().st_mtime
    trang_cu = so.trang.get(ngay_chot.isoformat()) or {}
    if any(v.nguon == "ban-tay" for v in trang_cu.values()):
        ghi_log(f"  Sổ ghi ngày {ngay_chot:%d/%m/%Y} đã lấy từ báo cáo trong thư mục gốc, giữ nguyên")
    else:
        so.ghi_trang(ngay_chot, {d.ten_chuan: tinhtoan.thanh_ho_so(d) for d in dong})
        so.ghi()

    can_xem = sum(1 for d in dong if d.cho_xem_lai)
    ghi_log(f"  -> {dich.name}: {len(dong)} khách, tổng {th.tong_du_no:,.0f}, "
            f"{can_xem} dòng cần xem lại")
    return dich


def chay_full(
    duong_tong_hop: Path,
    duong_chi_tiet: Path,
    thu_muc_ra: Path,
    duong_so_ghi: Path,
    mau_chuan: Path,
    duong_phai_tra: Optional[Path] = None,
    duong_sms: Optional[Path] = None,
    thu_muc_ghep: Optional[Path] = None,
    duong_ghep: Optional[Path] = None,
    duong_ban_hang: Optional[Path] = None,
    duong_so_ct: Optional[Path] = None,
) -> Path:
    """v0.4 — full credit report từ 4 file (WR-04). v0.5 — thêm Bán hàng + Sổ chi tiết TK131 (WR-06):
    job chưa thanh toán lấy từ Misa, đối chiếu với dư nợ từng khách."""
    th = misa.doc_tong_hop(duong_tong_hop)
    ct = chitiet.doc_chi_tiet(duong_chi_tiet)
    if ct.den_ngay > th.den_ngay:
        # FIN hay xuất Chi tiết đến hết tháng: cắt về đúng ngày cuối kỳ, không bắt xuất lại
        bo = ct.cat_den(th.den_ngay)
        ghi_log(f"  File Chi tiết 131 xuất tới sau ngày cuối kỳ: đã bỏ {bo} bút toán sau {th.den_ngay:%d/%m/%Y}")
    if ct.den_ngay != th.den_ngay:
        raise LoiDocFile(
            "File Chi tiết và file Tổng hợp phải thu không cùng ngày cuối kỳ, dừng để khỏi ra số sai.\n"
            f"  Tổng hợp: đến {th.den_ngay:%d/%m/%Y} · Chi tiết: đến {ct.den_ngay:%d/%m/%Y}")
    canh_bao: List[str] = []
    if ghep.FILE_MASTER and not ghep.FILE_MASTER.is_file():
        canh_bao.append(f"Không thấy file credit term chung: {ghep.FILE_MASTER}. OneDrive trên máy này chưa đồng bộ "
                        "file đó (hoặc file bị đổi tên/chuyển chỗ) — tuần này dùng bảng term trong sample/.")
        ghi_log(canh_bao[-1])
    if ct.tu_ngay > dt.date(2022, 1, 1):
        canh_bao.append(f"File Chi tiết 131 chỉ có từ {ct.tu_ngay:%d/%m/%Y}: hoá đơn cũ hơn không có ngày, "
                        "tuổi nợ của phần đó bị tính trẻ hơn thực tế. FIN xuất lại từ 01/01/2022.")
    pt = None
    if duong_phai_tra:
        pt = misa.doc_phai_tra(duong_phai_tra)
        if pt.den_ngay != th.den_ngay:
            canh_bao.append(f"File phải trả {Path(duong_phai_tra).name} đến ngày {pt.den_ngay:%d/%m/%Y}, "
                            f"khác kỳ phải thu ({th.den_ngay:%d/%m/%Y}) — không dùng.")
            pt = None
    else:
        canh_bao.append("Không có file Tổng hợp công nợ phải trả (TK331) — sheet Payable để trống.")

    so = sg.SoGhi(duong_so_ghi)
    ngay_chot = th.den_ngay + dt.timedelta(days=1)
    truoc = so.ngay_trang_truoc(ngay_chot)
    if truoc is None:
        ghi_log("  ! Sổ ghi chưa có tuần nào trước — không so được với tuần trước (cột Last week trống)")
        canh_bao.append("Chưa có báo cáo tuần trước trong sổ ghi: các cột so sánh tuần trước để trống, "
                        "dấu hiệu S1-S3 chưa tin được.")
    elif (ngay_chot - truoc).days > 10:
        canh_bao.append(f"Báo cáo tuần trước gần nhất là {truoc:%d/%m/%Y}, cách {(ngay_chot - truoc).days} "
                        "ngày: cột 'Last week' là của tuần đó.")

    dong_sms, bang, f_ghep = [], None, None
    if duong_sms:
        dong_sms = ghep.doc_sms(duong_sms)
        f_ghep = Path(duong_ghep) if duong_ghep else ghep.tim_file_ghep(thu_muc_ghep)
        bang = ghep.doc_bang_ghep(f_ghep) if f_ghep else None
        if bang is None:
            ghi_log("  ! Chưa có file ghép khách trong thư mục sample — chỉ ghép được khách trùng tên/mã")
    else:
        canh_bao.append("Không có file SMS: Job Number và Salesman lấy từ báo cáo tuần trước.")

    files = [Path(f) for f in (duong_tong_hop, duong_chi_tiet, duong_phai_tra, duong_sms) if f]
    gio = [f.stat().st_mtime for f in files]
    if max(gio) - min(gio) > LECH_XUAT_GIAY:
        canh_bao.append("Các file đầu vào tải về cách nhau hơn 1 ngày — số Misa đổi theo lúc xuất, "
                        "nên xuất cả 4 file cùng lúc.")

    thu_muc_mau = Path(duong_ghep).parent if duong_ghep else thu_muc_ghep
    job_no_cu = ghep.doc_job_no_cu(thu_muc_mau)
    job_misa = None
    if duong_ban_hang:
        job_misa = banhang.doc_job_misa(duong_ban_hang, duong_so_ct, [(k.ma, k.ten) for k in th.dong.values()])
        if job_misa is None:
            canh_bao.append(f"File Bán hàng {Path(duong_ban_hang).name} không có hoá đơn nào — dùng job từ SMS.")
        else:
            files += [Path(f) for f in (duong_ban_hang, duong_so_ct) if f]
            if not job_misa.co_job:
                canh_bao.append("Có file Bán hàng nhưng thiếu Sổ chi tiết TK131: biết hoá đơn nào chưa thanh "
                                "toán nhưng không biết job nào.")
            if job_misa.den_ngay < th.den_ngay - dt.timedelta(days=7):
                canh_bao.append(f"File Bán hàng chỉ có tới {job_misa.den_ngay:%d/%m/%Y}, kỳ báo cáo tới "
                                f"{th.den_ngay:%d/%m/%Y}: FIN xuất lại cùng lúc với các file khác.")
    elif duong_so_ct:
        canh_bao.append("Có Sổ chi tiết TK131 nhưng thiếu file Bán hàng: chưa biết trạng thái thanh toán, "
                        "job vẫn lấy từ SMS.")
    else:
        canh_bao.append("Không có file Bán hàng + Sổ chi tiết TK131: job lấy từ SMS (cách cũ), chưa có cột "
                        "Đối chiếu.")
    credit_terms = ghep.doc_credit_term(thu_muc_mau)
    # WR-08 (user 13/10): cảnh báo hợp đồng / credit term từ Client Master → sheet Contracts + Summary
    # đọc TRƯỚC khi tính — hạn mức nợ, "Đã có hợp đồng", link hợp đồng (WR-14 sẽ dùng cho tiêu chí FIN)
    from wr import khachhang
    # WR-12 (user 28/10): file credit term còn mẫu cũ (thiếu cột / sheet tên cũ) → tool tự nâng lên mẫu mới ngay đầu lần
    # chạy, kể cả tuần không có khách mới. Cùng điều kiện an toàn: chỉ thêm cột, sao lưu trước, file đang mở thì để sau.
    f_ct0 = khachhang.tim_file_credit_term(thu_muc_mau)
    if f_ct0 is not None and khachhang.thieu_mau(f_ct0):
        try:
            _n, cau = khachhang.them_khach_moi(f_ct0, [], Path(duong_so_ghi).parent / "luu-tru", ngay_chot)
        except Exception as loi:             # không bao giờ vì việc này mà không ra báo cáo
            cau = f"Chưa nâng {f_ct0.name} lên mẫu mới: {loi}"
        if cau:
            ghi_log(f"  {cau}")
            canh_bao.append(cau)
    master = khachhang.doc_master(thu_muc_mau, ngay_chot)
    co_cot_hd = any(m.get("co_hd") is not None for m in master.values())
    if master and not co_cot_hd:
        canh_bao.append("File credit term chưa có cột 'Đã có hợp đồng' (mẫu cũ): chưa đếm được khách chưa có hợp đồng. "
                        "Chạy tools/nang_cap_credit_term.py một lần.")
    # WR-13: hạn mức nợ chưa dùng để xếp mức (tiêu chí TC2 chờ Client Master, WR-14) → không cảnh báo nữa
    kq = tindung.tinh_full(th, ct, so, dong_sms, bang, pt, ngay_chot, job_no_cu=job_no_cu, job_misa=job_misa,
                           credit_terms=credit_terms, ma_no_cu=ghep.ma_job_no_cu(thu_muc_mau), master=master)
    kq.master = master
    term_truoc = {str(m).upper(): hs.credit_term for m, hs in (so.trang_truoc(ngay_chot) or {}).items()
                  if getattr(hs, "credit_term", "")}
    canh_bao_hd = khachhang.canh_bao_hop_dong(master, kq.rui_ro, ngay_chot, term_truoc)
    # WR-12 (FIN 27/10: "tự update vào đúng chỗ, không cần can thiệp"): khách đang nợ chưa có trong file credit term →
    # tool tự thêm dòng (chỉ thêm, sao lưu trước, file đang mở thì để lần sau). Báo cáo vẫn ra dù thêm được hay không.
    moi = [c for c in canh_bao_hd if c["loai"] == khachhang.LOAI_CHUA_MASTER]
    kq.them_khach = ""
    if moi:
        f_ct = khachhang.tim_file_credit_term(thu_muc_mau)
        if f_ct is None and not master:
            f_ct = Path(thu_muc_mau) / khachhang.TEN_FILE_MAC_DINH if thu_muc_mau else None
        if f_ct is None:
            kq.them_khach = ("Chưa thêm khách mới: không thấy file credit term đúng mẫu (sheet 'Credit term') trong "
                             "sample/ — chạy tools/nang_cap_credit_term.py với file đang dùng.")
        else:
            try:
                n_them, kq.them_khach = khachhang.them_khach_moi(f_ct, moi, Path(duong_so_ghi).parent / "luu-tru",
                                                                 ngay_chot)
            except Exception as loi:             # không bao giờ vì việc này mà không ra báo cáo
                n_them, kq.them_khach = 0, f"Chưa thêm được khách mới vào {f_ct.name}: {loi}"
            for c in moi:
                c["da_them"] = n_them > 0
        if kq.them_khach:
            ghi_log(f"  {kq.them_khach}")
            if not any(c.get("da_them") for c in moi):
                canh_bao.append(kq.them_khach)
    dem = lambda f: sum(1 for c in canh_bao_hd if f(c["loai"]))
    kq.chi_so.update({
        "hd_het_han": dem(lambda x: x == "Hợp đồng đã hết hạn"),
        "hd_sap_het_30": dem(lambda x: x == "Sắp hết hạn ≤30 ngày"),
        "khach_chua_term": dem(lambda x: x == "Chưa có credit term"),
        "term_doi": dem(lambda x: x == "Credit term vừa đổi"),
        "khach_chua_han_hd": sum(1 for r in kq.rui_ro if not (master.get(r.ma.upper()) or {}).get("han_dung")),
        # None (ô trống trên Summary) khi Client Master chưa có cột — không đếm sai (lỗi tự soi WR-09)
        "khach_chua_hd": (sum(1 for r in kq.rui_ro if r.co_hd is False) if co_cot_hd else None),
        "khach_chua_master": dem(lambda x: x == khachhang.LOAI_CHUA_MASTER),
        "khach_chua_han_muc": sum(1 for r in kq.rui_ro if not r.han_muc),
        **{f"muc_{i}": sum(1 for r in kq.rui_ro if r.muc_fin == i) for i in (1, 2, 3)},
        "no_rui_ro_cao": round(sum(r.cuoi_ky for r in kq.rui_ro if r.muc_fin == 3), 2),
        "hang_d": sum(1 for r in kq.rui_ro if r.hang == "D"),
        "sales_chua_giai_trinh": len({r.salesman or "(chưa có Sales)" for r in kq.rui_ro
                                      if r.qua_han > 0.5 and r.s[3]}),
    })
    # Lịch sử tuần: tuần đầu dùng v0.4 thì dựng lại điểm của tuần trước từ báo cáo cũ (sổ ghi)
    truoc_ngay = so.ngay_trang_truoc(ngay_chot)
    if (truoc_ngay and (ngay_chot - truoc_ngay).days >= 5
            and truoc_ngay.isoformat() not in so.lich_su_tuan and kq.chi_so_truoc):
        den_cu = truoc_ngay - dt.timedelta(days=1)
        so.lich_su_tuan[truoc_ngay.isoformat()] = {**kq.chi_so_truoc, "ngay_chot": truoc_ngay.isoformat(),
                                                  "tuan": f"W{den_cu.isocalendar()[1]:02d}",
                                                  "den_ngay": den_cu.isoformat(), "tu_bao_cao_cu": True}
    so.lich_su_tuan[ngay_chot.isoformat()] = kq.chi_so
    lich_su = [v for k, v in sorted(so.lich_su_tuan.items()) if dt.date.fromisoformat(k) <= ngay_chot]

    dich = thu_muc_ra / baocao.ten_file_bao_cao(th.tu_ngay, th.den_ngay)
    if dich.exists():
        # Báo cáo cũ của tuần này có thể đã được kế toán sửa tay: không ghi đè (RULES 5.1)
        dich = dich.with_name(f"{dich.stem} (chay lai {dt.datetime.now():%d-%m %Hh%M}){dich.suffix}")
    ky = f"{th.tu_ngay:%d/%m/%Y} - {th.den_ngay:%d/%m/%Y}"
    nguon = [_mo_ta(f) for f in files] + ([f"Bảng ghép khách: {f_ghep.name}"] if f_ghep else [])

    bang_them = [
        ("File Chi tiết 131 và Tổng hợp 131 lệch nhau — tiền theo Tổng hợp",
         ("Khách hàng", "Tổng hợp", "Chi tiết", "Chênh"),
         [(t, a, b, a - b) for t, a, b in kq.lech_chi_tiet]),
        ("Phải thu nhập bổ sung cho kỳ trước — dư đầu kỳ Misa khác dư cuối báo cáo tuần trước",
         ("Khách hàng", "Báo cáo tuần trước", "Dư đầu kỳ Misa", "Chênh"),
         [(r.ten, r.tt_total, r.dau_ky, r.nhap_bo_sung) for r in kq.rui_ro
          if r.nhap_bo_sung is not None and abs(r.nhap_bo_sung) >= 1]),
        ("Khách trả trước (dư Có TK131) — không đưa vào Receivables",
         ("Khách hàng", "Dư Có", "Mã", ""),
         [(k.ten, k.du_co, k.ma, "") for k in th.dong.values() if k.du_co > 0.5]),
        ("Vendor mình trả trước (dư Nợ TK331) — không đưa vào Payable",
         ("Vendor", "Dư Nợ", "Mã", ""), [(t, v, m, "") for m, t, v in kq.vendor_tra_truoc]),
    ]
    if kq.job_misa:
        # v0.5 (WR-06): lệnh báo cho FIN — gom theo màu, kèm gợi ý nguyên nhân
        for mau, tieu in (("do", "ĐỎ — tổng job chưa thanh toán LỚN hơn dư nợ: đã thu chưa đối trừ, hoặc "
                                 "thanh toán một phần Misa chưa tự trừ"),
                          ("vang", "VÀNG — lệch nhỏ: nghi quên đánh giá chênh lệch tỷ giá"),
                          ("cam", "CAM — dư nợ LỚN hơn tổng job chưa thanh toán: nợ chưa gắn job (nợ cũ → "
                                  "bảng job nợ cũ) hoặc hoá đơn chưa lên Bán hàng")):
            ds = [d for d in kq.dong if d.doi_chieu.get("mau") == mau]
            bang_them.append((tieu, ("Khách hàng", "Dư nợ Misa", "Tổng job chưa TT", "Lệch"),
                              [(d.ten, d.total, d.doi_chieu["tong_job"], d.doi_chieu["lech"])
                               for d in sorted(ds, key=lambda d: -abs(d.doi_chieu["lech"]))]))
        bang_them.append(("ĐỎ — Misa đã hết nợ nhưng hoá đơn vẫn ghi chưa thanh toán: đối trừ trên Misa",
                          ("Khách hàng", "Số hoá đơn", "Tiền ghi chưa TT", ""),
                          [(t, n, v, "") for t, n, v in kq.het_no_con_hd]))
        bang_them.append(("TÍM — tên khách trên file Bán hàng không tìm thấy trên Misa: kiểm tra tên",
                          ("Khách hàng", "Số hoá đơn", "Tiền ghi chưa TT", ""),
                          [(t, n, v, "") for t, n, v in kq.ten_khong_ghep]))

    def them_sheet(wb):
        ghifull.ghi_summary(wb, kq, ngay_chot, lich_su)
        ghifull.ghi_risk(wb, kq, ngay_chot)
        ghifull.ghi_contracts(wb, canh_bao_hd, ngay_chot, kq.them_khach)
        ghifull.ghi_payable(wb, kq, ngay_chot, ky)
        ghifull.ghi_cash(wb, kq, ngay_chot, ky)
        ghifull.ghi_invoices(wb, kq, ngay_chot)
        ghifull.ghi_methodology(wb, ngay_chot, nguon)
        ghifull.sap_xep(wb)

    baocao.ghi_bao_cao(
        kq.dong, ngay_chot, mau_chuan, dich, tong_misa=th.tong_du_no,
        doi_chieu={
            "che_do": "hoa-don", "canh_bao": canh_bao, "da_bo": kq.da_bo,
            "sms_cat": kq.sms_cat, "sms_khong_misa": kq.sms_khong_misa,
            "sms_chua_ghep": kq.sms_chua_ghep, "bang_them": bang_them,
            "nguon": [("Nguồn dữ liệu", nguon[0])] + [("", n) for n in nguon[1:]] +
                     [("", "Số Misa thay đổi theo thời điểm xuất. Chỉ so với bản làm tay khi "
                           "cả hai dùng cùng một lần xuất.")],
        },
        them_sheet=them_sheet,
    )

    so.file_tool[dich.name] = dich.stat().st_mtime
    trang_cu = so.trang.get(ngay_chot.isoformat()) or {}
    if any(v.nguon == "ban-tay" for v in trang_cu.values()):
        ghi_log(f"  Sổ ghi ngày {ngay_chot:%d/%m/%Y} đã lấy từ báo cáo trong thư mục gốc, giữ nguyên")
    else:
        so.ghi_trang(ngay_chot, {d.ten_chuan: tinhtoan.thanh_ho_so(d) for d in kq.dong})
    if pt is not None:
        so.phai_tra[ngay_chot.isoformat()] = {v.ma: round(v.du_co, 2) for v in pt.dong if v.du_co > 0.5}
    so.ghi()
    for c in canh_bao:
        ghi_log(f"  ! {c}")
    cs = kq.chi_so
    ghi_log(f"  -> {dich.name}: {len(kq.dong)} khách, phải thu {cs['tong_phai_thu']:,.0f}, "
            f"quá hạn {cs['qua_han']:,.0f}, nợ xấu {cs['no_xau']:,.0f}"
            + (f", phải trả {cs['tong_phai_tra']:,.0f}" if "tong_phai_tra" in cs else ""))
    return dich


TEN_TAC_VU = "Bao cao cong no tuan"


def sua_lich_chay_an() -> None:
    """Máy cài trước 09/10: Task Scheduler gọi chay_nen.bat nên mỗi lần chạy bật cửa sổ CMD.
    Tự đổi sang gọi thẳng pythonw.exe một lần, người dùng không phải cài lại."""
    if os.name != "nt":
        return
    import subprocess
    try:
        hien = subprocess.run(["schtasks", "/query", "/tn", TEN_TAC_VU, "/xml"], capture_output=True,
                              check=False, creationflags=0x08000000)
        raw = hien.stdout or b""
        xml = raw.decode("utf-16", "ignore") if b"\x00" in raw[:200] else raw.decode("utf-8", "ignore")
        if hien.returncode != 0 or "chay_nen.bat" not in xml.lower():
            return
        pythonw = Path(sys.executable).with_name("pythonw.exe")
        trinh_chay = pythonw if pythonw.exists() else Path(sys.executable)
        doi = subprocess.run(["schtasks", "/change", "/tn", TEN_TAC_VU, "/tr",
                              f'"{trinh_chay}" "{Path(__file__).resolve()}"'],
                             capture_output=True, text=True, check=False, creationflags=0x08000000)
        ghi_log("  Đã đổi lịch chạy nền sang chạy ẩn (không bật cửa sổ CMD)" if doi.returncode == 0 else
                f"  ! Không đổi được lịch chạy ẩn: {(doi.stderr or doi.stdout).strip()[:200]}")
    except Exception as loi:      # không bao giờ để việc phụ này làm hỏng lần chạy chính
        ghi_log(f"  ! Không đổi được lịch chạy ẩn: {loi}")


DASHBOARD = Path(__file__).resolve().parent / "dashboard" / "Credit_Dashboard.html"


def cap_nhat_dashboard(thu_muc: Path) -> None:
    """Đặt Credit_Dashboard.html ở thư mục gốc workspace, cạnh báo cáo tuần (WR-05).
    Bản trong mã nguồn mới hơn thì chép đè; không có gì đổi thì để nguyên."""
    if not DASHBOARD.exists():
        return
    dich = Path(thu_muc) / DASHBOARD.name
    try:
        if not dich.exists() or dich.read_bytes() != DASHBOARD.read_bytes():
            import shutil
            shutil.copyfile(DASHBOARD, dich)
            ghi_log(f"  Đã đặt {DASHBOARD.name} vào thư mục gốc — mở bằng trình duyệt, kéo báo cáo tuần vào")
    except OSError as loi:
        ghi_log(f"  ! Không chép được dashboard: {loi}")


def _phan_loai(files: List[Path]):
    """Nhận dạng từng file theo nội dung. Trả về dict loại -> [(ngày cuối kỳ, file)]."""
    ra = {"sms": [], "tong-hop": [], "chi-tiet": [], "phai-tra": [], "tuoi-no": [], "ban-hang": [],
          "so-chi-tiet": [], "la": []}
    for f in files:
        try:
            if ghep.la_file_sms(f):
                ra["sms"].append((None, f))
                continue
            loai = misa.nhan_dang(f)
            if loai == "tong-hop":
                ra[loai].append((misa.doc_tong_hop(f).den_ngay, f))
            elif loai == "chi-tiet":
                ra[loai].append((chitiet.ky_cua(f)[1], f))
            elif loai == "phai-tra":
                ra[loai].append((misa.doc_phai_tra(f).den_ngay, f))
            elif loai == "tuoi-no":
                ra[loai].append((misa.doc_tuoi_no(f).den_ngay, f))
            elif loai in ("ban-hang", "so-chi-tiet"):
                ra[loai].append((None, f))
            else:
                ra["la"].append((f, "không nhận ra là file Misa hay SMS nào"))
        except Exception as loi:
            ra["la"].append((f, str(loi)))
    return ra


def quet_thu_muc(cfg: caidat.CauHinh) -> int:
    cfg.tao_cac_thu_muc()
    ghep.dat_file_master(cfg.client_master)
    cap_nhat_dashboard(cfg.ra)
    files = [f for f in sorted(cfg.vao.glob("*.xlsx"))
             if not f.name.startswith(("[DONE]", "[LOI]", "~$"))]
    if not files:
        return 0

    so = sg.SoGhi(cfg.so_ghi)
    if nap_ban_tay(cfg.ban_tay, so):
        so.ghi()

    pl = _phan_loai(files)
    for f, vi_sao in pl["la"]:
        ghi_log(f"  Bỏ qua {f.name}: {vi_sao}")
    ds_sms = [f for _, f in pl["sms"]]
    sms = max(ds_sms, key=lambda f: f.stat().st_mtime) if ds_sms else None
    moi = lambda loai: (max((f for _, f in pl[loai]), key=lambda f: f.stat().st_mtime)
                        if pl[loai] else None)
    ban_hang, so_ct = moi("ban-hang"), moi("so-chi-tiet")

    # v0.4: Tổng hợp 131 + Chi tiết 131 cùng ngày cuối kỳ
    viec = []   # (ngày, loại, các file)
    ct_theo_ngay = {n: f for n, f in sorted(pl["chi-tiet"], key=lambda x: x[1].stat().st_mtime)}
    da_dung_ct = set()

    def chi_tiet_cho(n):
        """Chi tiết cùng ngày cuối kỳ; không có thì file Chi tiết xuất tới SAU ngày đó gần nhất
        (FIN W41: Chi tiết tới 31/10, Tổng hợp tới 09/10 — tool tự cắt)."""
        if n in ct_theo_ngay:
            return n
        sau = [m for m in ct_theo_ngay if m > n]
        return min(sau) if sau else None
    tra_theo_ngay = {n: f for n, f in sorted(pl["phai-tra"], key=lambda x: x[1].stat().st_mtime)}
    th_cu = []
    for n, f in sorted(pl["tong-hop"]):
        m = chi_tiet_cho(n)
        if m is not None:
            da_dung_ct.add(m)
            viec.append((n, "full", (f, ct_theo_ngay[m], tra_theo_ngay.pop(n, None))))
        else:
            th_cu.append(f)
    for n, f in ct_theo_ngay.items():
        if n in da_dung_ct:
            continue
        ghi_log(f"  Bỏ qua {f.name}: chưa có file Tổng hợp phải thu đến ngày {n:%d/%m/%Y} đi kèm")
    for n, f in tra_theo_ngay.items():
        ghi_log(f"  Bỏ qua {f.name}: chưa có file Tổng hợp + Chi tiết phải thu đến ngày {n:%d/%m/%Y} đi kèm")
    # v0.3: Tổng hợp + Tuổi nợ (giữ để không gãy khi FIN còn thả file cũ)
    if th_cu or pl["tuoi-no"]:
        cap, la = _ghep_cap(th_cu + [f for _, f in pl["tuoi-no"]])
        for f, vi_sao in la:
            if not (f in th_cu and not pl["tuoi-no"]):
                ghi_log(f"  Bỏ qua {f.name}: {vi_sao}")
        viec += [(n, "v03", (a, b)) for n, a, b in cap]
        for f in th_cu:
            if not any(f in v[2] for v in viec):
                ghi_log(f"  Bỏ qua {f.name}: chưa có file Chi tiết công nợ phải thu cùng kỳ đi kèm")
    if not viec:
        ghi_log(f"Chưa đủ file để chạy. Cần Tổng hợp 131 + Chi tiết 131 cùng kỳ. Đang có: "
                f"{[f.name for f in files]}")
        return 0
    viec.sort(key=lambda v: v[0])

    cuoi = viec[-1]
    thieu = []
    if sms is None:
        thieu.append("SMS (AR-AP)")
    if cuoi[1] == "full" and cuoi[2][2] is None:
        thieu.append("Tổng hợp công nợ phải trả")
    if cuoi[1] == "full" and ban_hang is None:
        thieu.append("Bán hàng")
    if cuoi[1] == "full" and so_ct is None:
        thieu.append("Sổ chi tiết tài khoản 131")
    if thieu:
        moi_nhat = max(f.stat().st_mtime for f in cuoi[2] if f)
        if time.time() - moi_nhat < CHO_SMS_GIAY:
            ghi_log(f"Đã có file Misa phải thu, đang chờ: {', '.join(thieu)}. Không có thì sau "
                    f"{CHO_SMS_GIAY // 3600} giờ tool tự chạy với những file đang có.")
            return 0
        ghi_log(f"  ! Chạy thiếu: {', '.join(thieu)}")
    if len(viec) > 1:
        ghi_log(f"Có {len(viec)} kỳ trong thư mục, chạy lần lượt từ kỳ cũ nhất"
                + (". File SMS chỉ dùng cho kỳ mới nhất" if sms else ""))

    xong = 0
    for i, (ngay, loai, ds) in enumerate(viec):
        sms_ky = sms if i == len(viec) - 1 else None
        bh_ky = ban_hang if i == len(viec) - 1 else None
        sct_ky = so_ct if i == len(viec) - 1 else None
        if loai != "full":
            bh_ky = sct_ky = None
        dung = [f for f in ds if f] + [f for f in (sms_ky, bh_ky, sct_ky) if f]
        ghi_log(f"Bắt đầu kỳ đến {ngay:%d/%m/%Y}: " + " + ".join(f.name for f in dung))
        try:
            if loai == "full":
                chay_full(ds[0], ds[1], cfg.ra, cfg.so_ghi, cfg.mau_chuan, duong_phai_tra=ds[2],
                          duong_sms=sms_ky, thu_muc_ghep=cfg.ghep, duong_ban_hang=bh_ky, duong_so_ct=sct_ky)
            else:
                chay_mot_lan(ds[0], ds[1], cfg.ra, cfg.so_ghi, cfg.mau_chuan,
                             duong_sms=sms_ky, thu_muc_ghep=cfg.ghep)
        except PermissionError as loi:
            ghi_log(f"  Tạm hoãn kỳ này: {loi}. Nhiều khả năng file kết quả đang "
                    "mở trong Excel. Đóng file đi, lần chạy sau tool làm tiếp.")
            continue
        except Exception as loi:  # không bao giờ để tiến trình chết im lặng
            ghi_log(f"  LỖI: {loi}")
            for f in ds:
                if f:
                    _doi_ten(f, "[LOI]")
            (cfg.vao / f"[LOI] {dt.date.today():%Y%m%d} - doc-vi-sao-hong.txt").write_text(
                f"{dt.datetime.now():%d/%m/%Y %H:%M}\n\n{loi}\n\n"
                "Sửa xong thì bỏ chữ [LOI] khỏi tên file, tool sẽ chạy lại ở lần kế tiếp.\n\n"
                + traceback.format_exc(),
                encoding="utf-8",
            )
            continue
        con_can = {g for _, _, ds2 in viec[i + 1:] for g in ds2 if g}   # một file Chi tiết dùng cho nhiều kỳ
        for f in dung:
            if f not in con_can:
                _doi_ten(f, "[DONE]")
        xong += 1
    return xong


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Sinh báo cáo công nợ tuần (full credit report) từ file Misa + SMS")
    p.add_argument("--tong-hop"); p.add_argument("--tuoi-no"); p.add_argument("--ra")
    p.add_argument("--so-ghi"); p.add_argument("--mau-chuan")
    p.add_argument("--nap-bao-cao"); p.add_argument("--ngay-chot")
    p.add_argument("--ngoai-le"); p.add_argument("--nhom"); p.add_argument("--vi-sao", default="")
    p.add_argument("--credit-term"); p.add_argument("--gia-tri")
    p.add_argument("--bo-qua"); p.add_argument("--nhan-lai")
    p.add_argument("--dat-thu-muc", help="thư mục đích trên máy này")
    p.add_argument("--khong-tao-thu-muc-con", action="store_true",
                   help="dùng thẳng thư mục đã trỏ, không tạo Cong-No-Workspace bên trong")
    p.add_argument("--xem-cau-hinh", action="store_true")
    p.add_argument("--client-master", help='file credit term chung trên OneDrive của máy này, nếu không để trong sample/ ("" = bỏ)')
    p.add_argument("--chi-tiet", help="file Chi tiết công nợ phải thu từ 2022 (chạy tay, v0.4)")
    p.add_argument("--phai-tra", help="file Tổng hợp công nợ phải trả (chạy tay, v0.4)")
    p.add_argument("--sms", help="file SMS AR-AP (chạy tay)")
    p.add_argument("--ban-hang", help="file Misa Bán hàng (chạy tay, v0.5)")
    p.add_argument("--so-chi-tiet", help="file Misa Sổ chi tiết tài khoản 131 (chạy tay, v0.5)")
    p.add_argument("--ghep", help="file ghép khách SMS ↔ Misa (chạy tay)")
    p.add_argument("--kiem-tra-ghep-khach", action="store_true",
                   help="đối chiếu file ghép khách với Misa/SMS mới nhất, mở file kiểm tra")
    p.add_argument("--khong-mo-file", action="store_true", help=argparse.SUPPRESS)
    a = p.parse_args(argv)

    if a.dat_thu_muc:
        cfg = caidat.dat_workspace(a.dat_thu_muc, not a.khong_tao_thu_muc_con)
        dat_nhat_ky(cfg.log)
        ghi_log(f"Thư mục đích: {cfg.workspace}")
        for d in cfg.cac_thu_muc():
            _in(f"    {d}")
        return 0

    if a.xem_cau_hinh:
        try:
            cfg = caidat.doc()
        except caidat.ChuaDatThuMuc as loi:
            _in(str(loi))
            return 2
        _in(f"Máy: {caidat.ten_may()}")
        _in(f"Thư mục mã nguồn: {caidat.GOC}")
        _in(f"Thư mục đích: {cfg.workspace}")
        for d in cfg.cac_thu_muc():
            _in(f"    {d}")
        _in(f"Mẫu chuẩn: {cfg.mau_chuan}")
        cm = cfg.client_master
        _in(f"File credit term: {cm if cm else '(chưa đặt — đọc bảng term trong sample/)'}"
            + ("" if not cm or cm.is_file() else "  ← KHÔNG THẤY FILE"))
        return 0

    if a.client_master is not None:
        p = caidat.dat_client_master(a.client_master)
        if p is None:
            _in("Đã bỏ đường dẫn file credit term chung: tool đọc bảng term trong sample/.")
            return 0
        _in(f"File credit term của máy này: {p}")
        if not p.is_file():
            _in("  Chưa thấy file ở đường dẫn này. Kiểm tra OneDrive đã đồng bộ chưa, rồi chạy lại.")
            return 2
        ghep.dat_file_master(p)
        _in(f"  Đọc được {len(ghep.doc_credit_term(None))} khách có credit term.")
        return 0

    # Luôn dùng cấu hình của máy nếu đã đặt thư mục đích, kể cả khi chạy tay,
    # để sổ ghi và nhật ký không bao giờ rơi vào thư mục mã nguồn.
    cfg = None
    try:
        if caidat.TEP_CAU_HINH.exists():
            cfg = caidat.doc()
    except caidat.ChuaDatThuMuc:
        cfg = None
    if cfg:
        dat_nhat_ky(cfg.log)
        ghep.dat_file_master(cfg.client_master)
    duong_so_ghi = Path(a.so_ghi) if a.so_ghi else (cfg.so_ghi if cfg else Path("so-ghi.json"))
    mau_chuan = Path(a.mau_chuan) if a.mau_chuan else (
        cfg.mau_chuan if cfg else caidat.GOC / "template" / "Bang_cong_no_MAU_CHUAN.xlsx")

    if a.kiem_tra_ghep_khach:
        if cfg is None:
            ghi_log("Chưa đặt thư mục đích. Chạy 1_Khoi_tao_workspace.bat trước.")
            return 2
        from wr import kiemtra
        cfg.tao_cac_thu_muc()
        so = sg.SoGhi(cfg.so_ghi)
        if nap_ban_tay(cfg.ban_tay, so):
            so.ghi()
        try:
            f = kiemtra.kiem_tra(cfg, so, ghi_log)
        except PermissionError as loi:
            ghi_log(f"Không ghi được file kiểm tra: {loi}. Đóng file KIEM_TRA đang mở rồi bấm lại.")
            return 1
        if f is None:
            return 2
        if os.name == "nt" and not a.khong_mo_file:
            try:
                os.startfile(str(f))       # mở bằng Excel cho FIN xem ngay
            except OSError:
                pass
        return 0

    if a.nap_bao_cao:
        if not a.ngay_chot:
            p.error("--nap-bao-cao cần kèm --ngay-chot YYYY-MM-DD")
        so = sg.SoGhi(duong_so_ghi)
        n = baocao.nap_vao_so_ghi(so, Path(a.nap_bao_cao), dt.date.fromisoformat(a.ngay_chot))
        so.ghi()
        ghi_log(f"Đã nạp {n} khách từ {Path(a.nap_bao_cao).name} vào sổ ghi, chốt {a.ngay_chot}")
        return 0

    if a.ngoai_le:
        if a.nhom not in sg.TEN_NHOM:
            p.error(f"--nhom phải là một trong {sg.TEN_NHOM}")
        so = sg.SoGhi(duong_so_ghi)
        so.ngoai_le[chuan_hoa_ten(a.ngoai_le)] = {"nhom": a.nhom, "ly_do": a.vi_sao}
        so.ghi()
        ghi_log(f"Đã khai ngoại lệ cho {a.ngoai_le}: xếp vào nhóm {a.nhom}")
        return 0

    if a.credit_term:
        if not a.gia_tri:
            p.error("--credit-term cần kèm --gia-tri, ví dụ --gia-tri \"15 days\"")
        so = sg.SoGhi(duong_so_ghi)
        so.dieu_chinh.setdefault(chuan_hoa_ten(a.credit_term), {})["credit_term"] = a.gia_tri
        so.ghi()
        ghi_log(f"Đã khai credit term cho {a.credit_term}: {a.gia_tri}")
        return 0

    if cfg is None and not any([a.tong_hop, a.nap_bao_cao, a.ngoai_le, a.credit_term,
                                a.bo_qua, a.nhan_lai]):
        ghi_log("Chưa đặt thư mục đích. Chạy 1_Khoi_tao_workspace.bat, hoặc: "
                'python runner.py --dat-thu-muc "<đường dẫn>"')
        return 2

    if a.bo_qua or a.nhan_lai:
        so = sg.SoGhi(duong_so_ghi)
        if a.bo_qua:
            so.bo_qua[chuan_hoa_ten(a.bo_qua)] = {"ly_do": a.vi_sao or "kế toán khai loại khỏi bảng"}
            ghi_log(f"Đã loại {a.bo_qua} khỏi báo cáo: {a.vi_sao or '(không ghi lý do)'}")
        else:
            so.bo_qua.pop(chuan_hoa_ten(a.nhan_lai), None)
            ghi_log(f"Đã nhận lại {a.nhan_lai} vào báo cáo")
        so.ghi()
        return 0

    if a.tong_hop and a.chi_tiet:
        if not a.ra:
            p.error("chạy tay cần --ra <thư mục ra>")
        try:
            chay_full(Path(a.tong_hop), Path(a.chi_tiet), Path(a.ra), duong_so_ghi, mau_chuan,
                      duong_phai_tra=Path(a.phai_tra) if a.phai_tra else None,
                      duong_sms=Path(a.sms) if a.sms else None,
                      duong_ghep=Path(a.ghep) if a.ghep else None,
                      thu_muc_ghep=cfg.ghep if cfg else None,
                      duong_ban_hang=Path(a.ban_hang) if a.ban_hang else None,
                      duong_so_ct=Path(a.so_chi_tiet) if a.so_chi_tiet else None)
        except LoiDocFile as loi:
            ghi_log(f"LỖI: {loi}")
            return 2
        return 0

    if a.tong_hop or a.tuoi_no:
        if not (a.tong_hop and a.tuoi_no and a.ra):
            p.error("chạy tay cần đủ --tong-hop, --chi-tiet (hoặc --tuoi-no) và --ra")
        try:
            chay_mot_lan(Path(a.tong_hop), Path(a.tuoi_no), Path(a.ra), duong_so_ghi, mau_chuan,
                         duong_sms=Path(a.sms) if a.sms else None,
                         duong_ghep=Path(a.ghep) if a.ghep else None,
                         thu_muc_ghep=cfg.ghep if cfg else None)
        except LoiDocFile as loi:
            ghi_log(f"LỖI: {loi}")
            return 2
        return 0

    sua_lich_chay_an()
    try:
        xong = quet_thu_muc(cfg)
    except Exception as loi:      # chạy nền: không bao giờ chết im lặng
        ghi_log(f"LỖI khi quét thư mục: {loi}\n{traceback.format_exc()}")
        return 1
    if not xong:
        _in("Không có báo cáo mới. Kiểm tra thư mục input đã đủ 4 file chưa (xem _tool/runner.log).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
