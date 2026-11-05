# MAP — weekly-report

Kiến trúc hiện tại. Mọi thay đổi ở tầng này phải vẽ ở đây TRƯỚC khi đổi mã.

Cập nhật: 2026-10-27 (WR-12). Trạng thái: v0.7.1 — full credit report: phải thu tính tuổi
từ ngày hoá đơn thật, phải trả, dòng tiền, nhóm rủi ro, hợp đồng (Client Master), lịch sử tuần, cách tính.

## Hai nơi, đừng nhầm

```
THƯ MỤC MÃ NGUỒN (chỉ có mã)            THƯ MỤC ĐÍCH (người dùng tự trỏ)
weekly-report/                           Cong-No-Workspace/
  1_Khoi_tao_workspace.bat                 *_Credit_Report_*.xlsx   báo cáo tuần, ngay ở gốc;
  2_Kiem_tra_ghep_khach.bat                                         chị KTT sửa thẳng vào đây
  runner.py  wr/  tools/  tests/           input/   mỗi tuần 4 file (xem dưới)
  template/   mẫu báo cáo chuẩn            sample/  file khai 1 lần: ghép khách, Master List
  settings.json  chỉ ghi đường dẫn         _tool/   (ẩn) sổ ghi, nhật ký, lưu trữ
  handoff/    phòng bàn giao (chuẩn v2.0)
```
Không còn nút "Chạy ngay": tool tự chạy khi đủ file; ép chạy sớm thì bấm Run trong
Task Scheduler. Đường dẫn chỉ nằm trong settings.json của từng máy.

## Sáu file mỗi tuần (input/), xuất CÙNG NGÀY

| File | Cho biết gì | Bắt buộc? |
|---|---|---|
| Misa Tổng hợp công nợ phải thu (TK131) | số dư từng khách = TIỀN đúng, kỳ báo cáo | có |
| Misa Chi tiết công nợ phải thu (TK131), **từ 01/01/2022 đến nay** | từng hoá đơn có ngày, từng phiếu thu có TK đối ứng | có |
| Misa Tổng hợp công nợ phải trả (TK331) | nợ vendor đầu kỳ, phát sinh, cuối kỳ | chờ 2 giờ rồi chạy thiếu |
| SMS AR-AP | Salesman của từng job (job + tiền đã lấy từ Misa khi có 2 file dưới) | chờ 2 giờ rồi chạy thiếu |
| Misa Bán hàng, **từ đầu năm** (WR-06) | trạng thái thanh toán từng hoá đơn | chờ 2 giờ; thiếu thì job lấy từ SMS (cách cũ) |
| Misa Sổ chi tiết tài khoản 131, **từ đầu năm** (WR-06) | job của từng hoá đơn ("Mã đối tượng THCP"), phiếu thu sau kỳ | chờ 2 giờ rồi chạy thiếu |

Mã: `wr/banhang.py` đọc 2 file WR-06; `wr/tindung.py` dựng dòng con + `doi_chieu()`.

Không có file Chi tiết mà có file Tuổi nợ cũ → chạy luật v0.3 (giữ nguyên để không gãy).

## Luồng chạy mỗi tuần

```
 input/ 6 file ──► nhận dạng theo NỘI DUNG (không tin tên file), ghép theo kỳ
 sample/ ghép khách, Master List ──► nhận dạng theo nội dung
 gốc/ báo cáo các tuần trước ──► đọc lại vào sổ ghi (bản người thắng bản máy)
                 │
   ① TIỀN mỗi khách   = Tổng hợp 131 (dư Nợ cuối kỳ). Luôn luôn.
   ② TUỔI             = Chi tiết 131: hoá đơn theo ngày hoá đơn; phiếu thu ghi số HĐ thì
                         trừ đúng HĐ đó (NVK tỷ giá cùng ngày gộp vào HĐ đó); còn lại trừ trước
                         HĐ Misa ghi "Đã thanh toán", rồi HĐ CŨ NHẤT (WR-09); hạn = ngày HĐ + term
                         ("14 ngày của tháng tiếp theo" = ngày 14 tháng sau);
                         quá hạn từ HÔM SAU ngày đến hạn (FIN 08/10)
                         khó đòi (bảng ghép / lý do có "khó đòi") → cả số vào 120+
                         Chi tiết ≠ Tổng hợp → tiền theo Tổng hợp, ghi Review
   ③ JOB, SALESMAN    = SMS (qua bảng ghép) > tuần trước
   ③b LÝ DO (PA1)     = tuần này: trống, trừ khi sổ ghi đã có trang của chính ngày chốt (chạy lại);
                         tuần trước: trang trước trong sổ ghi → cột chỉ đọc
   ④ RỦI RO           = luật của chị KTT (S1..S7, N0..N4), so với tuần trước trong sổ ghi
   ④b MÀU (WR-11)     = Client Master đọc TRƯỚC khi tính (hạn mức, ✓ hợp đồng, không nâng hạng)
                         + Chi tiết 131 (chitiet.hoa_don_tra_tre) → tindung.danh_gia_mau;
                         AR Risk ghi cùng luật bằng công thức (ghifull.cong_thuc_mau), tô cả hàng
   ④0 MẪU (WR-12, 28/10) = đầu lần chạy: file credit term thiếu cột / sheet tên cũ → khachhang.them_khach_moi(f, [])
                         nâng mẫu (cùng điều kiện an toàn) rồi mới doc_master
   ④c KHÁCH MỚI (WR-12) = khách đang nợ chưa có trong file credit term → khachhang.them_khach_moi: đọc cả file
                         (doc_file), thêm dòng, sao lưu _tool/luu-tru, ghi lại theo mẫu mới (ghi_master) qua file tạm
                         rồi os.replace; file đang mở (~$ hoặc không mở được để ghi) → bỏ qua, báo Contracts + Review
   ⑤ PHẢI TRẢ         = Tổng hợp 331: mọi vendor còn dư Có + vendor trả xong trong tuần
   ⑥ DÒNG TIỀN        = phiếu thu trong kỳ của Chi tiết 131 tách theo TK đối ứng
                         (112 ngân hàng, 111 tiền mặt, 331 cấn trừ, 515/635 tỷ giá, khác)
                         + phát sinh Nợ 331 (đã trả vendor)
   ⑦ ghi <gốc>/<năm>_W<tuần>_Credit_Report_<từ>-<đến>.xlsx, file cùng tên đã có thì
      KHÔNG ghi đè, đặt tên "(chay lai dd-mm HHhMM)"; cập nhật sổ ghi + lịch sử tuần
```

## File báo cáo tuần (đầu vào của dashboard WR-05)

| Sheet | Một dòng là | Nguồn |
|---|---|---|
| Summary | 1 tuần (tuần này + mọi tuần trước) | tổng hợp các sheet dưới + sổ ghi |
| Receivables | 1 khách còn nợ: ô Risk level (công thức INDEX sang AR Risk, WR-13), tuổi nợ, Salesman, Job Number, lý do tuần trước / tuần này, 3 ô Sales đánh giá (WR-13); dòng con (gập) = job + hoá đơn + hạn + Sales + số tiền, dòng cân đối. Vị trí cột: hằng số `thuonghieu.C_*`, `COT_DANH_GIA` — đừng gán cứng số cột | ①②③ |
| Invoices | 1 hoá đơn còn nợ: ngày, hạn, số ngày quá hạn, nhóm; cộng theo khách = Total | ② |
| AR Risk | 1 khách: đầu kỳ, tăng, giảm, cuối kỳ, so tuần trước, S1..S7, nhóm rủi ro N0–N4, xu hướng; WR-13: ngày quá hạn lớn nhất, nhóm khách, mức FIN (công thức `ghifull.cong_thuc_fin`), căn cứ, hình thức thanh toán, phương án, điểm + hạng A–D | ④ |
| Payable | 1 vendor: đầu kỳ, nợ mới, đã trả, cuối kỳ, trạng thái | ⑤ |
| Cash Flow | 1 khách đã trả / 1 vendor đã được trả trong tuần | ⑥ |
| Methodology | 1 con số: tính thế nào, lấy từ file nào cột nào | cố định trong mã |
| Review | đối chiếu và những chỗ tool không chắc | mọi bước |

Excel ghi số, riêng AR Risk ghi công thức của kế toán trưởng và dòng TOTAL Receivables là SUMIFS
(WR-08). Dashboard không đọc được kết quả công thức (file chưa mở bằng Excel thì chưa có số) nên
tự tính lại S1–S6, Tier, Trend, Required action bằng JS. Cách tính viết bằng lời ở Methodology.
WR-08 thêm: cột "Jobs by Salesman" ở AR Risk (`tindung.chia_theo_sales`); credit term đọc từ file
mẫu trong `sample/` (`ghep.doc_credit_term`), file mẫu dựng bằng `wr/khachhang.py` (Client Master List);
`khachhang.doc_master` + `canh_bao_hop_dong` → sheet Contracts (WR-11: cuối sheet có bảng khách đang nợ
chưa có trong Client Master, cột C–F, để FIN dán) (`ghifull.ghi_contracts`) + 5 chỉ số Summary
(hd_het_han, hd_sap_het_30, khach_chua_term, term_doi, khach_chua_han_hd); dashboard đọc sheet Contracts. File Client Master ngoài `sample/`: `caidat.CauHinh.client_master` →
`ghep.dat_file_master` (runner/kiemtra đặt trước khi chạy); `_cac_bang_term` chỉ đọc file đó khi có.

WR-13 (05/11): luật FIN ở `wr/danhgia.py` (mức theo ngày, nâng mức theo 3 ô Sales, nhóm khách, phương án, điểm A–D);
`tindung.danh_gia_fin` gắn vào từng `DongRuiRo`; `chitiet.thong_ke_cham_diem` cho điểm; 3 ô Sales lưu trong sổ ghi
(`HoSoKhach.danh_gia`), tuần mới lấy tuần trước. Ba chỗ cùng một luật: `danhgia.py` · `ghifull.cong_thuc_fin` (Excel) ·
`dashboard/src/app.js` (`tierOf`) — sửa một chỗ thì sửa cả ba; `tests/test_wr13.py` so Excel (LibreOffice) với Python.

## Dashboard (WR-05)

`dashboard/src` (HTML, CSS, JS, thư viện đọc Excel, font) → `dashboard/build.py` → một file
`Credit_Dashboard.html` chạy offline. Mỗi lần quét, tool chép file này vào thư mục gốc workspace.
Người xem kéo file báo cáo tuần vào trang; dashboard chỉ đọc số, riêng nhóm rủi ro tính lại khi
lý do trễ hạn trong Receivables khác bản tool viết. Nút In / Lưu PDF dùng hộp in của trình duyệt.
WR-07: mỗi tab là một trang (Tổng quan · Khách hàng · Sales · Nợ xấu · Phải trả · Dòng tiền). Bấm
một khách / một Sales → trang chi tiết; bộ lọc ghi trên thanh địa chỉ (`#sales?s=…&c=…`) nên đi
theo qua các tab và nút Back chạy. In / Lưu PDF in đủ mọi trang.
13/10: `tenNgan` rút tên khách cho nhãn biểu đồ; `ngayHen` + S7 lỡ hẹn (giống `tindung.ngay_hen`), đọc
thêm cột Promise date của AR Risk khi lý do không có ngày.

## Sổ ghi (_tool/so-ghi.json)

WR-12: runner.nap_ban_tay bỏ qua chỉnh sửa trên báo cáo của tuần đã có trang tuần sau (tuần đã chốt).

Một trang cho mỗi ngày chốt: từng khách (số dư, 6 nhóm, lô nợ, term, salesman, job, lý do)
+ `lich_su_tuan` (các con số của sheet Summary từng tuần). Báo cáo ở thư mục gốc được đọc
lại thành trang của đúng ngày chốt ghi trong file. Cùng ngày chốt: bản người làm (không có tab
Review, tên có "human", hoặc file tool đã bị sửa sau khi ghi — `so.file_tool`) thắng bản tool;
cùng loại thì bản sửa sau cùng thắng.

## Hai nút cho người dùng

1. `1_Khoi_tao_workspace.bat` — cài thư viện, chọn thư mục đích, tạo ngăn, đăng ký lịch 15 phút.
2. `2_Kiem_tra_ghep_khach.bat` — đối chiếu file ghép khách với Misa/SMS mới nhất, mở file KIEM_TRA.

## Ngoài phạm vi hiện tại

Nợ vendor theo job/loại phí từ SMS (cần ghép vendor SMS ↔ Misa).
Power Automate gửi email.
