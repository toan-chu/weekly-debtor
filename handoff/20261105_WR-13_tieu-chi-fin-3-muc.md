Mã: WR-13
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: WR-12
Người làm: claude
Người soát: user
Trạng thái: chờ-soát
Vòng: 1/2

## ĐỀ BÀI

Kỳ vọng của user:
(2026-11-05) "Đọc w44 & Đánh gái khách hàng phải thu và tiêu chí (2) 1 — Tuần 44 vừa rồi đã chạy, + tiêu chí đánh giá
nợ do FIN ban hành + màu sắc." · "bổ sung thêm tính năng ngày đến hạn của từng job" · "bữa họp thì cũng đồng ý là màu
của CEO quá mức khó đọc" · "Thôi khoan cái đó chưa làm, lưu thành phiếu. Vì tôi đang dựng bảng client master + contract
upload trên đó. Từ từ đi" · "bỏ qua cái vụ hợp đồng pdf, hạn mức nợ và lãi gộp (3 cái này thành phiếu, chờ list
master, còn lại ta xử lý trước)"

Mục tiêu: Thay luật màu 5 mức (ý CEO, WR-11) bằng bộ tiêu chí FIN ban hành: 3 mức theo ngày quá hạn, 3 tiêu chí nâng
hạng, nhóm khách × mức độ → phương án xử lý; màu kiểu Excel Good/Neutral/Bad chỉ tô ô; hiện ngày đến hạn của từng job.

Ngoài phạm vi (chuyển WR-14, chờ Client Master SharePoint List):
- Hợp đồng PDF · hạn mức nợ (TC2, cả hạn mức tạm tính theo billing) · lãi gộp (X1 của TC1) · "chưa có hợp đồng" của TC3 ·
  ô HĐ còn hiệu lực / hồ sơ đầy đủ · ngày gia hạn.
- Làm phần còn lại khi FIN trả lời 12 câu bên dưới.

Kế thừa:
- WR-11: màu 5 mức + T1–T4 đang chạy (sẽ bị thay). WR-12: file credit term tự thêm khách/cột, link hợp đồng.
- Nguồn tiêu chí: `handoff/docs/Đánh gái khách hàng phải thu và tiêu chí (2) 1.xlsx` (không lên git).

Kịch bản test (chạy lại W44 trên máy FIN — user bấm, khoảng 15 phút):
1. Đóng mọi file Excel trong workspace. GitHub Desktop → repo weekly-report → Fetch origin → Pull origin.
   Không chạy lệnh gì thêm; không xoá workspace.
2. Vào `input\`, bỏ chữ `[DONE] ` ở 6 file `2026_W44_…`. Task Scheduler → Bao cao cong no tuan → Run. Chờ 1 phút.
3. Mở báo cáo mới `2026_W44_Credit_Report_24.10-30.10.2026 (chay lai …).xlsx` bằng Excel. Thấy:
   a. Summary: 55 khách, phải thu 13.099.776.819 (y bản W44 cũ).
   b. Receivables: cột D "Risk level" cạnh tên khách, tô màu nhạt — Bình thường 36 · Cảnh báo 5 · Rủi ro cao 14
      (ví dụ KH00418 quá hạn 296 ngày = Rủi ro cao). Không còn tô cả hàng màu đậm.
   c. Receivables: sau cột lý do có 3 cột Sales (Adverse info? · Customer responds? · Broken promises…), mọi khách
      "Không · Có · Không", bấm vào ô có mũi tên chọn Có/Không. Lý do Sales đã ghi ở W44 cũ vẫn còn.
   d. Bấm nút "2" góc trái: dòng job dạng `└ JOB · HĐ 00000765 ngày 29/05/2026 · hạn 28/06/2026`.
   e. Thử: chọn "Có" ở ô Adverse info? của một khách đang Bình thường → ô Risk level của khách đó đổi sang Rủi ro cao
      ngay (không cần chạy tool). Đóng KHÔNG lưu.
   f. AR Risk, kéo sang phải: Max days overdue · Customer group · Risk level · Basis · Suggested payment terms ·
      Action plan (FIN) · Score · Grade (A–D tô màu).
4. Kéo báo cáo mới vào `Credit_Dashboard.html`: tab Nợ xấu có cột "Mức độ (FIN)"; trang một khách có nhãn mức độ.
Tuyệt đối không được xảy ra: tổng phải thu khác 13.099.776.819; ô #VALUE / #NAME / #N/A; lý do Sales của W44 mất;
file W44 cũ bị sửa; file Excel nào hiện trong danh sách thay đổi của GitHub Desktop.

### Tóm tắt tiêu chí FIN (đọc 05/11)
- 3 mức theo **số ngày quá hạn lớn nhất**: Bình thường (chưa quá hạn + 1–30) · Cảnh báo (31–60) · Rủi ro cao (>60).
- Nâng hạng (vi phạm 1 trong 3 → lên Rủi ro cao, tham số đổi được thành Cảnh báo):
  TC1 nợ quá hạn ≥ X, X = MIN(25% × lãi gộp BQ tháng quý gần nhất; nợ quá hạn của khách cuối nhóm top 15% khách quá hạn) ·
  TC2 dư nợ > credit limit (hợp đồng; chưa có thì billing BQ 6 tháng × số tháng kỳ hạn × 1,3; "cuối tháng" +0,5 tháng) ·
  TC3 chưa có hợp đồng hoặc khách Mới, dư nợ ≥ Y = X/2.
- Bỏ: số hoá đơn trả trễ (T3 cũ), ô "Không nâng hạng". Thêm ô định tính: thông tin bất lợi, HĐ còn hiệu lực, hồ sơ đủ,
  có phản hồi, thất hứa (đẩy lên Rủi ro cao / Cảnh báo).
- Nhóm KH: Mới (HĐ đầu trong 6 tháng) · Lâu năm (HĐ trước 01/01 năm nay) · Thường xuyên → bảng 'mức độ' ra Phương án xử lý +
  Hình thức thanh toán đề xuất. Điểm A–D chỉ tham khảo.
- Màu (CF của FIN): Bình thường nền C6EFCE chữ 006100 · Cảnh báo FFEB9C / 9C5700 · Rủi ro cao FFC7CE / 9C0006; ô TC vi phạm
  và "VƯỢT HẠN MỨC" FFC7CE / 9C0006. Chỉ tô ô, không tô cả hàng.
- Dữ liệu: billing 6 tháng, ngày HĐ đầu, ngày quá hạn lớn nhất tính được từ 6 file hiện có. Lãi gộp T7–T9 đã có trong file FIN.
- Ngày đến hạn từng job: tool đã có (hạn của hoá đơn = ngày HĐ + credit term), chỉ chưa hiện ở dòng job Receivables.
  Phương án khác: cột "Hạn thanh toán" của Misa → cần file thứ 7 "Chi tiết công nợ phải thu theo hoá đơn".

### FIN đã chốt (Head FIN + FIN, user chuyển 05/11)
1 Bỏ 5 màu, dùng 3 mức (A) · 2 tính cả hoá đơn quá hạn còn vài nghìn đồng (A) · 3 thêm ô "Mức độ" cạnh tên khách ở
Receivables (A) · 4 vi phạm tiêu chí → Rủi ro cao (A) · 5 TC1 CHỜ lãi gộp (B) → WR-14 · 6 top 15% làm tròn lên (A) ·
7 TC3 so tổng dư nợ (A); TC3 cũng CHỜ lãi gộp vì Y = X/2 (A) → WR-14 · 8 không cần cột sửa nhóm KH (B) · 9 GIỮ CẢ HAI:
"việc cần làm" N0–N4 và "Phương án xử lý" + "Hình thức thanh toán đề xuất" · 10 đưa điểm A–D vào báo cáo (A), chấm trên
hoá đơn từ đầu năm + hoá đơn năm trước còn nợ (A) · 11 ba ô đánh giá tay (thông tin bất lợi, khách có phản hồi, thất
hứa/trễ nhiều lần) nhập trong BÁO CÁO mỗi tuần, Sales cập nhật; tuần mới tool giữ giá trị tuần trước (A) · 12 hạn job =
ngày hoá đơn + credit term (A).

### Phạm vi làm ở WR-13 (sau khi user xác nhận)
- 3 mức theo ngày quá hạn lớn nhất; nâng mức chỉ theo 3 ô Sales (bất lợi / thất hứa → Rủi ro cao khi khách đang quá hạn
  (bất lợi: luôn); không phản hồi → Cảnh báo khi đang quá hạn) — TC1/TC2/TC3 chưa chạy (WR-14).
- Màu Good/Neutral/Bad chỉ tô ô; ô "Mức độ" cạnh tên khách (Receivables, chèn cột) + ở AR Risk.
- Nhóm KH tự động; Phương án xử lý + Hình thức thanh toán; giữ "Required action" N0–N4.
- Điểm A–D (4 thành phần theo 'hệ số hạn mức' của FIN).
- Ngày đến hạn từng job trên dòng job Receivables.
- Bỏ luật màu 5 mức + T1–T4 + cột "Không nâng hạng màu" khỏi báo cáo (cột trong file credit term giữ, không dùng).

### Phụ thuộc
- User đang dựng Client Master (SharePoint List) + upload hợp đồng: hạn mức, hợp đồng, ô định tính có thể lấy từ đó thay
  cho file credit term — chốt nguồn trước khi làm.

## TODO
- [x] FIN trả lời 12 câu + 3 câu bổ sung (05/11)
- [x] User xác nhận làm ("ok go, làm xong sẽ cài lại cho máy của bé FIN và rerun tuần 44")
- [x] Luật FIN (`wr/danhgia.py`), cột mới ở Receivables + AR Risk, công thức Excel, sổ ghi giữ 3 ô Sales
- [x] Hạn job trong nhãn dòng con; bỏ luật màu CEO, cột "Không nâng hạng màu", ô X/Y/N, đếm hoá đơn trả trễ
- [x] Dashboard đọc Receivables theo tiêu đề + nhãn mức độ FIN
- [x] Phép thử: 92 đạt / 20 bỏ qua; làm hỏng 4 lần đều đỏ; chạy lại W44 thật trong máy ảo
- [x] README (v0.8), RULES, MAP, STATE
- [ ] User commit + chạy kịch bản test trên máy FIN

## AUDIT
Thay đổi (05/11):
- `wr/danhgia.py` (mới): 3 mức, nâng mức theo 3 ô Sales, nhóm khách, hình thức thanh toán, phương án (nguyên văn sheet
  'mức độ' của FIN), điểm A–D.
- `wr/thuonghieu.py`: COT Receivables 19 cột (chèn D "Risk level"; 3 ô Sales cuối), hằng số `C_*`, `COT_DANH_GIA`,
  `MUC_FIN`, `MAU_MUC_FIN`. `wr/baocao.py`: ghi/đọc theo hằng số + theo tiêu đề; CF ô mức; DV Có/Không; nhãn job có hạn.
- `wr/tindung.py`: `danh_gia_fin`; bỏ `danh_gia_mau`/`muc_goc`/NGUONG_X…; 3 ô Sales: tuần này → tuần trước → mặc định.
- `wr/ghifull.py`: AR Risk thay khối màu bằng khối FIN (`cong_thuc_fin`), CF chỉ ô Risk level + Grade; Summary
  muc_1..3, no_rui_ro_cao, hang_d; Methodology. `wr/chitiet.py`: `thong_ke_cham_diem`, bỏ `hoa_don_tra_tre`.
- `wr/khachhang.py`: bỏ cột "Không nâng hạng màu" ở mẫu mới (file cũ có cột đó thì giữ như cột người dùng chèn).
- `runner.py`: chỉ số Summary mới; bỏ cảnh báo "màu chưa nâng theo hạn mức". `dashboard/src/app.js` + `app.css`.
- Tests: `tests/test_wr13.py` (9 mới); `test_wr11.py` bỏ 5 phép thử màu CEO; `test_wr4.py`, `test_wr6.py` đổi số cột
  gán cứng sang `th.C_*`.

Bằng chứng:
```
$ python3 -B -m pytest -p no:cacheprovider --no-header
92 passed, 20 skipped in 63.92s
$ node --check dashboard/src/app.js && python3 -B dashboard/build.py
OK · Đã dựng Credit_Dashboard.html (1.35 MB)
```
Làm hỏng thử (mỗi lần trả lại file gốc, sau đó `tests/test_wr13.py` 9 passed):
```
bỏ luật "thông tin bất lợi" trong danhgia.muc_cuoi        → 2 failed, 7 passed
công thức Excel ngưỡng 60 → 30 (Python giữ nguyên)        → 1 failed (-k excel, LibreOffice)
tuần mới không mang ô Sales tuần trước                    → 1 failed (-k giu)
bỏ "han" khỏi dòng job                                    → 1 failed (-k han)
```
Chạy lại W44 thật (bản sao `handoff/docs/Cong-No-Workspace (W44)` → ~/wk/ws44, bỏ [DONE], `runner.quet_thu_muc`):
```
Đã đọc lại báo cáo 2026_W44_…xlsx: 55 khách (bản người làm)
-> 2026_W44_Credit_Report_24.10-30.10.2026 (chay lai 05-11 05h06).xlsx: 55 khách, phải thu 13,099,776,819,
   quá hạn 4,585,168,278, nợ xấu 2,643,065,730, phải trả 2,709,076,069
LibreOffice tính lại: lỗi công thức 0 · Receivables Risk level == AR Risk Risk level: True (55)
phân bố Bình thường 36 · Cảnh báo 5 · Rủi ro cao 14 = Summary muc_1/2/3 36/5/14 · lệch luật ngày: []
nhóm: Lâu năm 38 · Mới 13 · Thường xuyên 4 · hạng A 11 · B 10 · C 7 · D 27 · Collection rate > 1: []
dòng job 181, có hạn 181 · VD "└ EXSANA26040057 · HĐ 00000765 ngày 29/05/2026 · hạn 28/06/2026"
dashboard (Chromium headless): errors 0 · bảng Nợ xấu có "Mức độ (FIN)"
```
Tự soi được trong lúc làm: tỷ lệ thu hồi chia cho giá trị hoá đơn trong năm nhưng phần đã thu tính cả hoá đơn năm
trước → có thể > 100%. Đã sửa mẫu số (gồm phần năm trước còn nợ ở đầu năm), có phép thử.

Giả định kỹ thuật:
[GIẢ ĐỊNH KỸ THUẬT] Khách không có hoá đơn theo sổ (khó đòi / ngoại lệ đã khai) lấy số ngày quá hạn = cận trên nhóm tuổi cũ nhất (0 · 30 · 60 · 90 · 120 · 121).
Vì sao chọn: không có ngày hoá đơn để tính; khó đòi luôn ở nhóm 120+ nên ra Rủi ro cao, đúng ý FIN.
Sai thì làm lại: `tindung.danh_gia_fin` (nhánh `else`) + `dashboard/src/app.js` `tierOf`.

[GIẢ ĐỊNH KỸ THUẬT] Điểm A–D: "đã thu" = phần hoá đơn giảm đi trong năm, gồm cả giảm trừ / cấn trừ / tỷ giá (sổ Chi tiết không tách được).
Vì sao chọn: cùng cách trừ tiền thu đang dùng cho tuổi nợ; điểm chỉ để tham khảo (FIN chốt).
Sai thì làm lại: `chitiet.thong_ke_cham_diem`.

[GIẢ ĐỊNH KỸ THUẬT] Khách chưa có hoá đơn nào theo sổ xếp nhóm "Mới".
Vì sao chọn: như cột 'Nhóm KH áp dụng' trong file FIN (ô trống → Mới).
Sai thì làm lại: `danhgia.nhom_khach`.

Chưa làm (WR-14, chờ Client Master): TC1 / TC2 / TC3, hạn mức, lãi gộp, hợp đồng còn hiệu lực, hồ sơ đầy đủ.

## HISTORY
### 2026-11-05
Mở phiếu, chưa làm theo lệnh user ("Từ từ đi").
Failure: không
Báo lỗi giả: không
Lessons: không
-- claude - 2026-11-05 10:20 UTC+7

### 2026-11-05 (vòng 1)
Làm theo lệnh user "ok go". Xong mã + phép thử + chạy lại W44 trong máy ảo; chờ user commit và chạy kịch bản trên máy FIN.
Failure: không
Báo lỗi giả: không
Lessons: phép thử gán cứng số cột (test_wr4, test_wr6) vỡ khi chèn cột — đã đổi sang hằng số `th.C_*`.
-- claude - 2026-11-05 12:10 UTC+7
