Mã: WR-09
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: WR-08
Người làm: claude
Người soát: codex
Trạng thái: đóng-đạt
Vòng: 1/2

## ĐỀ BÀI

Kỳ vọng của user: (user, 2026-10-20, theo ghi chú FIN sau W42)
1. STA CENTER: dòng "HĐ 1560 ngày 04/09 — 29,4tr" không có trên sổ; 29,4tr là của 2 hoá đơn khác.
2. Tên công ty: TRUSTANA VIETNAM COMPANY LIMITED.
3. Client Master thêm cột "Đã có hợp đồng": có thì tích, chưa có để trống (Sales đề nghị).
4. Chạy W42 từ file FIN gửi, so với bản "(human)" FIN sửa tay, liệt kê thay đổi.

Mục tiêu: Sửa các chỗ FIN chỉ ra sau W42 (phân bổ tiền thu, tỷ giá, tên công ty, cột Đã có hợp đồng) và chạy lại W42 so bản human.

Ngoài phạm vi: (phiếu mở trước v2.0, không ghi)

Kế thừa: (phiếu mở trước v2.0, không ghi; xem phiếu ở trường Nối tiếp)

Kịch bản test: (phiếu mở trước v2.0, không ghi riêng; xem Kỳ vọng của user)

## TODO
- [x] Tiền thu không ghi số HĐ trừ trước vào HĐ Misa ghi "Đã thanh toán" → STA: HĐ 1529 27.957.741 + HĐ 1322 1.463.336 (đúng ý FIN)
- [x] Chênh lệch tỷ giá NVK cùng ngày với phiếu thu ghi số HĐ → gộp vào HĐ đó (OPTIMALOG HĐ 579 = 2.115.321)
- [x] Khó đòi có trong bảng job nợ cũ vẫn hiện dòng HĐ (Minh Khôi HĐ 2519)
- [x] CONG_TY = TRUSTANA VIETNAM COMPANY LIMITED
- [x] Client Master: cột Đã có hợp đồng (✓, thả xuống), ngày gia hạn chỉ tính khi đã tới; Summary đếm khách chưa có HĐ
- [x] tools/nang_cap_client_master.py: nâng cấp file FIN đang dùng (82 khách, 13 ✓ theo Số hợp đồng)
- [x] tests/test_wr9.py

## AUDIT

Báo cáo của người làm:
```
python -m pytest            → 72 passed, 20 skipped
W42 (workspace test, mốc = W41 human) → 68 khách; phải thu 15.651.640.892; quá hạn 5.648.539.192;
                               nợ xấu 1.214.313.194; phải trả 7.574.718.156
So W42 (human)              → 68/68 khách, tổng khớp hết, term và Sales khớp hết; dòng job khớp 64/68
LibreOffice tính lại        → 0 lỗi công thức
```
4 khách còn khác, tool theo sổ + trạng thái Misa (chờ FIN xác nhận): OPTIMALOG (579 vs 1019), ACT 247
(13.898 của HĐ 1790), RUNWAY và AIRSPACE (tỷ giá gộp vào HĐ thay vì dòng riêng).

Chờ kiểm trên máy thật: [RUNTIME UNVERIFIED]


Người làm tự soi lại commit 5aff24e (2026-10-27), chưa sửa mã — phiếu này còn lỗi:
1. `khachhang.py:344` — `co_hd = bool(ô)`: gõ "Chưa", "Không", "x" cũng thành đã có hợp đồng (danh sách thả xuống
   không chặn vì `showErrorMessage=False`). Summary đếm thiếu khách chưa có hợp đồng.
2. `runner.py:362` — Client Master chưa nâng cấp (không có cột) → `co_hd=None` → mọi khách đang nợ bị đếm
   "chưa có HĐ", Review không cảnh báo (trái luật R.3a: không để tool im lặng).
3. `ghifull.py` Methodology (Ageing buckets, Rows under a customer) vẫn ghi "oldest first", chưa nói bước ưu tiên
   hoá đơn Misa "Đã thanh toán" và gộp chênh lệch tỷ giá (trái luật R.2: đổi luật tính phải sửa Methodology).
4. `tools/nang_cap_client_master.py` bỏ mất cột người dùng tự chèn (Hướng dẫn cho phép chèn cột) và tự tích ✓
   theo Số hợp đồng.
5. Luật "ngày gia hạn chỉ tính khi đã tới" không có trong Kỳ vọng của user — đã đưa lên "Đang treo, cần user".
6. Thiếu phép thử cho: nhánh khó đòi (Minh Khôi), cách dựng `da_tt` trong `tindung.py`, script nâng cấp,
   giá trị chữ ở cột ✓.

Đã sửa ở WR-11 (2026-10-27): mục 1, 2, 3, 4, 6 ở trên; mục 5 vẫn chờ user (STATE).

Giả định kỹ thuật: (phiếu mở trước v2.0, chưa ghi theo khuôn ba dòng)

Soát: user — đóng theo lệnh user 2026-11-05 ("các ticket đang mở cái nào xong thì close & move vào done"). Bằng chứng: Báo cáo W43 (chạy lại 29/10) và W44 (30/10) do Task Scheduler trên máy FIN sinh ra, dùng thật trong tuần (handoff/docs/Cong-No-Workspace (W44): runner.log, 2026_W44_Credit_Report_24.10-30.10.2026.xlsx — 55 khách, phải thu 13.099.776.819; Sales đã ghi lý do tuần này vào file). Không qua codex soát (lệch mục 0/9 của chuẩn, user quyết).
Kết quả: ĐẠT
Còn treo cho user (chuyển STATE): luật 'ngày gia hạn chỉ tính khi đã tới'.
-- claude - 2026-10-27 11:40 UTC+7

## HISTORY
### 2026-10-20
Failure: tiền thu không ghi số HĐ trừ theo ngày HĐ cũ nhất, trong khi Misa đã đối trừ theo cách khác
(HĐ 1560 hạch toán 22/07 nhưng ngày HĐ 04/09).
Lessons: trạng thái "Đã thanh toán" trên Misa là dấu vết đối trừ của kế toán — dùng nó để phân bổ trước FIFO.
### 2026-11-05
Failure: không
Báo lỗi giả: không
Lessons: đóng phiếu theo lệnh user sau khi tool chạy thật W43–W44 trên máy FIN.
-- claude - 2026-11-05 10:05 UTC+7
