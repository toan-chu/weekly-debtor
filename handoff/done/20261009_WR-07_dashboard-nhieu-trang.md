Mã: WR-07
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: WR-05, WR-06
Người làm: claude
Người soát: codex
Trạng thái: đóng-đạt
Vòng: 1/2

## ĐỀ BÀI

Kỳ vọng của user: (user, 2026-10-09)
Mỗi tab trên thanh đầu trang là một màn riêng, không cuộn lên xuống. Bấm tab nào hiện tab đó.
Vào chi tiết (một Sales, một khách) thì mới hiện rõ phần của người đó, và các tab khác chỉ còn
phần liên quan (kiểu dashboard nhiều trang + drill-through như Looker Studio / Power BI).
Kèm theo: Task Scheduler chạy mà không bật cửa sổ CMD.

Mục tiêu: Dashboard nhiều trang có bấm xuyên theo khách và Sales; lịch chạy nền không bật cửa sổ CMD.

Ngoài phạm vi:
- Nợ vendor theo job/loại phí; màu phải trả tăng — chờ FIN.

Kế thừa: (phiếu mở trước v2.0, không ghi; xem phiếu ở trường Nối tiếp)

Kịch bản test:
user: test ngay trên máy user.

## TODO
- [x] 6 trang: Tổng quan · Khách hàng · Sales · Nợ xấu · Phải trả · Dòng tiền; viên tím đứng ở trang đang mở
- [x] Trang chi tiết khách: tuổi nợ, ai phụ trách (bấm → trang Sales), job + trạng thái Misa, hoá đơn, S1–S6, Đối chiếu
- [x] Trang riêng Sales: phần phụ trách, quá hạn, nợ xấu, khách (bấm → trang khách), job, khách nợ xấu
- [x] Bộ lọc nằm trên thanh địa chỉ: đi theo qua các tab, nút Back của trình duyệt quay lại được
- [x] Tổng quan thêm "Cần chú ý" và "Theo Sales" để bấm xuyên
- [x] In / Lưu PDF in đủ mọi trang
- [x] Lịch chạy nền gọi thẳng pythonw.exe; máy cũ tự đổi lịch một lần khi runner chạy

## AUDIT

Báo cáo của người làm:
```
$ python dashboard/build.py                 Credit_Dashboard.html (1.34 MB)
$ python -m pytest -o addopts="" -q         55 passed, 20 skipped
Chromium headless với báo cáo W40 (bản WR-06): 0 lỗi console.
  Sales → bấm Trần Văn Hiếu → #sales?s=…; sang tab Nợ xấu chỉ còn 3 khách của Hiếu; Back quay lại đúng.
  Bấm khách trong trang Sales → trang khách LCC UPP (9 job, Đối chiếu: Khớp). In PDF: 11 trang.
```
[RUNTIME UNVERIFIED] Đổi lịch chạy ẩn (schtasks /change) trên Windows của user.

Giả định kỹ thuật: (phiếu mở trước v2.0, chưa ghi theo khuôn ba dòng)

Soát: user — đóng theo lệnh user 2026-11-05 ("các ticket đang mở cái nào xong thì close & move vào done"). Bằng chứng: Báo cáo W43 (chạy lại 29/10) và W44 (30/10) do Task Scheduler trên máy FIN sinh ra, dùng thật trong tuần (handoff/docs/Cong-No-Workspace (W44): runner.log, 2026_W44_Credit_Report_24.10-30.10.2026.xlsx — 55 khách, phải thu 13.099.776.819; Sales đã ghi lý do tuần này vào file). Không qua codex soát (lệch mục 0/9 của chuẩn, user quyết).
Kết quả: ĐẠT
-- claude - 2026-10-27 11:40 UTC+7

## HISTORY

### 2026-10-09
Dựng theo yêu cầu. Lỗi thấy khi soi ảnh chụp: cột "Việc cần làm" bị canh phải (bảng HTML coi cột
cuối là số tiền) → canh theo tên cột. Thuộc tính hidden bị CSS display:flex đè → thêm luật chung.
Failure: không. Lessons: thanh địa chỉ giữ trạng thái trang là cách rẻ nhất để có nút Back.

### 2026-10-09 (cuối ngày)
FIN chốt Payable: "làm y chang mẫu của em… chỉ cần lấy số tổng, không phân tích, không cần tô".
Sheet Payable còn 5 cột (No · Code · Vendor's Name · Current · Total), chỉ vendor còn số dư; tiền đã
trả vendor giữ ở Cash Flow. Dashboard tab Phải trả: bỏ thẻ trạng thái, đọc tiền đã trả từ Cash Flow.
Review bỏ bảng "phải trả nhập bổ sung". Giả lập máy FIN bằng gói handoff/docs/sample_w40: 6 file W40
+ W39 final + file ghép → chạy 1 lần ra W40 (70 khách, 14.251.295.535; Đối chiếu 57/6/2/5),
nút 2: 68 mã đã ghép, 0 khách mới, 0 dòng sửa, Tham_chieu_job_no_cu có 1 hoá đơn (Minh Khôi 32tr).
### 2026-11-05
Failure: không
Báo lỗi giả: không
Lessons: đóng phiếu theo lệnh user sau khi tool chạy thật W43–W44 trên máy FIN.
-- claude - 2026-11-05 10:05 UTC+7
