Mã: WR-05
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: WR-04
Người làm: claude
Người soát: codex
Trạng thái: đóng-đạt
Vòng: 1/2

## ĐỀ BÀI

Kỳ vọng của user: (lời user, 2026-10-08)
Excel đẹp và dashboard đẹp, vì chưa biết CEO thích cái nào. Excel: bảng Receivables có dòng
con gập/mở dưới mỗi khách — job nào, Sales nào phụ trách, còn nợ bao nhiêu ("Hoàng Kim nợ 300,
Hiếu job ABC 100, Hằng job CDE 200"). Dòng con KHÔNG tính tuổi, tuổi chỉ ở dòng khách (sếp cần
biết ai phụ trách để hỏi thẳng). Dashboard HTML đặt ở thư mục gốc workspace để chiếu trong
công ty, xuất được PDF cho sếp đọc.

Mục tiêu: Excel có dòng con job/Sales dưới mỗi khách và một dashboard HTML chạy offline, in được PDF cho sếp.

Ngoài phạm vi:
- Nợ vendor theo job/loại phí (chưa chốt nguồn: SMS + bảng ghép vendor, hay file Chi tiết 331).

Kế thừa: (từ WR-04)
- Báo cáo Excel 7 sheet; tiền theo Misa; tên Salesman đầy đủ như SMS (CEO 08/10).

Kịch bản test:
user 08/10: "Quất đi" — test ngay trên máy user.

## TODO
- [x] Receivables: dòng con job SMS (gập sẵn, nút 1|2), dòng cân đối "Nợ cũ / chưa có job"
- [x] Dòng TOTAL ghi số (không SUM, tránh cộng trùng dòng con); đọc lại bỏ qua dòng con
- [x] Sheet Invoices: mỗi hoá đơn còn nợ, cộng theo khách = Total
- [x] Summary thành trang bìa: 12 ô số (phải thu · phải trả · dòng tiền); in ngang vừa 1 trang
- [x] dashboard/: src + build.py (cùng cách incal) → Credit_Dashboard.html, tool chép vào gốc
- [x] Dashboard: ô số, tuổi nợ, xu hướng, bảng khách bấm mở job + hoá đơn, theo Sales,
      nợ xấu, phải trả, dòng tiền; nút In / Lưu PDF; tự tính lại nhóm rủi ro khi KTT sửa lý do

## AUDIT

Báo cáo của người làm:
```
$ python -m pytest tests/      48 passed, 20 skipped
$ dashboard/build.py           Credit_Dashboard.html (1.30 MB, chạy offline)
Chạy dashboard bằng Chromium headless với báo cáo W40 thật: 0 lỗi console, 70 khách,
bấm LLC "ILC" mở 8 job + dòng cân đối 424.617.216; xuất PDF A4 ngang.
Receivables W40: 70 dòng khách, TOTAL 14.251.295.535; Invoices 233 hoá đơn, cộng = 14.251.295.535.
```
Số W40: 44/70 khách có job SMS; 1 khách có 2 Sales (INTERFREIGHT) — nay hiện cả hai.
Kết luận cũ của người làm (chuẩn v1.1): chờ user. `[RUNTIME UNVERIFIED]` mở dashboard bằng Chrome/Edge trên Windows.

Giả định kỹ thuật: (phiếu mở trước v2.0, chưa ghi theo khuôn ba dòng)

Soát: user — đóng theo lệnh user 2026-11-05 ("các ticket đang mở cái nào xong thì close & move vào done"). Bằng chứng: Báo cáo W43 (chạy lại 29/10) và W44 (30/10) do Task Scheduler trên máy FIN sinh ra, dùng thật trong tuần (handoff/docs/Cong-No-Workspace (W44): runner.log, 2026_W44_Credit_Report_24.10-30.10.2026.xlsx — 55 khách, phải thu 13.099.776.819; Sales đã ghi lý do tuần này vào file). Không qua codex soát (lệch mục 0/9 của chuẩn, user quyết).
Kết quả: ĐẠT
-- claude - 2026-10-27 11:40 UTC+7

## HISTORY

### 2026-10-08
Failure    Bản đầu: bảng lịch sử ở Summary làm hẹp cột chứa ô số (hiện ###); thanh bar quá
           nhỏ vẽ path âm (lỗi console). Đã sửa, soi lại bằng ảnh chụp.
           Báo lỗi giả: không
Lessons    Dòng con trong Excel làm SUM cả cột cộng trùng → tổng phải ghi bằng số.
-- claude - 2026-10-08 UTC+7

### 2026-10-08 (tối) — vòng góp ý 1 của user
Failure    Mũi tên tăng/giảm tô ngược ý user (đang tô theo tốt/xấu). Dashboard dùng bảng quá nhiều.
           Báo lỗi giả: không
Lessons    Viết lại theo kiểu Looker Studio: mỗi thẻ có nút chọn Donut / Cột / Bảng; trỏ vào phần nào
           cũng hiện số và 5 khách nợ nhiều nhất; bấm vào một phần biểu đồ thì lọc cả trang (Sales,
           nhóm rủi ro, tuổi nợ), thanh lọc dính trên đầu có chip bỏ lọc. Thanh điều hướng: viên tím
           đậm trượt theo chuột, đứng ở mục đang xem. ▲ xanh, ▼ đỏ. Biểu đồ vẽ theo bề rộng thật
           của thẻ (vẽ lại khi đổi cỡ cửa sổ). Chạy Chromium headless: 0 lỗi console.
-- claude - 2026-10-08 UTC+7

### 2026-10-08 (tối, vòng 2)
Failure    không
           Báo lỗi giả: không
Lessons    user chốt LUẬT MÀU: đỏ và màu lân cận = tiêu cực; xanh lá = tích cực; xanh dương, tím
           = trung tính. Mũi tên tô theo ý nghĩa từng chỉ số (quá hạn/nợ xấu tăng = đỏ; thu được,
           dòng tiền tăng = xanh lá; tổng phải thu, phải trả, đã trả vendor = xanh dương). Tuổi nợ
           từ xanh lá (chưa đến hạn) sang đỏ đậm (trên 120). Màu phân biệt Sales/khách/vendor chỉ
           dùng họ xanh — đã chạy bộ kiểm tra màu (đạt mù màu). Chọn một Sales thì thẻ "theo Sales"
           tách riêng phần của người đó theo từng khách, trỏ vào thấy job.
           → luật màu đã lên RULES.md 5.5 ngày 2026-10-08.
-- claude - 2026-10-08 UTC+7

### 2026-10-08 (tối, vòng 3)
Failure    Đọc ngày chốt của chính báo cáo tool: sheet Summary ghi "At 03 Oct · Reporting period …
           02/10" trên một dòng, hàm đọc lấy 02/10 → tuần sau thấy "tuần trước" giả cách 1 ngày,
           nợ xấu tính sai (2,02 tỷ → 0,96 tỷ). Phát hiện khi chạy lại trên workspace thử.
           Báo lỗi giả: không
Lessons    Ưu tiên "At dd Mon yyyy" khi đọc ngày chốt; dựng điểm lịch sử tuần trước chỉ khi cách
           ≥ 5 ngày. Có phép thử khoá lại.
           user: tổng phải thu tăng = đỏ. Bấm khách ở bất kỳ đâu → chip "Khách" + mọi biểu đồ
           liên quan lọc theo (đồng bộ toàn trang). Bấm menu cuộn mượt. Dòng con job nằm đúng cột
           tuổi (phân bổ nhóm già cho nợ cũ / job cũ trước; dòng FIN khai theo hoá đơn lấy đúng
           nhóm của hoá đơn). File sample/Tham_chieu_job_no_cu.xlsx: nút 2 tạo lần đầu, FIN điền
           Job + Sales một lần; KIEM_TRA có sheet "Nợ chưa có job" cho dòng mới. Lịch sử tuần
           dựng lại W39 từ bản final → biểu đồ xu hướng thành đường. Dashboard thêm dải tóm tắt
           tuần, đường nhỏ trong ô số, hiệu ứng vẽ biểu đồ. 50 phép thử đạt.
-- claude - 2026-10-08 UTC+7

### 2026-10-09
user mở Credit_Dashboard.html trên Windows: chạy được. Phần [RUNTIME UNVERIFIED] mở dashboard
coi như đã qua; ký đóng phiếu vẫn chờ user.
### 2026-11-05
Failure: không
Báo lỗi giả: không
Lessons: đóng phiếu theo lệnh user sau khi tool chạy thật W43–W44 trên máy FIN.
-- claude - 2026-11-05 10:05 UTC+7
