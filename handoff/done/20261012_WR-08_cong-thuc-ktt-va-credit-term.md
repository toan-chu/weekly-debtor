Mã: WR-08
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: WR-06, WR-07
Người làm: claude
Người soát: codex
Trạng thái: đóng-đạt
Vòng: 1/2

## ĐỀ BÀI

Kỳ vọng của user: (user, 2026-10-12)
1. File Excel cuối có sẵn công thức của kế toán trưởng (bản W40 KTT gửi).
2. Credit term từng khách: file mẫu FIN điền cho khách cũ + phiếu Sales điền cho khách mới;
   tham khảo credit-term-draft của KTT.
3. Giải thích vì sao ATH bị đánh giá khác nhau giữa hai tuần.
4. Trả lời Customer Experience: AR Risk đã chính xác chưa · Sales ghi lý do ở đâu · căn cứ và
   người duyệt ngưng dịch vụ · khách nhiều Sales (Inter-Freight) phải thấy job nào của Sales nào.
5. Chạy lại W41 ở workspace test, so với bản W41 "(human)" FIN đã sửa.

Mục tiêu: File Excel mang công thức của kế toán trưởng, credit term từ Client Master, và trả lời các câu hỏi của CX về AR Risk.

Ngoài phạm vi:
- Khoá dịch vụ tự động — KTT 13/10: không làm, FIN tự quyết.

Kế thừa: (phiếu mở trước v2.0, không ghi; xem phiếu ở trường Nối tiếp)

Kịch bản test:
user: chạy lại W41 ở workspace test, so bản human.

## TODO
- [x] Dòng con = số còn nợ từng hoá đơn theo Chi tiết 131; job từ Sổ chi tiết 131 (chia theo tỷ lệ);
      hoá đơn trước năm nay lấy job từ bảng job nợ cũ; NVK ghi "Điều chỉnh / chênh lệch tỷ giá"
- [x] AR Risk ghi công thức KTT + tô màu theo Tier; TOTAL Receivables = SUMIFS
- [x] Dashboard tự tính lại S1–S6, Trend, Required action (công thức chưa có số khi chưa mở Excel)
- [x] Đọc credit term từ file mẫu trong sample/ theo tiêu đề cột (dùng nguyên file KTT)
- [x] PHIEU_khach_moi.xlsx theo thứ tự cột của KTT; tạo file credit term trống khi chưa có
- [x] Cột "Jobs by Salesman" cho mọi khách nhiều Sales
- [x] Bỏ khách số dư dưới 1.000đ, liệt kê ở Review
- [x] Bản FIN sửa thẳng trên file tool thắng (sổ ghi nhớ giờ sửa); tên có "human" thắng
- [x] Phép thử test_wr8.py; sửa test cũ theo nghĩa mới
- [x] (13/10) Sheet Contracts + 5 chỉ số Summary + ô "Hợp đồng & credit term" trên dashboard, nhãn trên trang khách
- [x] (13/10) S7 lỡ hẹn; tên ngắn trên nhãn biểu đồ
- [x] (13/10) Đọc thẳng file Client Master trên OneDrive (`--client-master`, hỏi trong cài đặt); bỏ PHIEU_khach_moi
- [x] (13/10) Client Master List theo theme (`wr/khachhang.py`): file chung OneDrive, trạng thái hợp đồng tự tính
      ≤30/60/90 ngày theo ngưỡng của KTT; file mẫu = 71 khách KTT + 11 khách W41; lần đầu tool cũng tạo bằng mẫu này

## AUDIT

Báo cáo của người làm:
```
python -m pytest -q          → 69 passed, 20 skipped (13/10)
W41 workspace test           → 66 khách; phải thu 14.468.170.527; quá hạn 5.663.131.210;
                               nợ xấu 1.730.203.676; phải trả 9.737.786.631
So bản (human) FIN sửa       → 66/66 khách; 58/66 dòng con y hệt (trước WR-08: 9 khách FIN phải sửa)
LibreOffice tính lại         → 931 công thức, 0 lỗi; AR Risk khớp Python 66/66
Dashboard headless           → 0 lỗi; Inter-Freight hiện tách 3 nhóm Sales
```

Lệch còn lại (8 khách, nhỏ): KH00418 khó đòi (dòng nợ cũ giữ chỗ) · OPTIMALOG, STA (FIFO chọn khác
hoá đơn) · ILC (chia 777k) · SPA (dòng NVK tỷ giá, đã gắn nhãn) · AIRSPACE, KH00140 (làm tròn tỷ
giá 150k / 23k) · DEALLOG (nhóm tuổi khác vì term).

Chờ kiểm trên máy thật: [RUNTIME UNVERIFIED] trên máy Windows của FIN.


Việc trên máy thật còn chờ (chuyển từ mục "Đang treo" của STATE cũ, 2026-10-27):
- Cài máy FIN kèm `.gitattributes`; đưa Client Master lên OneDrive và cho máy FIN đồng bộ file về máy.
- Chạy WR-08 trên máy user rồi máy FIN [RUNTIME UNVERIFIED].
- Dữ liệu: Sales điền ngày hết hạn hợp đồng (63 khách đang nợ còn trống), FIN điền term cho 22 khách đang nợ.

Giả định kỹ thuật: (phiếu mở trước v2.0, chưa ghi theo khuôn ba dòng)

Soát: user — đóng theo lệnh user 2026-11-05 ("các ticket đang mở cái nào xong thì close & move vào done"). Bằng chứng: Báo cáo W43 (chạy lại 29/10) và W44 (30/10) do Task Scheduler trên máy FIN sinh ra, dùng thật trong tuần (handoff/docs/Cong-No-Workspace (W44): runner.log, 2026_W44_Credit_Report_24.10-30.10.2026.xlsx — 55 khách, phải thu 13.099.776.819; Sales đã ghi lý do tuần này vào file). Không qua codex soát (lệch mục 0/9 của chuẩn, user quyết).
Kết quả: ĐẠT
-- claude - 2026-10-27 11:40 UTC+7

## HISTORY

### 2026-10-13
Client Master List (file chung OneDrive) theo theme. Dashboard: tên ngắn trên nhãn biểu đồ. ATH hẹn
"tháng 9" mà sang tháng 10 vẫn N2 → user chọn A: thêm S7 lỡ hẹn (Promise date đọc từ lý do),
đếm như S1–S3. Chạy lại W41: ATH lên N3, nợ xấu 1.730.203.676 → 1.852.888.263 (16 khách). LibreOffice
tính lại: 0 lỗi, nợ xấu theo công thức khớp số tool. Dashboard: ATH hiện S2 S6 S7, N3, 0 lỗi.
Failure: công thức KTT coi mọi câu Sales ghi là đã giải trình, kể cả lời hẹn đã quá ngày.
Lessons: lý do có ngày thì ngày đó là một mốc phải kiểm, không chỉ là chữ.
Chiều: KTT chốt N3 = câu nhắc, FIN sửa cuối Client Master, HĐ hết hạn theo term cũ → sheet Contracts.
W41: 2 HĐ hết hạn còn nợ, 1 sắp hết hạn, 22 chưa có term. Dashboard 0 lỗi.

### 2026-10-12
Làm toàn bộ TODO, chạy lại W41.
Failure: W41 ban đầu FIN phải sửa tay 9 khách vì dòng con lấy tiền hoá đơn gốc thay vì số còn nợ.
Lessons: tiền dòng con phải đi từ sổ (đã trừ phiếu thu), trạng thái Misa chỉ dùng để đối chiếu.
ATH: W40 bị S3 (lần đầu quá hạn) nên lên nhóm cao; W41 S3 hết, còn N2 term "thanh toán ngay".
Lý do cũ vì tool đọc lại file W40 của chính nó — nay bản FIN sửa thắng.
### 2026-11-05
Failure: không
Báo lỗi giả: không
Lessons: đóng phiếu theo lệnh user sau khi tool chạy thật W43–W44 trên máy FIN.
-- claude - 2026-11-05 10:05 UTC+7
