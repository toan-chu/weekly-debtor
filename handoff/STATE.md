Cập nhật: 2026-11-05 12:10 UTC+7 - bởi claude

## Việc đang mở
WR-15 - người làm claude - chờ-soát - vòng 1/2 - 5 mức màu Head FIN (v3), chỉ tô ô Risk level, sheet Tham số trong file credit term (v0.9; chờ user commit + chạy lại W44 trên máy FIN)
WR-13 - người làm claude - chờ-soát - vòng 1/2 - tiêu chí FIN 3 mức + 3 ô Sales + điểm A–D + hạn từng job (v0.8; chờ user commit + chạy lại W44 trên máy FIN)
WR-14 - người làm claude - chờ-đề-bài - vòng 0/2 - hợp đồng PDF, hạn mức từ Client Master (chờ SharePoint List; T1 lãi gộp + T2 tạm tính đã làm ở WR-15)
WR-12 - người làm claude - chờ-soát - vòng 1/2 - khách mới tự vào file credit term, link HĐ, khoá tuần đã chốt (chờ lần chạy 2 trên máy FIN: link + hạn mức)
WR-02 - người làm claude - chờ-đề-bài - vòng 0/2 - so bản tay trên cùng bộ file Misa (chờ user: còn cần không)

## Hôm qua
- 06/11: WR-15 (v0.9) màu 5 mức Head FIN, 99 phép thử đạt; W44 trong máy ảo: 54/55 khách bị +1 (T2 tạm tính 36, T4 38) — cần Head FIN xem. Repo đổi tên weekly-debtor (3_Sua_lich_chay.bat).
- 05/11: WR-13 (v0.8) xong: 92 phép thử đạt; W44 chạy lại trong máy ảo: tổng y bản cũ, 36/5/14 khách Bình thường/Cảnh báo/Rủi ro cao, Excel = Python 55/55.
- 05/11: user gửi W44 (chạy 30/10 trên máy FIN) + file tiêu chí đánh giá khách hàng phải thu do FIN ban hành (3 mức, 3 tiêu chí nâng hạng). Đóng WR-01, 03–11.
- 20/10: WR-09 sửa theo FIN sau W42; chạy lại W42 khớp bản human 64/68 dòng job, 4 khách tool theo sổ, chờ FIN xác nhận.
- 22/10: commit WR-09 (5aff24e).
- 27/10: người làm tự soi lại WR-09, thấy 2 lỗi đếm "khách chưa có hợp đồng", Methodology chưa sửa, script nâng cấp Client Master bỏ cột lạ. Chưa sửa mã.
- 27/10: WR-10 đưa repo từ chuẩn v1.1 lên v2.0; mọi phiếu chuyển sang khuôn mới, chờ codex soát.
- 28/10: WR-12 bổ sung: file credit term tự lên mẫu mới mỗi lần chạy kể cả không có khách mới; README cho FIN; 87 phép thử đạt. Chờ user commit + triển khai máy FIN (GitHub Desktop Pull).
- 27/10: WR-12 (v0.7.1): khách mới tự vào file credit term (W43 thật: 5 khách, 83 dòng cũ giữ nguyên), ô Mở hợp đồng, khoá tuần đã chốt; 85 phép thử đạt. Tên gọi: file credit term; Client Master = SharePoint List sau này.
- 27/10: WR-11 (v0.7) xong, 81 phép thử đạt; chạy lại W43 thật trong máy ảo: tổng y bản cũ, màu gốc khớp bản Head FIN 61/61. Kịch bản chạy lại W43 trên máy FIN: dùng bản trong WR-12 (thay WR-11).

## Vừa đóng
WR-01, WR-03 → WR-11 - đóng-đạt - 2026-11-05 (lệnh user; bằng chứng: W43 chạy lại + W44 chạy thật trên máy FIN)

## Đang treo, cần user
0. (Đã xong 05/11: FIN chốt 3 mức, WR-13 thay màu CEO. X, Y chuyển WR-14.)
1. Kỳ vọng: WR-02 (kế toán làm tay trên đúng bộ file tool đã chạy rồi so) còn cần không? WR-04 → WR-09 đã so W40–W42 với bản human. Không cần thì đóng-huỷ.
2. Kỳ vọng: "Ngày gia hạn chỉ có hiệu lực khi đã tới ngày đó" (thêm ở WR-09, luật R.2) có đúng ý không? Nếu Sales gia hạn sớm thì hạn mới tính từ ngày gia hạn hay từ ngày hết hạn cũ?
3. Đề xuất luật cho RULES (bẫy sống quá một tuần):
   a. Không được ghép phiếu thu với hoá đơn đầu tiên mang số đó; phải chọn hoá đơn cùng số gần nhất trước ngày thu (Misa đánh số lại mỗi năm).
   b. Không được tin tên file Misa (Windows thêm hậu tố (21), (22)); phải đọc ngày bên trong file.
   c. Không được so khớp khách giữa các file Misa theo mã; phải so theo tên đã chuẩn hoá.
   d. Gộp hai câu ở R.2 về thứ tự trừ tiền thu: câu v0.4 "còn lại trừ hoá đơn cũ nhất trước" nay đã bị câu 20/10 (ưu tiên hoá đơn Misa "Đã thanh toán") thay một phần.
4. Luật bị thay khi sinh lại RULES từ v2.0 (xem AUDIT WR-10): mục 1 cũ "Chỉ huy" và giải trừ giới hạn; mục 2 cũ quyền ký nghiệm thu theo vai; mục 3 cũ trần vòng lặp (nay A3); mục 4 cũ "Bốn câu hỏi kiểm tra của user" (nay A10).

## Bẫy đã biết
- Hoá đơn nhập lùi ngày làm dư đầu kỳ Misa khác dư cuối báo cáo tuần trước: xem cột "Backdated vs last report" ở AR Risk và Payable.
- File Tổng hợp và Chi tiết xuất khác giờ thì số dư lệch (W40: Thủy Ngân 308tr); tiền theo Tổng hợp, Review liệt kê.
- File Excel thật trong `handoff/docs/` và `handoff/archive/` bị .gitignore chặn: máy khác không có, phải chép tay.
- Phép thử: 20 phép thử bỏ qua khi thiếu dữ liệu thật trong `tests/fixtures/`; máy ảo Linux cần `pip install --user pytest openpyxl` trước. `tests/conftest.py` hạ ngưỡng số lẻ xuống 0,5đ — phép thử cần ngưỡng thật phải tự đặt lại.

## Báo lỗi giả
Tháng này: 0
