Mã: WR-14
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: WR-13
Người làm: claude
Người soát: codex
Trạng thái: chờ-đề-bài
Vòng: 0/2

## ĐỀ BÀI

Kỳ vọng của user:
(2026-11-05) "bỏ qua cái vụ hợp đồng pdf, hạn mức nợ và lãi gộp (3 cái này thành phiếu, chờ list master, còn lại ta xử lý
trước)" · "tôi đang dựng bảng client master + contract upload trên đó."

Mục tiêu: Khi Client Master (SharePoint List) + nơi upload hợp đồng xong: lấy link hợp đồng PDF, hạn mức nợ, tình trạng hợp
đồng từ List; áp TC2 (dư nợ > credit limit, chưa có thì tạm tính theo billing) và X1 (25% lãi gộp BQ tháng quý gần nhất)
theo tiêu chí FIN; khi có lãi gộp thì bật TC1 (X = MIN(X1, X2 top 15% làm tròn lên)) và TC3 (khách mới / chưa có HĐ, tổng dư nợ ≥ Y = X/2) — FIN chốt 05/11 chờ lãi gộp mới áp.

Ngoài phạm vi:
- Phần tiêu chí FIN không phụ thuộc hợp đồng / hạn mức / lãi gộp (WR-13).

Kế thừa:
- WR-12: link hợp đồng + hạn mức đang đọc từ file credit term (SAMPLE_credit_term_khach.xlsx).
- WR-13: tóm tắt tiêu chí FIN; file nguồn `handoff/docs/Đánh gái khách hàng phải thu và tiêu chí (2) 1.xlsx`.

Kịch bản test: viết khi đủ đề bài.

### Câu hỏi để dành (hỏi khi List xong)
- Link thư mục / cột hợp đồng PDF trên List · bảng hạn mức trong hợp đồng.
- Ai cập nhật lãi gộp mỗi quý, ngày nào (T7–T9/2026 đã có trong file FIN: ~15,2 / 9,8 / 10,5 tỷ).
- Hạn mức tạm tính: billing có trừ giảm trừ không · gốc kỳ hạn (ngày invoice / cuối tháng) đọc từ Thời hạn hay cột riêng.
- TC3 "chưa có hợp đồng" + ô HĐ còn hiệu lực / hồ sơ đầy đủ lấy từ List · ngày gia hạn có hiệu lực khi nào.
- Nguồn sự thật: List thay file credit term hay song song?

## TODO
- [ ] Chờ Client Master SharePoint List + upload hợp đồng

## AUDIT

## HISTORY
### 2026-11-05
Mở phiếu, chờ Client Master theo lệnh user.
Failure: không
Báo lỗi giả: không
Lessons: không
-- claude - 2026-11-05 10:30 UTC+7
