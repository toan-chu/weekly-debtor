Mã: WR-03
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: WR-01
Người làm: claude
Người soát: codex
Trạng thái: đóng-đạt
Vòng: 0/2

## ĐỀ BÀI

Kỳ vọng của user: (lời user, 2026-10-06)
Tuổi nợ W39 sai vì 2 file Misa không có ngày hoá đơn. Bổ sung file SMS (AR-AP)
làm nguồn ngày cho từng job, file ghép khách SMS ↔ Misa do FIN điền một lần, và
bản final đã duyệt của kế toán trưởng làm trí nhớ cho nợ cũ không có trong SMS.
Cài lên máy FIN với:
- một nút khởi tạo thư mục làm việc,
- một nút chạy file ghép khách (debtors) để FIN kiểm tra,
- sau đó mỗi tuần FIN chỉ việc bỏ 3 file (2 Misa + 1 SMS) vào là ra báo cáo.

Luật nghiệp vụ user đã chốt trong hội thoại 2026-09-28 → 2026-10-06 (thuộc Kỳ vọng của user):
- Tiền mỗi khách luôn theo Misa Tổng hợp.
- SMS > Misa: khách đã trả mà SMS quên gạch, cắt bớt job cũ nhất cho bằng Misa,
  liệt kê phần cắt để FIN gạch trên SMS (FIN xác nhận SMS hay quên ấn paid).
- SMS < Misa: phần dư là nợ cũ, lấy nhóm tuổi từ bản final/sổ ghi tuần trước.
- Khách FIN ghi "không có khả năng thu hồi": luôn ở nhóm 120+.

Mục tiêu: Sửa tuổi nợ bằng SMS AR-AP, file ghép khách SMS ↔ Misa và bản final tuần trước; cài được lên máy FIN bằng nút bấm.

Ngoài phạm vi:
- Power Automate, gửi email.
- Đổi mẫu báo cáo 14 cột. Số job ghi vào cột Ghi chú.

Kế thừa: (phiếu mở trước v2.0, không ghi; xem phiếu ở trường Nối tiếp)

Kịch bản test:
Theo lời user — CHỜ USER chốt ngưỡng số:
1. FIN nháy đúp `1_Khoi_tao_workspace.bat`, chọn thư mục → thấy 2 ngăn `input`, `sample`.
2. FIN bỏ file ghép khách đã điền vào `sample`, bản final gần nhất vào thư mục gốc,
   nháy đúp `2_Kiem_tra_ghep_khach.bat` → trong vòng 1 phút Excel mở file kiểm tra,
   liệt kê dòng cần sửa và khách mới chưa ghép.
3. FIN bỏ 3 file tuần vào `input` → trong vòng 15 phút (hoặc nháy đúp
   `3_Chay_ngay.bat`) báo cáo tuần xuất hiện ở thư mục gốc, tổng = tổng Misa.
   Tuần sau tool tự đọc báo cáo tuần trước ở gốc, không cần bỏ thêm file nào.
4. Ngưỡng đạt về tuổi nợ: ___/70 khách khớp bản kế toán trưởng tuần 40 — CHỜ USER.
Không bao giờ được xảy ra: tổng báo cáo khác tổng Misa; file của FIN bị xoá.

## TODO
- [x] Vẽ luồng mới vào MAP.md trước khi đổi mã
- [x] Đọc file SMS AR-AP, file ghép khách, bản final dạng sheet `data`
- [x] Luật ghép: Misa quyết tiền, SMS quyết ngày, final quyết nợ cũ, khó đòi 120+
- [x] Lệnh `--kiem-tra-ghep-khach`, 3 nút .bat
- [x] Chặn dữ liệu thật trong handoff/docs khỏi git
- [x] Phép thử mới, chạy thử trên dữ liệu thật W40
- [ ] user/codex chạy trên Windows: 3 nút, Task Scheduler   [RUNTIME UNVERIFIED]

## AUDIT

Báo cáo của người làm:
(xem HISTORY 2026-10-06 — lệnh và kết quả dán ở đó)

Kết luận cũ của người làm (chuẩn v1.1)
- Logic tính số (Claude):            xem AUDIT
- Chạy 3 nút trên Windows (user/codex): ☐
- Ngưỡng khớp bản tay tuần 40 (user): ☐

Giả định kỹ thuật: (phiếu mở trước v2.0, chưa ghi theo khuôn ba dòng)

Soát: user — đóng theo lệnh user 2026-11-05 ("các ticket đang mở cái nào xong thì close & move vào done"). Bằng chứng: Báo cáo W43 (chạy lại 29/10) và W44 (30/10) do Task Scheduler trên máy FIN sinh ra, dùng thật trong tuần (handoff/docs/Cong-No-Workspace (W44): runner.log, 2026_W44_Credit_Report_24.10-30.10.2026.xlsx — 55 khách, phải thu 13.099.776.819; Sales đã ghi lý do tuần này vào file). Không qua codex soát (lệch mục 0/9 của chuẩn, user quyết).
Kết quả: ĐẠT
-- claude - 2026-10-27 11:40 UTC+7

## HISTORY

### 2026-10-06
Làm      `wr/ghep.py` (đọc SMS AR-AP, đọc file ghép khách, gợi ý ghép theo tên),
         `tinhtoan.tinh_bang_sms` (Misa = tiền, SMS = ngày, sổ ghi/final = nợ cũ,
         khó đòi = 120+, SMS > Misa → bỏ job cũ nhất, vênh ≤1%/500k = tỷ giá),
         `wr/kiemtra.py` (nút 2), nạp tự động bản final trong 03_BanTay
         (đọc được dạng sheet `data` của kế toán trưởng), ngăn 04_GhepKhach,
         3 nút .bat, chờ SMS tối đa 2 giờ, bọc lỗi chạy nền, chặn handoff/docs/*.xlsx khỏi git.
AUDIT    `python -m pytest -p no:cacheprovider tests`
           → 27 passed, 20 skipped (20 cần dữ liệu thật trong fixtures/, máy này không có)
         Làm hỏng thử: đảo "lô cũ nhất trước" thành "mới nhất trước" và tắt hấp thụ tỷ giá
           → 2 phép thử đỏ (test_sms_it_hon_misa..., test_kho_doi..._chenh_ty_gia...). Trả lại → xanh.
         Chạy thật W40 trong thư mục tạm (Misa W40 + AR-AP + file ghép FIN điền + w39 final):
           nút 2 → "68 mã đã ghép, 0 SMS mới, 1 Misa mới (KH00361), 1 dòng cần sửa (TTND VIETNGA)"
           chạy  → "70 khách, tổng 14,251,295,535, 10 dòng cần xem lại" — tổng KHỚP Misa
           Nguồn tuổi: SMS 9,35 tỷ · bản final 3,75 tỷ · khó đòi 0,91 tỷ · ước tính 0,25 tỷ
           (ước tính: LLC "ILC" 244 triệu, PSL 1,98 triệu — không có trong SMS lẫn W39 final).
KHÔNG ĐO ĐƯỢC  Chạy 3 nút trên Windows, Task Scheduler, Excel mở file   [RUNTIME UNVERIFIED]
         Ngưỡng khớp bản tay tuần 40 — chưa có bản tay W40, chưa có ngưỡng.
[GIẢ ĐỊNH KỸ THUẬT] Tuổi tính từ ETD (Outbound/Logistics) hoặc ETA (Inbound) + credit term.
         Vì sao: SMS không có ngày hoá đơn, FIN chưa trả lời câu này.
         Sai thì: đổi một dòng thứ tự ETD/ETA trong `ghep.doc_sms`, chạy lại.
[GIẢ ĐỊNH KỸ THUẬT] SMS > Misa thì bỏ job cũ nhất trước (FIN: SMS hay quên ấn paid).
         Sai thì: đổi chiều sắp xếp trong `tinh_bang_sms`, phép thử sẽ chỉ chỗ.
Failure  `git status` để lại `.git/index.lock` vì máy ảo không có quyền xoá; đã xin
         quyền và xoá. Từ nay dùng `git --no-optional-locks status`.
         Báo lỗi giả: không
Lessons  File thật của FIN đặt tạm trong handoff/docs suýt đi theo git — đã chặn bằng .gitignore.
-- claude

### 2026-10-06 (chiều)
Đổi     user chốt bố cục thư mục đích: gốc chứa báo cáo + `input/` + `sample/`;
        tuần sau tool đọc lại báo cáo tuần trước ở gốc (kể cả bản kế toán sửa tay).
        Trí nhớ, nhật ký chuyển vào `_tool/` ẩn. Báo cáo cùng tên đã có thì không ghi đè.
AUDIT   `python -m pytest -p no:cacheprovider tests` → 28 passed, 20 skipped
        (thêm test_tuan_sau_doc_lai_bao_cao_tuan_truoc_o_goc_va_khong_ghi_de).
        Chạy thật W40 bố cục mới: w39 final ở gốc được đọc (71 khách, chốt 26/09),
        báo cáo W40 ra ở gốc, 70 khách, tổng 14,251,295,535 khớp Misa.
Failure không
        Báo lỗi giả: không
Lessons "Có cần file tuần trước không": chỉ lần đầu. Không có nó thì ~3,75 tỷ nợ cũ
        (26%) quay lại tình trạng đoán tuổi như tuần 39.
-- claude

### 2026-10-06 (tối)
Đổi     user: báo cáo có lúc gửi đối tác nước ngoài → chữ trên sheet Receivable để
        tiếng Anh: "TRUSTANA VIETNAM", cột "No.", "Notes", "Reason for late payment";
        sheet "Cần xem lại" đổi tên "Review"; bỏ 2 dòng trắng (tên cột ở dòng 5, dữ liệu
        từ dòng 6). Chữ tiêu đề gom vào `wr/thuonghieu.py`, tool ghi lại mỗi lần chạy.
AUDIT   pytest → 28 passed, 20 skipped. Chạy thật W40: sheet ['Receivable', 'Review'],
        dòng 1–5 đúng như trên, tổng 14,251,295,535 khớp Misa.
Còn treo "Accounts receivable" → user muốn thêm "s". Claude giữ nguyên chờ chốt:
        thuật ngữ chuẩn là "Accounts Receivable" (không s). Sửa 1 dòng TEN_BANG.
Failure không
-- claude
### 2026-11-05
Failure: không
Báo lỗi giả: không
Lessons: đóng phiếu theo lệnh user sau khi tool chạy thật W43–W44 trên máy FIN.
-- claude - 2026-11-05 10:05 UTC+7
