# weekly-debtor

Công cụ sinh báo cáo công nợ hằng tuần (phải thu, phải trả, dòng tiền, rủi ro) cho Trustana Việt Nam.
Kế toán thả 6 file (5 file Misa + 1 file SMS AR-AP) vào thư mục `input/` trên OneDrive; một tiến trình
Python chạy nền theo lịch đọc các file đó cùng sổ ghi của chính nó, rồi ghi file Excel báo cáo tuần và
dashboard HTML vào thư mục gốc. Không ai phải điền tay. Đây là tool đầu tiên trong chuỗi credit control.

## Trước khi làm bất cứ việc gì

1. Đọc `handoff/RULES.md` — luật đứng, sinh từ handoff-standard v2.0.
2. Đọc `handoff/STATE.md` — hiện trạng, việc đang mở.
3. Mở đúng một phiếu việc của mình trong `handoff/`.

Đụng tới kiến trúc thì đọc thêm `handoff/MAP.md` và vẽ ở đó trước khi đổi mã.
Cấm quét cả `handoff/`, cấm đọc phiếu trong `handoff/done/` khi không có trường `Nối tiếp` trỏ tới,
cấm đọc `handoff/docs/` khi không có dòng nào trỏ tới, cấm đọc `handoff/archive/` trừ khi user yêu cầu.

Chuẩn phối hợp: `handoff/handoff-standard-v2.0.md`.
