# 📊 Báo cáo công nợ tuần — weekly debtor report

[![License](https://img.shields.io/badge/license-nội%20bộ%20Trustana-4d148c?style=for-the-badge)](#)
[![Python](https://img.shields.io/badge/Python-3.9%2B-4d148c?style=for-the-badge&logo=python&logoColor=white)](#)
[![Hạ tầng](https://img.shields.io/badge/OneDrive%20%2B%20Task%20Scheduler-0078D4?style=for-the-badge)](#)
[![Windows](https://img.shields.io/badge/Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)](#)
[![Phép thử](https://img.shields.io/badge/phép%20thử-99%20đạt-2ea043?style=for-the-badge)](#️-dành-cho-người-bảo-trì)

> **Mỗi tuần tự ra một file Excel cho TRUSTANA Việt Nam: khách còn nợ bao nhiêu, đã trả bao nhiêu, mình còn nợ vendor bao nhiêu, khách nào đáng lo — FIN và kế toán trưởng không phải làm tay.**
> FIN thả 6 file (5 Misa + 1 SMS) vào `input/` -> tool chạy ngầm 15 phút một lần, tính tuổi nợ từng hoá đơn -> báo cáo tuần và dashboard xuất hiện ở thư mục gốc workspace.

## 🎯 Ai làm gì

| Vai trò | Việc | Cần biết code? |
|---|---|---|
| 👩‍💼 **FIN** | Mỗi tuần xuất 5 file Misa + 1 file AR-AP từ SMS **cùng lúc**, thả vào `input` (~10 phút) | Không |
| 👩‍💼 **FIN** | Xem cột **Đối chiếu** + tab Review, sửa chỗ ĐỎ / CAM trên Misa (tuỳ tuần) | Không |
| 👩‍💼 **FIN** | Một lần: điền file ghép khách + bảng job nợ cũ, bấm `2_Kiem_tra_ghep_khach.bat` (~30 phút) | Không |
| 👩‍💼 **FIN** | Giữ **file credit term** (`SAMPLE_credit_term_khach.xlsx`): điền Thời hạn, Hạn mức nợ cho khách tool vừa thêm (ô vàng); FIN là người sửa cuối (~5 phút/tuần) | Không |
| 👩‍💼 **FIN** | **Mỗi tháng một lần**: nhập lãi gộp 3 tháng của quý gần nhất vào sheet **Tham số** của file credit term (~2 phút) | Không |
| 🧑‍💼 **Sales** | Tích ✓ hợp đồng, dán link hợp đồng, ngày hết hạn, gia hạn trong file credit term (file chung OneDrive); mỗi tuần ghi lý do trễ hạn tuần này | Không |
| 👩‍💼 **Kế toán trưởng** | Đọc báo cáo, sửa lý do trễ hạn ngay trên file, trình sếp | Không |
| 🧑‍🔧 **Người cài đặt** | Nháy đúp `1_Khoi_tao_workspace.bat`, một lần duy nhất (~5 phút) | Không |
| 🛠️ **Người bảo trì** | Sửa mã, chạy phép thử, đẩy bản mới lên GitHub | Có |

Tool tự làm phần còn lại trong ~10 giây mỗi lần chạy: đọc file, tính tuổi nợ từng hoá đơn, phân nhóm rủi ro, ghi báo cáo.

## 🔄 Luồng chạy

```
  👩‍💼 FIN (mỗi tuần, xuất cùng lúc)              🤖 TOOL (15 phút thức dậy 1 lần, chạy ẩn)
  ─────────────────────────────────            ──────────────────────────────────
  Misa  Tổng hợp công nợ phải thu    ─┐        1️⃣ nhận dạng file theo nội dung; thiếu file → chờ 2 giờ
  Misa  Chi tiết công nợ phải thu     │        2️⃣ tiền mỗi khách ← Tổng hợp phải thu
        (từ 01/01/2022)               │        3️⃣ tuổi nợ ← từng hoá đơn trong file Chi tiết
  Misa  Tổng hợp công nợ phải trả     ├──► 📥  4️⃣ job chưa thanh toán ← Bán hàng + Sổ chi tiết TK131
  Misa  Bán hàng (từ đầu năm)         │  input    Salesman của job ← SMS
  Misa  Sổ chi tiết TK131 (đầu năm)   │        5️⃣ đối chiếu: tổng job vs dư nợ → tô màu cho FIN
  SMS   AR-AP                        ─┘        6️⃣ nhóm rủi ro ← luật của kế toán trưởng; mức độ ← tiêu chí FIN
                                               7️⃣ cảnh báo hợp đồng / credit term ← file credit term
                                               8️⃣ phải trả (số tổng), dòng tiền, so với tuần trước
  🔗 sample/: ghép khách + job nợ cũ (1 lần) + file credit term (term, hạn mức, hợp đồng — tool tự thêm khách mới)
  📊 gốc/: báo cáo các tuần trước ───────────►    (đọc lại bản kế toán đã sửa)
                                                          │
                                                          ▼
                                   📊 gốc/2026_W41_Credit_Report_....xlsx   ✅ input đổi tên [DONE]
```

## 🚀 Cài đặt — một lần duy nhất

```
1️⃣  Cài Python từ python.org            ⚠️ nhớ tick "Add Python to PATH"
2️⃣  Nháy đúp 1_Khoi_tao_workspace.bat   chọn thư mục đích, tool tạo các ngăn + lịch 15 phút chạy ẩn
3️⃣  Bỏ file ghép khách (FIN đã điền)    vào sample
    Bỏ bản final gần nhất (KTT duyệt)   vào thư mục gốc — chỉ lần đầu, để có mốc so sánh
4️⃣  Nháy đúp 2_Kiem_tra_ghep_khach.bat  Excel mở file KIEM_TRA; lần đầu tool tạo luôn
                                        sample/Tham_chieu_job_no_cu.xlsx để FIN điền job nợ cũ
```

**File credit term (credit term + hạn mức + hợp đồng) dùng chung trên OneDrive** — chọn một trong hai cách:

- **Cách 1 (đơn giản nhất):** để file `SAMPLE_credit_term_khach.xlsx` ngay trong `sample/` của workspace
  (workspace đã nằm trên OneDrive). Trên OneDrive bấm chuột phải vào file → **Share** → chọn Sales, quyền
  **Can edit**. Không phải cài đặt gì thêm.
- **Cách 2:** file nằm chỗ khác trên OneDrive (VD thư mục của chị KTT). Trên máy FIN: chuột phải thư mục
  chứa file → **Always keep on this device**; giữ Shift + chuột phải vào file → **Copy as path**; chạy lại
  `1_Khoi_tao_workspace.bat` và dán đường dẫn khi được hỏi (hoặc
  `python runner.py --client-master "<đường dẫn>"`). Có file chung thì tool chỉ đọc file đó, bỏ các bảng
  term trong `sample/`. Mất file (OneDrive chưa đồng bộ) → tool dùng tạm `sample/` và ghi cảnh báo ở Review.
  Xem đang dùng file nào: `python runner.py --xem-cau-hinh`.

### 📦 Bắt đầu trên máy FIN bằng gói tuần 40

Gói `sample_w40` (zip) có đủ mọi file để chạy lại tuần 40, làm mốc cho tuần 41:

| File trong gói | Bỏ vào | Là gì |
|---|---|---|
| `W39_bao-cao-final_nguoi-lam.xlsx` | thư mục gốc | Bản final W39 kế toán trưởng duyệt — mốc "tuần trước" |
| `SAMPLE_ghep-khach-SMS-Misa_FIN-da-dien.xlsx` | `sample/` | File ghép khách FIN đã điền (đã sửa TTND VIETNGA → KH00416) |
| `SAMPLE_Tham_chieu_job_no_cu.xlsx` | `sample/` | Bảng job nợ cũ tool đã dựng sẵn từ số W40 (1 hoá đơn: Minh Khôi 32tr) — FIN điền Job, Sales ở ô vàng |
| `SAMPLE_credit_term_khach.xlsx` | `sample/` | File credit term: 71 khách của chị KTT (credit-term-draft) + 11 khách đang nợ trong W41; 56 khách có term |
| `W40_misa-tong-hop-phai-thu-131.xlsx` | `input/` | File 1 tuần 40 |
| `W40_misa-chi-tiet-phai-thu-131_tu-2022.xlsx` | `input/` | File 2 tuần 40 |
| `W40_misa-tong-hop-phai-tra-331.xlsx` | `input/` | File 3 tuần 40 |
| `W40_misa-ban-hang_tu-dau-nam.xlsx` | `input/` | File 4 tuần 40 |
| `W40_misa-so-chi-tiet-131_tu-dau-nam.xlsx` | `input/` | File 5 tuần 40 |
| `W40_sms-AR-AP.xlsx` | `input/` | File 6 tuần 40 |

Thứ tự: cài đặt (bước 1–2 ở trên) → giải nén, bỏ file đúng chỗ → điền ô vàng trong
`SAMPLE_Tham_chieu_job_no_cu.xlsx` (làm sau cũng được) → Task Scheduler → **Run** → có báo cáo W40 ở
thư mục gốc. Bảng job nợ cũ không nằm trong git vì chứa số liệu khách; máy nào chưa có thì bấm
`2_Kiem_tra_ghep_khach.bat` là tool tự tạo.
Từ tuần 41: chỉ thả 6 file mới vào `input/`.

### 🔄 Cập nhật bản mới trên máy FIN (GitHub Desktop)

1. Người bảo trì đẩy code lên GitHub (GitHub Desktop: **Commit** → **Push origin**).
2. Trên máy FIN: mở **GitHub Desktop** → chọn repo `weekly-debtor` → **Fetch origin** → **Pull origin**.
   Cấu hình riêng của máy (`settings.json`) không đi theo git nên giữ nguyên; không cần cài lại.
   Lần chạy kế tiếp tool tự chép dashboard mới vào workspace.
3. **Không xoá workspace.** Thư mục `_tool` giữ trí nhớ các tuần trước; báo cáo cũ ở thư mục gốc là mốc so
   sánh. **Không file Excel nào đi theo git** (trừ mẫu chuẩn trong `template/`).
4. **File credit term không phải làm gì.** Bản mới cần thêm cột vào file? Lần chạy kế tiếp tool tự thêm cột (kể cả
   tuần không có khách mới), giữ nguyên mọi ô đã điền, và cất bản cũ vào `_tool/luu-tru` trước khi ghi. Chỉ cần
   **đóng file credit term** lúc tool chạy; file đang mở thì tool để lần sau và ghi một dòng ở sheet Review.
5. Muốn chạy lại một tuần bằng bản mới: vào `input/`, bỏ chữ `[DONE] ` ở 6 file của tuần đó → Task
   Scheduler → **Run**. Tool tạo file `… (chay lai …).xlsx`, không đè báo cáo FIN đã sửa.

### 📂 Đổi tên / chuyển thư mục mã nguồn

Git không mang tên thư mục, Pull không đổi tên giúp. Lịch chạy nền gọi đường dẫn đầy đủ, nên đổi tên là tool ngừng chạy
cho tới khi làm bước 3:

1. Đóng GitHub Desktop. 2. Đổi tên thư mục (VD `weekly-report` → `weekly-debtor`).
3. Nháy đúp **`3_Sua_lich_chay.bat`** trong thư mục mới: lịch chạy trỏ về thư mục mới, `settings.json` tự sửa đường dẫn
   cũ. Không hỏi gì, không đụng workspace OneDrive.
4. Mở GitHub Desktop → repo báo không thấy → **Locate…** → chọn thư mục mới.
5. Task Scheduler → *Bao cao cong no tuan* → **Run** → `_tool\runner.log` trong workspace có dòng mới là xong.

## 📋 Dùng thế nào

### 📁 Thư mục làm việc

```
📂 Cong-No-Workspace/
   ├── 📊 2026_W40_Credit_Report_....xlsx   ← 🎉 báo cáo tuần nằm ngay ở đây; sửa thẳng vào đây
   ├── 🌐 Credit_Dashboard.html             ← mở bằng Chrome/Edge, kéo báo cáo tuần vào
   ├── 📥 input/    ← mỗi tuần thả 6 file vào đây; xong tool đổi tên thành [DONE] ...
   ├── 🔗 sample/   ← những file FIN điền MỘT LẦN, tool đọc lại mỗi tuần (bảng dưới)
   └── 🙈 _tool/    ← ẩn: trí nhớ, nhật ký. Không cần mở, đừng xoá
```

**Các file trong `sample/`**

| File | Ai tạo | Ý nghĩa | FIN làm gì |
|---|---|---|---|
| `SAMPLE_ghep-khach-SMS-Misa_FIN-da-dien.xlsx` | FIN | Mã khách bên SMS ứng với mã khách nào bên Misa (tên hai bên khác nhau) | Có khách mới thì thêm dòng (file KIEM_TRA chỉ ra dòng nào) |
| `Tham_chieu_job_no_cu.xlsx` (tên nào cũng được, tool nhận theo nội dung) | Tool, lần đầu bấm nút 2 — gói W40 có sẵn | Hoá đơn **trước năm nay** còn nợ mà Misa không cho biết job | Điền cột Job, Sales (ô vàng) một lần |
| `SAMPLE_credit_term_khach.xlsx` (tên nào cũng được) | FIN + Sales, file chung trên OneDrive — gói W40 có sẵn | **File credit term**: credit term, hợp đồng, ngày hết hạn, trạng thái gia hạn (tự tính ≤30/60/90 ngày). Tool chỉ đọc cột "Mã kh" và "Thời hạn", thắng mọi nguồn khác. Sheet "Hướng dẫn" ghi ai điền gì | Ô vàng = còn thiếu; Sales thêm khách mới ở cuối, FIN điền Mã kh |
| `KIEM_TRA_ghep_khach_<ngày giờ>.xlsx` | Tool, mỗi lần bấm nút 2 | Kết quả kiểm tra: khách mới chưa ghép, dòng ghép sai, nợ cũ chưa có job | Chỉ đọc; sửa ở hai file trên |

> 🔁 Tuần sau tool đọc lại báo cáo tuần trước ở thư mục gốc để lấy lý do trễ hạn, credit term
> và làm mốc so sánh. FIN cứ **sửa thẳng trên file báo cáo** rồi lưu, để nguyên ở thư mục gốc:
> tool nhận ra file đã bị sửa và lấy bản đó làm chuẩn. Bản KTT tự làm (không có tab Review, hoặc
> tên có chữ "human") cũng luôn thắng bản tool cùng tuần. Đừng xoá báo cáo tuần gần nhất.

### 👩‍💼 FIN — mỗi tuần

**Xuất 6 file trong cùng một buổi** (số Misa đổi theo lúc xuất: tool so được hai file lệch nhau là
do xuất khác lúc, nhưng xuất cùng lúc thì không phải dò). Kỳ báo cáo: **thứ Bảy → thứ Sáu**.
Tên file là gì cũng được — tool đọc tiêu đề bên trong file.

| # | Lấy ở | Báo cáo (tiêu đề trong file) | Chọn khi xuất | Tool dùng để |
|---|---|---|---|---|
| 1 | Misa | **Tổng hợp công nợ phải thu khách hàng** | TK 131 · loại tiền Tổng hợp · thứ Bảy → thứ Sáu | Tiền còn nợ của từng khách (số chuẩn) |
| 2 | Misa | **Chi tiết công nợ phải thu khách hàng** | TK 131 · **từ 01/01/2022** → thứ Sáu | Tuổi nợ từng hoá đơn, tiền thu trong tuần |
| 3 | Misa | **Tổng hợp công nợ phải trả nhà cung cấp** | TK 331 · thứ Bảy → thứ Sáu | Tab Payable (số tổng), tiền đã trả vendor |
| 4 | Misa | **Bán hàng** (danh sách chứng từ bán hàng) | **Từ 01/01 năm nay** → hôm nay · có cột *TT thanh toán* | Hoá đơn nào chưa thanh toán |
| 5 | Misa | **Sổ chi tiết tài khoản** | TK 131 · **từ 01/01 năm nay** → hôm nay · có cột *Mã đối tượng THCP* | Hoá đơn thuộc job nào; phiếu thu sau kỳ |
| 6 | SMS | **AR-AP** | Như mọi tuần | Tên Sales của từng job |

Rồi:

1. Thả cả 6 file vào `📥 input`.
2. Chờ tối đa 15 phút (tool chạy ẩn, không hiện cửa sổ). Muốn có ngay: mở **Task Scheduler** →
   *Bao cao cong no tuan* → **Run**.
3. Mở báo cáo ở thư mục gốc → tab **Review** → xử lý theo màu:

| Màu ở cột *Đối chiếu job* | Nghĩa là | FIN làm gì |
|---|---|---|
| 🟩 Khớp | Tổng job chưa thanh toán = dư nợ | Không cần làm gì |
| 🟥 ĐỎ | Job chưa thanh toán **nhiều hơn** dư nợ | Đã thu mà chưa đối trừ vào hoá đơn, hoặc thanh toán một phần Misa chưa trừ → đối trừ trên Misa |
| 🟨 VÀNG | Lệch nhỏ (≤ 2%) | Nghi quên đánh giá chênh lệch tỷ giá |
| 🟧 CAM | Dư nợ **nhiều hơn** job chưa thanh toán | Nợ trước năm nay → điền bảng job nợ cũ; hoặc hai file xuất khác lúc (cột ghi rõ) → xuất lại cùng lúc |
| ⬜ Khó đòi | Nợ khó đòi đã khai | Không cần làm gì |

Review còn liệt kê: khách Misa đã hết nợ mà hoá đơn vẫn ghi "chưa thanh toán" (ĐỎ — đối trừ trên
Misa), và tên khách trên file Bán hàng không tìm thấy trên Misa (TÍM — kiểm tra tên).

### 📖 Đọc file kết quả

| Sheet | Để làm gì |
|---|---|
| **Summary** | Trang bìa: các ô số chính, tuần này so với tuần trước, lịch sử các tuần |
| **Receivables** | Bảng công nợ theo mẫu chuẩn. Bấm **+** (hoặc nút **2** góc trái) để mở các job chưa thanh toán của khách (lấy từ Misa): hoá đơn nào, **hạn thanh toán ngày nào**, Sales nào, còn bao nhiêu. Cột cuối **Đối chiếu job**: Khớp / ĐỎ / VÀNG / CAM — chỗ FIN cần xem lại |
| **AR Risk** | Đầu kỳ, tăng, giảm, cuối kỳ; dấu hiệu S1–S7 (S7 = lỡ hẹn: đã qua ngày Sales hẹn mà chưa thu); nhóm rủi ro N0–N4 và việc cần làm của kế toán trưởng (N3: chỉ là câu nhắc, FIN tự quyết ngưng dịch vụ). Bên phải: **tiêu chí FIN** — số ngày quá hạn lớn nhất, nhóm khách, mức độ, căn cứ, hình thức thanh toán đề xuất, phương án xử lý, điểm và hạng A–D |
| **Đánh giá** | Bảng nâng hạng từng khách như file của Head FIN: dư nợ, quá hạn, billing 6 tháng, hạn mức (hợp đồng / tạm tính), T1–T4, KẾT QUẢ "NÂNG HẠNG RỦI RO", ghi chú. Tool tự điền, không nhập tay |
| **Tham số** | Số tool đã dùng tuần đó: lãi gộp, X1, X2, X, Y, N… Chỉ để xem — sửa ở sheet Tham số của file credit term |
| **Contracts** | Cảnh báo từ file credit term: hợp đồng hết hạn mà còn nợ (đỏ), sắp hết hạn ≤30/60/90 ngày, khách nợ chưa có credit term, term vừa đổi so với tuần trước. Hợp đồng hết hạn vẫn tính tuổi nợ theo term cũ. **Cuối sheet: khách mới** — tool đã tự thêm vào file credit term (file đang mở thì ghi "NOT yet", lần chạy sau thêm). Cột cuối **Mở hợp đồng** bấm ra file PDF |
| **Payable** | Theo mẫu của FIN: vendor còn nợ, chỉ số tổng (Current = Total), không phân tích |
| **Cash Flow** | Tiền khách trả trong tuần (ngân hàng, tiền mặt, cấn trừ, tỷ giá) và tiền trả vendor |
| **Invoices** | Từng hoá đơn còn nợ — lọc một khách, cộng lại đúng bằng số ở Receivables |
| **Methodology** | Từng con số tính thế nào, lấy từ file nào — để trả lời khi sếp hỏi |
| **Review** | Đối chiếu tổng, cảnh báo, những chỗ tool không chắc |

```
Tuổi nợ   = hạn của từng hoá đơn (ngày hoá đơn + credit term), quá hạn tính từ HÔM SAU ngày đến hạn
Tiền thu  = phiếu thu ghi số hoá đơn thì trừ đúng hoá đơn đó; còn lại trừ trước vào hoá đơn Misa đã ghi
            "Đã thanh toán", rồi mới tới hoá đơn cũ nhất. Chênh lệch tỷ giá khi đối trừ đi theo hoá đơn của
            phiếu thu cùng ngày
Khó đòi   = luôn ở nhóm 120+
Tiền      = luôn theo file Tổng hợp; file Chi tiết lệch thì ghi ở Review
Job       = mỗi hoá đơn còn nợ theo sổ; tiền = số CÒN NỢ của hoá đơn (đã trừ phiếu thu), hoá đơn nhiều
            job chia theo dòng hàng trên Sổ chi tiết. Các dòng job luôn cộng bằng Total.
Nhóm rủi ro = công thức của chị KTT: N3 khi có 2 dấu hiệu trong S1, S2, S3, S7 hoặc có nợ trên 120 ngày
Lỡ hẹn (S7) = lý do có ngày hẹn ('15/9', 'tháng 9' = cuối tháng 9) đã qua mà quá hạn không giảm
```

**Mức màu rủi ro (Head FIN, 06/11/2026)** — ô *Risk level* ở sheet **AR Risk** (FIN 12/11: không đưa sang Receivables). **Chỉ tô ô này**, không tô cả hàng:

| Mức | Màu ô | Mức gốc (theo tuổi nợ cũ nhất còn tiền) |
|---|---|---|
| 0 | ⬜ **Trắng** | Chưa quá hạn |
| 1 | 🟨 **Vàng** | Quá hạn 1–30 ngày |
| 2 | 🟧 **Cam** | Quá hạn 31–60 ngày |
| 3 | 🟥 **Đỏ** | Quá hạn 61–90 ngày |
| 4 | 🟪 **Tím** (chữ trắng) | Quá hạn từ 91 ngày |

```
Nâng +1 mức  = CHỈ khách ĐANG CÓ NỢ QUÁ HẠN, có ÍT NHẤT MỘT (không cộng dồn), trừ khách ✓ "Không nâng hạng màu":
   T1  nợ quá hạn ≥ X      X = số nhỏ hơn giữa X1 = 25% lãi gộp bình quân/tháng của quý gần nhất
                               và X2 = nợ quá hạn của khách xếp cuối trong top 15% khách đang quá hạn (làm tròn lên)
                               (chưa nhập lãi gộp → X = X2)
   T2  dư nợ > hạn mức     hạn mức = cột "Hạn mức nợ" trong file credit term; trống thì TẠM TÍNH
                               = billing bình quân/tháng 6 tháng gần nhất × hệ số term (30 ngày = 1, 45 ngày = 1,5…) × 1,3
   T3  ≥ 2 hoá đơn trả trễ hơn hạn quá 15 ngày trong 6 tháng (tool tự đếm từ Chi tiết 131)
   T4  khách mới (tuần trước chưa nợ) nợ ≥ Y (50 triệu)
Khách TRONG HẠN có T1–T4 = vẫn Trắng, ô ghi thêm dấu theo dõi, VD "Trắng ↑T3" (Head FIN 12/11)
Chưa có credit term / chưa ✓ hợp đồng = KHÔNG nâng màu; là việc cần bổ sung, liệt kê ở sheet Contracts
Lên thẳng Đỏ = lý do Sales ghi (tuần này, hoặc tuần trước nếu tuần này trống) có "mất liên lạc / hứa không trả /
               từ chối thanh toán…". Đã Tím thì giữ Tím.
```

Ở AR Risk, mỗi khách có đủ các cột để soát: Billing avg, Credit limit (Hợp đồng / Tạm tính), Late invoices, Base level,
T1–T4, Upgrade +1, Straight to Red, Final level. Đầu sheet ghi X, Y, N của tuần đó và cách tính X.

```
Số ngày quá hạn = của hoá đơn quá hạn LÂU NHẤT còn tiền (kể cả hoá đơn chỉ còn vài nghìn đồng)
Hạn hoá đơn     = ngày hoá đơn + credit term (file credit term)
Phương án xử lý = bảng của FIN theo nhóm khách × mức: Trắng, Vàng → mức 1 · Cam → mức 2 · Đỏ, Tím → mức 3
Nhóm khách      = Mới (hoá đơn đầu tiên trong 6 tháng gần nhất, hoặc chưa có hoá đơn) · Thường xuyên · Lâu năm
                  (có hoá đơn từ năm trước) → quyết định câu "Phương án xử lý" và "Hình thức thanh toán đề xuất"
Điểm A–D        = tham khảo, chấm trên hoá đơn từ 01/01 năm nay + hoá đơn năm trước còn nợ ở đầu năm:
                  số ngày quá hạn (40) + tỷ lệ thu hồi (30) + quá hạn / dư nợ (20) + tỷ lệ hoá đơn đã trả đủ (10)
                  A ≥ 80 · B ≥ 60 · C ≥ 40 · D dưới 40
Nhóm N0–N4      = luật của kế toán trưởng, giữ nguyên, chạy song song (cột Required action)
```

**Lý do trễ hạn (PA1, 27/10):** mỗi tuần mới cột *Reason for late payment* để **trống** cho Sales ghi lý do tuần
này; cột *Reason last week (read only)* (nền xám) hiện lý do tuần trước. Tool không tự viết chữ "[TỰ ĐỘNG]" nữa.
Cuối sheet Summary có bảng **Sales chưa giải trình** (khách quá hạn mà tuần này chưa có lý do).

### 📝 File credit term (`SAMPLE_credit_term_khach.xlsx`)

> **👩‍💼 FIN chỉ cần nhớ 3 điều**
> 1. **Khách mới tự vào file.** Có khách đang nợ mà chưa có trong file → tool tự thêm dòng ở cuối bảng, ghi chú
>    *"Tool thêm dd/mm/yyyy"*. Chị chỉ điền **Thời hạn** và **Hạn mức nợ** ở các ô vàng.
> 2. **Cột mới tự có.** Khi người bảo trì cập nhật tool, file tự được thêm cột ở lần chạy sau. Không chạy lệnh gì.
> 3. **Đóng file khi tool chạy** (khoảng giờ thả 6 file). File đang mở thì tool không ghi, chờ lần sau — báo cáo vẫn ra
>    bình thường, sheet Review ghi rõ.
> 4. **Mỗi tháng: sheet "Tham số"** — nhập lãi gộp 3 tháng của quý gần nhất (từ P&L) vào 3 ô vàng. Các ô khác (25%, top
>    15%, Y, 1,3, 2 hoá đơn, 15 ngày, 6 tháng) là số Head FIN đã chốt, chỉ sửa khi Head FIN đổi. Bảng **hệ số term**
>    bên dưới: gặp kiểu term mới thì thêm một dòng (term ghi đúng chữ như cột Thời hạn, hệ số = số tháng).
>
> Tool **không bao giờ sửa** ô chị hay Sales đã điền. Mỗi lần ghi, bản cũ được cất trong `_tool/luu-tru` — cần lấy
> lại thì mở thư mục đó, chọn file theo ngày giờ.

File `SAMPLE_credit_term_khach.xlsx` (sheet **Credit term**; file cũ tên sheet *Client Master* vẫn đọc được) là danh sách khách hàng gốc. *Client Master* từ nay là tên dành cho SharePoint List chứa hợp đồng PDF (sau này): một khách
một dòng. Sales và FIN mở chung trên OneDrive, sửa trực tiếp. Tool đọc file này mỗi lần chạy để tính
hạn từng hoá đơn và báo hợp đồng sắp hết hạn.

**Ai điền cột nào**

| Cột | Ai điền | Ghi chú |
|---|---|---|
| Mã kh | **FIN** | Mã khách **đúng như trên Misa** (VD `KH00418`, `INTERFREIGHT`). Sai mã là tool không nhận |
| Khách hàng, Sales phụ trách, PIC, Mã số thuế | Sales | Sales ghi tên đầy đủ như trên SMS; nhiều Sales thì cách nhau dấu phẩy |
| Tài khoản ngân hàng (4 cột) | Sales / FIN | Không bắt buộc |
| **Đã có hợp đồng** | Sales | Bấm mũi tên trong ô, chọn **✓** nếu khách đã ký hợp đồng; chưa có thì để trống |
| Số hợp đồng, Loại hợp đồng | Sales | Chọn trong danh sách hoặc gõ |
| **Link hợp đồng** | Sales | Dán link file hợp đồng (SharePoint / OneDrive) |
| **Thời hạn** | Sales điền, **FIN xác nhận** | Cột tool dùng để tính quá hạn — cách ghi ở bảng dưới |
| **Hạn mức nợ (VND)** | FIN | Số tiền tối đa khách được nợ (hợp đồng). Dư nợ vượt → T2. Trống thì tool tạm tính (xem trên) |
| **Không nâng hạng màu** | FIN | ✓ nếu khách tốt (giữ được hàng, đi hàng đều, hồ sơ tốt): không nâng màu theo T1–T4 |
| Ngày hiệu lực, Ngày hết hạn | Sales | Gõ ngày dạng `dd/mm/yyyy` |
| Ngày gia hạn | Sales | Ngày hợp đồng được gia hạn. **Tới ngày đó** file mới cộng 12 tháng ra hạn mới; chưa tới thì hạn vẫn là Ngày hết hạn |
| Cột nền xám | — | Tự tính (STT, hạn đang áp dụng, còn lại bao nhiêu ngày, trạng thái). **Đừng gõ đè** |

**Cách ghi cột Thời hạn** — tool đọc được các kiểu sau:

| Ghi | Tool hiểu là |
|---|---|
| `thanh toán ngay` · `trả trước` · `COD` | Đến hạn ngay ngày hoá đơn |
| `30 days` · `30 ngày` · `Net 30` | 30 ngày sau ngày hoá đơn |
| `14 ngày của tháng tiếp theo` | Ngày 14 của tháng sau tháng hoá đơn |
| `15 ngày kể từ ngày gửi bảng công nợ` | 15 ngày sau **ngày hoá đơn** (tool không biết ngày gửi bảng công nợ) |
| để trống | Tạm tính 30 ngày, và báo "Chưa có credit term" trong sheet Contracts |

**Màu trong file**

| Màu | Nghĩa là |
|---|---|
| 🟨 Ô vàng ở Mã kh / Thời hạn | Khách đã có tên nhưng còn thiếu — cần điền |
| 🟥 Mã kh nền đỏ | Mã bị trùng ở hai dòng — giữ một dòng |
| 🟥 Trạng thái "Đã hết hạn" | Hợp đồng hết hạn — Sales làm gia hạn |
| 🟨 Trạng thái "Sắp hết hạn ≤30 ngày" | Hết hạn trong 30 ngày tới |

**Việc thường gặp**

- *Link hợp đồng:* dán link SharePoint/OneDrive của file PDF vào cột **Link hợp đồng** → báo cáo tuần có ô
  **Mở hợp đồng** ở AR Risk và Contracts, bấm là mở trình duyệt.
- *Khách sắp có giao dịch (chưa nợ):* Sales có thể tự thêm một dòng ở cuối bảng trước; FIN điền Mã kh khi khách đã có
  trên Misa. Khách đang nợ mà chưa ai thêm thì tool tự thêm.
- *Ký gia hạn hợp đồng:* Sales điền **Ngày gia hạn**. Không sửa Ngày hết hạn cũ.
- *Đổi thời hạn thanh toán:* sửa cột Thời hạn. Tuần sau sheet Contracts báo "Credit term vừa đổi" để FIN
  soát lại với hợp đồng.
- *Hợp đồng hết hạn chưa gia hạn:* tool **vẫn tính theo thời hạn cũ**, chỉ cảnh báo.

**Không được:** đổi chữ ở dòng tiêu đề (dòng 6), nhất là hai cột **Mã kh** và **Thời hạn**. Chèn thêm cột
thì được (tool giữ cột đó khi ghi lại file). Màu ô, chú thích tự tô tay có thể mất khi tool ghi lại — bản cũ vẫn có
trong `_tool/luu-tru`. Sheet **Hướng dẫn** trong file có bản tóm tắt này và danh sách Sales để chọn.

**Mỗi tuần FIN xem sheet Contracts** trong báo cáo: hợp đồng hết hạn mà khách còn nợ (đỏ), sắp hết hạn,
khách đang nợ chưa có term, term vừa đổi. Sửa ở file credit term → tuần sau hết cảnh báo.

> Lệnh `python runner.py --credit-term "<tên khách>" --gia-tri "15 days"` vẫn chạy, nhưng file credit
> term luôn thắng. Nên sửa ở file để mọi người cùng thấy.

### 📊 Dashboard

Thư mục gốc workspace luôn có `Credit_Dashboard.html`. Mở bằng Chrome hoặc Edge, kéo file báo cáo
tuần vào trang. Mỗi tab là một trang riêng: Tổng quan · Khách hàng · Sales · Nợ xấu · Phải trả ·
Dòng tiền. Bấm một khách hoặc một Sales để vào trang chi tiết; các tab khác lọc theo luôn, bỏ lọc
bằng dấu ✕ trên thanh lọc, nút Back của trình duyệt quay lại được. Trang Tổng quan có ô **Hợp đồng &
credit term** (cảnh báo từ sheet Contracts). Trang khách và bảng Nợ xấu có nhãn **mức độ FIN**. **In / Lưu PDF** in đủ mọi trang
để gửi sếp (chọn "Save as PDF").
Không cần mạng, dữ liệu không rời khỏi máy.

### 🏷️ Khai những thứ file Misa không có — khai một lần, tool nhớ mãi

```bat
python runner.py --credit-term "GRAND FORWARDING LIMITED" --gia-tri "15 days"
python runner.py --ngoai-le "BSF LLC" --nhom "1 - 30" --vi-sao "lo tau chim, xuat lai DN"
python runner.py --bo-qua "OPTIMALOG LLC" --vi-sao "no ao do loi xuat hoa don tien coc"
python runner.py --nhan-lai "OPTIMALOG LLC"
```

## ❓ Hỏng thì làm gì

> 🛟 Tool không bao giờ chặn không ra báo cáo. Chỗ nào không chắc thì vẫn ra file và ghi ở *Review*.

| Hiện tượng | Nghĩa là | Làm gì |
|---|---|---|
| 🟥 File đầu vào bị đổi thành `[LOI] ...` | Dữ liệu không hợp lệ | Đọc file `[LOI] ... doc-vi-sao-hong.txt` cùng thư mục, sửa rồi bỏ chữ `[LOI]` |
| ⏳ Đã thả file mà chưa có báo cáo | Đang chờ file còn thiếu trong 6 file (tối đa 2 giờ) | Thả nốt file còn thiếu; nhật ký ghi tên file đang chờ |
| 📅 File Chi tiết 131 xuất tới cuối tháng, Tổng hợp tới thứ Sáu | Bình thường (FIN W41) | Không cần làm gì: tool tự bỏ phần sau ngày cuối kỳ, nhật ký ghi số bút toán đã bỏ |
| 🟧 Đối chiếu ghi *"hai file xuất khác lúc"* | Misa sửa hoá đơn giữa hai lần xuất | Xuất lại cả 6 file cùng lúc |
| 🪟 Lịch chạy vẫn bật cửa sổ đen | Máy cài trước 09/10 | Lần chạy kế tiếp tool tự sửa; nếu vẫn bật, chạy lại `1_Khoi_tao_workspace.bat` |
| ⚠️ Review ghi *"Chi tiết 131 và Tổng hợp 131 lệch"* | Hai file xuất khác thời điểm | Xuất lại cả hai cùng lúc |
| 😴 Không thấy file kết quả | Máy tắt, hoặc lịch chưa gọi | Task Scheduler → *Bao cao cong no tuan* → Run |
| 🔒 Nhật ký ghi *"Tạm hoãn kỳ này"* | File kết quả đang mở trong Excel | Đóng file, lần chạy sau tool làm tiếp |
| 📄 Review ghi *"Không thấy file credit term chung"* | OneDrive chưa đồng bộ file, hoặc file bị đổi tên/chuyển chỗ | Mở OneDrive cho đồng bộ; `python runner.py --xem-cau-hinh` để xem tool đang tìm ở đâu |
| 📋 Contracts báo *"Chưa có credit term"* / *"Chưa có trong file credit term"* với khách đã điền | Mã kh trong file credit term khác mã trên Misa | Sửa Mã kh cho đúng mã Misa |
| ⬜ Ô *Risk level* ở Receivables trống | File mở bằng trình xem không tính công thức (VD xem trước trên OneDrive web) | Mở bằng Excel; hoặc xem cột Risk level ở AR Risk / dashboard |
| 🟨 Nhiều khách quá hạn bị +1 mức | Chưa có Hạn mức nợ nên T2 dùng hạn mức tạm tính; T3 trả trễ nhiều lần | FIN điền Hạn mức nợ theo hợp đồng; khách tốt thì ✓ Không nâng hạng màu |
| ⬜ Ô ghi *"Trắng ↑T3"* | Khách đang trong hạn nhưng có dấu hiệu (T1–T4): không đổi màu, chỉ để theo dõi | Không cần làm gì; xem sheet Đánh giá cột Ghi chú |
| 📄 Đầu AR Risk ghi *"CHƯA NHẬP lãi gộp"* | Sheet Tham số trống lãi gộp | FIN nhập lãi gộp 3 tháng; tuần đó X = X2 |
| 🟥 Khách chưa quá hạn mà vẫn *Đỏ* | Lý do (tuần này, hoặc tuần trước nếu tuần này trống) có chữ "mất liên lạc / hứa không trả…" | Sales ghi lại lý do đúng tình hình tuần này |
| 📄 Review ghi *"Chưa nâng … lên mẫu mới: file đang mở"* | File credit term đang mở lúc tool chạy | Đóng file; lần chạy sau tool tự thêm cột (hoặc chạy lại tuần đó) |
| 📄 Contracts ghi *"NOT yet in the credit term file … đang mở"* | Lúc tool chạy, file credit term đang mở trong Excel | Đóng file; lần chạy sau tool thêm (hoặc chạy lại tuần đó) |
| 👤 Job ghi *"Không thấy job trên SMS"* | Job có trên Misa nhưng không có trong file SMS AR-AP | Tra mã job trên SMS: SMS ghi đã thu (lệch Misa), job nằm dưới tên khách khác (sửa file ghép khách), hoặc job làm ngoài SMS |

📜 Nhật ký: `_tool/runner.log` trong thư mục đích.

## 🛠️ Dành cho người bảo trì

```
weekly-debtor/            chỉ chứa mã nguồn, không có dữ liệu chạy
├── runner.py             điểm chạy, cũng là chỗ nhận mọi lệnh
├── wr/                   misa · chitiet · banhang · tindung · ghifull · ghep · khachhang · soghi · tinhtoan · baocao · kiemtra
├── tools/                make_template · setup_wizard · nang_cap_credit_term
├── tests/                87 phép thử đạt + 20 bỏ qua khi thiếu dữ liệu thật trong tests/fixtures/
├── dashboard/            src + build.py → Credit_Dashboard.html (tool chép vào workspace)
├── template/             mẫu chuẩn sheet Receivables
└── handoff/              RULES · STATE · MAP · phiếu đang mở · done/ · docs/ · shots/ · archive/
```

```bat
python -B -m pytest                    :: 87 đạt, 20 bỏ qua khi thiếu tests/fixtures/
python runner.py --tong-hop A.xlsx --chi-tiet B.xlsx --phai-tra C.xlsx --sms D.xlsx ^
                 --ban-hang E.xlsx --so-chi-tiet F.xlsx --ra <thư mục>
python dashboard/build.py              :: sửa dashboard/src thì dựng lại rồi commit cả file html
```

📐 Kiến trúc: `handoff/MAP.md` · Luật: `handoff/RULES.md` (sinh từ `handoff/handoff-standard-v2.0.md`) · Hiện trạng: `handoff/STATE.md` — đọc trước khi sửa

### 🆕 v0.9.1 (12/11/2026) — WR-15, sửa theo FIN sau W45

- W45: 62/62 khách bị +1, khách trong hạn thành Vàng (FIN hỏi). Head FIN chốt: khách **trong hạn giữ Trắng**, chỉ gắn
  "↑T…"; **chưa có credit term / chưa ✓ hợp đồng không nâng màu**, chuyển sang danh sách "Chưa ✓ hợp đồng" ở Contracts.
- W44 chạy lại: Trắng 16 (14 có ↑) · Vàng 1 · Cam 19 · Đỏ 5 · Tím 14; nâng hạng 34/55 (trước 54/55).
- FIN 12/11: **bỏ cột Risk level và 3 ô Sales đánh giá khỏi Receivables** (nhiều cột, trùng AR Risk). Màu chỉ ở AR Risk;
  lên thẳng Đỏ chỉ còn theo lý do Sales ghi.

### 🆕 v0.9 (06/11/2026) — WR-15, 5 mức màu của Head FIN

- **Ô Risk level 5 màu** Trắng · Vàng · Cam · Đỏ · Tím theo file tiêu chí nâng hạng v3 của Head FIN; **chỉ tô ô**, không
  tô cả hàng.
- **+1 mức theo T1–T4**: X từ lãi gộp + top 15%, hạn mức hợp đồng hoặc tạm tính từ billing 6 tháng, tool tự đếm hoá đơn
  trả trễ, chưa có hợp đồng / term / khách mới ≥ Y. Lên thẳng Đỏ theo lý do và 3 ô Sales.
- **Báo cáo tuần thêm 2 sheet "Đánh giá" và "Tham số"** như file của Head FIN, tool tự điền.
- Dashboard: biểu đồ cột hiện đủ tên khách / vendor (tên dài thì nằm trên thanh), chụp màn hình rõ tên.
- **Sheet "Tham số" trong file credit term** (tool tự thêm): lãi gộp FIN nhập mỗi tháng, các tỷ lệ, bảng hệ số term.
- Cột **Không nâng hạng màu** quay lại file credit term. Giữ 3 ô Sales, nhóm khách, phương án xử lý, điểm A–D (WR-13).
- Đổi tên repo thành `weekly-debtor`; thêm `3_Sua_lich_chay.bat` để trỏ lịch chạy về thư mục mới.

### 🆕 v0.8 (05/11/2026) — WR-13, tiêu chí FIN 3 mức

- **Bỏ 5 màu của bản CEO** (khó đọc) → **3 mức FIN**: Bình thường / Cảnh báo / Rủi ro cao, màu kiểu Excel, chỉ tô ô.
- Ô **Risk level** cạnh tên khách ở Receivables; **3 ô Sales đánh giá** (giữ giá trị tuần trước) nâng mức theo luật FIN.
- AR Risk thêm: số ngày quá hạn lớn nhất, nhóm khách, căn cứ, hình thức thanh toán đề xuất, phương án xử lý (nguyên văn
  bảng của FIN), điểm và hạng A–D. Giữ nhóm N0–N4 + Required action của kế toán trưởng.
- **Hạn thanh toán từng job** trong nhãn dòng con: `└ JOB · HĐ … ngày … · hạn dd/mm/yyyy`.
- Bỏ cột "Không nâng hạng màu" và ô tham số X / Y / N. Dashboard đọc cột theo tiêu đề (mở được file cũ lẫn mới).

### 🆕 v0.7.1 (27/10/2026) — WR-12

- **Khách mới tự vào file credit term** (FIN không phải dán): chỉ thêm dòng, sao lưu trước, file đang mở thì để lần sau.
- **File credit term tự lên mẫu mới** mỗi khi tool cần cột mới, kể cả tuần không có khách mới — không ai phải chạy lệnh.
- **Ô "Mở hợp đồng"** ở AR Risk và Contracts: bấm mở file hợp đồng PDF từ cột Link hợp đồng.
- **Khoá lời Sales**: ô lý do tuần trước chặn gõ; tuần đã có báo cáo tuần sau thì tool không đọc lại chỉnh sửa.
- Tên gọi: **file credit term** (sheet "Credit term"); "Client Master" dành cho SharePoint List sau này. Script đổi tên `tools/nang_cap_credit_term.py`.

### 🆕 v0.7 (27/10/2026) — WR-11, màu tuổi nợ của CEO + ý FIN

- **Màu ở AR Risk**: màu gốc theo tuổi nợ, nâng 1 mức theo T1–T4, lên thẳng Đỏ khi hứa không trả / mất liên lạc; tool tự đếm hoá đơn trả trễ > 15 ngày trong 6 tháng.
- **Client Master** thêm cột Link hợp đồng, Hạn mức nợ (VND), Không nâng hạng màu; ô ✓ chỉ nhận ✓ / x / có (gõ "Chưa" không còn bị coi là đã có hợp đồng).
- **Lý do tuần này / tuần trước (PA1)**: cột lý do để trống mỗi tuần mới; lý do tuần trước ở cột xám; bỏ chữ "[TỰ ĐỘNG]"; Summary liệt kê Sales chưa giải trình.
- **Khách mới chưa có trong Client Master**: bảng cuối sheet Contracts để FIN dán.
- Client Master chưa nâng cấp thì Summary để trống ô "chưa có hợp đồng" và Review cảnh báo (không đếm sai). Methodology viết lại cách trừ tiền thu.
- Không file Excel nào lên git (trừ mẫu chuẩn).

### 🆕 v0.6.1 (20/10/2026) — WR-09, sửa theo FIN sau tuần 42

- Tiền thu không ghi số hoá đơn trừ trước vào hoá đơn Misa đã ghi "Đã thanh toán" (STA CENTER: hết dòng HĐ 1560 ma).
- Chênh lệch tỷ giá khi đối trừ gộp vào hoá đơn của phiếu thu cùng ngày, không thành dòng riêng.
- Khách khó đòi có tên trong bảng job nợ cũ vẫn hiện dòng hoá đơn (Minh Khôi HĐ 2519).
- Tên công ty: TRUSTANA VIETNAM COMPANY LIMITED.
- Client Master thêm cột **Đã có hợp đồng** (✓); Summary đếm khách đang nợ chưa có hợp đồng. Ngày gia hạn chỉ tính khi đã tới.
- `python tools/nang_cap_credit_term.py "<file>"` (tên cũ: nang_cap_client_master): đưa file Client Master đang dùng lên mẫu mới, giữ nguyên dữ liệu.

### 🆕 v0.6 (13/10/2026) — WR-08

- **Tiền của từng job = số còn nợ theo sổ** (đã trừ phiếu thu); dòng job luôn cộng bằng Total. W41: 58/66 khách y hệt bản FIN sửa tay.
- **AR Risk dùng công thức của chị KTT** (S1–S6, nhóm, xu hướng, việc cần làm); TOTAL ở Receivables là công thức.
- **S7 lỡ hẹn**: lý do có ngày hẹn đã qua mà chưa thu → tính thêm một dấu hiệu (ATH W41 lên N3).
- **Cột Jobs by Salesman**: khách nhiều Sales thấy job nào của ai (câu hỏi của Customer Experience).
- **Client Master** (credit term + hợp đồng) dùng chung trên OneDrive; **sheet Contracts** cảnh báo hợp đồng hết hạn còn nợ, sắp hết hạn, chưa có term, term vừa đổi; ô tương ứng trên dashboard.
- FIN sửa thẳng trên file báo cáo là tool nhận bản đó; bỏ phiếu khách mới; dashboard hiện tên công ty gọn.

### 🆕 v0.5 (09/10/2026)

- **Job lấy từ Misa** (Bán hàng + Sổ chi tiết TK131) thay cho SMS: tiền theo tỷ giá hoá đơn nên khớp sổ; SMS chỉ còn cho tên Sales (WR-06).
- **Cột Đối chiếu job** tô màu + gợi ý nguyên nhân; Review gom theo màu cho FIN (WR-06).
- **Payable theo mẫu FIN**: chỉ số tổng, không phân tích, không tô màu.
- **Dashboard nhiều trang**: mỗi tab một trang, bấm khách / Sales vào trang chi tiết, bộ lọc đi theo, nút Back chạy (WR-07).
- **Chạy ẩn**: Task Scheduler không bật cửa sổ CMD nữa (WR-07).

<div align="center">

**TRUSTANA VIỆT NAM** · tool đầu tiên trong chuỗi credit control

</div>
