Mã: WR-11
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: WR-09
Người làm: claude
Người soát: codex
Trạng thái: đóng-đạt
Vòng: 1/2

## ĐỀ BÀI

Kỳ vọng của user:
(2026-10-27, user chuyển ý CEO + Head FIN + FIN)
1. Ảnh 1 (CEO, Head FIN làm thử ở file `2026_W43_Credit_Report_17_10-23_10_2026_v2_mau_tuoi_no`):
   "Mark màu theo TUỔI NỢ -> có thể GIẢM xuống 3 MÀU cho đơn giản hơn -> đây sẽ là MÀU GỐC:
   Chưa quá hạn Xanh 0 · 1–30 Vàng 1 · 31–60 Cam 2 · 61–90 Đỏ 3 · 91–120 và trên 120 Đỏ đậm 4.
   Thay đổi màu gốc LÊN 1 NẤC dựa trên MỨC NÂNG HẠNG theo 1 trong các tiêu chí: Overdue Amount ≥ X ·
   Total Outstanding > credit limit · ≥ 2 hóa đơn trả trễ > 15 ngày trong 6 tháng · Chưa có hợp đồng, hoặc
   khách mới có nợ ≥ Y. NÂNG THẲNG: Hứa không trả / mất liên lạc → Đỏ. KHÔNG NÂNG HẠNG MÀU: Có khả năng
   giữ hàng, tần suất đi hàng đều đặn, dựa trên hồ sơ tổng quan tốt."
2. Ảnh 2: "Chốt link file hợp đồng sẽ nằm trong file Client Master · Credit term: 2 tiêu chí: Số ngày nợ
   và Hạn mức nợ · Thông báo sale nào chưa điền 'Lý do chưa thu hồi được công nợ' — PA 1: Thêm 2 cột lý do
   của tuần trước và tuần này. Khóa cột tuần trước, chỉ cho phép sửa cột tuần này; PA 2: Khóa luôn file
   cũ nếu quá thời gian điền lý do · Cột 'Lý do chưa thu hồi được công nợ' xóa toàn bộ chữ 'Tự động…'"
3. FIN (Ranee): "tuần này ví dụ có thêm các khách mới, không có trong file của tuần trước, vậy nó có thể
   tự chạy vào file khách hàng để cho em điền thông tin hạn nợ ko ạ"
4. User: "cậu chỉnh gitignore đừng đưa excel lên là đc. Toàn bộ excel trong repo" · sửa luôn các lỗi
   WR-09 đã tự soi · "cho tôi kịch bản mới chạy lại thử tuần 43 trên máy FIN với khách mới các các thay
   đổi mới, để xem file chạy lại có đc áp dụng toàn bộ thay đổi ko?"

Mục tiêu: Đưa luật màu tuổi nợ của CEO vào tool (dữ liệu khai một lần ở Client Master, phần đếm được thì
tool tự đếm), làm PA1 cho lý do của Sales, báo khách mới để FIN dán vào Client Master, sửa các lỗi WR-09,
và chặn mọi file Excel khỏi git.

Ngoài phạm vi:
- Tool không ghi vào Client Master (file chung OneDrive) — khách mới ra danh sách để FIN dán.
- Không khoá ô/sheet bằng mật khẩu (khoá sheet Receivables làm hỏng nút gập/mở dòng job).
- Dashboard chưa tô theo màu mới (chỉ đọc thêm lý do tuần trước để nhóm rủi ro khớp file Excel).
- Nối Client Master trên SharePoint (phiếu sau, khi có file thật).

Kế thừa:
- WR-08: AR Risk ghi công thức của Head of FIN; Client Master là file chung, tool chỉ đọc.
- WR-09: phân bổ tiền thu theo trạng thái Misa; cột "Đã có hợp đồng" (✓) trong Client Master.

Kịch bản test (chạy lại W43 trên máy FIN — user hoặc FIN bấm, khoảng 15 phút):
1. Đóng mọi file Excel trong workspace. Mở GitHub Desktop → repo weekly-report → Fetch origin → Pull origin.
2. Nâng cấp Client Master một lần: trong thư mục code, giữ Shift + chuột phải → Open in Terminal, gõ
   `python tools\nang_cap_client_master.py "<workspace>\sample\SAMPLE_credit_term_khach.xlsx"` → thấy dòng
   "83 khách; 25 đã ✓ hợp đồng; …" và file `SAMPLE_credit_term_khach_moi.xlsx` cạnh file cũ.
   Đổi tên file cũ thành `SAMPLE_credit_term_khach_cu.xlsx`, chuyển ra ngoài `sample\`; đổi tên file mới bỏ đuôi `_moi`.
3. Mở Client Master mới: thấy 3 cột mới Link hợp đồng · Hạn mức nợ (VND) · Không nâng hạng màu. Thử: điền Hạn mức
   nợ nhỏ hơn dư nợ cho 1 khách đang màu Vàng; tích ✓ "Không nâng hạng màu" cho 1 khách đang Cam. Lưu, đóng.
4. Vào `input\`, bỏ chữ `[DONE] ` ở đầu 6 file `2026_W43_…`. Task Scheduler → Bao cao cong no tuan → Run. Chờ 1 phút.
5. Ở thư mục gốc có file mới `2026_W43_Credit_Report_17.10-23.10.2026 (chay lai …).xlsx`. Mở bằng Excel, thấy:
   a. Summary: phải thu 14.324.678.762 · quá hạn 4.739.541.231 (y bản cũ). Dòng "Colour level 4 … 0" có số khách
      từng màu. Cuối sheet bảng "Reasons missing this week — by salesman".
   b. AR Risk: cả hàng tô màu (xanh lá / vàng / cam / đỏ / đỏ đậm). Bên phải: các cột T1–T4, Final level, Colour; ô
      vàng X, Y, N ở dòng 1–3. Khách điền hạn mức ở bước 3: T2 = 1, màu lên 1 nấc. Khách tích ✓: Upgrade +1 = 0.
   c. Receivables: cột Reason for late payment giữ lý do Sales đã ghi ở W43 (chạy lại cùng tuần); cột cuối nền xám
      "Reason last week (read only)" là lý do của W42. Không còn ô nào có chữ "[TỰ ĐỘNG]".
   d. Contracts: cuối sheet bảng "New customers to add to Client Master" — với dữ liệu W43 là 5 khách: GENTHERM_VN,
      BOTT - SHENZHEN, BAIYUE, FTS-HAN, DUONG-TTNDVN.
   e. Mở Credit_Dashboard.html, kéo file vào: không báo lỗi; trang một khách chưa ghi lý do tuần này hiện
      "Chưa có lý do tuần này · tuần trước: …".
Tuyệt đối không được xảy ra: file W43 cũ bị sửa hay mất; tổng phải thu khác bản cũ; ô lỗi #VALUE / #NAME ở AR Risk;
file Excel nào hiện trong danh sách thay đổi của GitHub Desktop.

## TODO
- [x] .gitignore chặn mọi *.xlsx (trừ mẫu chuẩn trong template/) và mọi thư mục workspace
- [x] Sửa lỗi WR-09: đọc ô ✓ chặt; Client Master thiếu cột → báo ở Review, không đếm sai; Methodology; script nâng cấp giữ cột lạ, không tự tích ✓
- [x] Client Master: thêm cột Link hợp đồng, Hạn mức nợ (VND), Không nâng hạng màu
- [x] Tool tự đếm hoá đơn trả trễ > 15 ngày trong 6 tháng
- [x] AR Risk: màu gốc + nâng hạng + nâng thẳng + không nâng hạng (công thức), tô màu hàng theo mức cuối
- [x] PA1: cột lý do tuần này để trống mỗi tuần mới, cột lý do tuần trước chỉ để đọc; Summary liệt kê Sales chưa giải trình
- [x] Bỏ chữ "[TỰ ĐỘNG]" khỏi cột lý do
- [x] Contracts: danh sách khách đang nợ chưa có trong Client Master, cột xếp đúng thứ tự để FIN dán
- [x] Dashboard đọc lý do tuần trước khi xét nhóm rủi ro
- [x] Phép thử mới; chạy lại W42 thật trong máy ảo, LibreOffice tính lại công thức
- [x] README, RULES (luật user yêu cầu), MAP, STATE; kịch bản chạy lại W43 trên máy FIN

## AUDIT

Báo cáo của người làm:
File đổi: `.gitignore`, `wr/chitiet.py` (đếm hoá đơn trả trễ), `wr/tindung.py` (PA1, màu, ngưỡng), `wr/tinhtoan.py`,
`wr/baocao.py` (cột lý do tuần trước, bỏ "[TỰ ĐỘNG]", đọc lại bỏ qua cột chỉ đọc), `wr/khachhang.py` (3 cột mới, đọc ô ✓
chặt, đọc số tiền, giữ cột lạ), `wr/ghifull.py` (cột + công thức màu, tô màu, Summary, Contracts, Methodology),
`wr/thuonghieu.py`, `runner.py` (đọc Client Master trước khi tính, chỉ số mới, cảnh báo Review), `tools/nang_cap_client_master.py`
(viết lại), `dashboard/src/app.js` + `dashboard/Credit_Dashboard.html` (dựng lại), `tests/test_wr4.py`, `tests/test_wr8.py`
(sửa theo luật mới), `tests/test_wr11.py` (mới, 9 phép thử), README, RULES (4 luật user yêu cầu, gộp dòng phần B), MAP, STATE.

```
$ git check-ignore -v template/Bang_cong_no_MAU_CHUAN.xlsx      → không bị chặn (mẫu chuẩn vẫn lên git)
$ git status --short -uall | grep -i xlsx                       → (trống) — không file Excel nào chờ commit
$ python3 -B -m pytest -p no:cacheprovider                      → 81 passed, 20 skipped in 28.11s   (trước: 72 passed)
$ node --check dashboard/src/app.js && python3 dashboard/build.py → jsok · Đã dựng Credit_Dashboard.html (1.35 MB)

Làm hỏng thử (mục 10):
  công thức "Upgrade +1" đổi <>"Yes" thành ="Yes"  → test_cong_thuc_mau_trong_excel_khop_python ĐỎ
     ("mức 0: công thức Excel = tool — assert 2 == 0"); trả lại → xanh
  công thức "Final level" đổi MIN(4 → MIN(3        → phép thử KHÔNG đỏ: dữ liệu giả không có khách mức 4.
     Trần mức 4 chỉ được phép thử Python (test_nang_hang_mot_muc_khong_cong_don) canh, công thức thì chưa.

Chạy lại W43 thật (bản sao workspace trong máy ảo, làm đúng bước 2 và 4 của kịch bản, chưa điền hạn mức):
  nang_cap_client_master → "83 khách; 25 đã ✓ hợp đồng; 80 có term; 0 ngày còn dạng chữ; giữ thêm cột: không có"
  runner → "Đã đọc lại báo cáo 2026_W43_Credit_Report_17.10-23.10.2026.xlsx: 61 khách, chốt 24/10/2026 (bản người làm)"
           "…v2_mau_tuoi_no.xlsx: cùng ngày chốt … — giữ bản người làm"
           "! Client Master chưa có Hạn mức nợ cho khách nào: màu chưa nâng theo tiêu chí 'dư nợ vượt hạn mức'."
           "-> … (chay lai 27-10 07h52).xlsx: 61 khách, phải thu 14,324,678,762, quá hạn 4,739,541,231" (y bản 24/10)
  LibreOffice tính lại → 0 ô lỗi công thức ở AR Risk
  Màu gốc so bản v2 Head FIN làm tay        → khớp 61/61 khách
  Final level (công thức) = Summary (Python) → 0:3 · 1:13 · 2:21 · 3:6 · 4:18
  T1 11 · T2 0 · T3 26 · T4 42 · lên thẳng Đỏ 3 (DLV, DTS, 131saoviet: "Công nợ khó đòi, mất liên lạc") · Upgrade +1 58/61
  Receivables: lý do tuần này 32 khách (giữ lý do Sales đã ghi W43) · lý do tuần trước 29 · ô "[TỰ ĐỘNG]" 0
  Contracts: 5 khách mới (GENTHERM_VN, BOTT - SHENZHEN, BAIYUE, FTS-HAN, DUONG-TTNDVN)
  Summary: Sales chưa giải trình 5 (Trần Văn Hiếu 8 khách 543tr …) · khách chưa có hợp đồng 42 · chưa có hạn mức 61
  Dashboard (Chromium headless, file W43 chạy lại) → 0 lỗi console; KH00274 hiện "Chưa có lý do tuần này · tuần trước: …"
```

Lệch so với đề bài:
- PA1 không khoá cột tuần trước bằng mật khẩu: khoá sheet Receivables làm hỏng nút gập/mở dòng job. Cột tuần trước
  nền xám, tool không đọc lại cột đó, nên Sales có sửa cũng không ảnh hưởng gì.
- Màu chưa có trên dashboard (dashboard chỉ đọc thêm lý do tuần trước để nhóm rủi ro khớp file Excel).

Câu hỏi còn treo cho user (kỳ vọng, đã đưa lên STATE):
- Với dữ liệu W43 thật, 58/61 khách bị nâng 1 mức: T4 vì 42 khách chưa ✓ hợp đồng, T3 vì 26 khách có từ 2 hoá đơn trả
  trễ trong 6 tháng (khách đi hàng nhiều, IDC 68 hoá đơn, STA 41). Như vậy màu gần như chỉ còn "gốc + 1". Head FIN cần
  chốt X, Y, N (N = 2 có thể quá thấp, hoặc đổi sang tỷ lệ hoá đơn trễ) và Sales tích ✓ cho khách đã ký.
- CEO ghi "có thể giảm xuống 3 màu": hiện giữ 5 mức như bản Head FIN làm thử.

Giả định kỹ thuật:
[GIẢ ĐỊNH KỸ THUẬT] "Hoá đơn trả trễ" = hoá đơn đã tất toán trong 6 tháng gần nhất, ngày tất toán trễ hơn hạn quá 15 ngày; ngày tất toán theo đúng cách tool phân bổ tiền thu (trạng thái Misa hiện tại dùng cho cả quá khứ).
Vì sao chọn: Misa không có cột ngày thanh toán của từng hoá đơn; dùng cùng luật phân bổ thì số khớp với tuổi nợ đang báo.
Sai thì làm lại: `chitiet.hoa_don_tra_tre` và dòng Methodology "Late invoices"; nếu FIN có nguồn ngày thanh toán thật thì đọc nguồn đó.

[GIẢ ĐỊNH KỸ THUẬT] Khách mới (T4) = tuần trước không có dư nợ trong báo cáo; khách không có dòng trong Client Master (khi file có cột ✓) tính là chưa có hợp đồng.
Vì sao chọn: giống bản Head FIN làm thử; khách chưa vào Client Master thì không thể đã ✓.
Sai thì làm lại: `tindung.danh_gia_mau` (T4) và `ghifull.cong_thuc_mau`.

[GIẢ ĐỊNH KỸ THUẬT] Lý do tuần này để trống thì nhóm rủi ro (X, N4), ngày hẹn (S7) và lên thẳng Đỏ đọc lý do tuần trước; S4 / CEO review chỉ xét lý do tuần này.
Vì sao chọn: không để khách "đang kiện" rơi khỏi N4 chỉ vì Sales chưa ghi lại tuần này; vẫn nhắc Sales ghi.
Sai thì làm lại: `DongRuiRo.ly_do_hieu_luc`, công thức `goc` trong `ghifull.cong_thuc_risk`, `tierOf` trong dashboard.

[GIẢ ĐỊNH KỸ THUẬT] X = 100tr, Y = 50tr, N = 2 lấy từ bản thử W43, đặt trong `tindung.py`; ô vàng trong file chỉ đổi màu của file đó.
Vì sao chọn: Head FIN chưa chốt; để ô sửa được cho Head FIN thử.
Sai thì làm lại: 3 hằng số trong `tindung.py` (hoặc chuyển thành cột cấu hình trong Client Master nếu FIN muốn tự đổi).

Ghi chú: lần chạy giả lập ghi thêm vài dòng vào `log/runner.log` của thư mục code (thư mục này đã bị .gitignore chặn).

Soát: user — đóng theo lệnh user 2026-11-05 ("các ticket đang mở cái nào xong thì close & move vào done"). Bằng chứng: Báo cáo W43 (chạy lại 29/10) và W44 (30/10) do Task Scheduler trên máy FIN sinh ra, dùng thật trong tuần (handoff/docs/Cong-No-Workspace (W44): runner.log, 2026_W44_Credit_Report_24.10-30.10.2026.xlsx — 55 khách, phải thu 13.099.776.819; Sales đã ghi lý do tuần này vào file). Không qua codex soát (lệch mục 0/9 của chuẩn, user quyết).
Kết quả: ĐẠT
Ghi chú: luật màu theo ý CEO (5 mức, T1–T4) sẽ thay bằng tiêu chí FIN ban hành 05/11 ở phiếu mới.
-- claude - 2026-10-27 15:05 UTC+7

## HISTORY
### 2026-10-27
Failure: phép thử mức màu gốc báo sai vì `tests/conftest.py` hạ ngưỡng số lẻ xuống 0,5đ cho dữ liệu giả — mã đúng, phép thử thiếu đặt lại ngưỡng thật. Phép thử "trần mức 4" của công thức không đỏ khi làm hỏng (dữ liệu giả không có khách mức 4).
Báo lỗi giả: không
Lessons: phép thử dựa vào ngưỡng tiền phải tự đặt ngưỡng nó cần; làm hỏng thử từng nhánh mới biết phép thử đối chiếu có phủ nhánh đó không.
-- claude - 2026-10-27 15:05 UTC+7
### 2026-11-05
Failure: không
Báo lỗi giả: không
Lessons: đóng phiếu theo lệnh user sau khi tool chạy thật W43–W44 trên máy FIN.
-- claude - 2026-11-05 10:05 UTC+7
