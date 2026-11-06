Mã: WR-12
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: WR-11
Người làm: claude
Người soát: codex
Trạng thái: chờ-soát
Vòng: 1/2

## ĐỀ BÀI

Kỳ vọng của user:
(2026-10-27, sau khi xem WR-11)
- "FIN mún tự update vào đúng chỗ lun, ko cần cô ấy can thiệp" (khách mới tự vào file credit term).
- "SAMPLE_credit_term_khach -> thêm 1 cột chứa url pdf -> chạy vào report cho CEO -> CEO đọc file report, thấy risk
  nợ -> bấm vào nhảy tới web browser chứa file hđ"
- "cột previous week sẽ bị khóa -> cột tuần này mở cho sales điền -> Tuần sau lại thế đôn lên đúng ko? Tránh sales
  sửa lại lời mình đã ghi"
- "Thống nhất ko gọi file đó là client master nữa, client master là sharepoint list chứa HĐ pdf trong đó" ·
  "file credit_term sample là đc rồi"
- "Ok thử đi cậu" (đồng ý cả gói 4 việc claude đề xuất).
- (28/10) "Làm lun đi rồi tui commit. Nhớ update README để FIN đọc hiểu" — file credit term tự thêm cột mới cả khi
  tuần đó không có khách mới.

Mục tiêu: Khách mới tự được thêm vào file credit term (an toàn: chỉ thêm dòng, sao lưu trước, file đang mở thì để lần
sau), link hợp đồng bấm mở được từ báo cáo, lời Sales đã ghi không sửa lại được sau khi tuần đó đã qua, và đổi tên gọi
file Excel thành "file credit term".

Ngoài phạm vi:
- Client Master trên SharePoint List (phiếu sau, khi có List thật).
- Khoá sheet bằng mật khẩu (hỏng nút gập/mở dòng job).

Kế thừa:
- WR-11: màu tuổi nợ, 3 cột mới trong file credit term, PA1 lý do tuần này / tuần trước.
- WR-08: file credit term là file chung OneDrive; FIN sửa lần cuối.

Kịch bản test (chạy lại W43 trên máy FIN; thay cho kịch bản của WR-11 — user hoặc FIN bấm, khoảng 20 phút):
1. Đóng mọi file Excel trong workspace. GitHub Desktop → repo weekly-debtor → Fetch origin → Pull origin.
   KHÔNG cần chạy script nâng cấp file credit term.
2. Lần chạy 1 — khách mới: vào `input\`, bỏ chữ `[DONE] ` ở 6 file `2026_W43_…`. Task Scheduler → Bao cao cong no tuan
   → Run. Chờ 1 phút. Thấy:
   a. `sample\SAMPLE_credit_term_khach.xlsx` giờ có sheet "Credit term"; cuối bảng có 5 khách mới (GENTHERM_VN,
      BOTT - SHENZHEN, BAIYUE, FTS-HAN, DUONG-TTNDVN), Ghi chú "Tool thêm 27/10/2026…", ô Thời hạn tô vàng. 83 dòng
      cũ giữ nguyên, 25 ô ✓ hợp đồng vẫn còn. Có 3 cột mới Link hợp đồng · Hạn mức nợ (VND) · Không nâng hạng màu.
   b. `_tool\luu-tru\` có file `SAMPLE_credit_term_khach_truoc-khi-tool-ghi_<ngày giờ>.xlsx` (bản trước khi tool ghi; file mẫu cũ + khách mới → 2 bản: trước khi nâng mẫu, trước khi thêm khách).
   c. Báo cáo `… (chay lai …).xlsx`: sheet Contracts cuối có "New customers — added to the credit term file (5)".
      Summary, AR Risk, Receivables như kịch bản WR-11 (tổng 14.324.678.762; cả hàng AR Risk có màu; cột lý do tuần
      trước nền xám; không còn chữ "[TỰ ĐỘNG]").
3. Lần chạy 2 — link + hạn mức: mở file credit term, dán link SharePoint của file PDF hợp đồng cho 1 khách đang nợ
   vào cột Link hợp đồng; điền Hạn mức nợ nhỏ hơn dư nợ cho 1 khách đang Vàng; tích ✓ Không nâng hạng màu cho 1 khách
   đang Cam. Lưu, ĐÓNG file. Bỏ `[DONE] ` ở 6 file W43 lần nữa → Run. Ở báo cáo `(chay lai …)` mới nhất:
   a. AR Risk (và Contracts nếu khách đó có cảnh báo): ô "Mở hợp đồng" chữ xanh gạch chân — bấm → trình duyệt mở PDF.
   b. Khách điền hạn mức: cột T2 = 1, màu lên 1 nấc. Khách tích ✓: Upgrade +1 = 0.
4. Khoá lời Sales: ở sheet Receivables, gõ thử vào một ô cột "Reason last week (read only)" → Excel báo "Chỉ để đọc",
   không nhận. Gõ vào cột Reason for late payment thì nhận bình thường.
Tuyệt đối không được xảy ra: dòng cũ trong file credit term bị đổi số / mất ✓; file W43 cũ bị sửa hay mất; tổng phải
thu khác bản cũ; ô lỗi #VALUE / #NAME; file Excel nào hiện trong danh sách thay đổi của GitHub Desktop.

## TODO
- [x] Khách đang nợ chưa có trong file credit term → tool tự thêm dòng (sao lưu trước; file đang mở thì bỏ qua lần này)
- [x] Link hợp đồng: ô "Mở hợp đồng" bấm được ở AR Risk và Contracts
- [x] Khoá lời Sales: ô "tuần trước" chặn gõ; báo cáo của tuần đã có tuần sau thì tool không đọc lại chỉnh sửa
- [x] Đổi tên gọi: "file credit term" (sheet "Credit term"), vẫn đọc file cũ có sheet "Client Master"
- [x] (28/10) File credit term mẫu cũ tự nâng lên mẫu mới ở đầu mỗi lần chạy, kể cả không có khách mới; README cho FIN
- [x] Phép thử, chạy lại W43 thật, README/RULES/MAP/STATE, kịch bản test

## AUDIT

Báo cáo của người làm:
File đổi: `wr/khachhang.py` (sheet "Credit term" + đọc tên cũ, `doc_file`, `tim_file_credit_term`, `dang_mo`,
`them_khach_moi`, số dòng mẫu tự dài theo dữ liệu, chữ hiển thị), `runner.py` (tự thêm khách sau khi tính; tuần đã chốt
không đọc lại; chữ hiển thị), `wr/ghifull.py` (cột "Contract (PDF)" ở AR Risk + Contracts, bảng khách mới ghi trạng thái
đã thêm / chưa, Methodology), `wr/baocao.py` (ô lý do tuần trước chặn gõ), `wr/tindung.py` (link_hd, them_khach),
`tools/nang_cap_client_master.py` → `tools/nang_cap_credit_term.py`, `dashboard/src/app.js` + html dựng lại,
`tests/test_wr8.py`, `tests/test_wr11.py` (chữ mới), `tests/test_wr12.py` (mới, 4 phép thử), README (v0.7.1), RULES, MAP, STATE.

```
$ python3 -B -m pytest -p no:cacheprovider          → 85 passed, 20 skipped in 55.46s   (trước WR-12: 81 passed)
$ node --check dashboard/src/app.js && python3 dashboard/build.py → ok · Đã dựng Credit_Dashboard.html (1.35 MB)

Làm hỏng thử (mục 10):
  tắt bước "tuần đã chốt" trong runner.nap_ban_tay → test_tuan_da_chot… ĐỎ ('Sales sửa lại lời cũ' == 'Khách hẹn trả 15/10')
  tắt kiểm tra file khoá ~$ trong khachhang.dang_mo → test_file_credit_term_dang_mo… ĐỎ ("file đang mở thì không đụng vào")
  trả lại mã → xanh

Chạy lại W43 thật (bản sao workspace trong máy ảo; file credit term CŨ sheet "Client Master", thêm tay một cột
Link hợp đồng cho HOÀNG KIM; không chạy script nâng cấp):
  runner → "Đã thêm 5 khách mới vào SAMPLE_credit_term_khach.xlsx: GENTHERM_VN, BOTT - SHENZHEN, BAIYUE, FTS-HAN,
            DUONG-TTNDVN. FIN điền Thời hạn, Hạn mức nợ ở các ô vàng."
           "-> … (chay lai 27-10 09h34).xlsx: 61 khách, phải thu 14,324,678,762, quá hạn 4,739,541,231" (y bản 24/10)
  _tool/luu-tru → SAMPLE_credit_term_khach_truoc-them-khach_20261027-093433.xlsx
  File credit term sau khi ghi → sheets ['Credit term', 'Hướng dẫn'] · 88 dòng (83 + 5) · ✓ hợp đồng 25 (giữ đủ) ·
                                   26 cột · 7 danh sách thả xuống · LibreOffice mở được, cột Trạng thái tính ra chữ
  AR Risk (LibreOffice tính lại) → 0 ô lỗi · Final level 0:3 · 1:13 · 2:21 · 3:6 · 4:18 (y WR-11)
  AR Risk ô link → HOÀNG KIM "Mở hợp đồng" → https://trustana.sharepoint.com/sites/FIN/HopDong/HOANG-KIM.pdf
  Contracts → "New customers — added to the credit term file (5)"
```

Lệch so với đề bài:
- "Khoá" cột lý do tuần trước là chặn gõ (Data Validation) + tool không đọc lại tuần đã chốt — không khoá bằng mật khẩu
  (khoá sheet làm hỏng nút gập/mở dòng job). Dán đè vẫn vào được ô, nhưng không ảnh hưởng gì tới báo cáo các tuần sau.
- Dashboard chưa có nút mở hợp đồng (CEO đọc file Excel; dashboard chỉ nhận thêm tên cảnh báo mới).
- Cờ `--client-master` và khoá `client_master_theo_may` trong settings.json giữ tên cũ để máy đã cài không phải khai lại.

Câu hỏi còn treo cho user: như WR-11 (X, Y, N; 5 hay 3 màu); xem STATE.

Giả định kỹ thuật:
[GIẢ ĐỊNH KỸ THUẬT] Thêm khách = đọc cả file credit term rồi ghi lại toàn bộ theo mẫu chuẩn (ghi_master) qua file tạm + os.replace, không sửa từng ô trên file cũ.
Vì sao chọn: openpyxl bỏ các phần mở rộng của Excel khi lưu đè (đã thấy cảnh báo "Data Validation extension … removed" với file FIN); ghi lại theo mẫu thì danh sách thả xuống, tô màu, công thức luôn đủ, và file cũ tự lên mẫu mới.
Sai thì làm lại: `khachhang.them_khach_moi`; mất định dạng tay FIN tự thêm (màu ô, chú thích) — bản gốc còn trong `_tool/luu-tru`.

[GIẢ ĐỊNH KỸ THUẬT] "File đang mở" chỉ phát hiện được trên chính máy chạy tool (file khoá `~$…` hoặc Windows không cho mở để ghi). Người khác đang sửa trên OneDrive ở máy khác thì không thấy.
Vì sao chọn: không có cách đọc trạng thái đồng soạn thảo của OneDrive từ Python thuần; tool ghi trong vài giây và chỉ khi có khách mới.
Sai thì làm lại: chuyển việc thêm khách sang Power Automate / SharePoint List (Client Master) khi có List.

[GIẢ ĐỊNH KỸ THUẬT] Tuần đã chốt = sổ ghi đã có trang của một ngày chốt muộn hơn.
Vì sao chọn: đúng nhịp "tuần sau đôn lên" user mô tả; chạy lại tuần hiện tại vẫn đọc được lý do Sales vừa ghi.
Sai thì làm lại: điều kiện trong `runner.nap_ban_tay`.

[RUNTIME UNVERIFIED] Chưa chạy trên Windows của FIN: Excel thật hiện thông báo "Chỉ để đọc", bấm ô "Mở hợp đồng" mở trình duyệt, file khoá `~$` khi FIN mở file credit term trên OneDrive.

Bổ sung 2026-10-28 (user: "Làm lun đi"):
File đổi: `wr/khachhang.py` (`thieu_mau`; `them_khach_moi` nâng mẫu khi không có khách mới; bản sao lưu đổi tên
`…_truoc-khi-tool-ghi_<ngày giờ>`), `runner.py` (nâng mẫu đầu lần chạy, TRƯỚC khi đọc file → cảnh báo thiếu cột không còn
sai), `tests/test_wr12.py` (+2 phép thử), README (khung "FIN chỉ cần nhớ 3 điều", mục cập nhật bản mới, bảng sự cố).
```
$ python3 -B -m pytest -p no:cacheprovider          → 87 passed, 20 skipped in 60.34s
Làm hỏng thử: tắt bước nâng mẫu đầu lần chạy → test_file_mau_cu_tu_nang_mau… ĐỎ ('Client Master' == 'Credit term'); trả lại → xanh
```
Tuần không có khách mới và file đã đúng mẫu: tool không ghi file, không tạo bản sao lưu (phép thử kiểm giờ sửa file).

Bổ sung 2026-10-29 — user chạy lại W43 trên máy FIN (file `handoff/docs/2026_W43_… (chay lai 29-10 14h01).xlsx`):
- Đạt: 5 khách mới tự vào file credit term + nâng mẫu (Review ghi "Đã nâng … thêm: Link hợp đồng, Hạn mức nợ (VND),
  Không nâng hạng màu, tên sheet 'Credit term'"), Contracts "New customers — added … (5)", phải thu 14.324.678.762,
  lý do tuần này 32 / tuần trước 29, "[TỰ ĐỘNG]" 0, Final level 0:3 · 1:13 · 2:21 · 3:6 · 4:18 (LibreOffice, 0 ô lỗi).
- HỎNG: Excel không tô màu hàng AR Risk. Nguyên nhân: nền của định dạng có điều kiện chỉ ghi fgColor; Excel đọc bgColor
  (đối chiếu styles.xml của bản v2 Head FIN làm bằng Excel: `<patternFill><bgColor rgb="FFC00000"/>`). Hàng Đỏ đậm còn
  chữ trắng trên nền trắng. Cùng lỗi ở tô màu file credit term (ô vàng, trạng thái đỏ) từ WR-08.
- Sửa: `thuonghieu.nen_cf` ghi cả fgColor + bgColor, dùng ở `ghifull` (AR Risk) và `khachhang` (file credit term); bảng
  màu theo đúng bản Head FIN (C6EFCE · FFEB9C · F8CBAD · FF7C80 · C00000). Phép thử mới `test_to_mau_co_dieu_kien_ghi_bgcolor_cho_excel`.
```
$ python3 -B -m pytest -p no:cacheprovider → 88 passed, 20 skipped
Làm hỏng thử: nen_cf chỉ ghi fgColor → test_to_mau_co_dieu_kien… ĐỎ ("mọi nền định dạng có điều kiện phải có bgColor"); trả lại → xanh
```
[RUNTIME UNVERIFIED] màu hiện trong Excel thật — chờ user chạy lại trên máy FIN.

Bổ sung 2026-10-29 (user: "Cột lý do tuần trước nằm bên trái, bên phải là reason tuần này"):
- Receivables: N = Reason last week (read only), O = Reason for late payment, P = Đối chiếu job (FIN) (`th.COT`,
  `th.COT_LY_DO_TRUOC_SO`, `th.COT_LY_DO_NAY`). AR Risk: "Reason last week" đứng trước "Reason for late payment";
  công thức lấy lý do đổi `$N:$N` → `$O:$O`. Dashboard tìm cột Receivables theo tiêu đề (file cũ vẫn đọc được).
- Hàng mức 4 "mất chữ" là chữ trắng (CF) trên nền không tô — cùng lỗi bgColor ở trên.
```
$ python3 -B -m pytest -p no:cacheprovider → 88 passed, 20 skipped (sửa test_wr3/4/6/8/11 tìm cột theo hằng số)
W43 thật (máy ảo): Receivables N/O/P = Reason last week · Reason for late payment · Đối chiếu job (FIN);
  AR Risk lý do tuần này có chữ 32 khách (LibreOffice); Final level 0:3 · 1:13 · 2:21 · 3:6 · 4:18; dxf có bgColor;
  dashboard 0 lỗi, KH00274 "Chưa có lý do tuần này · tuần trước: …"
```

Soát:
Kết quả:
-- claude - 2026-10-29 14:25 UTC+7

## HISTORY
### 2026-10-27
Failure: không
Báo lỗi giả: không
Lessons: muốn "khoá" một cột mà không phá tính năng khác của sheet thì khoá ở nguồn dữ liệu (tool không đọc lại tuần đã chốt), ô chỉ để báo người gõ.
-- claude - 2026-10-27 16:55 UTC+7
### 2026-10-29
Failure: màu ở AR Risk không hiện trong Excel trên máy FIN (CF chỉ có fgColor). Phép thử và LibreOffice đều không bắt được vì chỉ kiểm giá trị công thức, không kiểm cách Excel tô màu.
Báo lỗi giả: không
Lessons: kiểm "thứ thấy trên màn hình" phải đối chiếu với file do Excel thật tạo ra (hoặc chạy trên Excel); LibreOffice tính đúng không có nghĩa Excel hiện đúng.
-- claude - 2026-10-29 14:25 UTC+7
