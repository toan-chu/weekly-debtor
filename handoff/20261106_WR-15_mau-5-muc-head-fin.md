Mã: WR-15
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: WR-13
Người làm: claude
Người soát: user
Trạng thái: chờ-soát
Vòng: 2/2

## ĐỀ BÀI

Kỳ vọng của user:
(2026-11-06) "Sửa màu lun. Nhưng mà theo chị Head FIN 2026_W43_Credit_Report_tieu_chi_nang_hang_rui_ro_v3 (1)" ·
"Ko tô màu cả hàng, tô màu vào 1 button mà ta đang làm với cái hiện tại W44 chạy lại" · "màu từ 3 lên 4-5 loại như head
fin đề xuất" · "tham số và đánh giá sheets cả tháng mới thêm 1 lần, vậy bỏ vào sample credit nhé" · "2 cái ko có trong 6
file input bỏ vào sample credit để FIN sửa tay mỗi tháng vậy. Vẫn tô màu theo 1 cột nhé, ko tô cả row, nhìn đau mắt lắm" ·
"Update README lun, nhân tiện đổi cả … = 📊 Báo cáo công nợ tuần — weekly debtor report nữa"

Mục tiêu: Ô Risk level (Receivables cột D + AR Risk) theo 5 mức màu của Head FIN (file v3): mức gốc theo tuổi nợ, +1 theo
T1–T4, lên thẳng Đỏ; chỉ tô ô. Tham số tháng (lãi gộp…) ở sheet "Tham số" của file credit term, FIN sửa tay.

Ngoài phạm vi: hợp đồng PDF / Client Master SharePoint List (WR-14). Phần WR-13 không đổi: 3 ô Sales, nhóm khách,
phương án xử lý, điểm A–D, hạn job.

Kế thừa: file nguồn `handoff/docs/2026_W43_Credit_Report_tieu_chi_nang_hang_rui_ro_v3 (1).xlsx` (sheet AR Risk cột AL–AX,
Tham_số, Đánh giá, Hướng_dẫn); code màu WR-11 (đếm hoá đơn trả trễ, từ khoá lên Đỏ, cột Không nâng hạng màu).

Giả định user chưa trả lời (06/11, hỏi 3 câu A/B, user bảo "sửa đi"):
- 5 mức đúng như v3 (Trắng tách khỏi Vàng); giữ T3 + "Không nâng hạng màu" như v3; màu đúng mã màu v3.
- Ô "Hứa không trả / mất liên lạc – nhập Có" của v3 = 3 ô Sales WR-13 (bất lợi, thất hứa) + từ khoá trong lý do.

Kịch bản test (máy FIN, ~15 phút):
1. Đóng mọi file Excel. GitHub Desktop → Fetch → Pull. (Nếu đã đổi tên thư mục: làm mục "Đổi tên thư mục" ở README trước.)
2. Bỏ `[DONE] ` ở 6 file W44 trong `input\` → Task Scheduler → Run. Chờ 1 phút.
3. Mở `sample\SAMPLE_credit_term_khach.xlsx`: có sheet **Tham số** (ô vàng, lãi gộp trống, bảng hệ số term); sheet Credit
   term có cột **Không nâng hạng màu**; dữ liệu cũ y nguyên. Nhập lãi gộp T7/T8/T9 (5.478.046.688 · 3.426.794.774 ·
   2.790.667.531), Quý gần nhất "Q3/2026". Lưu, ĐÓNG file.
4. Bỏ `[DONE] ` lần nữa → Run. Mở báo cáo `(chay lai …)` mới nhất:
   a. Summary: 55 khách, phải thu 13.099.776.819. Risk level: Trắng 1 · Vàng 15 · Cam 20 · Đỏ 5 · Tím 14; X = 227.689.408.
   b. Receivables: chỉ ô cột D tô màu (Tím chữ trắng), hàng không tô.
   c. AR Risk: dòng 1–3 phía trên cột T1 ghi X, Y, N và cách tính X (X1 = 974.625.749, X2 = 227.689.408).
   d. Chọn "Có" ở Adverse info? của một khách Vàng → ô Risk level đổi Đỏ ngay. Đóng KHÔNG lưu.
Tuyệt đối không được xảy ra: tổng phải thu khác; ô #VALUE/#NAME/#N/A; mất số trong file credit term; file Excel hiện
trong GitHub Desktop.

## TODO
- [x] Đọc file v3 của Head FIN (công thức AR Risk AL–AX, Tham_số, Đánh giá, Hướng_dẫn)
- [x] Luật 5 mức (`wr/danhgia.py`), T1–T4 trong `wr/tindung.py`, đếm trả trễ + billing (`wr/chitiet.py`)
- [x] Sheet Tham số + cột Không nâng hạng màu trong file credit term; tool tự thêm, giữ số FIN khi ghi lại
- [x] AR Risk: cột mới, công thức, X/Y/N đầu sheet; Receivables CF 5 màu chỉ ô D; dashboard
- [x] Phép thử 99 đạt; làm hỏng thử; chạy lại W44 thật trong máy ảo
- [x] README (đổi tiêu đề "weekly debtor report"), RULES, MAP, STATE
- [ ] User commit + chạy kịch bản trên máy FIN; Head FIN xem số khách bị +1 (54/55 tuần W44)

## AUDIT
Thay đổi: `wr/danhgia.py` (MUC5, THAM_SO_MAC_DINH, HE_SO_TERM_MAC_DINH, muc_goc5, nguong_x, he_so_term, nang_thang,
muc_cuoi5, muc3_tu_muc5) · `wr/tindung.py` (`danh_gia_mau5`, T2/T3 trong vòng lặp, X/Y sau vòng lặp; KetQuaFull.x/x1/x2/y/
tham_so) · `wr/chitiet.py` (`hoa_don_tra_tre` lấy lại từ WR-11, `billing_binh_quan`) · `wr/khachhang.py` (cột khong_nang,
sheet Tham số: `_ghi_tham_so`, `doc_tham_so_file`, `doc_tham_so`; `thieu_mau` báo thiếu sheet; `them_khach_moi` giữ tham số) ·
`wr/ghifull.py` (COT_RISK, `cong_thuc_fin`, `_can_cu_cong_thuc`, X/Y/N đầu sheet, CHI_SO muc_0..4 + nguong_x, Methodology) ·
`wr/thuonghieu.py` (MUC_FIN 5 màu) · `runner.py` (đọc tham số, cảnh báo) · `tools/nang_cap_credit_term.py` ·
`dashboard/src/app.js` + `app.css` · tests: `test_wr15.py` (10 mới), `test_wr13.py` (bỏ 3 phép thử 3 mức), `test_wr11.py`.

Bằng chứng:
```
$ python3 -B -m pytest -p no:cacheprovider --no-header
99 passed, 20 skipped
```
Chạy lại W44 thật (bản sao workspace W44 → ~/wk/ws44b). Lần 1: "Đã nâng SAMPLE_credit_term_khach.xlsx lên mẫu mới (thêm:
sheet 'Tham số')", cảnh báo chưa có lãi gộp. Nhập lãi gộp Q3 lấy từ file v3, lần 2:
```
-> … (chay lai 06-11 09h38).xlsx: 55 khách, phải thu 13,099,776,819, quá hạn 4,585,168,278
LibreOffice tính lại: lỗi 0 · Receivables Risk level == AR Risk: True (55)
Trắng 1 · Vàng 15 · Cam 20 · Đỏ 5 · Tím 14 (= Summary) · X = 227,689,408 (X1 974,625,749 · X2 227,689,408)
Base 0:16 1:20 2:5 3:3 4:11 · T1 6 · T2 36 · T3 22 · T4 38 · Upgrade +1: 54/55 · Straight to Red 3
Credit limit: Tạm tính 49 · Thiếu dữ liệu 6 · Hợp đồng 0
dashboard (Chromium headless): errors 0 · cột "Mức màu" · KH00418 / HOÀNG KIM / VINHGIANG = Tím
```
Phát hiện cần Head FIN xem: 54/55 khách bị +1 (T4 38 khách chưa ✓ hợp đồng; T2 36 khách vì hạn mức tạm tính = billing
6 tháng × hệ số × 1,3 thường nhỏ hơn dư nợ khi khách đang nợ quá hạn). File v3 của chị chưa nhập billing nên T2 chưa chạy,
vì vậy W43 của chị chỉ 10 Trắng / 26 Vàng.

Giả định kỹ thuật:
[GIẢ ĐỊNH KỸ THUẬT] Billing bình quân = tổng bút toán Nợ có số hoá đơn trên Chi tiết 131 trong 6 tháng tính lùi từ ngày cuối kỳ / 6.
Vì sao chọn: Head FIN ghi "billing TB/tháng 6 tháng gần nhất, từ Misa" mà không nói nguồn; Chi tiết 131 đã có trong 6 file và có từ 2022.
Sai thì làm lại: `chitiet.billing_binh_quan`.

[GIẢ ĐỊNH KỸ THUẬT] Term không có trong bảng hệ số thì hệ số = (hạn của hoá đơn ngày 15 − ngày hoá đơn) / 30.
Vì sao chọn: khớp cách Head FIN tính các dòng có sẵn (số ngày / 30); FIN thêm dòng vào bảng thì bảng thắng.
Sai thì làm lại: `danhgia.he_so_term`.

[GIẢ ĐỊNH KỸ THUẬT] Từ khoá lên thẳng Đỏ dùng danh sách WR-11 (mất liên lạc, hứa không trả, từ chối thanh toán…), KHÔNG dùng "không trả" / "không thanh toán" trơn như v3.
Vì sao chọn: W43 đã bắt nhầm "khách không trả kịp vì bank" (phép thử WR-11 giữ lại).
Sai thì làm lại: `danhgia.TU_KHOA_NANG_THANG` (Excel + dashboard dùng cùng danh sách).

Bổ sung 06/11 16:45 (user: "chạy lại có lun 2 sheet tham số và đánh giá đc ko"; "Kéo bảng ra, cho rõ tên"):
- Báo cáo thêm sheet "Đánh giá" (bố cục v3: hạn mức, T1–T4, KẾT QUẢ) và "Tham số" (bản chụp số đã dùng) — `ghifull.ghi_danh_gia`.
- Dashboard `hbar` / `stackBar`: tên dài thì đặt nguyên một dòng trên thanh (trước bị cắt "AIRSPACE TECH…").
```
W44 chạy lại: sheet Đánh giá 55 khách, NÂNG HẠNG 54 · Tạm tính 49 / Thiếu dữ liệu 6 · Tham số X1 974,625,749 · X 227,689,408
LibreOffice: lỗi 0 · pytest 99 passed, 20 skipped · dashboard headless: errors 0, tên đủ (ảnh dash_cf.png)
```

Vòng 2 (12/11) — FIN hỏi sau W45: "sao thêm 3 cột này", "các nhà này toàn công nợ trong hạn nhưng sao chuyển qua vàng hết".
W45: 62/62 khách +1; 23 khách trong hạn thành Vàng (T4 19 — 16 chỉ vì T4 —, T3 11, T2 5). Head FIN chốt (user chuyển):
khách trong hạn (A) giữ Trắng + ký hiệu ↑ · T4 thiếu term / chưa ✓ HĐ (A) chuyển sang danh sách bổ sung, không nâng màu.
Sửa: `tindung.danh_gia_mau5` (T4 chỉ khách mới ≥ Y; nang chỉ khi Q > 0,5; `theo_doi` "↑T…") · AR Risk thêm cột
"Watch (not overdue)", Risk level = màu & " ↑T…" · CF đổi sang LEFT(…) (Receivables, AR Risk, Đánh giá) · Contracts thêm
"Chưa ✓ hợp đồng" (`khachhang.LOAI_CHUA_HD`) · Đánh giá: "THEO DÕI (trong hạn)" + ghi chú bổ sung · Summary
`trang_theo_doi` · dashboard nhãn "Trắng ↑T3".
```
pytest: 100 passed, 20 skipped · làm hỏng (bỏ điều kiện Q > 0,5) → test_t2… FAILED
W44 chạy lại (máy ảo): LibreOffice lỗi 0 · Receivables == AR Risk 55/55 · Trắng 16 (↑ 14) · Vàng 1 · Cam 19 · Đỏ 5 · Tím 14
Contracts: Chưa ✓ hợp đồng 35 · Hết hạn 8 · Chưa có term 3 · nâng hạng 34/55 (trước 54/55)
```
FIN 12/11 15:04 (user chuyển, "Ta chắc bỏ theo ý em ấy"): "3 cái cột thừa … bỏ giúp" · "ô màu em chỉ cần thể hiện bên
sheet ar risk thôi … ko cho vào sheet receivables vì nó nhiều cột quá rồi … trùng dữ liệu ở 2 sheet". Đã bỏ cột Risk level
+ 3 ô Sales khỏi Receivables (`thuonghieu.COT` còn 15 cột) và khỏi AR Risk; `tinh_full` luôn dùng giá trị mặc định (giá trị
cũ trong sổ ghi bỏ qua); lên thẳng Đỏ chỉ theo từ khoá lý do. pytest 100 passed; W44 chạy lại: Receivables 16 cột, lỗi 0,
Trắng 16 · Vàng 1 · Cam 19 · Đỏ 5 · Tím 14; dashboard errors 0.

## HISTORY
### 2026-11-06 (vòng 1)
Mở và làm theo lệnh user. Chờ user commit + chạy kịch bản trên máy FIN; Head FIN xem tỷ lệ khách bị +1.
Failure: không
Báo lỗi giả: không
Lessons: `ws.cell(r, c, None)` của openpyxl không xoá ô (phép thử tự viết sai, đã sửa).
-- claude - 2026-11-06 16:40 UTC+7
