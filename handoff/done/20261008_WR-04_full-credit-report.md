Mã: WR-04
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: WR-03
Người làm: claude
Người soát: codex
Trạng thái: đóng-đạt
Vòng: 1/2

## ĐỀ BÀI

Kỳ vọng của user: (lời user, 2026-10-07 → 08)
Ra full credit report trong một tool: máy FIN chạy ngầm, thả 4 file vào là ra một file
Excel lõi; chị kế toán trưởng kéo file đó vào dashboard (WR-05) để trình sếp. Sếp cần thấy
tiền thu từ khách (Receivables) và tiền trả vendor (Payable). Excel ít sheet hơn file của
chị KTT: chỉ giữ sheet lõi, có viện dẫn cách tính để trả lời khi sếp hỏi.

Luật nghiệp vụ user đã chốt (thuộc Kỳ vọng của user):
- 4 file mỗi tuần, xuất cùng ngày: Tổng hợp 131, Chi tiết 131 từ 01/01/2022 đến nay (FIN
  đồng ý 08/10, "để ensure không lọt"), Tổng hợp 331, SMS AR-AP.
- Tuổi nợ tính từ ngày đến hạn = ngày hoá đơn + credit term; quá hạn tính từ hôm sau ngày
  đến hạn (FIN 08/10).
- Payable = toàn bộ số dư phải trả trên Misa, không lọc (FIN 08/10: bản 114 vendor lệch vì
  hoá đơn nhập sau báo cáo, Misa đúng).
- Hoá đơn nhập lùi ngày làm dư đầu kỳ Misa khác dư cuối báo cáo tuần trước: tách thành
  dòng "nhập bổ sung kỳ trước", không trộn vào nợ mới.
- Luật rủi ro theo sheet "data W39 so W40" + "Tham chieu" của chị KTT (S1..S6, N0..N4,
  ngưỡng 2 dấu hiệu, ngưỡng CEO 100.000.000).
- Cột Notes → "Job Number" (bỏ chữ "job:"); "Customer's Name".
- Bỏ nút 3_Chay_ngay.bat (user 08/10: chạy bằng Task Scheduler).

Mục tiêu: Ra full credit report (phải thu, phải trả, dòng tiền, rủi ro, cách tính) trong một file Excel lõi, chạy ngầm trên máy FIN.

Ngoài phạm vi:
- Dashboard HTML (WR-05). Nợ vendor theo job/loại phí. Customer Master List SharePoint.

Kế thừa: (từ WR-03)
- Tiền theo Misa Tổng hợp; SMS chỉ cho job/salesman; bảng ghép khách FIN điền giữ nguyên.
- Khách khó đòi luôn 120+. Không ghi đè báo cáo đã có; bản người thắng bản máy.

Kịch bản test:
user 08/10: không viết kịch bản, làm xong test ngay trên máy user. Ghi nhận đây là
quyết định của user (RULES mục 4 câu 1 bị bỏ qua có chủ đích).

## TODO
- [x] Nhận dạng 4 loại file theo nội dung; chờ file thiếu 2 giờ
- [x] Đọc Chi tiết 131: hoá đơn, phiếu thu, TK đối ứng; tuổi nợ từ hoá đơn
- [x] Đọc Tổng hợp 331 → sheet Payable
- [x] Sheet AR Risk (luật chị KTT), Cash Flow, Summary + lịch sử tuần, Methodology
- [x] Review: đối chiếu Chi tiết ↔ Tổng hợp, ngày xuất lệch, nhập bổ sung kỳ trước
- [x] Cột Job Number, Customer's Name; đọc sheet Receivables(s) của bản người
- [x] Bản không có tab Review thắng bản tool khi cùng ngày chốt
- [x] Bỏ 3_Chay_ngay.bat; sample/ nhận file theo nội dung
- [x] Phép thử + chạy thật W40, so bản người

## AUDIT

Báo cáo của người làm:
**Báo cáo của người làm** — claude, 2026-10-08

Đã làm: đọc Chi tiết 131 (`wr/chitiet.py`), Tổng hợp 331 (`misa.doc_phai_tra`), tính full report
(`wr/tindung.py`), ghi 5 sheet mới (`wr/ghifull.py`), chạy cả vòng (`runner.chay_full`, quét
4 file trong `quet_thu_muc`). Đổi: `misa, soghi, baocao, thuonghieu, ghep, kiemtra, runner,
setup_wizard`, README, MAP, RULES. Bỏ `3_Chay_ngay.bat` (đã `git rm`, bản file nằm trong
`_to_delete/` vì máy ảo không có quyền xoá). Thêm `tests/test_wr4.py` (17 phép thử).

```
$ python -m pytest tests/
44 passed, 20 skipped        (20 cần dữ liệu thật trong handoff/docs/fixtures/)

$ runner.py --nap-bao-cao 3_DOI-CHIEU_W39_bao-cao-final_nguoi-lam.xlsx --ngay-chot 2026-09-26
$ runner.py --tong-hop 1_INPUT_..131 --chi-tiet 1_INPUT_..tu-2022 --phai-tra 1_INPUT_..331 --sms ..
  -> 2026_W40_Credit_Report_26.09-02.10.2026.xlsx: 70 khách, phải thu 14,251,295,535,
     quá hạn 4,780,211,318, nợ xấu 2,017,668,373, phải trả 7,663,887,492
So sheet Receivables với 3_DOI-CHIEU_W40_bao-cao_nguoi-lam.xlsx:
  70/70 khách, tổng 14,251,295,535 = 14,251,295,535
  tuổi nợ khớp từng đồng 61/70 khách, 97,8% số tiền đúng nhóm
  (WR-03 trên cùng dữ liệu: 38/70, 84%)
Summary "Last week": 11,416,674,587 phải thu, 4,968,004,735 quá hạn = sheet General của chị KTT.
Thu trong tuần 1,707,873,984 = Amount collected của chị KTT; ngân hàng 1,701,506,190.
Nhóm rủi ro so sheet "data W39 so W40": 58/70 khách cùng nhóm.
```

Làm hỏng thử: đổi luật "ngày đến hạn vẫn Current" thành "ngày đến hạn đã quá hạn" →
`test_tuoi_no_theo_hoa_don_va_luat_qua_han_tu_hom_sau` ĐỎ. Trả lại → xanh.

9 khách lệch tuổi nợ còn lại — đã soi từng hoá đơn, đều do bản người không theo luật FIN:
KH00149, RUSTAM (không có term, bản người tính lệch), KH00375, NOYTECH, ALLIANCE 21, STA
(bản người dồn cả khách vào một nhóm), KH00489 (bản người chia một hoá đơn 88,9tr ra 2 nhóm),
KH00274, KH00236 (lệch < 1tr do làm tròn nhóm). Không sửa tool cho khớp các dòng này.

12 khách lệch nhóm rủi ro: chủ yếu vì lý do trễ hạn TUẦN NÀY chị KTT mới gõ (tool dùng lý do
của tuần trước), phần còn lại theo 9 khách lệch tuổi nợ ở trên.

Lệch so với đề bài: không. Ngày tính tuổi = ngày cuối kỳ (02/10), ngày chốt in trên báo cáo
= hôm sau (03/10) như bản người — thử tính tuổi tại 03/10 thì khớp ít hơn (97,4%).

Câu hỏi còn treo cho user:
1. AR Risk tính theo lý do của tuần trước. Chị KTT gõ lý do mới vào file tuần này thì nhóm
   rủi ro chỉ đổi ở tuần sau — dashboard (WR-05) nên tự tính lại từ sheet. Đồng ý?
2. `[RUNTIME UNVERIFIED]` chạy trên Windows của user rồi máy FIN.

Kết luận cũ của người làm (chuẩn v1.1): chờ user chạy trên máy.

Giả định kỹ thuật: (phiếu mở trước v2.0, chưa ghi theo khuôn ba dòng)

Soát: user — đóng theo lệnh user 2026-11-05 ("các ticket đang mở cái nào xong thì close & move vào done"). Bằng chứng: Báo cáo W43 (chạy lại 29/10) và W44 (30/10) do Task Scheduler trên máy FIN sinh ra, dùng thật trong tuần (handoff/docs/Cong-No-Workspace (W44): runner.log, 2026_W44_Credit_Report_24.10-30.10.2026.xlsx — 55 khách, phải thu 13.099.776.819; Sales đã ghi lý do tuần này vào file). Không qua codex soát (lệch mục 0/9 của chuẩn, user quyết).
Kết quả: ĐẠT
-- claude - 2026-10-27 11:40 UTC+7

## HISTORY

### 2026-10-08
Failure    không
           Báo lỗi giả: không
Lessons    Đổi tên file đối chiếu trong handoff/docs/new theo nhóm 1_INPUT / 2_SAMPLE /
           3_DOI-CHIEU / 4_CU để user đọc được. .gitignore chặn mọi xlsx dưới handoff/docs.
           Tuổi nợ: thêm luật "phiếu thu ghi số hoá đơn thì trừ đúng hoá đơn đó" (UPP: phiếu thu
           23/09 ghi trả HĐ 1600..1643, chừa HĐ 1573 — bản người đúng, FIFO thuần sai). Số HĐ
           Misa đánh lại mỗi năm: chọn HĐ cùng số gần nhất trước ngày thu. Term "14 ngày của
           tháng tiếp theo" (IDC) = ngày 14 tháng sau.
           Bảng lịch sử ở Summary làm hẹp cột A — đặt lại độ rộng sau khi vẽ (soi bằng ảnh).
-- claude - 2026-10-08 UTC+7

### 2026-10-08 (chiều)
Failure    không
           Báo lỗi giả: không
Lessons    user chạy trên Windows (workspace thử trong tests/Cong-No-Workspace, log 14:49): ra
           đúng số như máy ảo; so với sheet Receivables bản đầy đủ của chị KTT: 61/70, 97,8%.
           Workspace thử nằm trong repo → thêm `Cong-No-Workspace/` vào .gitignore.
           CEO chốt Salesman dùng tên đầy đủ như SMS: bảng đổi tên ngắn ở thuonghieu.TEN_SALES_DAY_DU
           (Hiếu = Trần Văn Hiếu, Hiếu 1 = Phạm Trần Hiếu ...). "Admin,Admin" gộp thành "Admin".
           Lý do trễ hạn trống → tool viết "[TỰ ĐỘNG] <xu hướng>" (chữ xám nghiêng), theo cột
           "Reason for late payment" của chị KTT; đọc lại tuần sau thì bỏ qua câu này để S4 vẫn đúng.
           46 phép thử đạt.
-- claude - 2026-10-08 UTC+7
### 2026-11-05
Failure: không
Báo lỗi giả: không
Lessons: đóng phiếu theo lệnh user sau khi tool chạy thật W43–W44 trên máy FIN.
-- claude - 2026-11-05 10:05 UTC+7
