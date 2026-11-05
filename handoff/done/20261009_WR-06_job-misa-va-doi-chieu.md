Mã: WR-06
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: WR-05
Người làm: claude
Người soát: codex
Trạng thái: đóng-đạt
Vòng: 1/2

## ĐỀ BÀI

Kỳ vọng của user: (FIN + user, 2026-10-09)
Số tiền job chưa thanh toán lấy từ Misa, không lấy từ SMS (tỷ giá SMS khác tỷ giá xuất hoá đơn
nên không khớp tổng công nợ). FIN thêm 2 file xuất từ ĐẦU NĂM tới nay: Bán hàng (cột "TT thanh
toán") và Sổ chi tiết tài khoản 131 (cột "Mã đối tượng THCP" = số job). FIN không xuất từ 2022:
hoá đơn các năm trước trả khác loại tiền không gắn được phiếu thu, vẫn hiện "Chưa thanh toán".
Luật FIN: chỉ "Đã thanh toán" mới tính là đã trả; "Chưa thanh toán" và "Thanh toán một phần"
đều liệt kê. Cộng job chưa thanh toán, so với dư nợ của khách; lệch thì báo màu kèm gợi ý để FIN
xem lại (quên đánh giá chênh lệch tỷ giá, thanh toán một phần Misa không tự trừ...).
user: không ẩn số tiền job ở khách lệch — FIN tự kiểm, tự sửa.

Mục tiêu: Tiền job chưa thanh toán lấy từ Misa (Bán hàng + Sổ chi tiết 131), kèm cột đối chiếu tô màu để FIN tự soát.

Ngoài phạm vi:
- Dashboard nhiều trang + bấm xuyên (làm sau Excel, phiếu riêng).
- Nợ vendor theo job/loại phí; màu phải trả tăng — chờ FIN trả lời.

Kế thừa: (phiếu mở trước v2.0, không ghi; xem phiếu ở trường Nối tiếp)

Kịch bản test:
user: test ngay trên máy user (workspace trong tests/).

## TODO
- [x] Nhận dạng 2 file theo nội dung (ban-hang, so-chi-tiet); runner chờ đủ 6 file (tối đa 2 giờ)
- [x] Dòng con = mỗi job của hoá đơn chưa TT, nằm đúng cột tuổi theo hạn của chính hoá đơn;
      Sales của job lấy từ SMS; hoá đơn trước năm nay vẫn lấy job ở bảng job nợ cũ (sample/)
- [x] Hoá đơn Misa ghi "Đã thanh toán" nhưng phiếu thu nằm SAU ngày cuối kỳ → vẫn là nợ của kỳ
- [x] Cột 15 "Đối chiếu job (FIN)": Khớp / VÀNG / ĐỎ / CAM / Khó đòi, ghi số lệch + gợi ý
- [x] Review: danh sách theo màu; khách hết nợ mà hoá đơn vẫn chưa TT (ĐỎ); tên không ra khách (TÍM)
- [x] Invoices: thêm Misa status + Job (Misa). Methodology viết lại phần dòng con + đối chiếu
- [x] Nút 2: sheet "Nợ chưa có job" chỉ còn hoá đơn trước năm nay
- [x] Sửa file ghép khách (sample/ trong tests): TTND VIETNGA → KH00416; KH00361 = không có SMS

## AUDIT

Báo cáo của người làm:
```
$ python -m pytest -o addopts="" -q        55 passed, 20 skipped   (tests/test_wr6.py: 5 phép thử mới)
W40 thật (chạy trên máy ảo, bản sao sổ ghi):
  70 khách, phải thu 14.251.295.535 — dòng khách giữ nguyên số Misa
  Đối chiếu: 57 Khớp · 6 ĐỎ · 2 CAM · 5 Khó đòi
    ĐỎ lớn nhất: NOYTECH dư nợ 5,06tr, job chưa TT 1,96 tỷ (đã thu 27/05, 09/06, 23/06, chưa đối trừ)
    LCC UPP, CHAI KA: lệch đúng bằng phiếu thu 08/10 → đã xử lý "trả sau ngày cuối kỳ", nay Khớp
  Review: 18 khách Misa đã hết nợ nhưng hoá đơn vẫn "chưa TT" (TRANSTEAMLOGISTICS 8 HĐ 1,02 tỷ ...)
```
[RUNTIME UNVERIFIED] chạy bằng Task Scheduler trên Windows của user.

Giả định kỹ thuật: (phiếu mở trước v2.0, chưa ghi theo khuôn ba dòng)

Soát: user — đóng theo lệnh user 2026-11-05 ("các ticket đang mở cái nào xong thì close & move vào done"). Bằng chứng: Báo cáo W43 (chạy lại 29/10) và W44 (30/10) do Task Scheduler trên máy FIN sinh ra, dùng thật trong tuần (handoff/docs/Cong-No-Workspace (W44): runner.log, 2026_W44_Credit_Report_24.10-30.10.2026.xlsx — 55 khách, phải thu 13.099.776.819; Sales đã ghi lý do tuần này vào file). Không qua codex soát (lệch mục 0/9 của chuẩn, user quyết).
Kết quả: ĐẠT
-- claude - 2026-10-27 11:40 UTC+7

## HISTORY

### 2026-10-09
Dựng theo ghi chú FIN. Lần chạy đầu 48 Khớp / 16 CAM: phần lớn CAM là hoá đơn Misa đã ghi
"Đã thanh toán" vì phiếu thu 03–08/10 (file xuất 08/10, kỳ tới 02/10). Thêm luật trả-sau-kỳ
dựa vào phiếu thu nhắc số hoá đơn trong Sổ chi tiết → 57 Khớp.
Failure: ban đầu danh sách "Misa hết nợ" chỉ bắt khách không có trên Tổng hợp; khách có trên
Tổng hợp với dư 0 bị bỏ sót. Đã thêm.
Lessons: trạng thái thanh toán của Misa là tại lúc xuất file, không phải tại ngày cuối kỳ.

### 2026-10-09 (chiều) — FIN audit
FIN hỏi GOLDEN GLOBE: Tổng hợp Misa hiện 28tr mà báo cáo 31tr. Soát: file Tổng hợp 131 W40 ghi
31.552.382 (không phát sinh trong tuần); Chi tiết 131, Bán hàng, Sổ chi tiết đều ghi HĐ 00001799
(hạch toán 18/09, ngày HĐ 08/10) = 28.833.901 → Misa sửa hoá đơn sau khi xuất file Tổng hợp. Không
phải lỗi tính. Thêm gợi ý: tổng job khớp Chi tiết 131 mà chỉ Tổng hợp lệch → "xuất khác lúc".
Số tiền job = tiền hoá đơn (Bán hàng); hoá đơn nhiều job tách theo dòng hàng trên Sổ chi tiết.

### 2026-10-09 (tối) — W41 không chạy
FIN xuất Chi tiết 131 từ 01/01/2022 tới 31/10/2026, Tổng hợp tới 09/10. Tool đòi hai file cùng
ngày cuối kỳ nên bỏ cả kỳ, nhật ký chỉ ghi "chưa có file Tổng hợp đến 31/10". Sửa: Chi tiết xuất tới
SAU ngày cuối kỳ thì ghép với Tổng hợp gần nhất và cắt bỏ bút toán sau ngày đó (chạy & nút 2).
Một file Chi tiết dùng cho nhiều kỳ thì chỉ đổi tên [DONE] sau kỳ cuối. Phép thử mới. Chạy thật 6
file W41: ra báo cáo W41 (67 khách, 14.468.170.611). Failure: luật "cùng kỳ" quá chặt so với cách
FIN xuất. Lessons: đầu vào do người xuất tay thì nhận khoảng rộng hơn rồi tự cắt, đừng bắt khớp tuyệt đối.
### 2026-11-05
Failure: không
Báo lỗi giả: không
Lessons: đóng phiếu theo lệnh user sau khi tool chạy thật W43–W44 trên máy FIN.
-- claude - 2026-11-05 10:05 UTC+7
