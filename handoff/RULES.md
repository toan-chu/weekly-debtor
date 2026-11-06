# RULES — weekly-debtor
Sinh từ handoff-standard v2.0 (bản sao: `handoff/handoff-standard-v2.0.md`). Chỉ chứa luật đứng; trần 200 dòng.
Phần A chép nguyên từ chuẩn; phần B là luật riêng của tool này. Hai phần mâu thuẫn thì phần A thắng.

# A. Luật chung (chép từ handoff-standard v2.0)

## A0. Phạm vi và thứ tự ưu tiên

- Một việc là việc DEV khi đủ cả hai điều kiện: chạm mã nguồn hoặc sản phẩm phần mềm, và kết quả sống lâu hơn cuộc chat. Mã viết dùng một lần không phải việc DEV và không theo tài liệu này.
- Đích của tài liệu: user không đọc mã vẫn biết tool có chạy đúng hướng mình muốn hay chưa; agent tự lo phần còn lại và tự chịu trách nhiệm việc mình làm.
- User là khách hàng: đưa hướng đi và kỳ vọng cuối cùng, không duyệt từng mốc.
- Tài liệu này không quy định cách viết mã. Công nghệ, thư viện, kiến trúc, cách tổ chức mã và quy trình kỹ thuật là quyền của agent và các plugin lập trình nó dùng; ngoại lệ duy nhất là mục 14, chỉ quy định chỗ đặt file và hai kho dùng chung của repo agentic.
- Claude và Codex ngang cấp: bên nào cũng được đề xuất cách làm, quyết kiến trúc, viết mã, chạy thử, đóng phiếu và soát việc của bên kia.
- Hai bên làm theo cùng một bản tài liệu này, nên việc một bên làm mà bên kia soát đạt là xong, không chờ user duyệt.
- Trong việc DEV, tài liệu này đè UP và chỉ dẫn của mọi plugin ở chỗ hai bên khác nhau.
- Plugin nén câu chữ chỉ áp cho chat; file trong `handoff/` và README viết câu đầy đủ, vì user không đọc mã vẫn phải hiểu được.
- Mọi agent phải đọc `handoff/RULES.md` và `handoff/STATE.md` trước khi làm bất cứ việc gì trong repo.
- Giao một việc DEV là lệnh làm: dựng phòng bàn giao, mở phiếu và ghi các file trong `handoff/` không cần hỏi lại.
- **Một việc, một phiếu**: cả vòng đời của một việc nằm trong đúng một file.
- **Phiếu tự đủ**: đọc một phiếu là hiểu được việc, không phải leo ngược chuỗi.
- **Ai ghi phần người đó**: không ai sửa mục của người khác; bất đồng thì ghi thêm, không ghi đè.
- **Một phiếu, một người làm**: không sửa file đang thuộc phiếu mở của bên kia; cần đụng thì mở phiếu mới hoặc ghi vào AUDIT phiếu đó.

## A3. Trạng thái và trần vòng lặp
- Trạng thái đi theo chiều `chờ-đề-bài` → `đang-làm` → `chờ-soát` → `đóng-đạt`. Soát HỎNG thì quay về `đang-làm` ở vòng 2; hỏng lần thứ hai thì sang `chờ-user`, rồi `đóng-huỷ` hoặc `đóng-hỏng`. Ngoài lần quay vòng đó, trạng thái không lùi, không nhảy cóc.
- Trần vòng lặp là 2: HỎNG hai lần nghĩa là hai bên chưa thống nhất, bắt buộc chuyển `chờ-user` kèm hai phương án theo mục 10, cấm vòng thứ ba.
- KHÔNG ĐO ĐƯỢC không tính là một vòng: phép thử chưa chạy được thì sửa phép thử trước, không sửa sản phẩm.
- Đóng rồi không mở lại; làm tiếp thì mở phiếu mới với `Nối tiếp` trỏ về. Ghi trạng thái đóng thì làm trong cùng một lượt: đổi `Trạng thái`, chuyển phiếu vào `done/`, và `STATE.md` thôi liệt kê nó trong việc đang mở.

## A9. Vai trò, trách nhiệm và bằng chứng

- **user**: khách hàng — đưa hướng đi và kỳ vọng, quyết khi hai agent không thống nhất, và là người duy nhất cho sửa `RULES.md`.
- **claude**, **codex**: ngang cấp — dịch kỳ vọng thành đề bài, đề xuất cách làm, viết mã, chạy thử, đóng phiếu, và soát việc của bên kia.
- Người làm ký phần báo cáo và chịu trách nhiệm việc đó; người soát ký kết quả soát; không ai sửa mục đã ký của bên kia.
- Chữ ký có dạng `-- <claude | codex> - <YYYY-MM-DD HH:MM> <múi giờ>`, theo giờ hệ thống của chính bên đó; không ai quy đổi hộ ai.
- Claude và Codex được tự mở một chat khác dưới nhân dạng khác để test, như người dùng cuối hay người vận hành của ngành; chat test không nối thư mục mã nguồn, vì giá trị của nó là nhìn sản phẩm như người ngoài, và kết quả ghi vào AUDIT dưới tên `claude-<vai>` hoặc `codex-<vai>`.
- **Thứ thấy trên màn hình** (trang có hiện không, bấm có ra kết quả không, số trên màn hình đúng chưa): kết luận được khi chạy thử qua trình duyệt hoặc ứng dụng thật trên máy của user, bắt buộc kèm ảnh trong `shots/`.
- **Thứ chỉ máy thật phán được** (phép thử đạt hay hỏng, toàn vẹn file, mã hoá ký tự, kết quả lệnh chạy): chỉ bên đã chạy lệnh trên chính máy đó và dán lệnh cùng kết quả mới được kết luận.
- Kết luận runtime rút ra từ việc đọc file qua lớp trung gian (máy ảo, thư mục đồng bộ, bản sao) phải gắn cờ `[RUNTIME UNVERIFIED]` cho tới khi có lệnh chạy trên máy thật, vì đọc qua lớp trung gian có thể thấy file hỏng trong khi file lành, và bên kia bị ép nhận lỗi sẽ bịa báo cáo "đã sửa xong".
- Tranh chấp giữa hai bên: bằng chứng thắng — lệnh và kết quả dán vào, ảnh trong `shots/`, dòng mã trích đúng — không phải lập luận.
- Mỗi bên được ghi tối đa một vòng bảo lưu ở AUDIT; vẫn chưa ngã ngũ thì chuyển `chờ-user` với hai phương án theo mục 10.

## A10. Làm việc với user

- User hỏi bằng câu cảm tính như "tool chạy ok chưa?", "xong chưa?", "có gì đáng lo không?"; agent trả lời thẳng "rồi" hoặc "chưa" bằng lời thường trước, rồi mới kèm bằng chứng: lệnh đã chạy, ảnh, con số.
- Chỉ trả lời "ok" khi kịch bản test đã chạy đạt và đã có người soát; còn chỗ KHÔNG ĐO ĐƯỢC thì nói rõ chỗ nào và vì sao.
- Agent tự có hướng test cho mọi việc: tự viết kịch bản, tự chạy, và tự làm hỏng thử một chỗ để chắc phép thử biết đỏ; phép thử không đỏ khi sản phẩm hỏng là đồ trang trí.
- Cần user test tay thì giải thích như cho một đứa trẻ: từng bước một, bấm vào đâu, gõ gì, sẽ thấy gì, và thấy khác đi thì chụp màn hình hay báo lại câu gì.
- Không hỏi user câu kỹ thuật; chỉ hỏi về kỳ vọng: ai dùng, để làm gì, đúng thì trông thế nào, điều gì không được xảy ra.
- Plugin hay quy trình nào bắt hỏi câu kỹ thuật thì agent tự quyết và ghi cờ giả định kỹ thuật.
- Đưa user chọn khi hai bên không thống nhất: đúng hai phương án, mỗi phương án nói bằng lời thường kèm hậu quả nếu chọn, và một khuyến nghị.
- Mỗi quyết định kỹ thuật agent tự đưa ra mà sai thì tốn công làm lại phải ghi vào AUDIT đúng ba dòng:

```
[GIẢ ĐỊNH KỸ THUẬT] <quyết định gì>
Vì sao chọn: <lý do>
Sai thì làm lại: <những gì phải làm lại>
```

## A11. Ngôn ngữ

- Tài liệu bàn giao, README, phiếu và câu trả lời cho user viết tiếng Việt có dấu; tên file, lệnh, mã và đường dẫn giữ nguyên.
- Tên trường, tên khoang, trạng thái và nhân dạng viết đúng từng chữ như tài liệu này, vì hai bên cùng ghi vào một phiếu và một chữ dịch khác là một chỗ gãy.
- Ba nhân dạng duy nhất là `user`, `claude`, `codex`; người đặt hướng đi luôn gọi là user, không dùng tên vai nào khác cho cùng một người.
- Từ và luật của bản cũ còn sót trong repo — như `chairman`, `Chỉ huy`, quyền ký theo vai — là lỗi chuẩn, sửa theo mục 12.

# B. Luật riêng của weekly-debtor

## R.1 Ranh giới với người dùng

- Thao tác hằng tuần duy nhất của kế toán là thả 6 file, xuất cùng lúc, vào `input/`: Tổng hợp 131, Chi tiết 131 từ 01/01/2022, Tổng hợp 331, SMS AR-AP (user chốt 2026-10-08, WR-04) + Bán hàng và Sổ chi tiết tài khoản 131, hai file xuất TỪ ĐẦU NĂM (FIN 09/10, WR-06). Không có nút "chạy ngay": ép chạy bằng Run trong Task Scheduler. Thao tác một lần: điền file ghép khách vào `sample/`, bỏ bản final tuần gần nhất vào thư mục gốc. Bố cục thư mục đích: gốc (báo cáo) + `input/` + `sample/` + `_tool/` ẩn — user chốt 06/10. Mọi thiết kế ép kế toán mở file ra sửa, điền, hay chạy lệnh đều là vi phạm.
- Tool không bao giờ chặn không ra báo cáo vì thiếu thông tin. Không chắc chỗ nào thì vẫn ra file, đánh dấu chỗ đó vào sheet "Cần xem lại".
- Tool không xoá và không sửa đè file nào của kế toán. File cũ chuyển sang thư mục lưu trữ, không biến mất. Ngoại lệ duy nhất (user 27–28/10, WR-12): tool THÊM dòng khách mới và THÊM cột của mẫu mới vào file credit term — chỉ thêm, không sửa ô đã có, sao lưu file vào `_tool/luu-tru` trước khi ghi, file đang mở thì không ghi.

## R.2 Số liệu

- Tuổi nợ (v0.4) tính từ từng hoá đơn trong file Chi tiết 131: hạn = ngày hoá đơn + credit term; quá hạn từ HÔM SAU ngày đến hạn (FIN 08/10). Phiếu thu ghi số hoá đơn thì trừ đúng hoá đơn đó, còn lại trừ hoá đơn cũ nhất trước. Cấm suy tuổi từ ngày ETD/ETA khi đã có file Chi tiết.
- Payable lấy toàn bộ số dư phải trả trên Misa, không lọc (FIN 08/10). Sheet Payable y chang mẫu FIN: No · Code · Name · Current · Total, chỉ số tổng, không phân tích, không tô màu (FIN 09/10).
- Job chưa thanh toán (v0.5) lấy từ Misa: Bán hàng (trạng thái) + Sổ chi tiết 131 (job). Chỉ "Đã thanh toán" mới tính là đã trả; "Chưa" và "Một phần" liệt kê đủ tiền hoá đơn (FIN 09/10). Trạng thái Misa là tại lúc xuất file: phiếu thu sau ngày cuối kỳ thì hoá đơn vẫn là nợ của kỳ.
- (v0.6, WR-08) Tiền dòng con = số CÒN NỢ của từng hoá đơn theo sổ (Chi tiết 131 đã trừ phiếu thu), job lấy từ Sổ chi tiết 131, hoá đơn trước năm nay lấy job từ bảng job nợ cũ. Dòng con luôn cộng bằng Total. Chứng từ không phải số hoá đơn (NVK) = "Điều chỉnh / chênh lệch tỷ giá". Cột Đối chiếu chỉ còn để FIN dọn trạng thái trên Misa.
- (WR-08) Sheet AR Risk ghi CÔNG THỨC của kế toán trưởng (S1–S6, Signals, Tier, Trend, Reason, Required action); dòng TOTAL ở Receivables = SUMIFS chỉ cộng dòng khách. Các sheet khác ghi số. Dashboard tự tính lại phần công thức. Đổi luật tính thì phải sửa Methodology cùng lần sửa.
- (WR-08) Credit term ưu tiên: file mẫu credit term trong `sample/` (nhận theo tiêu đề cột, dùng được nguyên file của kế toán trưởng) > khai `--credit-term` > tuần trước. Khách chưa có term liệt kê ở sheet Contracts. Không còn phiếu khách mới: Sales điền thẳng file credit term (13/10).
- (13/10) File credit term (`SAMPLE_credit_term_khach.xlsx`, trước gọi là Client Master) là file chung trên OneDrive. Trong `sample/` thì không cần cài gì; nằm chỗ khác thì mỗi máy khai đường dẫn (`--client-master`, lưu `client_master_theo_may` trong settings.json, không đi theo git). Có file chung → chỉ đọc file đó. Không thấy file → dùng `sample/` + cảnh báo ở Review.
- (20/10, FIN W42) Tiền thu không ghi số HĐ: trừ trước vào HĐ Misa (Bán hàng) ghi "Đã thanh toán" tại ngày chốt, rồi mới FIFO. Chênh lệch tỷ giá (NVK) cùng ngày với phiếu thu ghi số HĐ → cộng vào HĐ đó.
- (20/10) File credit term có cột "Đã có hợp đồng" (✓). Ngày gia hạn chỉ có hiệu lực khi đã tới ngày đó. Đổi mẫu thì nâng cấp file đang dùng bằng `tools/nang_cap_credit_term.py`, không dựng lại từ đầu.
- (13/10, user chọn A) S7 lỡ hẹn: lý do có ngày/tháng hẹn trả ('15/9', 'tháng 9' = cuối tháng, 'đầu tháng' = ngày 10, 'giữa tháng' = ngày 20) đã qua ngày chốt mà quá hạn không giảm so với tuần trước → tính như S1–S3 (2 dấu hiệu = N3). Cột Promise date tool điền từ lý do, FIN sửa tay được. Luật đọc ngày phải giống nhau ở `tindung.ngay_hen` và `ngayHen` của dashboard.
- (13/10, KTT qua user) N3 chỉ hiện câu "việc cần làm" trong báo cáo; FIN tự kiểm và quyết có ngưng dịch vụ hay không. Tool không khoá, không gửi gì.
- (13/10) File credit term (file chung OneDrive): FIN là người sửa lần cuối. Tool không có cột "xác nhận"; thay vào đó sheet Contracts báo "Credit term vừa đổi" mỗi khi term khác tuần trước để FIN soát.
- (13/10) Hợp đồng hết hạn chưa gia hạn: tuổi nợ vẫn tính theo term cũ. Tool chỉ cảnh báo (sheet Contracts + Summary + dashboard). Hạn đang áp dụng = ngày gia hạn + 12 tháng, không có thì ngày hết hạn.
- (06/11, Head FIN qua user, WR-15) Ô Risk level theo 5 mức màu của Head FIN (file W43 tiêu chí nâng hạng v3): mức gốc = nhóm tuổi cũ nhất còn tiền (0 Trắng chưa quá hạn · 1 Vàng 1–30 · 2 Cam 31–60 · 3 Đỏ 61–90 · 4 Tím ≥ 91); +1 (không cộng dồn) khi có T1 quá hạn ≥ X (X = MIN(25% lãi gộp BQ tháng quý gần nhất, nợ quá hạn khách cuối top 15% làm tròn lên)) · T2 dư nợ > hạn mức (hợp đồng, không có thì billing BQ 6 tháng × hệ số term × 1,3) · T3 ≥ 2 hoá đơn trả trễ > 15 ngày / 6 tháng · T4 chưa ✓ hợp đồng / chưa có term / khách mới nợ ≥ Y — trừ khách ✓ Không nâng hạng màu; lên thẳng ít nhất Đỏ khi lý do mất liên lạc… hoặc ô Sales bất lợi / thất hứa; quá hạn + không phản hồi → ít nhất Cam. CHỈ tô ô Risk level, không tô cả hàng (user 06/11: đau mắt). Tham số tháng ở sheet "Tham số" của file credit term, FIN sửa tay; tool không bao giờ ghi đè số FIN đã nhập. Ba chỗ cùng luật: `danhgia.py` · `ghifull.cong_thuc_fin` · `app.js`.
- (05/11, Head FIN + FIN qua user, WR-13 — mức 3 cấp ĐÃ THAY bởi WR-15; 3 ô Sales, nhóm khách, phương án, điểm A–D vẫn dùng) Mức rủi ro theo tiêu chí FIN, THAY luật màu 5 mức của CEO (họp: màu CEO khó đọc): 3 mức theo số ngày quá hạn của hoá đơn quá hạn lâu nhất còn tiền (kể cả còn vài nghìn đồng): Bình thường ≤ 30 · Cảnh báo 31–60 · Rủi ro cao > 60. Nâng mức theo 3 ô Sales nhập trong báo cáo (tuần mới giữ giá trị tuần trước): thông tin bất lợi → Rủi ro cao; đang quá hạn + thất hứa / trễ nhiều lần → Rủi ro cao; đang quá hạn + không phản hồi → ít nhất Cảnh báo. Màu kiểu Excel Good / Neutral / Bad (C6EFCE/006100 · FFEB9C/9C5700 · FFC7CE/9C0006), chỉ tô ô, không tô cả hàng. Giữ song song nhóm N0–N4 + Required action. Hạn job = ngày hoá đơn + credit term. TC1 / TC2 / TC3 (quá hạn ≥ X, vượt hạn mức, chưa HĐ hoặc khách mới ≥ Y) chờ Client Master (WR-14).
- (27/10, CEO qua user, WR-11 — ĐÃ THAY bởi WR-13 ngày 05/11, giữ để tra) Màu ở AR Risk: màu gốc = nhóm tuổi cũ nhất còn từ 1.000đ (0 chưa quá hạn · 1 (1–30) · 2 (31–60) · 3 (61–90) · 4 (trên 90)); nâng 1 mức (không cộng dồn) khi có ít nhất một: T1 quá hạn ≥ X · T2 dư nợ > Hạn mức nợ · T3 ≥ N hoá đơn trả trễ > 15 ngày trong 6 tháng (tool tự đếm) · T4 chưa ✓ hợp đồng hoặc khách mới nợ ≥ Y — trừ khách ✓ "Không nâng hạng màu"; hứa không trả / mất liên lạc → ít nhất Đỏ. Công thức trong Excel và `tindung.danh_gia_mau` phải cùng một luật. Risk Tier N0–N4 giữ cho việc cần làm, không tô màu hàng nữa.
- (27/10, user chọn PA1) Cột lý do tuần này để TRỐNG mỗi tuần mới; lý do tuần trước ở cột riêng chỉ để đọc; chạy lại cùng tuần thì giữ lý do tuần này. Tool không tự viết lý do (bỏ "[TỰ ĐỘNG]"). Không khoá sheet bằng mật khẩu (hỏng nút gập/mở dòng job).
- (27/10, FIN + user, WR-12) Khách đang nợ chưa có trong file credit term: tool tự thêm dòng (Mã kh, Khách hàng, Sales phụ trách, Ghi chú 'Tool thêm dd/mm/yyyy') để FIN điền Thời hạn, Hạn mức nợ; không thêm được thì báo ở Contracts + Review, báo cáo vẫn ra. Link hợp đồng (http/https) hiện thành ô 'Mở hợp đồng' ở AR Risk và Contracts. Tên gọi: 'file credit term'; 'Client Master' dành cho SharePoint List chứa hợp đồng PDF.
- (27/10, user) Không file Excel nào lên git, trừ `template/Bang_cong_no_MAU_CHUAN.xlsx`.
- (WR-08) Số dư dưới 1.000đ (tiền lẻ) không đưa vào báo cáo, liệt kê ở Review.
- (WR-08) Khách nhiều Sales: cột "Jobs by Salesman" tách job + tiền theo từng Sales trên SMS; job không thấy trên SMS ghi "Không thấy job trên SMS".

- Có file SMS: SMS chỉ quyết NGÀY từng job, không bao giờ quyết tiền. SMS ghi nhiều hơn Misa thì bỏ job cũ nhất và liệt kê cho FIN gạch paid (`handoff/docs/NOTE-ke-toan-tra-loi-20261006.md`).
- Total của một khách luôn lấy từ file **Tổng hợp công nợ phải thu**, cột số dư cuối kỳ bên Nợ. Khi file Tổng hợp và file Tuổi nợ mâu thuẫn, tiền theo Tổng hợp.
- Danh sách khách = khách có số dư cuối kỳ bên Nợ lớn hơn 0. Khách trả trước (số dư bên Có) không đưa vào báo cáo.
- Tổng các cột tuổi nợ của một dòng phải bằng Total của dòng đó. Không thoả thì dòng đó bị đánh dấu, không được im lặng cho qua.
- Tuổi nợ tính theo **số ngày thật đã trôi qua**, không tính theo số lần tool chạy. Nghỉ mấy tuần rồi chạy lại vẫn phải ra tuổi đúng.
- Các file đầu vào phải cùng kỳ: Chi tiết 131 và Tổng hợp 131 cùng ngày cuối kỳ (v0.4); Tuổi nợ = Tổng hợp + 1 ngày (v0.3). Lệch thì dừng và báo lỗi, cấm ra file.

## R.2b Những thứ không suy ra được từ file Misa

- Hợp đồng mới ký, biên bản bù trừ, khách trả theo tháng — file Misa không có. Kế toán khai một lần bằng lệnh `--credit-term` hoặc `--ngoai-le`, tool nhớ mãi. Cấm bắt kế toán khai lại mỗi tuần.
- Số job và Salesman chưa lấy được tự động thì mang nguyên của tuần trước sang. Cấm xoá trắng: kế toán sửa nhanh hơn là điền lại từ đầu.

## R.2c Đối chiếu

- Chỉ so báo cáo của tool với bản kế toán làm tay khi hai bên dùng CÙNG một cặp file Misa. Số trên Misa đổi theo thời điểm xuất, nên so hai lần xuất khác nhau là kết luận sai về tool.
- Mỗi báo cáo phải ghi lại tên và thời điểm sửa của hai file nguồn.

## R.3 Sổ ghi

- Mỗi lần chạy ghi một trang theo ngày chốt số: từng khách, số dư, cách chia nhóm tuổi, và ngày mà mỗi khoản xuất hiện lần đầu.
- Trang của một ngày chốt đã có thì ghi đè đúng trang đó, không tạo thêm bản.
- Khi kế toán có file làm tay cho một tuần, bản của con người thắng bản của máy. Bản người làm = file không có tab Review, HOẶC tên có chữ "human", HOẶC file tool ghi mà giờ sửa đã đổi (FIN mở ra sửa thẳng rồi lưu — sổ ghi nhớ giờ ghi trong `file_tool`). Cùng ngày chốt thì bản người thắng mọi bản tool, kể cả bản sửa sau. (27/10, WR-12) Tuần đã có báo cáo của tuần sau là tuần đã chốt: tool không đọc lại chỉnh sửa trên báo cáo của tuần đó, để Sales không sửa lại lời đã ghi.
- Mất sổ ghi thì chạy ở chế độ khởi động lạnh và đánh dấu, không được dừng.

## R.3a Không bao giờ để tool im lặng

- Mọi nhánh dừng sớm đều phải ghi một dòng vào nhật ký nói rõ vì sao. Cấm `SystemExit` hay `return` trơn: chạy nền thì không ai thấy màn hình.
- Cấm xoá `settings.json` của máy người dùng khi dọn dẹp. Mất file đó là tool đứng im mà không ai biết.

## R.3b Hai thư mục, không lẫn

- Thư mục mã nguồn chỉ chứa mã, mẫu chuẩn và phòng bàn giao. Cấm để sổ ghi, nhật ký, file đầu vào hay báo cáo ở đó, kể cả khi chạy thử.
- Thư mục đích do người dùng trỏ, tool tự tạo đủ các ngăn bên trong. Không ép người dùng phải đặt ở OneDrive.

## R.4 Hai máy, hai môi trường

- Cấm ghi đường dẫn tuyệt đối vào bất kỳ file nào đi theo dự án. Chỗ nào cần chỉ đường thì mô tả dấu hiệu nhận dạng, để lúc cài đặt tự dò.
- Cấu hình theo từng máy không đi theo git.

## R.5 Dashboard

- Nhãn trên biểu đồ dùng tên ngắn (bỏ "CÔNG TY TNHH / CỔ PHẦN / THƯƠNG MẠI / XUẤT NHẬP KHẨU…" ở đầu, "VIỆT NAM / LLC / LTD…" ở cuối); tooltip và bảng giữ tên đầy đủ (user 13/10).
- Sửa bất cứ file nào trong `dashboard/src` thì phải chạy `python dashboard/build.py` rồi commit cả `dashboard/Credit_Dashboard.html`. Máy FIN không build; nó chỉ nhận bản đã đóng gói qua git.
- Không bắt người dùng chép dashboard bằng tay: tool tự đặt bản mới nhất vào thư mục gốc workspace mỗi lần quét.
- Lịch chạy nền gọi thẳng `pythonw.exe runner.py`, không qua .bat (user 09/10: không bật cửa sổ CMD). Máy cài trước đó tự được runner đổi lịch một lần (`sua_lich_chay_an`).
- Luật màu (user 08/10): đỏ và các màu lân cận (cam, vàng cam) chỉ dùng cho điều TIÊU CỰC; xanh lá cho điều TÍCH CỰC; xanh dương, tím cho TRUNG TÍNH. Cấm dùng đỏ/cam để phân biệt Sales, khách hay vendor.
