Mã: WR-10
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp:
Người làm: claude
Người soát: codex
Trạng thái: đóng-đạt
Vòng: 1/2

## ĐỀ BÀI

Kỳ vọng của user:
"Cậu làm lại repo theo handoff standard v2 đi." (2026-10-27)

Mục tiêu: Đưa phòng bàn giao của weekly-report từ chuẩn v1.1 lên handoff-standard v2.0 theo mục 12, không làm gãy việc chạy của tool.

Ngoài phạm vi:
- Không sửa lỗi nghiệp vụ của WR-09 (đã ghi ở AUDIT WR-09, làm ở phiếu riêng).
- Không đổi chữ in ra báo cáo hay dashboard cho FIN.

Kế thừa:

Kịch bản test:
1. Mở thư mục `handoff/` → thấy phẳng: RULES, STATE, MAP, bản sao `handoff-standard-v2.0.md`, các phiếu đang mở, và các ngăn `docs/`, `shots/`, `done/`, `archive/`. Không còn bản v1.1.
2. Mở `handoff/RULES.md` → dòng đầu ghi sinh từ v2.0; có đủ mục 0, 3 (trạng thái, trần vòng), 9, 10, 11; dưới 200 dòng.
3. Mở `handoff/STATE.md` → đúng khuôn mục 5, dưới 100 dòng; mở một phiếu bất kỳ → có trường Người làm, Người soát, Kỳ vọng của user, Mục tiêu, Kịch bản test, Soát, Kết quả.
4. Mở README → tiêu đề có phiên bản, huy hiệu phép thử ghi đúng số đã đạt, đủ 5 mục theo thứ tự chuẩn.
5. Chạy `python -m pytest` → số phép thử đạt không đổi so với trước khi chuyển.
Không được xảy ra: phép thử hỏng; file Excel dữ liệu thật bị đưa vào git; mất file dữ liệu thật; lệnh chạy ghi trong README hết chạy.

## TODO
- [x] Dựng `handoff/done/`, `handoff/archive/`; xoá `handoff/docs/handoff-standard-v1.1.md`
- [x] Làm phẳng `docs/`: `docs/w42/*` → `docs/REF-W42_*`; `docs/archive/` (dữ liệu thật cũ) → `handoff/archive/`
- [x] Dữ liệu thử thật: `handoff/docs/fixtures/` → `tests/fixtures/` (sửa `tests/test_wr.py`, `.gitignore`, README)
- [x] `.gitignore` chặn `handoff/archive/**/*.xlsx`
- [x] Sinh lại `RULES.md` (phần A chép chuẩn, phần B luật riêng còn đúng)
- [x] Viết lại `STATE.md` theo mục 5; bẫy sống quá một tuần chuyển thành đề xuất luật
- [x] AGENTS.md trỏ bản sao v2.0 trong repo; MAP.md cập nhật bước ② theo WR-09
- [x] Chuyển 9 phiếu WR-01 → WR-09 sang khuôn mục 3 và từ ngữ mục 11
- [x] Từ ngữ cũ trong `docs/` và chú thích mã (.py, `dashboard/src/app.js`), dựng lại `Credit_Dashboard.html`
- [x] README theo mục 8
- [x] Chạy phép thử, quét đường dẫn và từ ngữ cũ

## AUDIT

Báo cáo của người làm:
File đổi: `.gitignore`, `AGENTS.md`, `README.md`, `handoff/RULES.md`, `handoff/STATE.md`, `handoff/MAP.md`, 9 phiếu WR-01 → WR-09,
4 file `handoff/docs/NOTE-*`/`SPEC-*` (chữ Chairman → user), chú thích trong `runner.py`, `tests/test_wr.py`, `tests/test_wr8.py`,
`tools/setup_wizard.py`, `wr/caidat.py`, `wr/ghep.py`, `wr/ghifull.py`, `wr/khachhang.py`, `wr/tindung.py`, `dashboard/src/app.js`,
`dashboard/Credit_Dashboard.html` (dựng lại). Xoá: `handoff/docs/handoff-standard-v1.1.md`. Mới: phiếu này, `handoff/done/.gitkeep`,
`handoff/archive/.gitkeep`, `handoff/handoff-standard-v2.0.md` (đã có sẵn trên máy, chưa commit, trùng md5 với bản gốc ở thư mục `apps`).

```
$ python3 dashboard/build.py  (trước khi sửa app.js)   → git diff dashboard/Credit_Dashboard.html: không đổi — build tái lập được
$ python3 -B -m pytest -p no:cacheprovider -rs          → 72 passed, 20 skipped in 23.06s   (trước khi chuyển: 72 passed, 20 skipped)
  SKIPPED ... Không có dữ liệu thật trong tests/fixtures/ ...     ← đường dẫn mới đã có hiệu lực
$ git check-ignore -v
  .gitignore:15:handoff/archive/**/*.xlsx  handoff/archive/w39 final.xlsx
  .gitignore:14:handoff/docs/**/*.xlsx     handoff/docs/REF-W42_SAMPLE_credit_term_khach.xlsx
  .gitignore:11:tests/fixtures/            tests/fixtures/x.xlsx
$ wc -l AGENTS.md handoff/RULES.md handoff/STATE.md handoff/MAP.md → 18 · 195 · 41 · 113
$ grep đường dẫn cũ (handoff/docs/fixtures, handoff-standard-v1, docs/w42, docs/archive, docs/new, sample_w40)
  handoff/20261008_WR-04_full-credit-report.md:68    (dòng kết quả lệnh cũ trong AUDIT, giữ nguyên vì là bằng chứng đã dán)
  handoff/20261008_WR-04_full-credit-report.md:115   (HISTORY 08/10, giữ nguyên)
  handoff/20261009_WR-07_dashboard-nhieu-trang.md:65 (HISTORY 09/10, giữ nguyên)
$ grep từ ngữ cũ (chairman, logisticist, COWORK-EXEC, CHƯA KIỂM, Chỉ huy, Giao cho:, chờ-nghiệm-thu, nghiệm thu)
  wr/ghifull.py:502            chữ in ra sheet Methodology cho FIN — giữ, ngoài phạm vi
  handoff/RULES.md:65, STATE.md:32   chữ của chuẩn v2.0 / tên luật cũ bị thay — đúng ý
  5 chỗ "nghiệm thu" trong lời kể HISTORY/AUDIT cũ của WR-01, WR-02 và NOTE-ke-toan-tra-loi-20260924 — từ thường, không phải trường hay trạng thái
```

Luật bị thay khi sinh lại RULES (cũng ghi ở "Đang treo, cần user" của STATE):
- Mục 1 cũ: "Chỉ huy" `claude` [COWORK-EXEC] và giải trừ giới hạn → thay bằng A0/A9: claude và codex ngang cấp.
- Mục 2 cũ: quyền ký nghiệm thu chia theo vai → thay bằng A9: kết luận theo bằng chứng, `[RUNTIME UNVERIFIED]`.
- Mục 3 cũ: trần vòng lặp → nay A3 (cùng nội dung, thêm trạng thái `chờ-soát`, `chờ-user`).
- Mục 4 cũ: bốn câu hỏi kiểm tra của user → thay bằng A10.
Luật riêng (phần B) giữ nguyên nội dung, chỉ đổi từ Chairman → user và `[RUNTIME CHƯA KIỂM]` → `[RUNTIME UNVERIFIED]`.

Lệch so với đề bài:
- Không đóng phiếu nào: WR-01 → WR-09 đều do claude làm, cần codex soát mới đóng được. Trạng thái cũ `chờ-nghiệm-thu` → `chờ-soát`; WR-02 giữ `chờ-đề-bài`.
- Mục "Đang treo" cũ có việc trên máy thật (cài máy FIN, chạy WR-08) — không thuộc ba loại mục 5 cho phép, chuyển vào AUDIT WR-08.
- README giữ thêm mục 🛠️ "Dành cho người bảo trì" sau mục ❓ (chuẩn không cấm, chứa lệnh chạy phép thử và lịch sử phiên bản).

Câu hỏi còn treo cho user: xem "Đang treo, cần user" trong STATE (WR-02 còn cần không; luật ngày gia hạn; 4 đề xuất luật).

Giả định kỹ thuật:
[GIẢ ĐỊNH KỸ THUẬT] weekly-report không phải repo agentic (không có tác nhân AI chạy vòng lặp trong sản phẩm) → không tạo `memory/`, `CLAUDE.md`, không chép mục 14.
Vì sao chọn: sản phẩm là tiến trình Python chạy theo lịch, tính số theo luật cố định.
Sai thì làm lại: tạo `memory/`, `CLAUDE.md` (`@AGENTS.md`), chép mục 14 vào RULES, dựng ARSENAL/MEMORY.

[GIẢ ĐỊNH KỸ THUẬT] Dữ liệu thử thật chuyển từ `handoff/docs/fixtures/` sang `tests/fixtures/`, vì `docs/` phải phẳng và đây là dữ liệu của phép thử, không phải tài liệu.
Vì sao chọn: giữ `docs/` đúng mục 7, phép thử vẫn tự bỏ qua khi thiếu dữ liệu.
Sai thì làm lại: đổi `FIX` trong `tests/test_wr.py`, dòng `.gitignore`, câu trong README; máy nào đã đặt fixtures ở chỗ cũ thì chép sang chỗ mới.

[GIẢ ĐỊNH KỸ THUẬT] Phiếu cũ: trường ĐỀ BÀI còn thiếu được thêm kèm ghi chú "(phiếu mở trước v2.0, ...)"; HISTORY cũ giữ nguyên, không thêm chữ ký hay dòng "Báo lỗi giả" hồi tố.
Vì sao chọn: không viết hộ điều người làm khi đó chưa ghi, không sửa mục đã ký.
Sai thì làm lại: bổ sung tay từng phiếu.

Soát: user — đóng theo lệnh user 2026-11-05 ("các ticket đang mở cái nào xong thì close & move vào done"). Bằng chứng: Báo cáo W43 (chạy lại 29/10) và W44 (30/10) do Task Scheduler trên máy FIN sinh ra, dùng thật trong tuần (handoff/docs/Cong-No-Workspace (W44): runner.log, 2026_W44_Credit_Report_24.10-30.10.2026.xlsx — 55 khách, phải thu 13.099.776.819; Sales đã ghi lý do tuần này vào file). Không qua codex soát (lệch mục 0/9 của chuẩn, user quyết).
Kết quả: ĐẠT
-- claude - 2026-10-27 11:52 UTC+7

## HISTORY
### 2026-10-27
Failure: script chuyển phiếu lần đầu dùng `\s*` trong regex nên nuốt dấu xuống dòng, nối nhầm tiêu đề với dòng kế tiếp; phát hiện khi đọc lại WR-03, lấy lại 9 phiếu từ git (`git show HEAD:<file>`) rồi chạy lại bản sửa.
Báo lỗi giả: không
Lessons: chuyển hàng loạt file bằng script thì đọc lại ít nhất một file có cấu trúc khác nhau trước khi coi là xong.
-- claude - 2026-10-27 11:52 UTC+7
### 2026-11-05
Failure: không
Báo lỗi giả: không
Lessons: đóng phiếu theo lệnh user sau khi tool chạy thật W43–W44 trên máy FIN.
-- claude - 2026-11-05 10:05 UTC+7
