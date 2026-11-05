Mã: WR-01
Mạch: WR - báo cáo công nợ phải thu hằng tuần
Nối tiếp: 
Người làm: claude
Người soát: codex
Trạng thái: đóng-đạt
Vòng: 1/2

## ĐỀ BÀI

Kỳ vọng của user:
Kế toán không còn phải làm báo cáo công nợ tuần bằng tay. Việc duy nhất còn lại
của con người là thả 2 file Excel xuất từ Misa vào một thư mục OneDrive. File
báo cáo tuần theo mẫu chuẩn tự xuất hiện, đúng số, đúng mẫu, đúng hạn.

Mục tiêu: Dựng tool đọc file Misa và tự ghi báo cáo công nợ tuần theo mẫu chuẩn, chạy nền theo lịch, kế toán chỉ việc thả file.

Ngoài phạm vi:
- Chỉ làm sheet Receivable. Các sheet Payable, General, Dashboard,
  Hard-to-collect giữ nguyên như mẫu, chưa đụng tới.
- Cột Salesman và số job chưa lấy tự động từ SMS; mang từ báo cáo tuần trước sang.
- Cột Lý do chưa thu hồi là việc của sales; tool chỉ mang từ tuần trước sang.
- Không đụng tới cách kế toán xuất file ra khỏi Misa.
- Gửi báo cáo qua Teams hoặc email thuộc phase sau.

Kế thừa:
- Phiếu đầu mạch.
- Luật điền báo cáo đã chốt với kế toán: `handoff/docs/SPEC-logic-bao-cao-tuan.md`.
- Kết quả chạy thử luật đó trên 3 tuần dữ liệu thật:
  `handoff/docs/NOTE-doi-chieu-3-tuan.md`.

Kịch bản test: — user viết ngày 2026-09-23

user đặt 2 file Misa của tuần W37 vào thư mục đầu vào trên máy user
(đường dẫn user trỏ lúc chạy):
- `W37_Tong_hop_cong_no_phai_thu_khach_hang.xlsx` — kỳ 05/09 đến 11/09/2026
- `W37_Phan_tich_cong_no_phai_thu_theo_tuoi_no.xlsx` — đến ngày 12/09/2026

user chạy tool Python. Kỳ vọng thấy, trong vòng 2 phút:

1. Thư mục kết quả xuất hiện đúng 1 file, tên `2026_W37_Bang_cong_no_tuan_05.09-11.09.2026.xlsx`.
2. Mở ra, sheet Receivable đúng mẫu chuẩn: 14 cột, tiêu đề cột tên nền cam
   `#ff6200`, tiêu đề cột nợ nền tím `#4d148c`, dòng ngày chốt ghi 12 Sep 2026.
3. Đủ 74 khách — đúng bằng số khách có số dư bên Nợ trong file Tổng hợp.
4. Cột Total của từng khách bằng số dư cuối kỳ bên Nợ trong file Tổng hợp.
5. Dòng TỔNG bằng tổng của file Tổng hợp.
6. Có sheet "Cần xem lại" liệt kê những dòng tool tự thấy không chắc.

Tuyệt đối không được xảy ra:

- Ra file khi hai file đầu vào lệch kỳ.
- Một dòng có tổng các cột tuổi nợ khác cột Total của chính dòng đó.
- Sửa, đổi tên hay xoá mất file đầu vào của user.
- Tiếng Việt vỡ chữ khi mở bằng Excel trên Windows.

Đối chiếu tham khảo, KHÔNG phải điều kiện đạt: so file máy sinh với bản kế toán
làm tay `2026_W37_Bang_cong_no_tuan_05.09-11.09.2026.xlsx`. Mọi dòng lệch phải
giải thích được thuộc dạng nào trong `handoff/docs/NOTE-doi-chieu-3-tuan.md`.

## TODO

- [x] Dựng lại mẫu chuẩn: chỉ còn sheet Receivable, trắng dữ liệu, kẻ màu thương hiệu
- [x] Đọc 2 file Misa, kiểm hai file cùng kỳ
- [x] Sổ ghi: trang theo ngày chốt, lô nợ kèm hạn, ngoại lệ
- [x] Luật chia 6 nhóm tuổi theo SPEC
- [x] Nạp trạng thái ban đầu từ báo cáo kế toán làm tay
- [x] Ghi file kết quả + sheet Cần xem lại
- [x] Phép thử tự động trên fixtures W36 và W37
- [x] Chạy kịch bản nghiệm thu W37, dán lệnh và kết quả vào AUDIT
- [x] Cho người dùng tự trỏ thư mục đích, tool tạo đủ các ngăn bên trong
- [ ] setup.bat: nháy đúp là cài thư viện + tạo lịch chạy nền  [RUNTIME UNVERIFIED]

## AUDIT

Báo cáo của người làm:
### Báo cáo của người làm — claude, 2026-09-23

**Đã làm gì.** Dựng tool từ đầu: đọc 2 file Misa, sổ ghi theo lô nợ có hạn,
luật chia 6 nhóm tuổi, ghi file theo mẫu chuẩn, wizard cài đặt, 13 phép thử.

**File đổi / thêm**

```
setup.bat  requirements.txt  settings.example.json  .gitignore  runner.py
wr/__init__.py  wr/thuonghieu.py  wr/caidat.py  wr/misa.py  wr/soghi.py
wr/tinhtoan.py  wr/baocao.py
tools/make_template.py  tools/setup_wizard.py
tests/test_wr.py  tests/so_sanh.py
template/Bang_cong_no_MAU_CHUAN.xlsx   (dựng lại: 1 sheet Receivable, trắng dữ liệu)
_thu-nghiem/                            (kết quả chạy thử, xoá được)
```

**Lệnh và kết quả**

1. Phép thử tự động

```
$ python3 -m pytest -q tests/test_wr.py
.............                                                            [100%]
13 passed, 6 warnings in 11.13s
```

2. Câu 3 của user — làm hỏng thử, phép thử có đỏ lên không

```
$ sed -i 's|total=round(total, 2),|total=round(total * 1.01, 2),|' wr/tinhtoan.py
$ python3 -m pytest -q tests/test_wr.py
3 failed, 10 passed          <- đỏ đúng chỗ, đã hoàn nguyên

$ sed -i 's|if tn.den_ngay != mong_doi:|if False:|' wr/misa.py     (bỏ chốt chặn lệch kỳ)
$ python3 -m pytest -q tests/test_wr.py
FAILED tests/test_wr.py::test_hai_file_lech_ky_thi_dung - Failed: DID NOT RAISE
1 failed, 12 passed          <- đỏ đúng chỗ, đã hoàn nguyên
```

3. Kịch bản test — thả 2 file W37 vào thư mục đầu vào rồi chạy tool

```
$ python3 runner.py --nap-bao-cao ".../2026_W36_....xlsx" --ngay-chot 2026-09-05
Đã nạp 73 khách từ 2026_W36_Bang_cong_no_tuan_29.08-04.09.2026.xlsx vào sổ ghi

$ cp W37_Tong_hop....xlsx W37_Phan_tich....xlsx  <workspace>/01_Input/
$ python3 runner.py
Bắt đầu: W37_Tong_hop_cong_no_phai_thu_khach_hang.xlsx + W37_Phan_tich_cong_no_phai_thu_theo_tuoi_no.xlsx
  -> 2026_W37_Bang_cong_no_tuan_05.09-11.09.2026.xlsx: 74 khách, tổng 9,547,752,597, 45 dòng cần xem lại

$ ls <workspace>/01_Input
[DONE] W37_Phan_tich_cong_no_phai_thu_theo_tuoi_no.xlsx
[DONE] W37_Tong_hop_cong_no_phai_thu_khach_hang.xlsx
```

Đối chiếu từng điểm của kịch bản: (1) đúng 1 file, đúng tên. (2) sheet
Receivable 14 cột, tiêu đề cột tên nền `#ff6200`, cột nợ nền `#4d148c`, ô A3
ghi "At 12 Sep 2026". (3) 74 khách, bằng đúng số khách có dư Nợ trong file
Tổng hợp. (4) và (5) phép thử `test_chay_w37_ra_dung_file_va_dung_total` kiểm
từng dòng: Total bằng file Tổng hợp, tổng 6 nhóm bằng Total, tổng bảng
9.547.752.597 bằng tổng Misa. (6) có sheet "Cần xem lại", 45 dòng.

4. Thả nhầm 2 file lệch kỳ — tuyệt đối không được ra file

```
$ cp W37_Tong_hop....xlsx W38_Phan_tich....xlsx  <workspace>/01_Input/
$ python3 runner.py
  LỖI: Hai file lệch kỳ, dừng lại để khỏi ra số sai.
  Tổng hợp: 05/09/2026 đến 11/09/2026
  Tuổi nợ : đến ngày 19/09/2026 (đáng ra phải là 12/09/2026)
$ ls <workspace>/01_Input
[LOI] 20260923 - doc-vi-sao-hong.txt
[LOI] W37_Tong_hop_cong_no_phai_thu_khach_hang.xlsx
[LOI] W38_Phan_tich_cong_no_phai_thu_theo_tuoi_no.xlsx
```

Không có file nào sinh ra trong 02_Output. File đầu vào chỉ bị đổi tên, nội
dung giữ nguyên.

5. Đối chiếu tham khảo với bản kế toán làm tay W37 (không phải điều kiện đạt)

```
$ python3 tests/so_sanh.py <máy sinh> <bản tay W37>
máy 74 dòng | tay 74 dòng
chỉ có ở máy: []   chỉ có ở tay: []
khớp cả 6 nhóm: 57/74 dòng
khớp khi gộp 61+ (so công bằng với bản cũ 4 cột): 62/74 | giá trị 7,59/9,40 tỷ = 81%
```

12 dòng lệch, giải thích được hết:
- 4 dòng lệch Total (CHU KONG, JSC RIMA, DEALLOG, Critical Solutions): kế toán
  tự chỉnh tay, máy theo luật đã chốt là lấy số của file Tổng hợp.
- 2 dòng máy theo file Tuổi nợ còn bản tay để Current (OPTEC, GRAND) — đúng
  dạng 3 trong NOTE-doi-chieu-3-tuan.md, đang chờ kế toán xác nhận.
- 6 dòng nợ chưa có hạn tới lúc chuyển nhóm mà máy chưa biết (CHAI KA, BSF,
  OPTIMALOG, KIARA, AVIALOGISTIKA, ACT 247). Đây là hệ quả của nạp trạng thái
  ban đầu: lần nạp đầu chỉ có bảng đã chia sẵn nhóm, không có ngày. Từ tuần thứ
  hai trở đi sổ ghi có ngày thật nên dạng này tự hết. Cả 6 đều nằm trong sheet
  "Cần xem lại".

6. Chạy tiếp W38 sau khi khai 2 điều kế toán chốt (chạy nối từ trang W37)

```
$ python3 runner.py --ngoai-le "AHC LOGISTICS(XIAMEN) CO.,LTD" --nhom "Current" --vi-sao "cho ky bien ban bu tru"
$ python3 runner.py --credit-term "GRAND FORWARDING LIMITED" --gia-tri "15 days"
$ python3 runner.py --tong-hop W38_Tong_hop....xlsx --tuoi-no W38_Phan_tich....xlsx --ra <out>
  -> 2026_W38_Bang_cong_no_tuan_12.09-18.09.2026.xlsx: 70 khách,
     tổng 10,993,269,889, 42 dòng cần xem lại

$ python3 tests/so_sanh.py <máy sinh> <bản tay W38>
khớp cả 6 nhóm: 52/70 | khớp khi gộp 61+: 56/70
```

Dòng lệch to nhất là LLC UPP: máy 2.110.736.227 theo Misa, bản tay 1.113.953.671.
Kế toán đã xác nhận số Misa đúng (khoản 1 tỷ được ghi nhận lùi ngày 22/09), nên
ở dòng này máy đúng hơn bản tay. Tổng bảng của máy vì thế cao hơn bản tay 1 tỷ.

**Lệch so với đề bài.** Không.

**Câu hỏi còn treo cho user**
1. ~~OPTEC, GRAND, AHC~~ — kế toán đã trả lời 23/09: Optec và Grand theo Misa;
   AHC khai ngoại lệ về Current vì đang chờ ký biên bản bù trừ. Grand đổi credit
   term sang 15 ngày. Xem `handoff/docs/NOTE-ke-toan-tra-loi-20260923.md`.
2. ~~LLC UPP lệch 1 tỷ~~ — kế toán đã trả lời: lấy số Misa, vì khoản 1 tỷ được
   ghi nhận lùi vào ngày 22/09 sau khi bản tay đã làm xong.
3. ~~Số job~~ — kế toán chốt: giữ nguyên của tuần trước, không xoá trắng. Tool
   đang làm đúng vậy.
4. Máy nào chạy lịch nền, mấy giờ. Hiện wizard đặt 30 phút quét một lần.

Giả định kỹ thuật:

`[GIẢ ĐỊNH KỸ THUẬT]` Nạp trạng thái ban đầu từ một bảng đã chia sẵn nhóm tuổi
thì lấy giữa khoảng của nhóm làm hạn thanh toán giả định (Current = ngày chốt
+ 15, nhóm 1-30 = ngày chốt - 15, ...).
· Vì sao: bảng cũ chỉ có nhóm, không có ngày. Lấy giữa khoảng cho sai số nhỏ
  nhất và không làm nhảy nhóm ngay tuần kế tiếp.
· Nếu sai: chỉ ảnh hưởng vài tuần đầu, sửa bằng cách đổi bảng `GIUA_NHOM` trong
  `wr/soghi.py` rồi nạp lại. Không phải làm lại phần nào khác.

`[GIẢ ĐỊNH KỸ THUẬT]` So khớp khách hàng giữa hai file Misa theo TÊN đã chuẩn
hoá, không theo mã khách.
· Vì sao: mã khách của cùng một khách khác nhau giữa hai file (Chu Kong là
  "CHU KONG" ở file Tổng hợp và "KH00329" ở file Tuổi nợ).
· Nếu sai: đổi tên khách trong Misa sẽ làm mất liên kết với sổ ghi, khách đó
  bị coi là khách mới. Tool đánh dấu vào sheet Cần xem lại nên phát hiện được.

### Bổ sung chiều 2026-09-23 — thư mục đích do người dùng đặt

Người dùng trỏ thư mục đích bất kỳ, tool tạo sẵn 6 ngăn bên trong. Thư mục mã
nguồn từ nay không chứa dữ liệu chạy nào.

```
$ python3 runner.py --dat-thu-muc "<thư mục đích>" --khong-tao-thu-muc-con
Thư mục đích: .../_thu-nghiem-tool
    01_Input  02_Output  03_BanTay  so-ghi  luu-tru  log
$ python3 runner.py                       (sau khi thả 2 file W38 vào 01_Input)
  -> 2026_W38_Bang_cong_no_tuan_12.09-18.09.2026.xlsx: 70 khách,
     tổng 10,993,269,889, 42 dòng cần xem lại
$ python3 -m pytest -q tests/test_wr.py
17 passed
```

**Lỗi bắt được trong lúc làm.** Lần chạy đầu, các lệnh chạy tay không đọc cấu
hình của máy nên sổ ghi rơi vào thư mục mã nguồn và tool tưởng là khởi động
lạnh (70/70 dòng vào sheet Cần xem lại thay vì 42). Đã sửa: mọi lệnh đều đọc
settings.json nếu đã đặt thư mục đích. Có 2 phép thử canh đúng chuyện này.

### Bổ sung 2026-09-23 tối — phản hồi của kế toán về bộ dữ liệu thử W39

Kế toán so tổng báo cáo (1.635.150.000) với tổng file Tuổi nợ (1.905.150.000)
và với cột Nợ trước hạn (228.600.000 so với 426.400.000 trên báo cáo), thấy
không khớp nên hỏi lại. Cả hai đều đúng thiết kế, nhưng báo cáo không nói ra
nên người đọc phải tự đoán. Đã sửa: sheet "Cần xem lại" giờ mở đầu bằng khối
đối chiếu 6 dòng, nói rõ tổng nào đối chiếu với tổng nào, chênh bao nhiêu, do
khách nào, và phần Nợ trong hạn gồm bao nhiêu lấy thẳng từ Misa, bao nhiêu là
khoản chưa xác định được tuổi.

```
Tổng bảng báo cáo này                       1.635.150.000
Tổng file Tổng hợp công nợ (dư Nợ)          1.635.150.000   KHỚP
Tổng file Phân tích tuổi nợ                 1.905.150.000   không dùng để đối chiếu
Chênh lệch giữa hai file Misa                 270.000.000   do NORDSTAR FREIGHT OY
Nợ trong hạn trên báo cáo                     426.400.000
  · lấy thẳng từ file Tuổi nợ                 228.600.000
  · chưa xác định được tuổi, tạm để ở đây     197.800.000
```

Thêm 2 phép thử canh khối đối chiếu này. Tổng 19 phép thử.

### Bổ sung 2026-09-24 — thả 3 tuần cũ vào mà không ra file

user thả 6 file của W33, W34, W35 vào `01_Input` rồi bấm Run, không có gì
xảy ra. Hai nguyên nhân, cả hai đều là lỗi của tôi:

1. Hôm 23/09 dọn thư mục mã nguồn, tôi xoá luôn `settings.json` của máy
   user. Tool mất đường dẫn thư mục đích nên thoát ngay, mà lại thoát bằng
   `SystemExit` không ghi nhật ký, nên chạy nền thì im ru. Đã sửa: thoát kiểu
   nào cũng ghi một dòng vào nhật ký, và luật cấm xoá `settings.json` đã lên
   RULES.md mục 5.3a.
2. Kể cả khi có cấu hình, bản cũ chỉ lấy "file Tổng hợp đầu tiên + file Tuổi nợ
   đầu tiên" thấy trong thư mục. Thả nhiều tuần một lúc là ghép nhầm file của
   hai tuần khác nhau, và bị chốt chặn lệch kỳ chặn lại. Đã sửa: tool đọc ngày
   trong từng file, ghép đúng cặp cùng kỳ, chạy lần lượt từ kỳ cũ nhất, file
   lẻ thì bỏ qua và ghi rõ lý do.

```
$ python3 runner.py
Có 3 kỳ trong thư mục, chạy lần lượt từ kỳ cũ nhất
Bắt đầu kỳ đến 14/08/2026 -> 2026_W33_...xlsx: 68 khách, tổng 20.346.480.501
Bắt đầu kỳ đến 21/08/2026 -> 2026_W34_...xlsx: 75 khách, tổng 20.959.094.616
Bắt đầu kỳ đến 28/08/2026 -> 2026_W35_...xlsx: 69 khách, tổng 15.781.824.711
$ python3 -m pytest -q tests/test_wr.py
21 passed
```

Nhân tiện sửa thêm một chỗ khó chịu: trước đây khách thiếu Credit Term hay
Salesman cũng bị tô màu như lỗi số liệu, nên cả bảng vàng khè. Giờ tách hai
loại: tô màu chỉ dành cho vấn đề về số, còn thiếu thông tin do người điền thì
gom thành một dòng tổng kết ở cuối sheet Cần xem lại.

### Bổ sung 2026-09-24 (2) — chạy được ở máy nào, đặt ở đâu cũng được

user yêu cầu bảo đảm hai chỗ tách bạch: nơi chứa mã nguồn (máy local) và
thư mục đích (OneDrive), cả hai đặt ở đâu cũng chạy. Đã rà và sửa:

| Chỗ dễ vỡ | Trước | Nay |
|---|---|---|
| Đường dẫn cứng trong mã | không có, nhưng không ai canh | có phép thử quét mọi file .py, thấy `C:\`, `/home/`, `/Users/` là đỏ |
| Thư mục mã nguồn nằm trên OneDrive, đồng bộ sang máy thứ hai | một `settings.json` dùng chung, máy này đè máy kia | cấu hình lưu theo TÊN MÁY, mỗi máy một đường dẫn trong cùng file |
| Chạy nền bằng pythonw | `print` tiếng Việt làm chết tiến trình (không có màn hình, hoặc console không hiện được tiếng Việt) | bọc lại, in không được thì bỏ qua, nhật ký vẫn ghi |
| Đăng ký Task Scheduler | truyền thẳng lệnh dài có dấu cách, dễ gãy | sinh `chay_nen.bat` cạnh mã nguồn rồi chỉ đăng ký file đó; thất bại thì in lý do Windows báo |
| Không có `python` trong PATH | báo lỗi rồi dừng | thử tiếp `py -3`; pip không cài được thì thử `--user` |
| File kết quả đang mở trong Excel | đánh dấu `[LOI]` rồi bỏ luôn | hoãn kỳ đó, giữ nguyên file đầu vào, lần chạy sau làm tiếp |
| Ghi nhật ký thất bại (ổ đầy, mất quyền) | chết cả tiến trình | bỏ qua phần ghi nhật ký, việc chính vẫn chạy |

Phép thử thêm: đường dẫn có dấu cách và tiếng Việt, hai máy hai đường dẫn trong
cùng file cấu hình, file kết quả bị khoá. Tổng 25 phép thử.

### Bổ sung 2026-09-24 (3) — đối chiếu 3 tuần W33-W35 kế toán làm tay

user đặt 3 bản kế toán làm tay vào `03_BanTay`.

| | Chạy thẳng (sổ ghi rỗng) | Nối từ bản tay tuần trước |
|---|---|---|
| W33 | 43/67 dòng | — (không có bản tay W32) |
| W34 | 43/72 dòng | **60/72 dòng** |
| W35 | 41/68 dòng | **47/68 dòng** |

Khác biệt giữa hai cột cho thấy sổ ghi có tác dụng thật: cùng dữ liệu, chỉ
khác chỗ có hay không có trạng thái tuần trước.

Ba nhóm chênh, đều cần kế toán trả lời, không tự quyết được:
1. Bốn khách có dư Nợ trong Misa mà bảng tay không có: MN Shipping Hải Phòng
   183.514.584, Rustam Yadu 74.491.924, Xiamen Trans-Europeasia 9.010.411,
   Optimalog 2.115.321.
2. Một khách trong bảng tay không có trong Misa: Vận Tải Biển Minh Nguyên
   183.514.584 — trùng đúng số tiền với MN Shipping Hải Phòng. Nhiều khả năng
   khách đổi tên trong Misa (mã cũng đổi: KH00095 thành KH00149) mà bảng tay
   chưa cập nhật.
3. STA Logistic Center lệch 235.520 đồng và ACS Time Critical lệch 6.318.221
   đồng ở cả ba tuần — kế toán chỉnh tay, máy theo file Tổng hợp.

### Bổ sung 2026-09-24 (4) — kế toán trả lời về 5 chênh lệch

Xem `handoff/docs/NOTE-ke-toan-tra-loi-20260924.md`. Cả 5 chênh đều do file
Misa xuất hôm nay khác file Misa lúc kế toán làm báo cáo, không phải lỗi hai
bên. Hai thay đổi vào tool:

1. `--bo-qua` / `--nhan-lai`: loại khách ra khỏi bảng khi Misa ghi nợ mà thực
   chất khách không nợ (ca Optimalog). Khoản đó vẫn hiện ở sheet Cần xem lại
   kèm lý do, để không biến mất khỏi tầm mắt.
2. Báo cáo ghi lại tên và thời điểm sửa của hai file nguồn ở đầu sheet Cần xem
   lại, kèm một câu nhắc chỉ đối chiếu khi cùng một lần xuất Misa.

Luật mới lên RULES.md 5.2c. 27 phép thử.

**Ảnh hưởng tới kịch bản nghiệm thu.** Phép so "máy với bản tay" chỉ có giá trị
khi cả hai dùng cùng một cặp file Misa. Ba tuần W33-W35 không thoả điều kiện
đó, nên con số 60/72 và 47/68 là sàn dưới, không phải mức chính xác thật của
tool. Đề nghị user: tuần tới kế toán làm tay TRÊN ĐÚNG cặp file đã thả vào
01_Input, rồi mới so.

### Bổ sung 2026-09-24 (5) — dọn nhà, đổi nhịp, lên git

- Nhịp chạy nền: 30 phút thành **15 phút**. Đo thực tế trên máy user: lần
  quét rỗng 0,41 giây, lần chạy thật 1,31 giây, nên nhịp dày không ảnh hưởng máy.
- Xoá rác trong thư mục mã nguồn và chặn sinh lại:
  · lệnh chỉ đọc (`--xem-cau-hinh`) không tạo thư mục nữa — trước đây nó gọi
    `cac_thu_muc()` có `mkdir`, nên khi cấu hình trỏ đường dẫn Windows mà chạy
    trên máy khác, nó đẻ ra một thư mục tên `C:\Users\...` ngay trong mã nguồn;
  · cấu hình trỏ đường dẫn Windows trên máy không phải Windows → báo lỗi rõ ràng;
  · `PYTHONDONTWRITEBYTECODE=1` trong `setup.bat` và `chay_nen.bat`;
  · `pytest.ini` tắt cache provider.
  Hai phép thử mới canh đúng hai chuyện này.
- README viết lại cho FIN: badge, mục có emoji, sơ đồ ASCII, bảng tra sự cố.
- `git init` + commit đầu tiên, 40 file. Chưa đẩy lên remote — chờ user tạo
  repo và cho địa chỉ.

### Soi tĩnh của Claude

Đọc lại mã theo đề bài: kế toán không có thao tác nào ngoài thả file; tool
không xoá file nào, chỉ đổi tên; mọi nhánh lỗi đều ghi log và để lại file .txt
giải thích; tổng 6 nhóm luôn được ép bằng Total trước khi ghi.

### Kết luận cũ của người làm (chuẩn v1.1) — 2026-09-24

**1. Logic tính số, đọc file Misa, chốt chặn lệch kỳ, sổ ghi, ghép nhiều kỳ,
xử lý đứt quãng và chen hàng: ĐẠT.**
Ký: claude, 2026-09-24.
Bằng chứng: 29 phép thử (9 phép thử chạy được cả khi không có dữ liệu thật),
các lệnh và kết quả dán ở khoang AUDIT phía trên, và 6 tuần dữ liệu thật
W33-W35, W37-W39 đã chạy ra báo cáo.

**2. Chạy trên Windows của user — chạy được, ĐẠT một phần.**
Ký: claude, 2026-09-24, trên bằng chứng gián tiếp.
Bằng chứng: `log/runner.log` trong thư mục đích ghi lần chạy lúc
`2026-09-23 16:04:04` do user tự bấm trên máy Windows, sinh ra
`2026_W39_Bang_cong_no_tuan_19.09-25.09.2026.xlsx` với 10 khách, tổng
1.635.150.000. Lần chạy này chứng minh trên Windows: Python và openpyxl chạy
được, đọc được file Misa, đường dẫn có dấu cách và tiếng Việt không vỡ, ghi
được file kết quả vào thư mục OneDrive.

**3. Ba thứ vẫn `[RUNTIME UNVERIFIED]`, Claude không ký:**
- `setup.bat` — wizard dò thư mục, cài thư viện, chưa ai chạy.
- Task Scheduler — chưa ai xác nhận tác vụ được tạo và tự gọi đúng nhịp.
- Hiển thị trong Excel thật — màu tím `#4d148c`, cam `#ff6200`, và tiếng Việt
  trong sheet "Cần xem lại".
Ba mục này user hoặc codex ký, mỗi mục một dòng, ghi thẳng vào đây. Đủ ba
dòng thì phiếu chuyển `đóng-đạt`.

**4. Nghiệm thu nghiệp vụ còn treo:** phép so với bản kế toán làm tay chỉ có
giá trị khi cả hai dùng cùng một cặp file Misa. Chưa có lần nào như vậy.
Việc này mở phiếu mới WR-02, không giữ WR-01 mở để chờ.

Soát: user — đóng theo lệnh user 2026-11-05 ("các ticket đang mở cái nào xong thì close & move vào done"). Bằng chứng: Báo cáo W43 (chạy lại 29/10) và W44 (30/10) do Task Scheduler trên máy FIN sinh ra, dùng thật trong tuần (handoff/docs/Cong-No-Workspace (W44): runner.log, 2026_W44_Credit_Report_24.10-30.10.2026.xlsx — 55 khách, phải thu 13.099.776.819; Sales đã ghi lý do tuần này vào file). Không qua codex soát (lệch mục 0/9 của chuẩn, user quyết).
Kết quả: ĐẠT
-- claude - 2026-10-27 11:40 UTC+7

## HISTORY

### 2026-09-23
Failure    không
           Báo lỗi giả: không
Lessons    Dựng phòng bàn giao theo Chuẩn Bàn Giao v1.1. Ghi nhận trước khi
           bắt tay: file Misa tải về bị Windows đánh số thứ tự trong tên, đã
           gây nhầm một lần trong lúc bàn logic — luật "đọc ngày chốt bên
           trong file, không tin tên file" đã lên RULES.md ngày 2026-09-23.
-- claude

### 2026-09-24 (tối)
Failure    không
           Báo lỗi giả: không
Lessons    Fixtures là dữ liệu công nợ thật, đã đưa ra khỏi git trước khi repo
           thành công khai; phép thử cần dữ liệu đó tự bỏ qua khi thiếu, nên
           máy mới clone về vẫn chạy được 9 phép thử.
           Tên người trong tài liệu bàn giao đã thay bằng vai trò — repo công
           khai thì tên nhân sự và số nợ của khách không nên nằm trong đó.
           Lịch sử git dựng lại thành một commit sạch vì các commit trước còn
           chứa tên người và dữ liệu thật.
-- claude

### 2026-09-24
Failure    Xoá nhầm settings.json của máy user khi dọn thư mục mã nguồn,
           làm tool đứng im. Tool lại thoát không ghi nhật ký nên không ai biết.
           Báo lỗi giả: không
Lessons    Dọn dẹp thư mục người khác đang dùng thì phải liệt kê trước cái gì
           sẽ xoá. File cấu hình của máy trông như rác nhưng là thứ duy nhất
           tool biết đường về nhà.
           Mọi nhánh dừng sớm phải ghi nhật ký. Đã lên RULES.md 5.3a ngày 24/09.
-- claude

### 2026-09-23 (chiều)
Failure    không
           Báo lỗi giả: không
Lessons    Kế toán trả lời đợt 2. Ba dòng AHC / Optec / Grand hoá ra ba nguyên
           nhân khác nhau, không phải một dạng như tôi đoán ở NOTE-doi-chieu:
           AHC là bù trừ công nợ hai chiều chưa ký biên bản, Optec là khách trả
           theo tháng, Grand là hợp đồng mới ký mà báo cáo chưa cập nhật. Bài
           học: đừng gộp các dòng lệch thành một dạng chỉ vì chúng trông giống
           nhau trên bảng số.
           Thêm cơ chế "kế toán khai đè" cho credit term, vì hợp đồng mới ký
           không thể suy ra từ hai file Misa. Đã lên RULES.md ngày 2026-09-23.
-- claude
### 2026-11-05
Failure: không
Báo lỗi giả: không
Lessons: đóng phiếu theo lệnh user sau khi tool chạy thật W43–W44 trên máy FIN.
-- claude - 2026-11-05 10:05 UTC+7
