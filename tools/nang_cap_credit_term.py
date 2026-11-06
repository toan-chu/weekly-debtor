"""Nâng cấp file credit term đang dùng (SAMPLE_credit_term_khach.xlsx) lên mẫu mới nhất — thêm cột mới, đổi tên sheet
"Client Master" cũ thành "Credit term". Giữ nguyên dữ liệu, giữ cả cột người dùng tự chèn (ghi thêm sau cột cuối), đổi
ngày gõ dạng chữ thành ngày thật. Không tự điền cột mới nào (ô ✓ hợp đồng để Sales tự tích).

Thường KHÔNG cần chạy tay: lần đầu tool tự thêm khách mới vào file là file được ghi lại theo mẫu mới (WR-12).

Dùng: python tools/nang_cap_credit_term.py "<file cũ>" ["<file mới>"]
Không ghi file mới thì ghi cạnh file cũ, thêm đuôi _moi. Không bao giờ ghi đè file cũ."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from wr import khachhang as kh  # noqa: E402


def nang_cap(src, dst=None) -> Path:
    src = Path(src)
    dst = Path(dst) if dst else src.with_name(src.stem + "_moi.xlsx")
    if dst.resolve() == src.resolve():
        raise SystemExit("File mới phải khác file cũ (tool không ghi đè file của người dùng).")
    dong, cot_them, sales = kh.doc_file(src)
    kh.ghi_master(dst, dong, sales, cot_them=cot_them, tham_so=kh.doc_tham_so_file(src))
    print("Đã ghi", dst)
    print(len(dong), "khách;", sum(1 for d in dong if kh.la_tich(d.get("co_hd"))), "đã ✓ hợp đồng;",
          sum(1 for d in dong if d.get("term")), "có term;",
          sum(1 for d in dong if isinstance(d.get("het_han"), str) or isinstance(d.get("hieu_luc"), str)),
          "ngày còn dạng chữ;", "giữ thêm cột:", ", ".join(cot_them) or "không có")
    return dst


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    nang_cap(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
