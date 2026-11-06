# -*- coding: utf-8 -*-
"""Sau khi đổi tên / chuyển thư mục mã nguồn: trỏ lịch chạy nền (Task Scheduler) về runner.py ở thư mục hiện tại,
sửa đường dẫn cũ trong settings.json. Không hỏi gì, không đụng workspace. Gọi qua 3_Sua_lich_chay.bat."""
import json
import os
import subprocess
import sys
from pathlib import Path

GOC = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(GOC))
from runner import TEN_TAC_VU  # noqa: E402
from tools.setup_wizard import lenh_chay_an, tao_file_khoi_dong  # noqa: E402

TEN_CU = ("weekly-report",)          # tên thư mục cũ của repo


def sua_settings() -> int:
    """Đường dẫn trong settings.json còn chứa thư mục cũ mà không còn tồn tại → đổi sang thư mục hiện tại."""
    tep = GOC / "settings.json"
    if not tep.exists():
        return 0
    raw = json.loads(tep.read_text(encoding="utf-8"))
    doi = 0

    def sua(v):
        nonlocal doi
        if isinstance(v, dict):
            return {k: sua(x) for k, x in v.items()}
        if isinstance(v, str) and not Path(v).exists():
            for cu in TEN_CU:
                for tach in ("\\", "/"):
                    mau = tach + cu + tach
                    k = v.lower().find(mau.lower())
                    if k >= 0:
                        moi = str(GOC) + tach + v[k + len(mau):]
                        if Path(moi).exists():
                            doi += 1
                            return moi
        return v

    moi = sua(raw)
    if doi:
        tep.write_text(json.dumps(moi, ensure_ascii=False, indent=2), encoding="utf-8")
    return doi


def main() -> int:
    print(f"Thư mục mã nguồn: {GOC}")
    n = sua_settings()
    print(f"settings.json: sửa {n} đường dẫn cũ." if n else "settings.json: không có đường dẫn cũ cần sửa.")
    if os.name != "nt":
        print("Không phải Windows: bỏ qua lịch chạy.")
        return 0
    tao_file_khoi_dong()
    kq = subprocess.run(["schtasks", "/create", "/tn", TEN_TAC_VU, "/tr", lenh_chay_an(),
                         "/sc", "minute", "/mo", "15", "/f"], capture_output=True, text=True, check=False)
    if kq.returncode == 0:
        print(f'XONG. Lịch "{TEN_TAC_VU}" giờ chạy {GOC / "runner.py"}, 15 phút một lần.')
        print("Kiểm: Task Scheduler → chọn tác vụ → Run → xem _tool\\runner.log trong workspace có dòng mới.")
        return 0
    print("LỖI: không sửa được lịch chạy. Windows báo:")
    print((kq.stderr or kq.stdout).strip()[:400])
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
