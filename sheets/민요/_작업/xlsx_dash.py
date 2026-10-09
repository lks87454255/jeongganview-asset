#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx_dash.py — 검수 xlsx 의 '―'(U+2015, 연음) 를 짧은 '–'(U+2013 en dash) 로 바꾼다 (엑셀에서 보기 좋게)
    python3 xlsx_dash.py                 # 검수/*.xlsx 전체
    python3 xlsx_dash.py 강강술래 ...    # 특정 곡만
    python3 xlsx_dash.py --reverse       # 되돌리기 (– → ―)

- 열려 있는 파일(LibreOffice .~lock / 엑셀 ~$ 잠금 파일)은 건너뛴다.
- xlsx 안 XML(xl/*.xml) 의 문자만 바꾸고 나머지 항목·압축 방식은 그대로 둔다.
- CSV 는 그대로 '―' : xlsx_to_csv.py 가 읽을 때 '–' → '―' 로 되돌린다.
"""
import os
import shutil
import sys
import tempfile
import unicodedata
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX_DIR = os.path.join(os.path.dirname(HERE), "검수")
LONG, SHORT = "\u2015", "\u2013"


def nfc(s):
    return unicodedata.normalize("NFC", s)


def convert(path, src, dst):
    """바꾼 글자 수를 돌려준다 (0 이면 파일을 건드리지 않음)."""
    with zipfile.ZipFile(path) as zin:
        items = [(info, zin.read(info.filename)) for info in zin.infolist()]
    total, out = 0, []
    for info, data in items:
        if info.filename.startswith("xl/") and info.filename.endswith(".xml"):
            text = data.decode("utf-8")
            n = text.count(src)
            if n:
                total += n
                data = text.replace(src, dst).encode("utf-8")
        out.append((info, data))
    if not total:
        return 0
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(path))
    os.close(fd)
    with zipfile.ZipFile(tmp, "w") as zout:
        for info, data in out:
            zout.writestr(info, data, compress_type=info.compress_type)
    shutil.copystat(path, tmp)
    os.replace(tmp, path)
    return total


def main():
    reverse = "--reverse" in sys.argv
    src, dst = (SHORT, LONG) if reverse else (LONG, SHORT)
    want = {nfc(a[:-5] if a.lower().endswith(".xlsx") else a) for a in sys.argv[1:] if not a.startswith("--")}
    files = sorted(f for f in os.listdir(XLSX_DIR) if f.lower().endswith(".xlsx") and not f.startswith(("~$", ".")))
    lo_locks = {nfc(f[len(".~lock."):-1]) for f in os.listdir(XLSX_DIR) if f.startswith(".~lock.")}
    ms_locks = {nfc(f[2:]) for f in os.listdir(XLSX_DIR) if f.startswith("~$")}   # 엑셀은 앞 글자를 ~$ 로 덮기도 함
    changed = skipped = 0
    for f in files:
        name = nfc(f)
        if want and name[:-5] not in want:
            continue
        if name in lo_locks or any(name == x or name[2:] == x or name[1:] == x for x in ms_locks):
            print(f"  ⚠️  열려 있어 건너뜀: {name}")
            skipped += 1
            continue
        n = convert(os.path.join(XLSX_DIR, f), src, dst)
        if n:
            changed += 1
            print(f"  ✏️  {name}: {n}개")
    print(f"{src} → {dst}: {changed}개 파일 변경, {skipped}개 건너뜀")


if __name__ == "__main__":
    main()
