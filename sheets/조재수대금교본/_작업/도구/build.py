#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build.py — 조재수대금교본: _작업/tsv → csv/·검수/ (민요 build_all.py 를 이 폴더용으로 임시 복사해 실행)

    python3 _작업/도구/extract.py      # 1) PDF 글자 층 → TSV·원본 PNG·곡목록.tsv
    python3 _작업/도구/build.py        # 2) TSV → CSV·검수 xlsx·_목록.xlsx (+ 머리 정보, 글꼴, A4 레이아웃)

- build_all.py 는 sheets-index.json 까지 쓰므로 반드시 임시 폴더에서 돌리고 csv·검수만 가져온다.
- 율명 칸 꾸밈 기호(앱 악상기호 카탈로그 ʒ ∼ ▾ c s ㄱ ㅅ …)는 팔레트 외로 세지 않게 KNOWN 에 더한다.
- CSV 머리에 곡번호·작사·작곡·노래·정보를 더한다(앱은 모르는 머리 행을 보존).
- 이미 사람이 검수한 xlsx 가 있으면 덮어쓰므로, 검수 시작 뒤에는 실행하지 말 것 (그 뒤로는 xlsx_to_csv.py).
"""
import csv
import importlib.util
import io
import os
import shutil
import subprocess
import sys
import tempfile
import unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.dirname(HERE)
ROOT = os.path.dirname(WORK)
REPO = os.path.dirname(os.path.dirname(ROOT))
GROUP = os.path.basename(ROOT)
BUILD_ALL = os.path.join(REPO, "sheets", "민요", "_작업", "build_all.py")
nfc = lambda s: unicodedata.normalize("NFC", s)


def main():
    if os.path.isdir(os.path.join(ROOT, "검수")) and "--force" not in sys.argv:
        locked = [f for f in os.listdir(os.path.join(ROOT, "검수")) if f.endswith(".xlsx")]
        if locked:
            sys.exit("검수/ 에 xlsx 가 이미 있음 — 검수 내용을 덮어쓸 수 있어 중단 (정말 다시 만들려면 --force)")
    spec = importlib.util.spec_from_file_location("gsi", os.path.join(REPO, "generate_sheets_index.py"))
    gsi = importlib.util.module_from_spec(spec); spec.loader.exec_module(gsi)
    catalog = sorted({ch for m in gsi.CATALOG_MARKS for ch in m})

    tmp = tempfile.mkdtemp()
    T = os.path.join(tmp, "sheets", GROUP)
    os.makedirs(os.path.join(T, "_작업"))
    shutil.copytree(os.path.join(WORK, "tsv"), os.path.join(T, "_작업", "tsv"))
    os.symlink(os.path.join(ROOT, "원본"), os.path.join(T, "원본"))
    src = open(BUILD_ALL, encoding="utf-8").read()
    src = src.replace('CATEGORY = "민요"', f'CATEGORY = "{GROUP}"', 1)
    src = src.replace("KNOWN = set(YUL) | set(MARKS)", f"KNOWN = set(YUL) | set(MARKS) | set({catalog!r})", 1)
    src = src.replace("판독 초안(검수 전)", "PDF 글자 추출(검수 전)")
    open(os.path.join(tmp, "build.py"), "w", encoding="utf-8").write(src)
    r = subprocess.run([sys.executable, os.path.join(tmp, "build.py"), T], capture_output=True, text=True)
    print(r.stdout.strip().splitlines()[-1] if r.stdout.strip() else r.stderr[-800:])
    if r.returncode:
        sys.exit(r.stderr[-2000:])
    for sub in ("csv", "검수"):
        dst = os.path.join(ROOT, sub)
        if os.path.isdir(dst): shutil.rmtree(dst)
        shutil.copytree(os.path.join(T, sub), dst)
    shutil.copy(os.path.join(T, "_작업", "생성리포트.txt"), os.path.join(WORK, "생성리포트.txt"))
    shutil.rmtree(tmp)

    # CSV 머리 정보 더하기 (곡목록.tsv)
    info = {}
    for row in csv.DictReader(open(os.path.join(WORK, "곡목록.tsv"), encoding="utf-8"), delimiter="\t"):
        info[nfc(row["곡명"])] = row
    for f in os.listdir(os.path.join(ROOT, "csv")):
        p = os.path.join(ROOT, "csv", f); name = nfc(f[:-4])
        row = info.get(name)
        if not row: continue
        raw = open(p, "rb").read().decode("utf-8-sig")
        lines = raw.split("\r\n") if "\r\n" in raw else raw.split("\n")
        extra = [(k, row.get(k, "")) for k in ("곡번호", "작사", "작곡", "노래")] + [("정보", row.get("정보", "").replace(",", " "))]
        i = next(k for k, l in enumerate(lines) if l.startswith("상태,"))
        add = [f"{k},{v}" for k, v in extra if v and not any(l.startswith(k + ",") for l in lines)]
        lines[i:i] = add
        open(p, "wb").write(("\ufeff" + ("\r\n" if "\r\n" in raw else "\n").join(lines)).encode("utf-8"))

    # 검수 xlsx 모양: 글꼴(제목 24·율명 24·가사 18) → – 표시 → A4 레이아웃(제목 40pt 동일·행 높이 균일)
    for cmd in (["xlsx_font.py"], ["xlsx_dash.py"]):
        subprocess.run([sys.executable, os.path.join(WORK, cmd[0])], check=True, capture_output=True)
    subprocess.run([sys.executable, os.path.join(HERE, "a4fit.py"), os.path.join(ROOT, "검수")], check=True, capture_output=True)
    subprocess.run([sys.executable, os.path.join(HERE, "layout_a4.py"), os.path.join(ROOT, "검수")], check=True, capture_output=True)
    subprocess.run([sys.executable, os.path.join(WORK, "xlsx_to_csv.py")], capture_output=True)   # 끝 빈 행 정리(엑셀 기준으로 맞춤)
    r = subprocess.run([sys.executable, os.path.join(WORK, "xlsx_to_csv.py"), "--check"], capture_output=True, text=True)
    print(r.stdout.strip().splitlines()[0])


if __name__ == "__main__":
    main()
