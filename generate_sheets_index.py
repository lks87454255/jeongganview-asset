#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_index.py  (jeongganview-sheets)
=======================================
sheets/<카테고리>/*.csv (정간보 CSV v2) 를 스캔해 sheets-index.json 을 만들고,
앱 불러오기 규칙(util/JeongganCsv.kt)대로 모든 CSV 를 검사합니다.

    python3 generate_index.py          # 검사 + 목록 생성
    python3 generate_index.py --check  # 검사만 (파일 변경 없음)

오류(ERROR)가 하나라도 있으면 종료 코드 1 — 고친 뒤 push 하세요.
  ERROR : 파일명 NFC 아님 · UTF-8 아님 · "정간번호" 행 없음 · 행수/페이지 블록 불일치 ·
          = + - @ 로 시작하는 칸(엑셀 수식 위험)
  WARN  : 율명 칸(대)에 율명·기호표에 없는 글자 (위치 출력)

URL 인코딩은 jeongganview-asset/generate_index.py 와 같다:
  Java  URLEncoder.encode(name, "UTF-8").replace("+", "%20")
  Python urllib.parse.quote(name, safe="")
"""
import csv
import hashlib
import io
import json
import os
import sys
import unicodedata
from datetime import datetime, timezone
from urllib.parse import quote

OWNER = "lks87454255"
REPO = "jeongganview-asset"
BRANCH = "main"
BASE = f"https://raw.githubusercontent.com/{OWNER}/{REPO}/{BRANCH}"
IGNORE = {".DS_Store", ".gitkeep", ".gitignore", "Thumbs.db"}
COLUMNS = 20

# 앱 Yulmyeong.NAMES + JeongganSymbols.MARKS 와 동일
YUL = list("㣴㣕㣖㣣㣨㣡㣸㣩") + ["𢓡"] + list("㣮㣳㣹僙㐲㑀俠㑬㑖") + ["𠐭"] + list(
    "㑣侇㑲㒇㒣黃大太夾姑仲㽔林夷南無應潢汏汰浹㴌㳞㶋淋洟湳潕㶐㶂") + ["𣴘"] + list("㳲㴺㵈㴢㶙㵉㴣㵜㶃㶝")
MARKS = ["―", "△", "И", "ﾉ", "^", "⌝", "ㅋ", "⌞", "է", "⊏", "⊔", "·", "○", "⁚", "‹", "／"]
KNOWN = set(YUL) | set(MARKS)
assert len(YUL) == 60, len(YUL)


def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def raw_url(*parts: str) -> str:
    return f"{BASE}/" + "/".join(quote(nfc(p), safe="") for p in parts)


def unguard(v: str) -> str:
    return v[1:] if len(v) > 1 and v[0] == "'" and v[1] in "=+-@'" else v


def check_csv(path: str, rel: str, errors: list, warns: list) -> dict:
    """CSV v2 를 앱과 같은 규칙으로 읽어 메타 정보를 돌려준다."""
    info = {"title": "", "beat": "", "rows": 0, "pages": 0}
    raw = open(path, "rb").read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        errors.append(f"UTF-8 아님: {rel}")
        return info
    records = [r for r in csv.reader(io.StringIO(text, newline="")) if any(c for c in r)]
    if not any(r and r[0].strip() == "정간번호" for r in records):
        errors.append(f"'정간번호' 행 없음: {rel}")
        return info
    page, max_row, pages, headers, page_rows = 1, 0, set(), 0, {}
    for r in records:
        key = r[0].strip()
        val = unguard(r[1]) if len(r) > 1 else ""
        for c in r:
            if c and c[0] in "=+-@":
                errors.append(f"수식으로 해석될 칸 '{c[:10]}': {rel}")
        if key == "타이틀":
            info["title"] = val.strip()
        elif key == "박자":
            info["beat"] = val.strip().removesuffix("박자").strip()
        elif key == "행수":
            info["rows"] = int(val) if val.strip().isdigit() else 0
        elif key == "페이지":
            page = int(val) if val.strip().isdigit() else page + 1
            pages.add(page)
        elif key == "정간번호":
            headers += 1
            pages.add(page)
        elif key.isdigit():
            n = int(key)
            max_row = max(max_row, n)
            page_rows.setdefault(page, set()).add(n)
            for i, cell in enumerate(r[1:], start=1):
                col = COLUMNS - i + 1
                cell = unguard(cell.strip())
                if col < 1 or not cell or col % 2:
                    continue  # 홀수 열 = 가사
                bad = sorted({ch for ch in cell if not ch.isspace() and ch not in KNOWN})
                if bad:
                    warns.append(f"{rel} {page}페이지 {(col + 1) // 2}줄 {n}정간: {' '.join(bad)}")
    info["pages"] = len(pages) or 1
    if info["rows"] < 1:
        errors.append(f"행수 없음: {rel}")
    elif max_row > info["rows"]:
        errors.append(f"행 번호 {max_row} > 행수 {info['rows']}: {rel}")
    if headers != info["pages"]:
        errors.append(f"정간번호 행 {headers}개 ≠ 페이지 {info['pages']}개: {rel}")
    for p, rows in page_rows.items():
        if len(rows) != info["rows"]:
            warns.append(f"{rel} {p}페이지 데이터 행 {len(rows)}개 ≠ 행수 {info['rows']}")
    return info


def scan(sheets_dir: str):
    errors, warns, categories = [], [], []
    for cat in sorted(n for n in os.listdir(sheets_dir) if os.path.isdir(os.path.join(sheets_dir, n)) and n not in IGNORE):
        files = []
        for fn in sorted(os.listdir(os.path.join(sheets_dir, cat))):
            if fn in IGNORE or fn.startswith(".") or not fn.lower().endswith(".csv"):
                continue
            rel = f"sheets/{cat}/{fn}"
            if fn != nfc(fn) or cat != nfc(cat):
                errors.append(f"파일명이 NFC(완성형)가 아님 — 앱 URL 과 달라 404: {rel}")
            path = os.path.join(sheets_dir, cat, fn)
            info = check_csv(path, rel, errors, warns)
            data = open(path, "rb").read()
            files.append({
                "name": os.path.splitext(fn)[0],
                "title": info["title"],
                "fileName": fn,
                "path": rel,
                "url": raw_url("sheets", cat, fn),
                "sizeBytes": len(data),
                "beat": info["beat"],
                "rows": info["rows"],
                "pages": info["pages"],
                "sha256": hashlib.sha256(data).hexdigest(),
            })
        if files:
            categories.append({"name": cat, "files": files})
            print(f"  📄 [{cat}] {len(files)}곡")
    return categories, errors, warns


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    sheets_dir = os.path.join(base_dir, "sheets")
    out = os.path.join(base_dir, "sheets-index.json")
    print(f"📁 레포 경로: {base_dir}\n🔗 Base URL : {BASE}\n")
    categories, errors, warns = scan(sheets_dir)
    for w in warns:
        print(f"  ⚠️ {w}")
    for e in errors:
        print(f"  ❌ {e}")
    print(f"\n경고 {len(warns)}건 / 오류 {len(errors)}건")
    if "--check" in sys.argv:
        sys.exit(1 if errors else 0)
    data = {
        "version": 1,
        "generated": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "owner": OWNER, "repo": REPO, "branch": BRANCH,
        "categories": categories,
        "_total": {"categories": len(categories), "files": sum(len(c["files"]) for c in categories)},
    }
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print(f"✅ sheets-index.json {data['_total']['categories']}개 카테고리 / {data['_total']['files']}곡")
    if errors:
        print("❌ 오류를 고친 뒤 다시 실행하세요(이 상태로 push 하면 앱에서 열리지 않을 수 있습니다).")
        sys.exit(1)
    print("\n다음 단계:\n  git add sheets sheets-index.json\n  git commit -m 'chore: update sheets index'\n  git push")


if __name__ == "__main__":
    main()
