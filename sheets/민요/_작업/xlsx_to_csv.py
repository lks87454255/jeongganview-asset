#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx_to_csv.py — 검수 엑셀의 '악보 N' 시트를 앱 CSV(v2)에 반영
=================================================================
    sheets/민요/검수/{곡}.xlsx  ─(악보 1..N 시트)→  sheets/민요/csv/{곡}.csv

    python3 xlsx_to_csv.py                  # 바뀐 곡만 CSV 갱신
    python3 xlsx_to_csv.py --check          # 바뀔 곡만 출력 (파일 변경 없음)
    python3 xlsx_to_csv.py "강강술래" ...   # 특정 곡만 (xlsx 파일명, 확장자 생략 가능)

규칙
  - 고칠 곳은 '악보 N' 시트뿐이다. '악보 N' = CSV '페이지,N' 블록.
    A열 = 정간 번호, 그 오른쪽은 'k줄'(율명) · '가사' 쌍 → CSV k열(대) · k열(소).
    (11줄 이상 이어지는 시트는 11줄 → 1열, 12줄 → 2열 …)
  - 한 칸에 여러 음은 칸 안 줄바꿈(Alt+Enter / Option+Enter).
  - CSV 머리 정보(타이틀·박자·행수·원본·상태 …)는 CSV 쪽 값을 그대로 둔다.
  - '정보' · '정간목록' 시트는 읽지 않는다(고쳐도 반영 안 됨).
  - 표준 라이브러리만 사용 (openpyxl 불필요). 엑셀·Numbers·구글시트에서 .xlsx 로 저장한 파일 지원.

반영 후 레포 루트에서:  python3 generate_sheets_index.py   (sha256·크기 갱신)
"""
import csv
import io
import os
import re
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                      # sheets/민요
XLSX_DIR = os.path.join(ROOT, "검수")
CSV_DIR = os.path.join(ROOT, "csv")
MAX_LINES = 10                                     # CSV 한 페이지 최대 줄(열) 수

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "pr": "http://schemas.openxmlformats.org/package/2006/relationships"}


def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


# ───────────────────────────── xlsx 읽기 ─────────────────────────────
def _text(el) -> str:
    """<si>/<is> 안의 <t> (리치 텍스트 포함) 를 이어 붙인다."""
    return "".join(t.text or "" for t in el.iter(f"{{{NS['m']}}}t"))


def _col_index(ref: str) -> int:
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n                                        # A=1


def read_sheets(path: str) -> dict:
    """{시트 이름: {(행, 열): 문자열}}"""
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        shared = [_text(si) for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS)]
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    out = {}
    for sh in ET.fromstring(z.read("xl/workbook.xml")).find("m:sheets", NS):
        target = rels[sh.get(f"{{{NS['r']}}}id")].lstrip("/")
        target = target if target.startswith("xl/") else "xl/" + target
        cells = {}
        for c in ET.fromstring(z.read(target)).iter(f"{{{NS['m']}}}c"):
            t, v = c.get("t"), c.find("m:v", NS)
            if t == "inlineStr":
                val = _text(c.find("m:is", NS))
            elif t == "s" and v is not None:
                val = shared[int(v.text)]
            elif v is not None and v.text is not None:
                val = v.text
                if t is None and re.fullmatch(r"-?\d+\.0+", val):   # 숫자 1 → "1.0" 방지
                    val = val.split(".")[0]
            else:
                continue
            ref = c.get("r")
            row = int(re.search(r"\d+", ref).group(0))
            cells[(row, _col_index(ref))] = val
        out[nfc(sh.get("name"))] = cells
    return out


def score_pages(sheets: dict, path: str) -> list:
    """'악보 N' 시트 → [ {정간: {k줄: (율명, 가사)}} ... ] (페이지 순)"""
    pages = sorted(((int(m.group(1)), name) for name in sheets
                    if (m := re.fullmatch(r"악보\s*(\d+)", name))), key=lambda x: x[0])
    if not pages:
        raise ValueError(f"'악보 N' 시트가 없음: {path}")
    if [p for p, _ in pages] != list(range(1, len(pages) + 1)):
        raise ValueError(f"'악보' 시트 번호가 1부터 연속이 아님: {[p for p, _ in pages]}")
    result = []
    for _, name in pages:
        cells = sheets[name]
        # 머리 행: A열이 '정간' 인 행
        head = next((r for (r, c), v in cells.items() if c == 1 and v.strip() == "정간"), None)
        if head is None:
            raise ValueError(f"[{name}] '정간' 머리 행을 찾을 수 없음")
        line_cols = {}                              # k줄 → (율명 열, 가사 열)
        for (r, c), v in cells.items():
            if r == head and (m := re.fullmatch(r"(\d+)\s*줄", v.strip())):
                line_cols[int(m.group(1))] = (c, c + 1)
        if not line_cols:
            raise ValueError(f"[{name}] 'k줄' 머리가 없음")
        # 한 이미지가 10줄을 넘으면 다음 페이지는 11줄부터 → 페이지 안 열 = (k-1) % 10 + 1
        base = (min(line_cols) - 1) // MAX_LINES * MAX_LINES
        if max(line_cols) - base > MAX_LINES:
            raise ValueError(f"[{name}] 한 시트에 {MAX_LINES}줄 초과: {sorted(line_cols)}")
        line_cols = {k - base: cols for k, cols in line_cols.items()}
        page = {}
        for (r, c), v in cells.items():
            if r <= head or c != 1 or not v.strip():
                continue
            if not v.strip().isdigit():
                raise ValueError(f"[{name}] {r}행 A열 정간 번호가 숫자가 아님: {v!r}")
            jg = int(v.strip())
            page[jg] = {k: (cells.get((r, yc), ""), cells.get((r, gc), "")) for k, (yc, gc) in line_cols.items()}
        result.append(page)
    return result


# ───────────────────────────── CSV 쓰기 ─────────────────────────────
def guard(v: str) -> str:
    """엑셀 수식 위험 방지 — build_all.ps1 / generate_sheets_index.py 규칙과 동일."""
    v = nfc(v.replace("\r\n", "\n").replace("\r", "\n"))
    if v and (v[0] in "=+-@" or (len(v) > 1 and v[0] == "'" and v[1] in "=+-@'")):
        v = "'" + v
    return v


def field(v: str) -> str:
    if any(ch in v for ch in ',"\n\r') or (v and (v[0].isspace() or v[-1].isspace())):
        return '"' + v.replace('"', '""') + '"'
    return v


def build_csv(csv_path: str, pages: list) -> str:
    raw = open(csv_path, "rb").read().decode("utf-8-sig")
    records = list(csv.reader(io.StringIO(raw, newline="")))
    meta = []
    for r in records:
        if r and r[0].strip() == "페이지":
            break
        meta.append(r)
    rows_meta = next((r[1] for r in meta if r and r[0].strip() == "행수" and len(r) > 1), "")
    if not rows_meta.strip().isdigit():
        raise ValueError(f"CSV '행수' 행이 없음: {csv_path}")
    nrows = int(rows_meta)
    header = ["정간번호"] + [f"{k}열({s})" for k in range(MAX_LINES, 0, -1) for s in ("대", "소")]
    lines = [",".join(field(c) for c in r) for r in meta]
    for i, page in enumerate(pages, 1):
        extra = sorted(j for j in page if j < 1 or j > nrows)
        if extra:
            raise ValueError(f"악보 {i}: 정간 번호 {extra} 가 행수 {nrows} 범위를 벗어남")
        lines.append(f"페이지,{i}")
        lines.append(",".join(header))
        for jg in range(1, nrows + 1):
            row = [str(jg)]
            for k in range(MAX_LINES, 0, -1):
                y, g = page.get(jg, {}).get(k, ("", ""))
                row += [guard(y), guard(g)]
            lines.append(",".join(field(c) for c in row))
    return "\ufeff" + "\n".join(lines) + "\n"


def same(a: str, b: str) -> bool:
    return a.replace("\r\n", "\n") == b.replace("\r\n", "\n")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    check = "--check" in sys.argv
    names = sorted(nfc(f[:-5]) for f in os.listdir(XLSX_DIR)
                   if f.lower().endswith(".xlsx") and not f.startswith(("_", "~$", ".")))
    if args:
        want = {nfc(a[:-5] if a.lower().endswith(".xlsx") else a) for a in args}
        missing = want - set(names)
        if missing:
            sys.exit(f"❌ 검수 폴더에 없음: {sorted(missing)}")
        names = [n for n in names if n in want]
    changed, errors = [], []
    for name in names:
        xlsx = next(os.path.join(XLSX_DIR, f) for f in os.listdir(XLSX_DIR) if nfc(f) == name + ".xlsx")
        csv_path = os.path.join(CSV_DIR, name + ".csv")
        try:
            if not os.path.exists(csv_path):
                raise ValueError(f"짝이 되는 CSV 없음: csv/{name}.csv")
            new = build_csv(csv_path, score_pages(read_sheets(xlsx), xlsx))
        except Exception as e:                     # noqa: BLE001 — 곡 단위로 보고하고 계속
            errors.append(f"{name}: {e}")
            continue
        old = open(csv_path, "rb").read().decode("utf-8-sig")
        if same("\ufeff" + old, new):
            continue
        changed.append(name)
        if not check:
            with open(csv_path, "w", encoding="utf-8", newline="") as f:
                f.write(new)
    for e in errors:
        print(f"  ❌ {e}")
    verb = "바뀔" if check else "갱신"
    print(f"{len(names)}곡 확인 / {verb} {len(changed)}곡 / 오류 {len(errors)}곡")
    for n in changed:
        print(f"  ✏️  csv/{n}.csv")
    if changed and not check:
        print("\n다음 단계 (레포 루트):\n  python3 generate_sheets_index.py\n"
              "  git add sheets/민요 sheets-index.json && git commit -m 'fix(sheets): 검수 반영' && git push")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
