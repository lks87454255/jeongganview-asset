#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx_to_csv.py — 검수 엑셀의 '악보' 시트를 앱 CSV(v2)에 반영
=================================================================
    sheets/민요/검수/{곡}.xlsx  ─(악보 시트)→    sheets/민요/csv/{곡}.csv

    python3 xlsx_to_csv.py                  # 바뀐 곡만 CSV 갱신
    python3 xlsx_to_csv.py --check          # 바뀔 곡만 출력 (파일 변경 없음)
    python3 xlsx_to_csv.py "강강술래" ...   # 특정 곡만 (xlsx 파일명, 확장자 생략 가능)

규칙
  - 고칠 곳은 '악보' 시트뿐이다. 가로로 병합된 제목 행이 나올 때마다 한 페이지
    (위에서부터 CSV '페이지,1', '페이지,2' …). 1행·A열(여백)·페이지 사이 빈 행은 무시한다.
  - 제목 행 아래 n번째 행 = 정간 n (정간 번호 열·머리 행 없음).
    행을 끼워 넣거나 지워도 되지만, 페이지마다 정간 행 수는 '행수'와 같아야 한다
    (끼워 넣었으면 맨 아래 빈 행을 지우고, 지웠으면 빈 행을 넣는다). 행수 아래로 밀려난 내용은 오류로 알려 준다.
  - 행수 자체를 바꾸려면 '정보' 시트 '박자 / 행수' 칸을 고친다 (예: 3/4 / 12정간 → 4/4 / 16정간).
    CSV 박자·행수도 그 값으로 바뀐다. 지원: 3/4=12, 4/4=16, 2/4=16, 정악=20.
  - 열: 제목 병합 범위의 오른쪽 끝 두 열 = 1줄(율명, 가사), 그 왼쪽 두 열 = 2줄 … → CSV k열(대) · k열(소).
    줄을 늘리거나 줄이려면 제목 병합 안쪽에서 열 2개(율명·가사)를 함께 끼워 넣거나 지운다(최대 10줄).
    병합 밖에 쓴 내용은 오류.
  - 구형('k줄' 머리 행 있음, A열 '정간' 번호 열, '악보 1' · '악보 2' … 시트)도 읽는다.
    (11줄 이상 이어지는 시트는 11줄 → 1열, 12줄 → 2열 …)
  - 한 칸에 여러 음은 칸 안 줄바꿈(Alt+Enter / Option+Enter).
  - CSV 머리 정보(타이틀·원본·상태 …)는 CSV 쪽 값을 그대로 둔다 (박자·행수만 정보 시트를 따름).
  - '정보' 시트는 '박자 / 행수' 칸만, '정간목록' 시트는 읽지 않는다.
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


class Cells(dict):
    """{(행, 열): 문자열} + merges: [(r1, c1, r2, c2)] (병합 셀)"""
    merges: list


def _ref(ref: str):
    return int(re.search(r"\d+", ref).group(0)), _col_index(ref)


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
        ws = ET.fromstring(z.read(target))
        cells = Cells()
        cells.merges = []
        for mc in ws.iter(f"{{{NS['m']}}}mergeCell"):
            a, _, b = mc.get("ref").partition(":")
            cells.merges.append((*_ref(a), *_ref(b or a)))
        for c in ws.iter(f"{{{NS['m']}}}c"):
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


LINE_RE = re.compile(r"(\d+)\s*줄")


def _parse_block(cells: dict, head: int, end: int, nrows: int, label: str) -> dict:
    """head 행('k줄 | 가사 …') 아래 ~ end 행 전까지 → {정간: {페이지 안 열: (율명, 가사)}}

    - A열이 '정간' 인 형식(구형): A열 숫자 = 정간 번호.
    - 정간 번호 열이 없는 형식(현재): 머리 행 아래 n번째 행 = 정간 n.
    """
    line_cols = {}                                  # k줄 → (율명 열, 가사 열)
    for (r, c), v in cells.items():
        if r == head and (m := LINE_RE.fullmatch(v.strip())):
            line_cols[int(m.group(1))] = (c, c + 1)
    # 한 이미지가 10줄을 넘으면 다음 페이지는 11줄부터 → 페이지 안 열 = (k-1) % 10 + 1
    base = (min(line_cols) - 1) // MAX_LINES * MAX_LINES
    if max(line_cols) - base > MAX_LINES:
        raise ValueError(f"[{label}] 한 페이지에 {MAX_LINES}줄 초과: {sorted(line_cols)}")
    line_cols = {k - base: cols for k, cols in line_cols.items()}

    def row_of(r):
        return {k: (cells.get((r, yc), ""), cells.get((r, gc), "")) for k, (yc, gc) in line_cols.items()}

    page = {}
    if cells.get((head, 1), "").strip() == "정간":
        for (r, c), v in cells.items():
            # A열이 숫자인 행만 정간 행. 빈 행·다음 페이지 제목 행은 건너뜀
            if not head < r < end or c != 1 or not v.strip().isdigit():
                continue
            jg = int(v.strip())
            if jg in page:
                raise ValueError(f"[{label}] 정간 {jg} 이(가) 두 번 나옴 ({r}행)")
            page[jg] = row_of(r)
        return page
    # 정간 번호 열 없음: 다음 페이지는 [빈 행][제목 행][머리 행] 이므로 head+nrows 는 end-2 보다 위여야 함
    if end < 10 ** 9 and head + nrows > end - 2:
        raise ValueError(f"[{label}] 정간 행이 행수 {nrows} 보다 적음 — 행을 지웠는지 확인 ({head + 1}~{end - 2}행)")
    for jg in range(1, nrows + 1):
        page[jg] = row_of(head + jg)
    return page


def _parse_titled(cells: "Cells", nrows: int, name: str) -> list:
    """현재 형식: 머리 행 없음. 한 행 전체 폭 병합 칸 = 제목 행, 그 아래 n번째 행 = 정간 n.

    병합 범위 [c1..c2] 가 악보 폭: 오른쪽 끝 두 열 = 1줄(율명, 가사), 그 왼쪽 두 열 = 2줄 …
    """
    titles = sorted((r1, c1, c2) for r1, c1, r2, c2 in cells.merges if r1 == r2 and c2 > c1)
    if not titles:
        raise ValueError(f"[{name}] 제목 행(가로 병합 칸)도 'k줄' 머리 행도 없음")
    pages = []
    for i, (t, c1, c2) in enumerate(titles, 1):
        label = f"{name} {i}번째 페이지"
        if (c2 - c1 + 1) % 2:
            raise ValueError(f"[{label}] 제목 병합 폭 {c2 - c1 + 1}열이 짝수가 아님 (율명·가사 쌍)")
        nlines = (c2 - c1 + 1) // 2
        if nlines > MAX_LINES:
            raise ValueError(f"[{label}] 한 페이지에 {MAX_LINES}줄 초과: {nlines}줄")
        if i < len(titles) and t + nrows >= titles[i][0]:
            raise ValueError(f"[{label}] 정간 행이 행수 {nrows} 보다 적음 — 행을 지웠다면 정보 시트 '박자 / 행수'도 "
                             f"바꾸거나 빈 행을 넣어 {nrows}행을 맞추세요 ({t + 1}행~)")
        # 행을 끼워 넣어 행수보다 아래로 밀려난 내용이 있으면 조용히 버리지 않고 오류
        stop = titles[i][0] if i < len(titles) else 10 ** 9
        over = sorted({r for (r, c), v in cells.items() if t + nrows < r < stop and v.strip()})
        if over:
            raise ValueError(f"[{label}] 행수 {nrows} 아래({over[0]}행)에도 내용이 있음 — 행을 끼워 넣었다면 "
                             f"맨 아래 빈 행을 지우거나 정보 시트 '박자 / 행수'를 바꾸세요")
        # 제목 병합 폭 밖(왼쪽·오른쪽)에 쓴 내용도 오류 — 줄을 늘렸다면 열을 병합 안쪽에 끼워 넣을 것
        outside = sorted({c for (r, c), v in cells.items() if t < r <= t + nrows and v.strip() and not c1 <= c <= c2})
        if outside:
            raise ValueError(f"[{label}] 제목 병합 범위 밖 열({outside})에 내용이 있음 — 줄을 늘리려면 "
                             f"제목 병합 안쪽에서 열 2개(율명·가사)를 끼워 넣으세요")
        page = {}
        for jg in range(1, nrows + 1):
            r = t + jg
            page[jg] = {k: (cells.get((r, c2 - 2 * k + 1), ""), cells.get((r, c2 - 2 * k + 2), ""))
                        for k in range(1, nlines + 1)}
        pages.append(page)
    return pages


def score_pages(sheets: dict, path: str, nrows: int) -> list:
    """악보 시트 → [ {정간: {열: (율명, 가사)}} ... ] (페이지 순)

    - '악보' 한 시트: 'k줄' 칸이 있는 머리 행마다 한 페이지 (위→아래 순서). 페이지 사이 빈 행/제목 행은 무시.
    - (구형) '악보 1' · '악보 2' … 시트: 시트 번호 순서대로 위와 같이 읽음.
    """
    named = []
    for name in sheets:
        if name.strip() == "악보":
            named.append((0, name))
        elif m := re.fullmatch(r"악보\s*(\d+)", name):
            named.append((int(m.group(1)), name))
    if not named:
        raise ValueError(f"'악보' 시트가 없음: {path}")
    if any(n == 0 for n, _ in named) and len(named) > 1:
        raise ValueError(f"'악보' 와 '악보 N' 시트가 함께 있음 — 하나만 남기세요: {[x for _, x in named]}")
    nums = sorted(n for n, _ in named)
    if nums != [0] and nums != list(range(1, len(nums) + 1)):
        raise ValueError(f"'악보 N' 시트 번호가 1부터 연속이 아님: {nums}")
    result = []
    for _, name in sorted(named):
        cells = sheets[name]
        heads = sorted({r for (r, c), v in cells.items() if LINE_RE.fullmatch(v.strip())})
        if not heads:
            result.extend(_parse_titled(cells, nrows, name))
            continue
        ends = heads[1:] + [10 ** 9]
        for i, (h, e) in enumerate(zip(heads, ends), 1):
            label = f"{name} {i}번째 페이지" if len(heads) > 1 else name
            result.append(_parse_block(cells, h, e, nrows, label))
    return result


BEAT_ROWS = {"3/4": {12}, "4/4": {16}, "2/4": {16}, "정악": {20}}   # 앱이 지원하는 조합


def info_beat_rows(sheets: dict):
    """정보 시트 '박자 / 행수' 칸 → (박자, 행수). 칸이 없으면 None (CSV 값 유지)."""
    cells = sheets.get("정보")
    if not cells:
        return None
    for (r, c), v in cells.items():
        if c == 1 and v.strip() == "박자 / 행수":
            raw = cells.get((r, 2), "").strip()
            m = re.fullmatch(r"(3/4|4/4|2/4|정악)\s*/\s*(\d+)\s*(?:정간)?", raw)
            if not m:
                raise ValueError(f"정보 시트 '박자 / 행수' 값 '{raw}' 를 읽을 수 없음 — 예: 3/4 / 12정간, 4/4 / 16정간")
            beat, rows = m.group(1), int(m.group(2))
            if rows not in BEAT_ROWS[beat]:
                raise ValueError(f"정보 시트 '박자 / 행수' {beat} 인데 {rows}정간 — "
                                 f"3/4=12, 4/4=16, 2/4=16, 정악=20 만 지원")
            return beat, rows
    return None


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


def build_csv(csv_path: str, xlsx_path: str) -> str:
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
    sheets = read_sheets(xlsx_path)
    # 정보 시트 '박자 / 행수'(예: "4/4 / 16정간")를 고치면 CSV 박자·행수도 바꾼다
    info = info_beat_rows(sheets)
    if info:
        beat, nrows = info
        for r in meta:
            if r and r[0].strip() == "박자" and len(r) > 1:
                r[1] = f"{beat}박자"
            elif r and r[0].strip() == "행수" and len(r) > 1:
                r[1] = str(nrows)
    pages = score_pages(sheets, xlsx_path, nrows)
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
            new = build_csv(csv_path, xlsx)
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
