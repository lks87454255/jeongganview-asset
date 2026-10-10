#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx_layout.py — 검수 xlsx '악보' 시트 모양 일괄 정리
    python3 xlsx_layout.py                 # 검수/*.xlsx 전체
    python3 xlsx_layout.py 강강술래 ...    # 특정 곡만
    python3 xlsx_layout.py --check         # 무엇이 바뀔지만 보기 (파일 변경 없음)

'악보' 시트에 적용한다 (정보 시트 설명 글은 그대로).
  1) 칸 글자 '·' → '–' (xlsx_to_csv.py 가 CSV 로 옮길 때 '–' → '―' 로 바꾼다)
     율명·가사 칸의 '‹' '⁚' '○' 는 지운다 (율명 칸은 그 기호만 있던 줄도 없앰, 가사 칸은 줄 그대로).
     정간목록 시트 율명·가사·확인 필요 글자 열도 같이.
  2) 열 너비: 정간보는 오른쪽 → 왼쪽으로 읽으므로 제목 병합 [c1..c2] 의 오른쪽 끝부터
     홀수 번째 = 가사, 짝수 번째 = 율명. 율명 열 너비 = 바로 오른쪽 가사 열 너비 × 2 (가사 열은 그대로)
  3) 행 높이: 율명이 있는 행을 LibreOffice '최적 행 높이'(자동 높이)로 맞췄을 때 가장 큰 값을
     제목 행 + 정간 행(제목 아래 행수만큼, 다음 제목 앞 빈 행 제외) 전체에 고정 높이로 적용
열려 있는 파일(.~lock / ~$)은 건너뛴다. 여러 번 실행해도 결과는 같다.
LibreOffice(/Applications/LibreOffice.app) 필요 — 자동 높이 측정에만 쓰고, 파일은 xml 만 고쳐서 저장한다.
"""
import csv
import html
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_escape

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import xlsx_to_csv as X  # noqa: E402
from xlsx_font import CELL_RE, COL_RE, FULL_CELL_RE, SI_RE, nfc, role_of, roles, set_attr  # noqa: E402

LO_PY = "/Applications/LibreOffice.app/Contents/Resources/python"
SRC, DST = "\u00b7", "\u2013"                       # '·' → '–'
ROW_RE = re.compile(r"<row\b[^>]*?(?:/>|>.*?</row>)", re.S)
ROW_TAG_RE = re.compile(r"<row\b[^>]*?/?>")


def fmt(x: float) -> str:
    return f"{x:.2f}".rstrip("0").rstrip(".")


def score_target(data: dict) -> str:
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(data["xl/_rels/workbook.xml.rels"])}
    for sh in ET.fromstring(data["xl/workbook.xml"]).find("m:sheets", X.NS):
        if nfc(sh.get("name")).strip() == "악보":
            t = rels[sh.get(f"{{{X.NS['r']}}}id")].lstrip("/")
            return t if t.startswith("xl/") else "xl/" + t
    return ""


def nrows_of(path: str, sheets: dict) -> int:
    info = X.info_beat_rows(sheets)
    if info:
        return info[1]
    name = nfc(os.path.basename(path))[:-5]
    raw = open(os.path.join(X.CSV_DIR, name + ".csv"), "rb").read().decode("utf-8-sig")
    for r in csv.reader(io.StringIO(raw, newline="")):
        if r and r[0].strip() == "행수" and len(r) > 1 and r[1].strip().isdigit():
            return int(r[1])
    raise ValueError("행수를 알 수 없음 (정보 시트 '박자 / 행수', CSV '행수' 모두 없음)")


# ───────────────────── 1) '·' → '–' (악보 칸만) ─────────────────────
def replace_dots(sheet: str, ss: str):
    """(새 sheet xml, 새 sharedStrings xml, 바꾼 칸 수). 다른 시트와 같이 쓰는 공유 문자열은 새로 추가."""
    sis = SI_RE.findall(ss) if ss else []
    added, n = [], 0

    def sub_cell(m):
        nonlocal n
        whole = m.group(0)
        tag = CELL_RE.match(whole).group(0)
        body = whole[len(tag):]
        if 't="s"' in tag:
            v = re.search(r"<v>(\d+)</v>", body)
            if not v or SRC not in sis[int(v.group(1))]:
                return whole
            new_si = sis[int(v.group(1))].replace(SRC, DST)
            all_si = sis + added
            if new_si in all_si:
                idx = all_si.index(new_si)
            else:
                added.append(new_si)
                idx = len(sis) + len(added) - 1
            n += 1
            return tag + body.replace(v.group(0), f"<v>{idx}</v>", 1)
        if SRC in body:                                    # inlineStr / str
            n += 1
            return tag + body.replace(SRC, DST)
        return whole

    sd = sheet.find("<sheetData")
    new_sheet = sheet[:sd] + FULL_CELL_RE.sub(sub_cell, sheet[sd:])
    if added:
        end = ss.rindex("</sst>")
        ss = ss[:end] + "".join(added) + ss[end:]
        head = re.search(r"<sst\b[^>]*>", ss).group(0)
        nh = set_attr(head, "uniqueCount", len(sis) + len(added)) if "uniqueCount=" in head else head
        ss = ss.replace(head, nh, 1)
    return new_sheet, ss, n


# ───────────────────── 1-2) 손 채보 기호 '‹' '⁚' 지우기 ─────────────────────
STRIP = "\u2039\u205a\u25cb"                              # ‹ ⁚ ○
T_RE = re.compile(r"(<t\b[^>]*>)(.*?)(</t>)", re.S)


def strip_marks(xml: str, seps: str) -> str:
    """<si>/<is> 안 <t> 들의 글자에서 STRIP 기호를 지우고, 그 때문에 빈 줄(구분자 seps 사이가 빈 것)이 생기면
    그 구분자도 지운다. 리치 텍스트 조각(<r>)의 서식은 그대로."""
    parts = T_RE.findall(xml)
    if not parts or not any(ch in html.unescape(p[1]) for p in parts for ch in STRIP):
        return xml
    chars = [(ch, i) for i, p in enumerate(parts) for ch in html.unescape(p[1]) if ch not in STRIP]
    out = []
    for ch, i in chars:                                    # 맨 앞 / 연속 구분자 버림
        if ch in seps and (not out or out[-1][0] in seps):
            continue
        out.append((ch, i))
    while out and out[-1][0] in seps:                      # 맨 끝 구분자 버림
        out.pop()
    texts = ["".join(ch for ch, j in out if j == i) for i in range(len(parts))]
    it = iter(texts)

    def sub_t(m):
        t = next(it)
        head = m.group(1)
        if ("\n" in t or t != t.strip()) and "xml:space" not in head:
            head = head[:-1] + ' xml:space="preserve">'
        return head + xml_escape(t) + m.group(3)
    return T_RE.sub(sub_t, xml)


def edit_cells(sheet: str, ss: str, pick, seps: str):
    """pick(r, c) 가 참인 칸에 strip_marks 적용. (새 sheet, 새 sharedStrings, 바꾼 칸 수)"""
    sis = SI_RE.findall(ss) if ss else []
    added, n = [], 0

    def sub_cell(m):
        nonlocal n
        whole = m.group(0)
        tag = CELL_RE.match(whole).group(0)
        ref = re.search(r'\br="([A-Z]+\d+)"', tag)
        if not ref or not pick(*X._ref(ref.group(1))):
            return whole
        body = whole[len(tag):]
        if 't="s"' in tag:
            v = re.search(r"<v>(\d+)</v>", body)
            if not v:
                return whole
            new_si = strip_marks(sis[int(v.group(1))], seps)
            if new_si == sis[int(v.group(1))]:
                return whole
            all_si = sis + added
            if new_si in all_si:
                idx = all_si.index(new_si)
            else:
                added.append(new_si)
                idx = len(sis) + len(added) - 1
            n += 1
            return tag + body.replace(v.group(0), f"<v>{idx}</v>", 1)
        new_body = strip_marks(body, seps) if "<is>" in body else body
        if new_body != body:
            n += 1
        return tag + new_body

    sd = sheet.find("<sheetData")
    new_sheet = sheet[:sd] + FULL_CELL_RE.sub(sub_cell, sheet[sd:])
    if added:
        end = ss.rindex("</sst>")
        ss = ss[:end] + "".join(added) + ss[end:]
        head = re.search(r"<sst\b[^>]*>", ss).group(0)
        nh = set_attr(head, "uniqueCount", len(sis) + len(added)) if "uniqueCount=" in head else head
        ss = ss.replace(head, nh, 1)
    return new_sheet, ss, n


def list_target(data: dict) -> str:
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(data["xl/_rels/workbook.xml.rels"])}
    for sh in ET.fromstring(data["xl/workbook.xml"]).find("m:sheets", X.NS):
        if nfc(sh.get("name")).strip() == "정간목록":
            t = rels[sh.get(f"{{{X.NS['r']}}}id")].lstrip("/")
            return t if t.startswith("xl/") else "xl/" + t
    return ""


LIST_COLS = ("율명(판독)", "율명(앱)", "확인 필요 글자")         # 정간목록 시트에서 기호를 지울 열 (비고 등 설명 글은 그대로)


# ───────────────────── 2) 율명 열 = 가사 열 × 2 ─────────────────────
def fix_widths(sheet: str, pages) -> str:
    fmt_pr = re.search(r"<sheetFormatPr\b[^>]*>", sheet)
    dflt = 8.43
    if fmt_pr and (m := re.search(r'defaultColWidth="([\d.]+)"', fmt_pr.group(0))):
        dflt = float(m.group(1))
    cm = re.search(r"<cols>(.*?)</cols>", sheet, re.S)
    per = {}                                               # 열 번호 → <col> 태그 (min=max=열)
    if cm:
        for tag in COL_RE.findall(cm.group(1)):
            lo = int(re.search(r'\bmin="(\d+)"', tag).group(1))
            hi = int(re.search(r'\bmax="(\d+)"', tag).group(1))
            for c in range(lo, hi + 1):
                per[c] = set_attr(set_attr(tag, "min", c), "max", c)

    def width(c):
        if c in per and (m := re.search(r'\bwidth="([\d.]+)"', per[c])):
            return float(m.group(1))
        return dflt

    want = {}
    for _, _, c1, c2 in pages:
        for c in range(c1, c2 + 1):
            if (c2 - c) % 2 == 1:                          # 오른쪽 끝부터 짝수 번째 = 율명
                want[c] = round(width(c + 1) * 2, 4)
    for c, w in want.items():
        tag = per.get(c, f'<col min="{c}" max="{c}"/>')
        flag = "true" if re.search(r'customWidth="(true|false)"', tag) else "1"
        per[c] = set_attr(set_attr(tag, "width", fmt(w)), "customWidth", flag)
    cols = "<cols>" + "".join(per[c] for c in sorted(per)) + "</cols>"
    if cm:
        return sheet[:cm.start()] + cols + sheet[cm.end():]
    sd = sheet.find("<sheetData")
    return sheet[:sd] + cols + sheet[sd:]


# ───────────────────── 3) 행 높이 ─────────────────────
def layout_rows(cells, pages, nrows):
    """(율명이 있는 행, 높이를 맞출 행: 제목 + 정간)"""
    yul, target = set(), set()
    filled = {r for (r, c), v in cells.items() if v.strip()}
    for t, stop, c1, c2 in pages:
        target.add(t)
        last = min(t + nrows, stop - 1)
        if stop < 10 ** 9 and stop - 1 > t and stop - 1 not in filled:   # 다음 제목 앞 빈 행(페이지 구분)은 제외
            last = min(last, stop - 2)
        target.update(range(t + 1, last + 1))
        for (r, c), v in cells.items():
            if t < r <= last and c1 <= c <= c2 and (c2 - c) % 2 == 1 and v.strip():
                yul.add(r)
    return sorted(yul), sorted(target)


def set_heights(sheet: str, rows, ht: float) -> str:
    sd_open = re.search(r"<sheetData\b[^>]*?(/?)>", sheet)
    if sd_open.group(1):                                   # <sheetData/>
        sheet = sheet[:sd_open.start()] + "<sheetData></sheetData>" + sheet[sd_open.end():]
        sd_open = re.search(r"<sheetData\b[^>]*>", sheet)
    start = sd_open.end()
    end = sheet.index("</sheetData>", start)
    existing = {}
    for m in ROW_RE.finditer(sheet[start:end]):
        existing[int(re.search(r'\br="(\d+)"', m.group(0)).group(1))] = m.group(0)
    style_true = any('customHeight="true"' in r or 'customHeight="false"' in r for r in existing.values())
    for r in rows:
        el = existing.get(r, f'<row r="{r}"/>')
        tag = ROW_TAG_RE.match(el).group(0)
        nt = set_attr(set_attr(tag, "ht", fmt(ht)), "customHeight", "true" if style_true else "1")
        existing[r] = nt + el[len(tag):]
    body = "".join(existing[r] for r in sorted(existing))
    return sheet[:start] + body + sheet[end:]


def measure(jobs: list) -> dict:
    with tempfile.TemporaryDirectory() as td:
        jin, jout = os.path.join(td, "in.json"), os.path.join(td, "out.json")
        json.dump(jobs, open(jin, "w", encoding="utf-8"), ensure_ascii=False)
        subprocess.run([LO_PY, os.path.join(HERE, "lo_row_height.py"), jin, jout], check=True)
        return json.load(open(jout, encoding="utf-8"))


def write_zip(path: str, items, new_data: dict, dest: str):
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(dest))
    os.close(fd)
    with zipfile.ZipFile(tmp, "w") as zout:
        for info, d in items:
            zout.writestr(info, new_data.get(info.filename, d), compress_type=info.compress_type)
    if os.path.exists(dest):
        shutil.copystat(dest, tmp)
    os.replace(tmp, dest)


def main():
    check = "--check" in sys.argv
    want = {nfc(a[:-5] if a.lower().endswith(".xlsx") else a) for a in sys.argv[1:] if not a.startswith("--")}
    D = X.XLSX_DIR
    files = sorted(f for f in os.listdir(D) if f.lower().endswith(".xlsx") and not f.startswith(("~$", ".", "_")))
    lo_locks = {nfc(f[len(".~lock."):-1]) for f in os.listdir(D) if f.startswith(".~lock.")}
    ms_locks = {nfc(f[2:]) for f in os.listdir(D) if f.startswith("~$")}
    work, errors, skipped = [], 0, 0
    tmpdir = tempfile.mkdtemp(prefix="jgv_layout_")
    try:
        # 1·2단계 → 임시 파일 (자동 높이는 바뀐 열 너비 기준으로 잰다)
        for f in files:
            name = nfc(f)
            if want and name[:-5] not in want:
                continue
            if name in lo_locks or any(name == x or name[2:] == x or name[1:] == x for x in ms_locks):
                print(f"  ⚠️  열려 있어 건너뜀: {name}")
                skipped += 1
                continue
            path = os.path.join(D, f)
            try:
                with zipfile.ZipFile(path) as z:
                    items = [(i, z.read(i.filename)) for i in z.infolist()]
                data = {i.filename: d for i, d in items}
                target = score_target(data)
                if not target:
                    continue
                sheets = X.read_sheets(path)
                cells = next(v for k, v in sheets.items() if nfc(k).strip() == "악보")
                pages = roles(cells)
                if not pages:
                    raise ValueError("악보 시트에 제목(가로 병합) 행이 없음")
                yul, rows = layout_rows(cells, pages, nrows_of(path, sheets))
                if not yul:
                    raise ValueError("율명이 있는 행이 없음")
                ss = data.get("xl/sharedStrings.xml", b"").decode("utf-8")
                sheet, ss, ndot = replace_dots(data[target].decode("utf-8"), ss)
                yul_cells = {k for k, v in cells.items() if role_of(pages, *k) == "yul"}
                sheet, ss, nmark = edit_cells(sheet, ss, lambda r, c: (r, c) in yul_cells, "\n")
                lyric_cells = {k for k in cells if role_of(pages, *k) == "lyric"}   # 가사 칸: 빈 줄은 일부러 둔 것 → 줄은 그대로
                sheet, ss, nl = edit_cells(sheet, ss, lambda r, c: (r, c) in lyric_cells, "")
                nmark += nl
                new_data = {}
                lt = list_target(data)
                if lt and (lcells := next((v for k, v in sheets.items() if nfc(k).strip() == "정간목록"), None)):
                    cols = {c for (r, c), v in lcells.items() if r == 1 and v.strip() in LIST_COLS}
                    lsheet, ss, nl = edit_cells(data[lt].decode("utf-8"), ss, lambda r, c: r > 1 and c in cols, "\n|")
                    nmark += nl
                    lyr = {c for (r, c), v in lcells.items() if r == 1 and v.strip() == "가사"}
                    lsheet, ss, nl = edit_cells(lsheet, ss, lambda r, c: r > 1 and c in lyr, "")
                    nmark += nl
                    new_data[lt] = lsheet.encode("utf-8")
                sheet = fix_widths(sheet, pages)
                new_data[target] = sheet.encode("utf-8")
                if ss:
                    new_data["xl/sharedStrings.xml"] = ss.encode("utf-8")
                tmp = os.path.join(tmpdir, f"{len(work)}.xlsx")
                write_zip(path, items, new_data, tmp)
                work.append(dict(name=name, path=path, tmp=tmp, items=items, data=data, target=target,
                                 new=new_data, yul=yul, rows=rows, ndot=ndot, nmark=nmark))
            except Exception as e:  # noqa: BLE001
                errors += 1
                print(f"  ❌ {name}: {e}")
        if not work:
            print(f"대상 없음 (건너뜀 {skipped}, 오류 {errors})")
            sys.exit(1 if errors else 0)
        # 3단계: LibreOffice 로 율명 행 자동 높이 측정 → 최댓값을 제목·정간 행 전체에
        hts = measure([{"path": w["tmp"], "sheet": "악보", "rows": w["yul"]} for w in work])
        changed = 0
        for w in work:
            ht = max(hts[w["tmp"]].values())
            sheet = set_heights(w["new"][w["target"]].decode("utf-8"), w["rows"], ht)
            w["new"][w["target"]] = sheet.encode("utf-8")
            diff = any(w["data"].get(k) != v for k, v in w["new"].items())
            print(f"  {'✏️ ' if diff else '  '} {w['name']}: 행 높이 {fmt(ht)}pt × {len(w['rows'])}행, '·'→'–' {w['ndot']}칸, '‹⁚○' 지움 {w['nmark']}칸")
            if diff:
                changed += 1
                if not check:
                    write_zip(w["path"], w["items"], w["new"], w["path"])
        verb = "바뀔" if check else "변경"
        print(f"{len(work)}개 확인 / {verb} {changed}개 / 건너뜀 {skipped}개 / 오류 {errors}개")
        sys.exit(1 if errors else 0)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    main()
