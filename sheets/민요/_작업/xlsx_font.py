#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx_font.py — 검수 xlsx 글꼴 일괄 변경
    python3 xlsx_font.py                 # 검수/*.xlsx 전체
    python3 xlsx_font.py 강강술래 ...    # 특정 곡만

- 모든 글꼴 이름(기본 글꼴 포함) → FONT_NAME
- '악보' 시트 크기: 제목(가로 병합 행) TITLE_SIZE, 율명 열 YUL_SIZE, 가사 열 LYRIC_SIZE
  (제목 병합 범위 [c1..c2] 에서 오른쪽 끝부터 율명·가사 쌍 — xlsx_to_csv.py 규칙과 같음)
  해당 칸의 서식(xf)을 복제해 글꼴만 바꾼다 → 테두리·채우기·정렬·굵게는 그대로.
- 열려 있는 파일(.~lock / ~$)은 건너뛴다. 여러 번 실행해도 결과는 같다.
"""
import os
import re
import shutil
import sys
import tempfile
import unicodedata
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsx_to_csv as X  # noqa: E402

FONT_NAME = "궁서체"
TITLE_SIZE, YUL_SIZE, LYRIC_SIZE = 24, 24, 12

XLSX_DIR = X.XLSX_DIR
FONTS_RE = re.compile(r"<fonts\b[^>]*>(.*?)</fonts>", re.S)
FONT_RE = re.compile(r"<font\b[^>]*?(?:/>|>.*?</font>)", re.S)
XFS_RE = re.compile(r"<cellXfs\b[^>]*>(.*?)</cellXfs>", re.S)
XF_RE = re.compile(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", re.S)
CELL_RE = re.compile(r"<c\b[^>]*?/?>")
COL_RE = re.compile(r"<col\b[^>]*?/>")
FULL_CELL_RE = re.compile(r"<c\b[^>]*?(?:/>|>.*?</c>)", re.S)
SI_RE = re.compile(r"<si\b[^>]*?(?:/>|>.*?</si>)", re.S)
RPR_RE = re.compile(r"<rPr>.*?</rPr>", re.S)


def rename_runs(xml: str) -> str:
    """리치 텍스트 조각(<rPr>)의 글꼴 이름 → FONT_NAME (칸 서식보다 우선하므로)."""
    return re.sub(r'<rFont val="[^"]*"/>', f'<rFont val="{FONT_NAME}"/>', xml)


def strip_run_size(xml: str) -> str:
    """악보 칸의 리치 텍스트 조각 글자 크기를 지워 칸 서식(제목·율명·가사 크기)을 따르게 한다."""
    return RPR_RE.sub(lambda m: re.sub(r'<sz val="[^"]*"/>', "", m.group(0)), xml)


def nfc(s):
    return unicodedata.normalize("NFC", s)


def set_attr(tag: str, name: str, val) -> str:
    """여는 태그 문자열에 속성 값 설정 (없으면 추가)."""
    if re.search(rf'\b{name}="[^"]*"', tag):
        return re.sub(rf'\b{name}="[^"]*"', f'{name}="{val}"', tag, count=1)
    end = 2 if tag.endswith("/>") else 1
    return tag[:-end] + f' {name}="{val}"' + tag[-end:]


def rename_font(font: str) -> str:
    font = re.sub(r"<scheme\b[^>]*/>", "", font)            # 테마 글꼴 연결 끊기 (이름이 우선되도록)
    if "<name " in font:
        return re.sub(r'<name val="[^"]*"/>', f'<name val="{FONT_NAME}"/>', font)
    if font.endswith("/>"):
        return f'<font><name val="{FONT_NAME}"/></font>'
    return font.replace("</font>", f'<name val="{FONT_NAME}"/></font>')


def resize_font(font: str, size: int) -> str:
    if "<sz " in font:
        return re.sub(r'<sz val="[^"]*"/>', f'<sz val="{size}"/>', font)
    return font.replace("<name ", f'<sz val="{size}"/><name ', 1)


def roles(cells) -> dict:
    """제목 병합 행마다 (제목 행, 다음 제목 행, c1, c2)."""
    titles = sorted((r1, c1, c2) for r1, c1, r2, c2 in cells.merges if r1 == r2 and c2 > c1)
    pages = []
    for i, (t, c1, c2) in enumerate(titles):
        stop = titles[i + 1][0] if i + 1 < len(titles) else 10 ** 9
        pages.append((t, stop, c1, c2))
    return pages


def role_of(pages, r, c):
    for t, stop, c1, c2 in pages:
        if not c1 <= c <= c2:
            continue
        if r == t:
            return "title"
        if t < r < stop:
            return "yul" if (c2 - c) % 2 else "lyric"
    return None


def col_role(pages, c):
    roles_ = {("yul" if (c2 - c) % 2 else "lyric") for _, _, c1, c2 in pages if c1 <= c <= c2}
    return roles_.pop() if len(roles_) == 1 else None


def convert(path: str) -> int:
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info.filename)) for info in z.infolist()]
    data = {info.filename: d for info, d in items}
    sheets = X.read_sheets(path)
    has_score = "악보" in sheets                            # _목록.xlsx 처럼 악보 시트가 없으면 글꼴 이름만
    pages = roles(sheets["악보"]) if has_score else []
    if has_score and not pages:
        raise ValueError("악보 시트에 제목(가로 병합) 행이 없음")
    # 악보 시트 xml 경로
    import xml.etree.ElementTree as ET
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(data["xl/_rels/workbook.xml.rels"])}
    target = None
    for sh in ET.fromstring(data["xl/workbook.xml"]).find("m:sheets", X.NS):
        if nfc(sh.get("name")).strip() == "악보":
            target = rels[sh.get(f"{{{X.NS['r']}}}id")].lstrip("/")
            target = target if target.startswith("xl/") else "xl/" + target
    styles = data["xl/styles.xml"].decode("utf-8")
    sheet = data[target].decode("utf-8") if target else ""

    # 1) 모든 글꼴 이름 변경
    fm = FONTS_RE.search(styles)
    fonts = [rename_font(f) for f in FONT_RE.findall(fm.group(1))]
    xm = XFS_RE.search(styles)
    xfs = XF_RE.findall(xm.group(1))
    size_of = {"title": TITLE_SIZE, "yul": YUL_SIZE, "lyric": LYRIC_SIZE}
    font_cache, xf_cache = {}, {}

    def font_for(fid: int, size: int) -> int:
        key = (fid, size)
        if key not in font_cache:
            new = resize_font(fonts[fid], size)
            if new in fonts:                                   # 같은 글꼴이 이미 있으면 재사용
                font_cache[key] = fonts.index(new)
            else:
                fonts.append(new)
                font_cache[key] = len(fonts) - 1
        return font_cache[key]

    def xf_for(s: int, role: str) -> int:
        key = (s, role)
        if key not in xf_cache:
            xf = xfs[s]
            head = re.match(r"<xf\b[^>]*?/?>", xf).group(0)
            fid = int(re.search(r'fontId="(\d+)"', head).group(1)) if "fontId=" in head else 0
            nfid = font_for(fid, size_of[role])
            new_head = set_attr(set_attr(head, "fontId", nfid), "applyFont", "1")
            new = new_head + xf[len(head):]
            if new == xf:
                xf_cache[key] = s
            elif new in xfs:
                xf_cache[key] = xfs.index(new)
            else:
                xfs.append(new)
                xf_cache[key] = len(xfs) - 1
        return xf_cache[key]

    # 2) 악보 시트 칸 서식 바꾸기
    shared_in_score = set()                                 # 악보 칸이 쓰는 공유 문자열 번호

    def sub_cell(m):
        whole = m.group(0)
        tag = CELL_RE.match(whole).group(0)
        ref = re.search(r'\br="([A-Z]+\d+)"', tag)
        if not ref:
            return whole
        r, c = X._ref(ref.group(1))
        role = role_of(pages, r, c)
        if not role:
            return whole
        s = int(re.search(r'\bs="(\d+)"', tag).group(1)) if re.search(r'\bs="', tag) else 0
        body = whole[len(tag):]
        if 't="s"' in tag and (v := re.search(r"<v>(\d+)</v>", body)):
            shared_in_score.add(int(v.group(1)))
        return set_attr(tag, "s", xf_for(s, role)) + strip_run_size(body)

    def sub_col(m):
        tag = m.group(0)
        if 'style="' not in tag:
            return tag
        lo, hi = int(re.search(r'min="(\d+)"', tag).group(1)), int(re.search(r'max="(\d+)"', tag).group(1))
        rs = {col_role(pages, c) for c in range(lo, hi + 1)}
        if len(rs) != 1 or None in rs:
            return tag
        return set_attr(tag, "style", xf_for(int(re.search(r'style="(\d+)"', tag).group(1)), rs.pop()))

    sd = sheet.find("<sheetData") if target else 0
    new_sheet = COL_RE.sub(sub_col, sheet[:sd]) + FULL_CELL_RE.sub(sub_cell, sheet[sd:])

    # 3) styles.xml 다시 조립
    fonts_xml = re.sub(r'count="\d+"', f'count="{len(fonts)}"', styles[fm.start():fm.start(1)], count=1)
    styles = styles[:fm.start()] + fonts_xml + "".join(fonts) + "</fonts>" + styles[fm.end():]
    xm = XFS_RE.search(styles)
    xfs_xml = re.sub(r'count="\d+"', f'count="{len(xfs)}"', styles[xm.start():xm.start(1)], count=1)
    styles = styles[:xm.start()] + xfs_xml + "".join(xfs) + "</cellXfs>" + styles[xm.end():]

    new_data = {"xl/styles.xml": styles.encode("utf-8")}
    if target:
        new_data[target] = new_sheet.encode("utf-8")
    # 4) 리치 텍스트: 모든 시트·공유 문자열 글꼴 이름, 악보 칸이 쓰는 공유 문자열은 크기 제거
    if "xl/sharedStrings.xml" in data:
        ss = rename_runs(data["xl/sharedStrings.xml"].decode("utf-8"))
        idx = iter(range(10 ** 9))
        ss = SI_RE.sub(lambda m: strip_run_size(m.group(0)) if next(idx) in shared_in_score else m.group(0), ss)
        new_data["xl/sharedStrings.xml"] = ss.encode("utf-8")
    for k in data:
        if k.startswith("xl/worksheets/") and k.endswith(".xml"):
            new_data[k] = rename_runs(new_data[k].decode("utf-8") if k in new_data else data[k].decode("utf-8")).encode("utf-8")
    if all(data[k] == v for k, v in new_data.items()):
        return 0
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(path))
    os.close(fd)
    with zipfile.ZipFile(tmp, "w") as zout:
        for info, d in items:
            zout.writestr(info, new_data.get(info.filename, d), compress_type=info.compress_type)
    shutil.copystat(path, tmp)
    os.replace(tmp, path)
    return 1


def main():
    want = {nfc(a[:-5] if a.lower().endswith(".xlsx") else a) for a in sys.argv[1:] if not a.startswith("--")}
    files = sorted(f for f in os.listdir(XLSX_DIR)
                   if f.lower().endswith(".xlsx") and not f.startswith(("~$", ".")))
    lo_locks = {nfc(f[len(".~lock."):-1]) for f in os.listdir(XLSX_DIR) if f.startswith(".~lock.")}
    ms_locks = {nfc(f[2:]) for f in os.listdir(XLSX_DIR) if f.startswith("~$")}
    changed, skipped, errors = 0, 0, 0
    for f in files:
        name = nfc(f)
        if want and name[:-5] not in want:
            continue
        if name in lo_locks or any(name == x or name[2:] == x or name[1:] == x for x in ms_locks):
            print(f"  ⚠️  열려 있어 건너뜀: {name}")
            skipped += 1
            continue
        try:
            changed += convert(os.path.join(XLSX_DIR, f))
        except Exception as e:  # noqa: BLE001
            errors += 1
            print(f"  ❌ {name}: {e}")
    print(f"글꼴 {FONT_NAME} / 제목 {TITLE_SIZE} · 율명 {YUL_SIZE} · 가사 {LYRIC_SIZE}: "
          f"{changed}개 파일 변경, {skipped}개 건너뜀, 오류 {errors}개")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
