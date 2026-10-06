#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
merge_score_sheets.py — 검수 엑셀의 악보 시트를 '인쇄용 악보 한 시트' 형식으로 바꿈
====================================================================================
    python3 merge_score_sheets.py            # 검수/*.xlsx 중 바꿀 것만 변환
    python3 merge_score_sheets.py --check    # 변환 대상만 출력
    python3 merge_score_sheets.py "새 타령"  # 특정 곡만

바꾸는 내용 (build_all.ps1 이 새로 만드는 형식과 같음)
  - '악보 1' · '악보 2' … 시트를 '악보' 한 시트로 합침 (페이지 사이 빈 행 1개)
  - A열(정간 번호) 삭제
  - 오른쪽 맞춤: 1줄이 항상 맨 오른쪽 두 열 (줄이 적은 페이지도 오른쪽부터). 열은 율명·가사 순 쌍
  - 머리 행('10줄 | 가사 | … | 1줄 | 가사') 삭제 — 제목 행 바로 아래 n번째 행 = 정간 n
  - 제목·정간 칸 테두리 — 줄이 덜 채워진 페이지도 페이지 폭 전체. 1행·A열은 여백(비움) — 페이지마다 [빈 행][제목][정간 …]
  - 3/4 박자는 3정간, 4/4 박자는 4정간 묶음으로 가사 열 안쪽 가로선을 흰색 (박자는 csv/{곡}.csv 기준)
  - 율명 글자 22pt · 가사 글자 13pt · 율명 안 '―'(장음)는 11pt 작은 글씨, 내용 행 높이는 줄 수와 무관하게 고정 56pt
  - 제목 행: 전체 폭 병합 · 가운데 · 곡 제목만 (원본 이미지 이름은 '정보' 시트에 있음)
  - 인쇄: A4 세로 · 가운데 · 페이지마다 새 종이(페이지 나누기) · 배율 고정(가장 넓은/긴 페이지가 한 장에 들어가게)
  - 제목 행 높이 60pt = 정간 행 기본(30pt)의 두 배
  - 셀 값·스타일(노란/주황 칸)·행 높이는 그대로. 틀 고정은 없앰
  - 엑셀/LibreOffice 로 열려 있는 파일(~$ · .~lock 있음)은 건너뜀 — 닫고 다시 실행
"""
import copy
import os
import re
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX_DIR = os.path.join(os.path.dirname(HERE), "검수")
M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", M)
ET.register_namespace("r", R)
LINE_RE = re.compile(r"\d+\s*줄")
CSV_DIR = os.path.join(os.path.dirname(HERE), "csv")
# 박자별 묶음(정간 수): 가사 열은 묶음 안쪽 가로선을 흰색으로 → 박 단위로 한 칸처럼 보임
BEAT_GROUP = {"3/4": 3, "4/4": 4}
WHITE = "FFFFFFFF"
LYRIC_PT = 13                        # 가사 글자 크기 (조금 더 크게)
PITCH_PT = 22                        # 율명 글자 크기 (가사보다 큼)
DASH_CH = "―"                        # 장음(이음) 기호 — 작은 글씨로
DASH_PT = 11                         # 장음 기호 글자 크기 = 작게
CONTENT_HT = 56                      # 내용(정간) 행 높이(pt) — 율명 줄 수와 무관하게 고정
ROW_HT = 30                          # 정간 행 기본 높이(pt)
TITLE_HT = ROW_HT * 2                # 제목 행 높이 = 일반(정간) 행의 두 배
# 인쇄: A4 세로(595×842pt) − 여백(좌우 0.4in, 위아래 0.5in + 머리말/꼬리말 여유)
PRINT_W, PRINT_H = 595 - 0.8 * 72, 842 - 1.0 * 72 - 20


def nfc(s):
    return unicodedata.normalize("NFC", s)


def q(tag, ns=M):
    return f"{{{ns}}}{tag}"


def col_num(ref):
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n):
    s = ""
    while n:
        n, m = divmod(n - 1, 26)
        s = chr(65 + m) + s
    return s


def set_text(c, txt):
    """셀을 inlineStr 문자열로 바꾼다 (스타일 s 는 유지)."""
    for ch in list(c):
        c.remove(ch)
    c.attrib.pop("t", None)
    c.set("t", "inlineStr")
    t = ET.SubElement(ET.SubElement(c, q("is")), q("t"))
    t.text = txt
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


class Book:
    def __init__(self, path):
        self.path = path
        self.z = zipfile.ZipFile(path)
        self.wb = ET.fromstring(self.z.read("xl/workbook.xml"))
        self.rels = ET.fromstring(self.z.read("xl/_rels/workbook.xml.rels"))
        self.rid2target = {r.get("Id"): r.get("Target") for r in self.rels}
        self.shared = []
        if "xl/sharedStrings.xml" in self.z.namelist():
            for si in ET.fromstring(self.z.read("xl/sharedStrings.xml")).findall(q("si")):
                self.shared.append("".join(t.text or "" for t in si.iter(q("t"))))

    def _styles(self):
        if getattr(self, "styles", None) is None:
            self.styles = ET.fromstring(self.z.read("xl/styles.xml"))
        return self.styles

    def _thin_border(self):
        """얇은 테두리(상하좌우) 번호 — styles.xml 에 추가."""
        if getattr(self, "_border", None) is None:
            borders = self._styles().find(q("borders"))
            b = ET.Element(q("border"))
            for side in ("left", "right", "top", "bottom"):
                ET.SubElement(ET.SubElement(b, q(side), {"style": "thin"}), q("color"), {"auto": "1"})
            ET.SubElement(b, q("diagonal"))
            self._border = self._find_or_add(borders, b)
        return self._border

    def _add_xf(self, font_id, wrap):
        xfs = self._styles().find(q("cellXfs"))
        xf = ET.Element(q("xf"), {"numFmtId": "0", "fontId": str(font_id), "fillId": "0",
                                          "borderId": str(self._thin_border()), "xfId": "0",
                                          "applyFont": "1", "applyBorder": "1", "applyAlignment": "1"})
        al = {"horizontal": "center", "vertical": "center"}
        if wrap:
            al["wrapText"] = "1"
        ET.SubElement(xf, q("alignment"), al)
        return self._find_or_add(xfs, xf)

    def title_style(self):
        """제목: 굵게 14pt · 가운데 · 테두리."""
        if getattr(self, "_title_style", None) is None:
            fonts = self._styles().find(q("fonts"))
            font = ET.SubElement(fonts, q("font"))
            ET.SubElement(font, q("b"))
            ET.SubElement(font, q("sz"), {"val": "14"})
            ET.SubElement(font, q("name"), {"val": "맑은 고딕"})
            fonts.set("count", str(len(fonts)))
            self._title_style = self._add_xf(len(fonts) - 1, wrap=False)
        return self._title_style

    def grid_style(self):
        """빈 정간 칸: 기본 글꼴 · 가운데 · 줄바꿈 · 테두리."""
        if getattr(self, "_grid_style", None) is None:
            self._grid_style = self._add_xf(0, wrap=True)
        return self._grid_style

    def _find_or_add(self, parent, el):
        key = ET.tostring(el)
        for i, x in enumerate(parent):
            if ET.tostring(x) == key:
                return i
        parent.append(el)
        parent.set("count", str(len(parent)))
        return len(parent) - 1

    def border_state(self, s):
        """스타일 s 의 (위 흰색?, 아래 흰색?) — 이미 원하는 상태면 다시 만들지 않기 위해."""
        st = self._styles()
        xf = st.find(q("cellXfs"))[int(s)]
        border = st.find(q("borders"))[int(xf.get("borderId", "0"))]
        def white(side):
            el = border.find(q(side))
            col = el.find(q("color")) if el is not None else None
            return col is not None and (col.get("rgb") or "").upper().endswith("FFFFFF")
        return white("top"), white("bottom")

    def font_size(self, s):
        st = self._styles()
        xf = st.find(q("cellXfs"))[int(s)]
        sz = st.find(q("fonts"))[int(xf.get("fontId", "0"))].find(q("sz"))
        return float(sz.get("val")) if sz is not None else None

    def sized_style(self, s, pt):
        """스타일 s 를 바탕으로 글자 크기만 pt 로 바꾼 스타일 번호."""
        cache = self.__dict__.setdefault("_size_cache", {})
        if (s, pt) in cache:
            return cache[(s, pt)]
        st = self._styles()
        xfs, fonts = st.find(q("cellXfs")), st.find(q("fonts"))
        xf = copy.deepcopy(xfs[int(s)])
        font = copy.deepcopy(fonts[int(xf.get("fontId", "0"))])
        sz = font.find(q("sz"))
        if sz is None:
            sz = ET.Element(q("sz"))
            font.insert(0, sz)
        sz.set("val", f"{pt:g}")
        xf.set("fontId", str(self._find_or_add(fonts, font)))
        xf.set("applyFont", "1")
        cache[(s, pt)] = self._find_or_add(xfs, xf)
        return cache[(s, pt)]

    def _font_name(self, s):
        """스타일 s 의 글꼴 이름(없으면 기본)."""
        st = self._styles()
        xf = st.find(q("cellXfs"))[int(s)]
        name = st.find(q("fonts"))[int(xf.get("fontId", "0"))].find(q("name"))
        return name.get("val") if name is not None else "맑은 고딕"

    def dash_runs(self, c, big_pt, small_pt):
        """셀 안 '―' 는 작은 글씨, 나머지 율명 글자는 큰 글씨로.
        '―' 가 있으면 셀을 리치텍스트(inlineStr)로 다시 쓰고 True, 없으면 그대로 False."""
        txt = self.text(c)
        if DASH_CH not in txt:
            return False
        name = self._font_name(c.get("s") or "0")
        # '―' 를 경계로 조각내되 '―' 자체도 조각으로 보존
        parts = re.split(f"({re.escape(DASH_CH)}+)", txt)
        parts = [p for p in parts if p != ""]
        # 이미 원하는 리치텍스트 형태면 그대로 둠 (멱등성)
        if c.get("t") == "inlineStr":
            runs = c.findall(f"{q('is')}/{q('r')}")
            cur = [("".join(t.text or "" for t in r.iter(q("t"))),
                    (r.find(f"{q('rPr')}/{q('sz')}").get("val") if r.find(f"{q('rPr')}/{q('sz')}") is not None else None))
                   for r in runs]
            want = [(p, f"{(small_pt if p[0] == DASH_CH else big_pt):g}") for p in parts]
            if cur == want:
                return False
        # 기존 셀 내용 비우고 inlineStr 리치텍스트로 교체
        c.set("t", "inlineStr")
        for child in list(c):
            c.remove(child)
        is_el = ET.SubElement(c, q("is"))
        for p in parts:
            r = ET.SubElement(is_el, q("r"))
            rpr = ET.SubElement(r, q("rPr"))
            ET.SubElement(rpr, q("sz"), {"val": f"{(small_pt if p[0] == DASH_CH else big_pt):g}"})
            ET.SubElement(rpr, q("rFont"), {"val": name})
            t = ET.SubElement(r, q("t"), {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
            t.text = p
        return True

    def beat_style(self, s, top_white, bottom_white):
        """스타일 s 를 바탕으로 위/아래 테두리만 흰색(또는 검정)으로 바꾼 스타일 번호."""
        key = (s, top_white, bottom_white)
        cache = self.__dict__.setdefault("_beat_cache", {})
        if key in cache:
            return cache[key]
        st = self._styles()
        xfs, borders = st.find(q("cellXfs")), st.find(q("borders"))
        xf = copy.deepcopy(xfs[int(s)])
        border = copy.deepcopy(borders[int(xf.get("borderId", "0"))])
        for side, white in (("top", top_white), ("bottom", bottom_white)):
            el = border.find(q(side))
            idx = list(border).index(el)
            border.remove(el)
            new = ET.Element(q(side), {"style": "thin"})
            ET.SubElement(new, q("color"), {"rgb": WHITE} if white else {"auto": "1"})
            border.insert(idx, new)
        for side in ("left", "right"):                                # 좌우는 항상 검정 얇은 선
            el = border.find(q(side))
            if el is None or el.get("style") is None:
                idx = list(border).index(el) if el is not None else 0
                if el is not None:
                    border.remove(el)
                new = ET.Element(q(side), {"style": "thin"})
                ET.SubElement(new, q("color"), {"auto": "1"})
                border.insert(idx, new)
        xf.set("borderId", str(self._find_or_add(borders, border)))
        xf.set("applyBorder", "1")
        cache[key] = self._find_or_add(xfs, xf)
        return cache[key]

    def part(self, sh):
        t = self.rid2target[sh.get(q("id", R))].lstrip("/")
        return t if t.startswith("xl/") else "xl/" + t

    def text(self, c):
        """셀 표시 문자열 (판별용)"""
        t = c.get("t")
        if t == "inlineStr":
            return "".join(x.text or "" for x in c.iter(q("t")))
        v = c.find(q("v"))
        if v is None or v.text is None:
            return ""
        return self.shared[int(v.text)] if t == "s" else v.text


def score_sheets(book):
    out = []
    for sh in book.wb.find(q("sheets")):
        name = nfc(sh.get("name"))
        if name.strip() == "악보":
            out.append((0, sh))
        elif m := re.fullmatch(r"악보\s*(\d+)", name):
            out.append((int(m.group(1)), sh))
    return [sh for _, sh in sorted(out, key=lambda x: x[0])]


def needs_convert(book, sheets):
    if len(sheets) != 1:
        return True
    ws = ET.fromstring(book.z.read(book.part(sheets[0])))
    if ws.find(q("mergeCells")) is None:
        return True                                           # 제목 행 병합 전(왼쪽 맞춤) 형식
    if not any(m.get("ref", "").startswith("B") for m in ws.find(q("mergeCells"))):
        return True                                           # 여백(A열) 전 형식
    for row in ws.find(q("sheetData")):
        if any(LINE_RE.fullmatch(book.text(c).strip()) for c in row):
            return True                                       # 머리 행('k줄') 이 남아 있음
    for row in ws.find(q("sheetData")):
        for c in row:
            if c.get("r", "").startswith("A") and re.fullmatch(r"A\d+", c.get("r")) and book.text(c).strip() == "정간":
                return True                                   # A열 정간 번호 열이 남아 있음
    return False


def convert(book, sheets):
    rows_out = []                                             # (row 속성 dict, [cell element])
    widths = {}
    breaks = []

    for si, sh in enumerate(sheets):
        ws = ET.fromstring(book.z.read(book.part(sh)))
        cols = ws.find(q("cols"))
        if cols is not None:
            for c in cols:
                for i in range(int(c.get("min")), int(c.get("max")) + 1):
                    if i >= 2:                                # A열 너비는 버림 → 한 칸 왼쪽
                        widths[i - 1] = max(widths.get(i - 1, 0), float(c.get("width", 0)))
        src_rows = sorted(ws.find(q("sheetData")), key=lambda r: int(r.get("r")))
        # 이 시트에 A열 정간 번호 열('정간' 머리)이 있을 때만 A열을 지운다
        has_num_col = any(col_num(c.get("r")) == 1 and book.text(c).strip() == "정간"
                          for row in src_rows for c in row)
        # 행 종류 판별: 머리 행(A='정간'), 정간 행(A=숫자), 그 밖(제목·빈 행)
        if si > 0:
            rows_out.append(({}, []))                         # 시트 사이 빈 행
        prev_r = None
        for row in src_rows:
            rn = int(row.get("r"))
            if prev_r is not None:
                for _ in range(rn - prev_r - 1):              # 원본의 빈 행 유지
                    rows_out.append(({}, []))
            prev_r = rn
            cells = list(row)
            a = next((c for c in cells if col_num(c.get("r")) == 1), None)
            a_txt = book.text(a).strip() if a is not None else ""
            is_grid = a_txt == "정간" or a_txt.isdigit()
            if not has_num_col:
                is_grid = None                                # 이미 번호 열 없음 → 열 그대로
            new_cells = []
            for c in cells:
                n = col_num(c.get("r"))
                c = copy.deepcopy(c)
                if is_grid is None:
                    c.set("_col", str(n))
                elif is_grid:
                    if n == 1:
                        continue                              # 정간 번호 칸 삭제
                    c.set("_col", str(n - 1))
                else:                                         # 제목 행: A 그대로, B 삭제, 나머지 왼쪽
                    if n == 2:
                        continue
                    c.set("_col", str(n if n == 1 else n - 1))
                new_cells.append(c)
            attrs = {k: v for k, v in row.attrib.items() if k in ("ht", "customHeight", "hidden")}
            rows_out.append((attrs, new_cells))

    # 페이지 나누기: 제목 행(머리 행 바로 위) 중 첫 번째를 뺀 나머지 위에서
    def is_head(cells):
        return any(LINE_RE.fullmatch(book.text(c).strip()) for c in cells)
    heads = [i for i, (_, cells) in enumerate(rows_out) if is_head(cells)]
    breaks = [h - 1 for h in heads[1:] if h >= 1]             # 0부터 행 번호 = 제목 행 (1부터 h)

    # ── 오른쪽 맞춤 배치 ──
    #   정간보는 오른쪽 → 왼쪽: 1줄이 항상 맨 오른쪽 두 열. 줄이 적은 페이지(이어지는 2쪽 등)도 오른쪽부터 채움.
    #   머리 행 'k줄' 은 페이지 안 번호(11줄 → 1줄). 제목 행은 전체 폭 병합 + 가운데, 곡 제목만.
    def lines_of(cells):
        return [int(re.match(r"\d+", book.text(c).strip()).group(0)) for c in cells
                if LINE_RE.fullmatch(book.text(c).strip())]
    width = 2 * max((len(lines_of(rows_out[h][1])) for h in heads), default=1)
    title_style = book.title_style()
    title_rows = []
    for j, h in enumerate(heads):
        end = heads[j + 1] - 2 if j + 1 < len(heads) else len(rows_out)   # 다음 [빈 행][제목 행] 앞까지
        nlines = len(lines_of(rows_out[h][1]))
        # 머리 행의 1줄 '가사' 열(= 가장 오른쪽 'k줄' 열 + 1)이 맨 오른쪽 열이 되도록 (이미 오른쪽 맞춤이면 0)
        right = max(int(c.get("_col")) for c in rows_out[h][1] if LINE_RE.fullmatch(book.text(c).strip())) + 1
        shift = width - right
        for i in range(h, end):
            for c in rows_out[i][1]:
                c.set("_col", str(int(c.get("_col")) + shift))
        # 내용 테두리: 줄이 덜 채워진 페이지도 페이지 폭 전체를 테두리 칸으로
        used = range(1, width + 1)
        for i in range(h + 1, end):
            attrs, cells = rows_out[i]
            have = {int(c.get("_col")): c for c in cells}
            for col in used:
                c = have.get(col)
                if c is None:
                    cells.append(ET.Element(q("c"), {"_col": str(col), "s": str(book.grid_style())}))
                elif c.get("s") in (None, "0"):
                    c.set("s", str(book.grid_style()))
        for c in rows_out[h][1]:                                         # 11줄 → 1줄
            txt = book.text(c).strip()
            if LINE_RE.fullmatch(txt):
                k = (int(re.match(r"\d+", txt).group(0)) - 1) % 10 + 1
                set_text(c, f"{k}줄")
        if h >= 1:                                                       # 제목 행
            attrs, cells = rows_out[h - 1]
            a = next((c for c in cells if c.get("_col") == "1"), None)
            raw = book.text(a) if a is not None else ""
            parts = raw.split(" — ")
            title = parts[1] if parts and parts[0].startswith("페이지 ") and len(parts) > 1 else parts[0]
            new = []
            for col in range(1, width + 1):
                c = ET.Element(q("c"), {"_col": str(col), "s": str(title_style)})
                if col == 1:
                    set_text(c, title.strip())
                new.append(c)
            rows_out[h - 1] = (attrs, new)
            title_rows.append(h - 1)
    widths = {}
    for col in range(1, width + 1):
        widths[col] = 9 if col % 2 == 1 else 6                           # 홀수 = 율명, 짝수 = 가사

    # 머리 행('k줄 | 가사 …') 삭제 — 제목 행 바로 아래부터 정간 1..행수
    drop = set(heads)
    keep_idx = [i for i in range(len(rows_out)) if i not in drop]
    new_index = {old: new for new, old in enumerate(keep_idx)}           # 0부터
    rows_out = [rows_out[i] for i in keep_idx]
    titles = [new_index[t] for t in title_rows]

    # 여백: 1행·A열은 비움 → 맨 위에 빈 행 1개, 모든 칸 한 열 오른쪽(B열부터)
    rows_out = [({}, [])] + rows_out
    for _, cells in rows_out:
        for c in cells:
            c.set("_col", str(int(c.get("_col")) + 1))
    titles = [t + 1 for t in titles]                                     # 0부터
    last_col = col_letters(width + 1)
    merges = [f"B{t + 1}:{last_col}{t + 1}" for t in titles]
    breaks = [t - 1 for t in titles[1:]]                                 # 다음 페이지 = [빈 행(여백)][제목 …] — 빈 행 위에서 나눔
    widths = {1: 2, **{col + 1: w for col, w in widths.items()}}         # A열 = 여백

    # 새 워크시트
    ws = ET.Element(q("worksheet"))
    maxc = 1
    sd_rows = []
    for i, (attrs, cells) in enumerate(rows_out, 1):
        row = ET.Element(q("row"), {"r": str(i), **attrs})
        for c in sorted(cells, key=lambda c: int(c.get("_col"))):
            col = int(c.attrib.pop("_col"))
            maxc = max(maxc, col)
            c.set("r", f"{col_letters(col)}{i}")
            row.append(c)
        sd_rows.append(row)
    ET.SubElement(ws, q("dimension"), {"ref": f"A1:{col_letters(maxc)}{max(len(rows_out), 1)}"})
    sv = ET.SubElement(ET.SubElement(ws, q("sheetViews")), q("sheetView"), {"workbookViewId": "0"})
    sv.set("tabSelected", "0")
    if widths:
        cols = ET.SubElement(ws, q("cols"))
        for i in sorted(widths):
            ET.SubElement(cols, q("col"), {"min": str(i), "max": str(i), "width": f"{widths[i]:g}", "customWidth": "1"})
    sd = ET.SubElement(ws, q("sheetData"))
    sd.extend(sd_rows)
    if merges:
        mc = ET.SubElement(ws, q("mergeCells"), {"count": str(len(merges))})
        for ref in merges:
            ET.SubElement(mc, q("mergeCell"), {"ref": ref})
    ET.SubElement(ws, q("printOptions"), {"horizontalCentered": "1"})
    ET.SubElement(ws, q("pageMargins"), {"left": "0.4", "right": "0.4", "top": "0.5", "bottom": "0.5",
                                         "header": "0.3", "footer": "0.3"})
    ET.SubElement(ws, q("pageSetup"), {"paperSize": "9", "orientation": "portrait", "scale": "100"})
    if breaks:
        rb = ET.SubElement(ws, q("rowBreaks"), {"count": str(len(breaks)), "manualBreakCount": str(len(breaks))})
        for b in breaks:
            ET.SubElement(rb, q("brk"), {"id": str(b), "max": "16383", "man": "1"})
    return ws


def save(book, sheets, new_ws):
    keep = sheets[0]
    keep.set("name", "악보")
    keep_part = book.part(keep)
    sheets_el = book.wb.find(q("sheets"))
    drop_parts, drop_rids = set(), set()
    for sh in sheets[1:]:
        drop_parts.add(book.part(sh))
        drop_rids.add(sh.get(q("id", R)))
        sheets_el.remove(sh)
    for r in list(book.rels):
        if r.get("Id") in drop_rids:
            book.rels.remove(r)
    # 시트별 정의된 이름(인쇄 영역 등)은 시트 순서가 바뀌므로 모두 제거
    dn = book.wb.find(q("definedNames"))
    if dn is not None:
        for d in list(dn):
            if d.get("localSheetId") is not None:
                dn.remove(d)
        if not len(dn):
            book.wb.remove(dn)
    ct = book.z.read("[Content_Types].xml").decode("utf-8")
    for p in drop_parts:
        ct = re.sub(rf'<Override[^>]*PartName="/{re.escape(p)}"[^>]*/>', "", ct)
    sheet_rels = {f"xl/worksheets/_rels/{os.path.basename(p)}.rels" for p in drop_parts | {keep_part}}
    tmp = book.path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for info in book.z.infolist():
            name = info.filename
            if name in drop_parts or name in sheet_rels:
                continue
            if name == "xl/workbook.xml":
                buf = ET.tostring(book.wb, xml_declaration=True, encoding="UTF-8")
            elif name == "xl/_rels/workbook.xml.rels":
                buf = ET.tostring(book.rels, xml_declaration=True, encoding="UTF-8")
            elif name == keep_part:
                buf = ET.tostring(new_ws, xml_declaration=True, encoding="UTF-8")
            elif name == "[Content_Types].xml":
                buf = ct.encode("utf-8")
            elif name == "xl/styles.xml" and getattr(book, "styles", None) is not None:
                buf = ET.tostring(book.styles, xml_declaration=True, encoding="UTF-8")
            else:
                buf = book.z.read(name)
            out.writestr(info, buf)
    book.z.close()
    os.replace(tmp, book.path)


def csv_meta(name):
    """csv/{곡}.csv 의 박자('3/4' …)·행수"""
    path = os.path.join(CSV_DIR, name + ".csv")
    beat, rows = "", 0
    if os.path.exists(path):
        for line in open(path, encoding="utf-8-sig"):
            k, _, v = line.rstrip("\n").partition(",")
            if k == "박자":
                beat = v.replace("박자", "").strip()
            elif k == "행수" and v.strip().isdigit():
                rows = int(v)
            elif k == "페이지":
                break
    return beat, rows


def apply_beat_lines(book, sh, group, nrows):
    """페이지 레이아웃 테두리 + 박 묶음.

    - 페이지마다 제목 병합 폭 전체(줄이 덜 채워진 페이지도) 정간 칸에 테두리.
    - group 이 있으면 가사 열: 박 묶음(group 정간) 안쪽 가로선을 흰색으로.
    바뀌면 새 워크시트 XML(bytes), 아니면 None.
    """
    part = book.part(sh)
    raw = book.z.read(part)
    ws = ET.fromstring(raw)
    rows = {int(r.get("r")): r for r in ws.find(q("sheetData"))}
    titles = []
    mc = ws.find(q("mergeCells"))
    for m in (mc if mc is not None else []):
        a, _, b = m.get("ref").partition(":")
        ra, rb = int(re.search(r"\d+", a).group(0)), int(re.search(r"\d+", b or a).group(0))
        if ra == rb:
            titles.append((ra, col_num(a), col_num(b or a)))
    changed = False
    for t, c1, c2 in sorted(titles):
        for jg in range(1, nrows + 1):
            row = rows.get(t + jg)
            if row is None:
                continue
            # 빈 칸·테두리 없는 칸 → 테두리 칸 (페이지 폭 전체)
            have = {col_num(c.get("r")): c for c in row}
            for col in range(c1, c2 + 1):
                c = have.get(col)
                if c is None:
                    c = ET.Element(q("c"), {"r": f"{col_letters(col)}{t + jg}"})
                    have[col] = c
                    changed = True
                if c.get("s") in (None, "0"):
                    c.set("s", str(book.grid_style()))
                    changed = True
            for c in list(row):
                row.remove(c)
            for col in sorted(have):
                row.append(have[col])
            # 율명 열(오른쪽 끝에서 홀수 번째)= PITCH_PT·'―'는 작게 / 가사 열(짝수 번째)= LYRIC_PT
            for col in range(c1, c2 + 1):
                c = have[col]
                is_pitch = (c2 - col) % 2 == 1
                want_pt = PITCH_PT if is_pitch else LYRIC_PT
                if book.font_size(c.get("s")) != want_pt:
                    c.set("s", str(book.sized_style(c.get("s"), want_pt)))
                    changed = True
                if is_pitch and book.dash_runs(c, PITCH_PT, DASH_PT):  # 율명 안 '―' → 작은 글씨 (리치텍스트)
                    changed = True
            ht = f"{CONTENT_HT:g}"                                     # 줄 수와 무관하게 고정 높이
            if row.get("ht") != ht:
                row.set("ht", ht)
                row.set("customHeight", "1")
                changed = True
            if not group:
                continue
            pos = (jg - 1) % group                                    # 0 = 묶음 첫 정간
            top_white = pos != 0
            bottom_white = pos != group - 1 and jg != nrows
            for c in row:
                col = col_num(c.get("r"))
                # 가사 열 = 병합 범위 오른쪽 끝에서 짝수 번째(…, c2-2, c2). 테두리 있는(쓰는) 칸만
                if not c1 <= col <= c2 or (c2 - col) % 2 or c.get("s") in (None, "0"):
                    continue
                base = c.get("s")
                if book.border_state(base) == (top_white, bottom_white):
                    continue
                new = str(book.beat_style(base, top_white, bottom_white))
                if new != base:
                    c.set("s", new)
                    changed = True
    # 제목 행 높이 = 일반 행의 두 배
    for t, _, _ in titles:
        row = rows.get(t)
        if row is not None and row.get("ht") != f"{TITLE_HT:g}":
            row.set("ht", f"{TITLE_HT:g}")
            row.set("customHeight", "1")
            changed = True
    # 인쇄 배율: 가장 넓은 폭·가장 긴 페이지가 A4 한 장에 들어가게 (맞춤 인쇄는 엑셀이 페이지 나누기를 무시하므로 배율 고정)
    width_pt = 0.0
    cols = ws.find(q("cols"))
    for c in (cols if cols is not None else []):
        n = int(c.get("max")) - int(c.get("min")) + 1
        width_pt += n * (float(c.get("width", 8.43)) * 7 + 5) * 0.75      # 열 너비(글자) → px → pt
    def ht(r):
        row = rows.get(r)
        if row is not None and row.get("ht") and row.get("customHeight") in ("1", "true"):
            return float(row.get("ht"))
        return 15.0                                                     # 기본 높이 (프로그램마다 계산값이 달라 고정)
    page_h = max((sum(ht(r) for r in range(t - 1, t + nrows + 1)) for t, _, _ in titles), default=1.0)
    scale = int(max(10, min(100, PRINT_W / max(width_pt, 1) * 100, PRINT_H / page_h * 100)))
    ps = ws.find(q("pageSetup"))
    if ps is not None:
        want = {"paperSize": "9", "orientation": "portrait", "scale": str(scale)}
        if any(ps.get(k) != v for k, v in want.items()):              # fitTo* 는 fitToPage 가 꺼져 있으면 무시됨
            for k, v in want.items():
                ps.set(k, v)
            changed = True
    sp = ws.find(q("sheetPr"))
    if sp is not None:
        pu = sp.find(q("pageSetUpPr"))
        if pu is not None and pu.get("fitToPage") in ("1", "true"):
            pu.set("fitToPage", "0")
            changed = True
    if not changed:
        return None
    return ET.tostring(ws, xml_declaration=True, encoding="UTF-8")


def write_parts(book, parts):
    """zip 안 파일 일부만 바꿔 다시 쓴다. parts = {이름: bytes}"""
    if getattr(book, "styles", None) is not None:
        parts["xl/styles.xml"] = ET.tostring(book.styles, xml_declaration=True, encoding="UTF-8")
    tmp = book.path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for info in book.z.infolist():
            out.writestr(info, parts.get(info.filename, book.z.read(info.filename)))
    book.z.close()
    os.replace(tmp, book.path)


def main():
    args = [nfc(a[:-5] if a.lower().endswith(".xlsx") else a) for a in sys.argv[1:] if not a.startswith("--")]
    check = "--check" in sys.argv
    files = sorted(f for f in os.listdir(XLSX_DIR)
                   if f.lower().endswith(".xlsx") and not f.startswith(("_", "~$", ".")))
    if args:
        files = [f for f in files if nfc(f[:-5]) in args]
    locked = {nfc(f[len(".~lock."):-1]) for f in os.listdir(XLSX_DIR) if f.startswith(".~lock.")} | \
             {nfc(f[2:]) for f in os.listdir(XLSX_DIR) if f.startswith("~$")}
    done, skipped = [], []
    for f in files:
        path = os.path.join(XLSX_DIR, f)
        book = Book(path)
        sheets = score_sheets(book)
        if not sheets:
            book.z.close()
            continue
        if nfc(f) in locked:
            if needs_convert(book, sheets):
                skipped.append(f)
            book.z.close()
            continue
        touched = False
        if needs_convert(book, sheets):
            touched = True
            if check:
                book.z.close()
                done.append(f)
                continue
            save(book, sheets, convert(book, sheets))
            book = Book(path)
            sheets = score_sheets(book)
        # 박자 묶음 가로선 (3/4 → 3정간, 4/4 → 4정간)
        beat, nrows = csv_meta(nfc(f[:-5]))
        if nrows and len(sheets) == 1:
            new = apply_beat_lines(book, sheets[0], BEAT_GROUP.get(beat), nrows)
            if new is not None:
                touched = True
                if not check:
                    write_parts(book, {book.part(sheets[0]): new})
        if not book.z.fp is None:
            book.z.close()
        if touched:
            done.append(f)
    for f in skipped:
        print(f"  ⏭️  열려 있음(건너뜀): {f}")
    print(f"{len(files)}개 확인 / {'변환 대상' if check else '변환'} {len(done)}개")
    for f in done:
        print(f"  🔀 {f}")


if __name__ == "__main__":
    main()
