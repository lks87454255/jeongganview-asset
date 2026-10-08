#!/usr/bin/env python3
"""
민요채보 생성 스크립트 (macOS/Linux용, build_all.ps1 을 그대로 옮긴 것. 표준 라이브러리만 사용)

  폴더 구조 (jeongganview-asset 레포):
    sheets-index.json                    ← 앱 목록
    sheets/민요/                         ← ROOT (기본값: 이 스크립트의 상위 폴더)
      csv/{곡}.csv                       ← 앱 CSV v2 (util/JeongganCsv.kt)
      원본/민요채보-1~3/*.png|jpg         ← 채보 스캔 (CSV '원본' 행은 원본/ 기준 상대 경로)
      검수/{곡}.xlsx, 검수/_목록.xlsx      ← 검수 엑셀
      _작업/tsv/*.tsv                     ← 입력 (판독규격.md 의 TSV v2)
      _작업/생성리포트.txt

  규칙
   - #이어짐 으로 연결된 이미지들은 한 곡(한 문서)의 연속 페이지.
   - 이미지 한 장 = CSV 1페이지. 한 장이 10줄을 넘으면 10줄씩 다음 페이지로 나눔.
   - 줄 n(페이지 안 1..10) → 앱 내부 열: 율명 = 2n(대), 가사 = 2n-1(소).
   - 율명 칸은 비슷한 글자를 대표 글자로 바꾸고(app JeongganSymbols.ALIASES 와 동일),
     율명·팔레트 기호가 아닌 글자는 "확인 필요"로 보고한다.

  사용 : python3 build_all.py [ROOT]
"""
import datetime
import hashlib
import math
import os
import re
import sys
import unicodedata
import urllib.parse
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.dirname(HERE)   # sheets/민요
TSV_DIR = os.path.join(ROOT, "_작업", "tsv")
REVIEW_DIR = os.path.join(ROOT, "검수")
REPO_DIR = os.path.dirname(os.path.dirname(ROOT))                                     # jeongganview-asset
CSV_DIR = os.path.join(ROOT, "csv")
REPORT_PATH = os.path.join(ROOT, "_작업", "생성리포트.txt")
CATEGORY = "민요"
MAX_LINES_PER_PAGE = 10
BEATS = ["4/4", "3/4", "2/4", "정악"]


def nfc(s):
    return unicodedata.normalize("NFC", s or "")


# ───────────────────── 기호표 (앱 Yulmyeong.kt / JeongganSymbols.kt 와 동일) ─────────────────────
YUL = list("㣴㣕㣖㣣㣨㣡㣸㣩𢓡㣮㣳㣹"
           "僙㐲㑀俠㑬㑖𠐭㑣侇㑲㒇㒣"
           "黃大太夾姑仲㽔林夷南無應"
           "潢汏汰浹㴌㳞㶋淋洟湳潕㶐"
           "㶂𣴘㳲㴺㵈㴢㶙㵉㴣㵜㶃㶝")
MARKS = ["―", "△", "И", "ﾉ", "^", "⌝", "ㅋ", "⌞", "է", "⊏", "⊔", "·", "○", "⁚", "‹", "／"]
KNOWN = set(YUL) | set(MARKS)
ALIAS = {}
for to, frm in [("―", "-‐–—−ㅡ"), ("△", "▵"), ("·", ".ㆍ•∙⋅"), ("○", "oO◯〇"),
                ("⁚", ":︰："), ("‹", "<〈＜"), ("／", "/∕")]:
    for f in frm:
        ALIAS[f] = to
if len(YUL) != 60:
    raise SystemExit(f"율명표 60자 아님: {len(YUL)}")


def convert_pitch(raw):
    """TSV 율명(| 구분) → 앱 셀 텍스트(\\n 구분, 대표 글자)"""
    if not raw or not raw.strip():
        return ""
    parts = []
    for seg in raw.split("|"):
        parts.append("".join(ALIAS.get(cp, cp) for cp in seg.strip() if not cp.isspace()))
    while parts and parts[0] == "":
        parts.pop(0)
    while parts and parts[-1] == "":
        parts.pop()
    return "\n".join(parts)


def get_unknown(text):
    out = []
    for cp in text:
        if not cp.isspace() and cp not in KNOWN and cp not in out:
            out.append(cp)
    return out


# ───────────────────── TSV 읽기 ─────────────────────
PROBLEMS = []   # 리포트용


class Obj:
    def __init__(self, **kw):
        self.__dict__.update(kw)


def to_int(s):
    try:
        return int(s.strip())
    except (ValueError, AttributeError):
        return None


def read_tsv(path):
    name = nfc(os.path.basename(path))
    meta, cells, header = {}, [], False
    text = open(path, encoding="utf-8-sig").read()
    for n, ln in enumerate(text.splitlines(), 1):
        if ln.strip() == "":
            continue
        if ln.startswith("#"):
            k, _, v = ln[1:].partition("\t")
            meta[k.strip()] = v.strip()
            continue
        f = ln.split("\t")
        if not header and f[0].strip() == "줄":
            header = True
            continue
        jul = to_int(f[0]) if len(f) >= 2 else None
        jg = to_int(f[1]) if len(f) >= 2 else None
        if jul is None or jg is None or jul < 1 or jg < 1:
            PROBLEMS.append(f"[형식] {name} {n}행 읽을 수 없음: {ln}")
            continue
        raw = f[2].strip() if len(f) > 2 else ""
        lyr = f[3].strip() if len(f) > 3 else ""
        note = " ".join(f[4:]).strip() if len(f) > 4 else ""
        pitch = convert_pitch(raw)
        if not pitch and not lyr:
            continue   # 빈 정간 행(검사용)은 건너뜀
        cells.append(Obj(Jul=jul, Jg=jg, Raw=raw, Pitch=pitch, Lyric=nfc(lyr), Note=nfc(note), Unknown=get_unknown(pitch)))
    base = name[:-4] if name.lower().endswith(".tsv") else name
    folder, _, stem = base.partition("_")
    g = lambda k: nfc(meta.get(k, ""))
    rows = to_int(g("행수")) or 0
    lines = to_int(g("줄수")) or 0
    max_jg = max((c.Jg for c in cells), default=0)
    max_jul = max((c.Jul for c in cells), default=0)
    seen, dedup = set(), []
    for c in cells:   # 같은 (줄,정간)이 두 번 나오면 뒤의 것을 버리고 보고
        k = (c.Jul, c.Jg)
        if k in seen:
            PROBLEMS.append(f"[중복] {name} 줄{c.Jul} 정간{c.Jg} 두 번 — 뒤의 것 무시")
            continue
        seen.add(k)
        dedup.append(c)
    return Obj(File=name, Folder=folder, Stem=stem,
               Title=g("제목"), Sub=g("부제"), Key=g("키"), Page=g("쪽"),
               Beat=g("박자"), Rows=max(rows, max_jg), RowsMeta=rows, LinesMeta=lines, MaxJul=max_jul,
               Source=g("원본"), Next=g("이어짐"), Memo=g("메모"),
               Cells=dedup, PresentLines=len({c.Jul for c in dedup}))


tsvs = {}
for fn in sorted((nfc(f) for f in os.listdir(TSV_DIR) if f.lower().endswith(".tsv"))):
    real = next(f for f in os.listdir(TSV_DIR) if nfc(f) == fn)
    tsvs[fn] = read_tsv(os.path.join(TSV_DIR, real))

# ───────────────────── 이어짐 → 곡(문서) 묶기 ─────────────────────


def resolve_next(v):
    if not v or not v.strip():
        return None
    v = v.strip()
    if not v.endswith(".tsv"):
        v += ".tsv"
    for k in tsvs:
        if k.lower() == v.lower():
            return k
    return None


next_of, has_prev = {}, {}
for t in tsvs.values():
    n = resolve_next(t.Next)
    if t.Next and not n:
        PROBLEMS.append(f"[이어짐] {t.File} → '{t.Next}' 파일 없음 — 무시")
    elif n == t.File:
        PROBLEMS.append(f"[이어짐] {t.File} 자기 자신 — 무시")
    elif n:
        if n in has_prev:
            PROBLEMS.append(f"[이어짐] {n} 앞장이 둘({has_prev[n]}, {t.File}) — 뒤의 것 무시")
        else:
            next_of[t.File] = n
            has_prev[n] = t.File

docs, used, broken = [], set(), set()


def add_doc(head):
    chain, cur = [], head
    while cur and cur not in used:
        t = tsvs[cur]
        if chain and (t.Beat.strip() != chain[0].Beat.strip() or t.Rows != chain[0].Rows):
            PROBLEMS.append(f"[이어짐] {chain[-1].File} → {cur} : 박자/행수 다름({chain[0].Beat}/{chain[0].Rows} ≠ {t.Beat}/{t.Rows}) — 따로 곡으로 분리")
            broken.add(cur)
            break
        used.add(cur)
        chain.append(t)
        cur = next_of.get(cur)
    docs.append(Obj(Images=chain))


for k in sorted(tsvs):
    if k not in has_prev:
        add_doc(k)
for k in sorted(tsvs):
    if k not in used:
        if k not in broken:
            PROBLEMS.append(f"[이어짐] {k} 순환 연결 — 따로 처리")
        add_doc(k)

# ───────────────────── 문서 내용 계산 ─────────────────────


def base_title(t):
    if t.Title:
        return t.Title.strip()
    return f"제목없음 ({re.sub(r'\.jpg$', '', t.Stem)})"


def clean_file_name(s):
    s = re.sub(r"\s+", " ", re.sub(r'[\\/:*?"<>|\r\n\t]', " ", s or "")).strip().rstrip(".")
    return nfc(s)


def clean_tag(s):
    return clean_file_name(re.sub(r"[()\[\]]", " ", s or ""))


for d in docs:
    h = d.Images[0]
    d.Title = base_title(h)
    d.Sub = h.Sub
    d.Key = next((im.Key for im in d.Images if im.Key), "")
    d.PageNo = ", ".join(im.Page for im in d.Images if im.Page)
    beat = h.Beat.strip()
    if beat not in BEATS:
        PROBLEMS.append(f"[박자] {h.File} '{beat}' 미지원 → 3/4")
        beat = "3/4"
    for im in d.Images:
        if im.Beat.strip() != h.Beat.strip():
            PROBLEMS.append(f"[박자] {im.File} '{im.Beat}' ≠ 첫 장 '{h.Beat}' — 첫 장 기준")
    d.Beat = beat
    rows = max((im.Rows for im in d.Images), default=0)
    d.RowCount = rows if rows and rows >= 1 else 12

    # 페이지: 이미지마다, 10줄 넘으면 나눔
    pages = []
    for im in d.Images:
        n_pages = max(1, math.ceil(max(im.MaxJul, 1) / MAX_LINES_PER_PAGE))
        for p in range(n_pages):
            cells, items = {}, []
            for c in im.Cells:
                if (c.Jul - 1) // MAX_LINES_PER_PAGE != p:
                    continue
                local = (c.Jul - 1) % MAX_LINES_PER_PAGE + 1
                if c.Pitch:
                    cells[(2 * local, c.Jg)] = c.Pitch
                if c.Lyric:
                    cells[(2 * local - 1, c.Jg)] = c.Lyric
                items.append(Obj(Cell=c, Local=local))
            lines_on_page = min(MAX_LINES_PER_PAGE, max(1, im.MaxJul - p * MAX_LINES_PER_PAGE))
            pages.append(Obj(Image=im, Part=p + 1, Parts=n_pages, Cells=cells, Items=items, Lines=lines_on_page))
    d.Pages = pages

    all_cells = [c for im in d.Images for c in im.Cells]
    d.CellCount = len(all_cells)
    d.NoteCount = sum(1 for c in all_cells if c.Note)
    d.UnknownCount = sum(len(c.Unknown) for c in all_cells)
    present = sum(im.PresentLines for im in d.Images)
    sparse = len(all_cells) < 0.3 * present * d.RowCount
    ratio = d.NoteCount / len(all_cells) if all_cells else 1
    memo = " ".join(im.Memo for im in d.Images)
    reasons = []
    if d.UnknownCount:
        reasons.append(f"팔레트 외 글자 {d.UnknownCount}")
    if ratio >= 0.15:
        reasons.append(f"비고 {math.floor(ratio * 100 + 0.5)}%")
    if sparse:
        reasons.append("정간 적음(누락 의심)")
    if re.search("자진모리|근사|가까운", memo):
        reasons.append("박자 근사")
    if re.search("두 곡", memo):
        reasons.append("한 장에 두 곡")
    if not h.Title:
        reasons.append("제목 없음")
    if not all_cells:
        reasons.append("내용 없음")
    if d.UnknownCount == 0 and ratio < 0.15 and not sparse and all_cells:
        d.Grade = "상"
    elif d.UnknownCount <= 10 and ratio < 0.5 and not sparse and all_cells:
        d.Grade = "중"
    else:
        d.Grade = "하"
    d.Reasons = ", ".join(reasons)

# 이름 정하기 (겹치면 부제 → 채보 폴더 → 파일명 순으로 구분)
groups = {}
for d in docs:
    groups.setdefault(clean_file_name(d.Title).lower(), []).append(d)
for grp in groups.values():
    gname = clean_file_name(grp[0].Title)
    if len(grp) == 1:
        grp[0].Name = gname
        continue
    for d in grp:
        sub = clean_tag(d.Sub)
        same_sub = sum(1 for x in grp if clean_tag(x.Sub) == sub)
        tag = sub if (sub and same_sub == 1 and not re.fullmatch(r"\d+", sub)) else f"채보{d.Images[0].Folder}"
        d.Name = f"{gname} ({tag})"
by_name = {}
for d in docs:
    by_name.setdefault(d.Name.lower(), []).append(d)
for grp in by_name.values():
    if len(grp) > 1:
        for d in grp:
            d.Name = clean_file_name(f"{d.Name} {d.Images[0].Stem}")
docs.sort(key=lambda d: d.Name)

# ───────────────────── CSV v2 ─────────────────────


def quote_csv(v):
    if v and ("=+-@".find(v[0]) >= 0 or (len(v) > 1 and v[0] == "'" and "=+-@'".find(v[1]) >= 0)):
        v = "'" + v
    if any(ch in v for ch in ',"\n\r') or (v and (v[0].isspace() or v[-1].isspace())):
        return '"' + v.replace('"', '""') + '"'
    return v


def header_name(c):
    return f"{(c + 1) // 2}열({'대' if c % 2 == 0 else '소'})"


def beat_label(b):
    return f"{b}박자" if "/" in b else b


def build_csv(d):
    rows = [["형식", "정간보CSV", "2"], ["타이틀", d.Title], ["박자", beat_label(d.Beat)], ["행수", str(d.RowCount)]]
    if d.Sub:
        rows.append(["부제", d.Sub])
    if d.Key:
        rows.append(["키", d.Key])
    if d.PageNo:
        rows.append(["쪽", d.PageNo])
    rows.append(["원본", " / ".join(im.Source for im in d.Images)])
    rows.append(["상태", "판독 초안(검수 전)"])
    header = ["정간번호"] + [header_name(c) for c in range(20, 0, -1)]
    for p, pg in enumerate(d.Pages):
        rows.append(["페이지", str(p + 1)])
        rows.append(header)
        for r in range(1, d.RowCount + 1):
            rows.append([str(r)] + [pg.Cells.get((c, r), "") for c in range(20, 0, -1)])
    # 줄바꿈 LF: 저장소(git index)가 LF 로 보관하므로 index 의 sha256 이 GitHub raw 와 일치하게 (앱 파서는 CRLF/LF 모두 허용)
    return "\n".join(",".join(quote_csv(x) for x in row) for row in rows) + "\n"


def parse_csv(text):
    """검증용 RFC 4180 파서 (앱 JeongganCsv.parseRecords 와 같은 규칙)"""
    records, row, field, in_q, i = [], [], [], False, 0
    while i < len(text):
        ch = text[i]
        if in_q:
            if ch == '"':
                if i + 1 < len(text) and text[i + 1] == '"':
                    field.append('"')
                    i += 1
                else:
                    in_q = False
            else:
                field.append(ch)
        else:
            if ch == '"' and not field:
                in_q = True
            elif ch == ",":
                row.append("".join(field))
                field = []
            elif ch in "\r\n":
                row.append("".join(field))
                field = []
                if any(row):
                    records.append(row)
                row = []
                if ch == "\r" and i + 1 < len(text) and text[i + 1] == "\n":
                    i += 1
            else:
                field.append(ch)
        i += 1
    return records


def unguard(v):
    return v[1:] if len(v) > 1 and v[0] == "'" and "=+-@'".find(v[1]) >= 0 else v


def test_csv(d, text):
    recs = parse_csv(text)
    errs = []
    n_head = sum(1 for r in recs if r[0] == "정간번호")
    n_page = sum(1 for r in recs if r[0] == "페이지")
    if n_head != len(d.Pages):
        errs.append(f"정간번호 행 {n_head} ≠ 페이지 {len(d.Pages)}")
    if n_page != len(d.Pages):
        errs.append(f"페이지 행 {n_page} ≠ 페이지 {len(d.Pages)}")
    for rec in recs:
        for cell in rec:
            if cell and "=+-@".find(cell[0]) >= 0:
                errs.append(f"수식 위험 칸: {cell}")
    page, back = 1, 0
    for rec in recs:
        if rec[0] == "페이지":
            page = int(rec[1])
            continue
        r = to_int(rec[0])
        if r is None:
            continue
        for i in range(1, len(rec)):
            v = unguard(rec[i].strip())
            if not v:
                continue
            c = 21 - i
            if d.Pages[page - 1].Cells.get((c, r)) != v:
                errs.append(f"왕복 불일치 p{page} 열{c} 행{r}")
            back += 1
    expect = sum(len(pg.Cells) for pg in d.Pages)
    if back != expect:
        errs.append(f"칸 수 불일치 {back} ≠ {expect}")
    return errs


# ───────────────────── xlsx ─────────────────────
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4"><font><sz val="11"/><name val="맑은 고딕"/><family val="2"/></font><font><b/><sz val="11"/><name val="맑은 고딕"/><family val="2"/></font><font><b/><sz val="14"/><name val="맑은 고딕"/><family val="2"/></font><font><sz val="22"/><name val="맑은 고딕"/><family val="2"/></font></fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF8CBAD"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="5"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>
<!-- 2~4: 박 묶음 가사 칸 (2 = 묶음 첫 칸: 아래 흰색, 3 = 가운데: 위·아래 흰색, 4 = 마지막: 위 흰색) -->
<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color rgb="FFFFFFFF"/></bottom><diagonal/></border>
<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color rgb="FFFFFFFF"/></top><bottom style="thin"><color rgb="FFFFFFFF"/></bottom><diagonal/></border>
<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color rgb="FFFFFFFF"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="20">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<!-- 8~16: 박 묶음 가사 칸 = (일반 8~10 · 비고 11~13 · 팔레트 외 14~16) × (첫 · 가운데 · 마지막) -->
<xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="2" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="3" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="4" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="4" borderId="2" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="4" borderId="3" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="4" borderId="4" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<!-- 17~19: 악보 율명 칸 22pt (가사 11pt 의 두 배) = 일반 · 비고 · 팔레트 외 -->
<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"""
# 스타일 번호
S_HEAD, S_CELL, S_NOTE, S_LEFT, S_BAD, S_TITLE, S_TITLEC = 1, 2, 3, 4, 5, 6, 7
# 박자별 묶음(정간 수): 악보 시트 가사 열은 묶음 안쪽 가로선을 흰색으로 (박 단위로 한 칸처럼 보임)
BEAT_GROUP = {"3/4": 3, "4/4": 4}
ROW_HT = 30
TITLE_HT = ROW_HT * 2          # 제목 행(두 배)
CONTENT_HT = 56                # 내용 행(줄 수 무관 고정)
PRINT_W = 595 - 0.8 * 72       # A4 세로 인쇄 영역(pt)
PRINT_H = 842 - 1.0 * 72 - 20


def pitch_style(st):           # 율명 22pt
    return {S_CELL: 17, S_NOTE: 18, S_BAD: 19}.get(st, st)


def beat_style(st, jg, rows, g):
    if not g:
        return st
    base = {S_CELL: 8, S_NOTE: 11, S_BAD: 14}.get(st, -1)
    if base < 0:
        return st
    pos = (jg - 1) % g
    if pos == 0:
        off = 0
    elif pos == g - 1 or jg == rows:
        off = 2
    else:
        off = 1
    if pos == 0 and (g == 1 or jg == rows):   # 한 칸짜리 묶음은 그대로
        return st
    return base + off


def esc(s):
    if s is None:
        return ""
    s = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(s))
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")


def col_name(n):
    s = ""
    while n > 0:
        m = (n - 1) % 26
        s = chr(65 + m) + s
        n = (n - 1) // 26
    return s


def new_sheet(name):
    return Obj(Name=name, Rows=[], Widths={}, Merges=[], FreezeRows=0, FreezeCols=0, Heights={}, Filter=None,
               Print=False, Scale=100, Breaks=[])


def add_row(sheet, values, style=S_CELL):
    """행 추가: 값 배열 + 스타일(하나 또는 배열)"""
    sheet.Rows.append([(v, style[i] if isinstance(style, (list, tuple)) else style) for i, v in enumerate(values)])


def sheet_xml(sh):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
    if sh.FreezeRows or sh.FreezeCols:
        tl = f"{col_name(sh.FreezeCols + 1)}{sh.FreezeRows + 1}"
        attr = ""
        if sh.FreezeCols:
            attr += f' xSplit="{sh.FreezeCols}"'
        if sh.FreezeRows:
            attr += f' ySplit="{sh.FreezeRows}"'
        pane = "bottomRight" if (sh.FreezeRows and sh.FreezeCols) else ("bottomLeft" if sh.FreezeRows else "topRight")
        out.append(f'<sheetViews><sheetView workbookViewId="0"><pane{attr} topLeftCell="{tl}" activePane="{pane}" state="frozen"/></sheetView></sheetViews>')
    if sh.Widths:
        out.append("<cols>")
        for c in sorted(sh.Widths):
            out.append(f'<col min="{c}" max="{c}" width="{sh.Widths[c]}" customWidth="1"/>')
        out.append("</cols>")
    out.append("<sheetData>")
    for r, row in enumerate(sh.Rows):
        rn = r + 1
        ht = f' ht="{sh.Heights[rn]}" customHeight="1"' if rn in sh.Heights else ""
        out.append(f'<row r="{rn}"{ht}>')
        for c, cell in enumerate(row):
            if cell is None:
                continue
            v, st = cell
            ref = f"{col_name(c + 1)}{rn}"
            if v is None or str(v) == "":
                out.append(f'<c r="{ref}" s="{st}"/>')
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                out.append(f'<c r="{ref}" s="{st}"><v>{v}</v></c>')
            else:
                out.append(f'<c r="{ref}" s="{st}" t="inlineStr"><is><t xml:space="preserve">{esc(v)}</t></is></c>')
        out.append("</row>")
    out.append("</sheetData>")
    if sh.Filter:
        out.append(f'<autoFilter ref="{sh.Filter}"/>')
    if sh.Merges:
        out.append(f'<mergeCells count="{len(sh.Merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in sh.Merges) + "</mergeCells>")
    if sh.Print:
        out.append('<printOptions horizontalCentered="1"/>')
        out.append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
        out.append(f'<pageSetup paperSize="9" orientation="portrait" scale="{sh.Scale}"/>')   # A4 세로, 배율 고정
    if sh.Breaks:
        out.append(f'<rowBreaks count="{len(sh.Breaks)}" manualBreakCount="{len(sh.Breaks)}">')
        out.extend(f'<brk id="{b}" max="16383" man="1"/>' for b in sh.Breaks)   # id = 이 행(0부터) 위에서 나눔
        out.append("</rowBreaks>")
    out.append("</worksheet>")
    return "".join(out)


def save_xlsx(path, sheets):
    ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
    rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    for i, sh in enumerate(sheets, 1):
        ct += f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        wb += f'<sheet name="{esc(sh.Name)}" sheetId="{i}" r:id="rId{i}"/>'
        rels += f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
    ct += "</Types>"
    wb += "</sheets></workbook>"
    rels += f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    root = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    if os.path.exists(path):
        os.remove(path)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", root)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", rels)
        z.writestr("xl/styles.xml", STYLES)
        for i, sh in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", sheet_xml(sh))


def cell_style(c):
    return S_BAD if c.Unknown else (S_NOTE if c.Note else S_CELL)


LEGEND = [
    ("줄 순서", "원본처럼 오른쪽이 1줄. 악보 시트도 같은 배치. 여러 페이지는 악보 시트 아래쪽으로 이어짐(빈 행 1개, 인쇄 시 페이지마다 새 종이). 1줄은 항상 맨 오른쪽 두 열(율명, 가사). 3/4 박자는 3정간, 4/4 박자는 4정간마다 가사 칸 가로선을 묶음. 율명 글자는 가사의 두 배. 인쇄는 A4 세로, 원본 한 페이지 = 종이 한 장. 정간 번호 열·머리 행 없음 — 제목 행 아래 n번째 행이 정간 n이므로 행을 지우거나 끼워 넣지 말고, 제목 병합도 풀지 말 것. 한 장이 10줄을 넘으면 10줄씩 다음 페이지"),
    ("율명 칸", "정간 안에서 위→아래 순서, 칸 안 줄바꿈으로 구분 (앱과 같음)"),
    ("기호", "― 연음 · △ 쉼 · 점 · ○ 동그라미 · ⁚ 두 점 · ‹ 정간 경계 표시 · ／ 사선 (· ○ ⁚ ‹ ／ 는 화면 표시 전용)"),
    ("노란 칸", "판독자가 비고를 단 칸 — 원본과 대조 필요"),
    ("주황 칸", "율명·기호표에 없는 글자 (? 포함) — 반드시 수정"),
    ("수정 방법", "악보 시트를 고쳐 .xlsx 로 저장한 뒤 _작업/xlsx_to_csv.py 실행 → csv 반영. 정보·정간목록 시트는 반영 안 됨"),
]


def build_review_book(d):
    sheets = []

    # 정보
    info = new_sheet("정보")
    info.Widths[1] = 16
    info.Widths[2] = 90
    add_row(info, [d.Title, ""], S_TITLE)
    kv = [("CSV 파일", f"sheets/{CATEGORY}/csv/{d.Name}.csv"), ("부제", d.Sub), ("키", d.Key), ("쪽", d.PageNo),
          ("박자 / 행수", f"{d.Beat} / {d.RowCount}정간"), ("페이지", str(len(d.Pages))),
          ("정간 수", str(d.CellCount)), ("비고(노란 칸)", str(d.NoteCount)), ("팔레트 외 글자(주황 칸)", str(d.UnknownCount)),
          ("판독 신뢰도", d.Grade), ("확인 필요", d.Reasons)]
    for k, v in kv:
        add_row(info, [k, v], [S_HEAD, S_LEFT])
    add_row(info, ["", ""], 0)
    add_row(info, ["원본 이미지", "판독 메모"], S_HEAD)
    for im in d.Images:
        add_row(info, [im.Source, im.Memo], [S_LEFT, S_LEFT])
    add_row(info, ["", ""], 0)
    add_row(info, ["범례", ""], S_TITLE)
    for k, v in LEGEND:
        add_row(info, [k, v], [S_HEAD, S_LEFT])
    sheets.append(info)

    # 정간목록
    ls = new_sheet("정간목록")
    ls.FreezeRows = 1
    for i, w in enumerate([6, 30, 6, 6, 6, 5, 14, 14, 12, 34, 10]):
        ls.Widths[i + 1] = w
    add_row(ls, ["페이지", "원본 이미지", "줄", "앱 줄", "정간", "박", "율명(판독)", "율명(앱)", "가사", "비고", "확인 필요 글자"], S_HEAD)
    beat_size = {"4/4": 4, "2/4": 2, "3/4": 3}.get(d.Beat, 0)
    for p, pg in enumerate(d.Pages):
        for it in sorted(pg.Items, key=lambda x: (x.Cell.Jul, x.Cell.Jg)):
            c = it.Cell
            st = cell_style(c)
            bak = (c.Jg - 1) // beat_size + 1 if beat_size else ""
            add_row(ls, [p + 1, pg.Image.Source, c.Jul, it.Local, c.Jg, bak, c.Raw, c.Pitch, c.Lyric, c.Note, " ".join(c.Unknown)],
                    [S_CELL, S_LEFT, S_CELL, S_CELL, S_CELL, S_CELL, st, st, st, S_LEFT, S_BAD if c.Unknown else S_CELL])
    ls.Filter = f"A1:K{len(ls.Rows)}"
    sheets.append(ls)

    # 악보 (한 시트, 인쇄용 겸 편집용)
    #   1행·A열 = 여백(비움). 악보는 B열부터, 정간보는 오른쪽 → 왼쪽: 1줄이 항상 맨 오른쪽 두 열(율명, 가사).
    #   페이지마다 [빈 행(여백)][제목 행: 전체 폭 병합·가운데·곡 제목만·테두리][정간 1..행수 (위→아래)·악보 폭 전체 테두리],
    #   다음 페이지는 빈 행 위에서 인쇄 페이지 나눔. 정간 번호 열·'k줄' 머리 행 없음 — 제목 행 아래 n번째 행 = 정간 n.
    #   xlsx_to_csv.py 는 가로 병합된 제목 행을 페이지 시작으로, 병합 범위 오른쪽 끝을 1줄로 읽는다.
    sh = new_sheet("악보")
    sh.Print = True
    max_l = max(pg.Lines for pg in d.Pages)
    W = max_l * 2                     # 악보 폭(열 수), B열부터
    C0 = 2                            # 악보 첫 열 = B (A = 여백)
    sh.Widths[1] = 2
    for i in range(1, W + 1):
        sh.Widths[C0 + i - 1] = 9 if i % 2 == 1 else 6   # 홀수 = 율명, 짝수 = 가사
    last_col = col_name(C0 + W - 1)
    g = BEAT_GROUP.get(d.Beat)
    for p, pg in enumerate(d.Pages):
        L = pg.Lines
        add_row(sh, [""], 0)                               # 1행 여백 / 페이지 사이 빈 행
        if p > 0:
            sh.Breaks.append(len(sh.Rows) - 1)             # 이 빈 행(0부터) 위에서 인쇄 나눔
        add_row(sh, [""] + [d.Title] + [""] * (W - 1), [0] + [S_TITLEC] * W)
        tr = len(sh.Rows)
        sh.Merges.append(f"B{tr}:{last_col}{tr}")
        sh.Heights[tr] = TITLE_HT                          # 제목 행 = 일반 행 두 배
        idx = {(it.Local, it.Cell.Jg): it.Cell for it in pg.Items}
        for jg in range(1, d.RowCount + 1):
            vals = [""] * (W + 1)                          # [0] = A열 여백
            sts = [0] + [S_CELL] * W                       # 악보 폭 전체 테두리 (줄이 덜 채워진 페이지도)
            for k in range(1, L + 1):
                pc = W - 2 * k + 1                         # 악보 안 k줄 율명 열
                c = idx.get((k, jg))
                if not c:
                    continue
                vals[pc] = c.Pitch
                vals[pc + 1] = c.Lyric
                st = cell_style(c)
                sts[pc] = st
                sts[pc + 1] = st
            for k in range(1, max_l + 1):                  # 율명 열 22pt
                pi = W - 2 * k + 1
                sts[pi] = pitch_style(sts[pi])
            for k in range(1, max_l + 1):                  # 가사 열 (페이지 폭 전체)
                gi = W - 2 * k + 2
                sts[gi] = beat_style(sts[gi], jg, d.RowCount, g)
            add_row(sh, vals, sts)
            sh.Heights[len(sh.Rows)] = CONTENT_HT          # 줄 수와 무관하게 고정
    # 인쇄 배율: 가장 넓은 폭·가장 긴 페이지가 A4 세로 한 장에 들어가게
    wpt = sum((w * 7 + 5) * 0.75 for w in sh.Widths.values())
    hmax = hcur = 0
    breaks = set(sh.Breaks)
    for r in range(1, len(sh.Rows) + 1):
        if r > 1 and (r - 1) in breaks:
            hmax = max(hmax, hcur)
            hcur = 0
        hcur += sh.Heights.get(r, 15)
    hmax = max(hmax, hcur)
    sh.Scale = int(max(10, math.floor(min(100, min(PRINT_W / wpt * 100, PRINT_H / hmax * 100)))))
    sheets.append(sh)
    save_xlsx(os.path.join(REVIEW_DIR, f"{d.Name}.xlsx"), sheets)


def jstr(s):
    if s is None:
        return "null"
    out = ['"']
    for ch in s:
        if ch == '"':
            out.append('\\"')
        elif ch == "\\":
            out.append("\\\\")
        elif ord(ch) < 0x20:
            out.append(f"\\u{ord(ch):04x}")
        else:
            out.append(ch)
    out.append('"')
    return "".join(out)


def main():
    # ───────────────────── 실행 ─────────────────────
    for dr in (REVIEW_DIR, CSV_DIR):
        os.makedirs(dr, exist_ok=True)
    for f in os.listdir(REVIEW_DIR):
        if f.lower().endswith(".xlsx"):
            os.remove(os.path.join(REVIEW_DIR, f))
    for f in os.listdir(CSV_DIR):
        if f.lower().endswith(".csv"):
            os.remove(os.path.join(CSV_DIR, f))

    csv_errors = 0
    for d in docs:
        text = build_csv(d)
        for e in test_csv(d, text):
            PROBLEMS.append(f"[CSV] {d.Name}: {e}")
            csv_errors += 1
        with open(os.path.join(CSV_DIR, f"{d.Name}.csv"), "wb") as fh:
            fh.write(b"\xef\xbb\xbf" + text.encode("utf-8"))
        build_review_book(d)

    # 목록
    lst = new_sheet("곡 목록")
    lst.FreezeRows = 1
    for i, w in enumerate([5, 30, 20, 14, 6, 8, 6, 6, 6, 8, 6, 8, 6, 40, 50]):
        lst.Widths[i + 1] = w
    add_row(lst, ["번호", "CSV 파일", "제목", "부제", "키", "쪽", "박자", "행수", "페이지", "정간 수", "비고", "팔레트 외", "신뢰도", "확인 필요", "원본 이미지"], S_HEAD)
    for no, d in enumerate(docs, 1):
        gs = {"상": S_CELL, "중": S_NOTE}.get(d.Grade, S_BAD)
        add_row(lst, [no, f"{d.Name}.csv", d.Title, d.Sub, d.Key, d.PageNo, d.Beat, d.RowCount, len(d.Pages), d.CellCount,
                      d.NoteCount, d.UnknownCount, d.Grade, d.Reasons, "\n".join(im.Source for im in d.Images)],
                [S_CELL, S_LEFT, S_LEFT, S_LEFT, S_CELL, S_CELL, S_CELL, S_CELL, S_CELL, S_CELL, S_CELL, S_CELL, gs, S_LEFT, S_LEFT])
    lst.Filter = f"A1:O{len(lst.Rows)}"

    imgs = new_sheet("이미지별")
    imgs.FreezeRows = 1
    for i, w in enumerate([34, 30, 20, 12, 6, 6, 6, 6, 8, 6, 8, 26, 30, 60]):
        imgs.Widths[i + 1] = w
    add_row(imgs, ["원본 이미지", "TSV", "제목", "부제", "박자", "행수", "줄수", "실제 줄", "정간 수", "비고", "팔레트 외", "이어짐", "CSV 파일", "판독 메모"], S_HEAD)
    for d in docs:
        for im in d.Images:
            unk = sum(len(c.Unknown) for c in im.Cells)
            add_row(imgs, [im.Source, im.File, im.Title, im.Sub, im.Beat, im.RowsMeta, im.LinesMeta, im.PresentLines, len(im.Cells),
                           sum(1 for c in im.Cells if c.Note), unk, im.Next, f"{d.Name}.csv", im.Memo],
                    [S_LEFT, S_LEFT, S_LEFT, S_LEFT, S_CELL, S_CELL, S_CELL, S_CELL, S_CELL, S_CELL, S_BAD if unk else S_CELL, S_LEFT, S_LEFT, S_LEFT])
    imgs.Filter = f"A1:N{len(imgs.Rows)}"

    iss = new_sheet("확인 필요 글자")
    iss.FreezeRows = 1
    for i, w in enumerate([30, 30, 6, 6, 6, 8, 16, 40]):
        iss.Widths[i + 1] = w
    add_row(iss, ["CSV 파일", "원본 이미지", "페이지", "줄", "정간", "글자", "율명 칸", "비고"], S_HEAD)
    issue_count = 0
    for d in docs:
        for p, pg in enumerate(d.Pages):
            for it in sorted(pg.Items, key=lambda x: (x.Cell.Jul, x.Cell.Jg)):
                for u in it.Cell.Unknown:
                    issue_count += 1
                    add_row(iss, [f"{d.Name}.csv", pg.Image.Source, p + 1, it.Cell.Jul, it.Cell.Jg, u, it.Cell.Pitch, it.Cell.Note],
                            [S_LEFT, S_LEFT, S_CELL, S_CELL, S_CELL, S_BAD, S_CELL, S_LEFT])
    iss.Filter = f"A1:H{len(iss.Rows)}"

    unk_freq = {}
    for d in docs:
        for im in d.Images:
            for c in im.Cells:
                for u in c.Unknown:
                    unk_freq[u] = unk_freq.get(u, 0) + 1
    freq = new_sheet("글자 빈도")
    freq.FreezeRows = 1
    freq.Widths.update({1: 8, 2: 10, 3: 12, 4: 50})
    add_row(freq, ["글자", "횟수", "코드", "제안"], S_HEAD)
    suggest = {"古": "姑 (고선)으로 보임", "神": "㳞 (氵仲, 청중려)으로 보임", "?": "원본 확인"}
    for ch, cnt in sorted(unk_freq.items(), key=lambda kv: -kv[1]):
        add_row(freq, [ch, cnt, f"U+{ord(ch):04X}", suggest.get(ch, "")], [S_BAD, S_CELL, S_CELL, S_LEFT])

    save_xlsx(os.path.join(REVIEW_DIR, "_목록.xlsx"), [lst, imgs, iss, freq])

    # sheets-index.json (generate_index.py 와 같은 형식)
    owner, repo, branch = "lks87454255", "jeongganview-asset", "main"
    base = f"https://raw.githubusercontent.com/{owner}/{repo}/{branch}"
    entries = []
    for d in docs:
        fn = f"{d.Name}.csv"
        data = open(os.path.join(CSV_DIR, fn), "rb").read()
        url = base + "/" + "/".join(urllib.parse.quote(nfc(x), safe="") for x in ("sheets", CATEGORY, "csv", fn))
        entries.append("        {\n"
                       f'          "name": {jstr(d.Name)},\n'
                       f'          "title": {jstr(d.Title)},\n'
                       f'          "fileName": {jstr(fn)},\n'
                       f'          "path": {jstr(f"sheets/{CATEGORY}/csv/{fn}")},\n'
                       f'          "url": {jstr(url)},\n'
                       f'          "sizeBytes": {len(data)},\n'
                       f'          "beat": {jstr(d.Beat)},\n'
                       f'          "rows": {d.RowCount},\n'
                       f'          "pages": {len(d.Pages)},\n'
                       f'          "sha256": {jstr(hashlib.sha256(data).hexdigest())}\n'
                       "        }")
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    js = ("{\n  \"version\": 1,\n"
          f"  \"generated\": {jstr(now)},\n  \"owner\": {jstr(owner)},\n  \"repo\": {jstr(repo)},\n  \"branch\": {jstr(branch)},\n"
          "  \"categories\": [\n    {\n"
          f"      \"name\": {jstr(CATEGORY)},\n      \"files\": [\n" + ",\n".join(entries) + "\n      ]\n    }\n  ],\n"
          f"  \"_total\": {{ \"categories\": 1, \"files\": {len(docs)} }}\n}}\n")
    with open(os.path.join(REPO_DIR, "sheets-index.json"), "w", encoding="utf-8", newline="\n") as fh:
        fh.write(js)

    # 리포트
    rep = [f"민요채보 생성 리포트  {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}",
           f"TSV {len(tsvs)}개 → 곡 {len(docs)}개 (페이지 {sum(len(d.Pages) for d in docs)}), 정간 {sum(d.CellCount for d in docs)}",
           f"신뢰도  상 {sum(d.Grade == '상' for d in docs)} / 중 {sum(d.Grade == '중' for d in docs)} / 하 {sum(d.Grade == '하' for d in docs)}",
           f"팔레트 외 글자 {issue_count}곳 (종류 {len(unk_freq)}), CSV 검증 오류 {csv_errors}건",
           "", "── 곡 ──"]
    for d in docs:
        rep.append(f"{d.Name}.csv | {d.Grade} {d.Beat}/{d.RowCount} p{len(d.Pages)} | 정간 {d.CellCount} 비고 {d.NoteCount} 외 {d.UnknownCount} | "
                   f"{' → '.join(im.File for im in d.Images)} | {d.Reasons}")
    rep += ["", "── 처리 중 문제 ──"] + (PROBLEMS or ["없음"])
    with open(REPORT_PATH, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(rep) + "\n")

    print(f"DONE docs={len(docs)} tsv={len(tsvs)} issues={issue_count} csvErrors={csv_errors} problems={len(PROBLEMS)}")


if __name__ == "__main__":
    main()
