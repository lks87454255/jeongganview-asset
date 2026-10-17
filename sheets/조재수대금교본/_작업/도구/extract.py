#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract.py — 「조재수 대금교본」 PDF 글자 층 → TSV(판독규격 v2)

    python3 extract.py                # 전체 → _작업/tsv/*.tsv, 원본/pNNN.png, _작업/곡목록.tsv
    python3 extract.py --show 11 41   # 한 쪽 격자 출력 (저장 안 함)

이 책의 특징 (2026-10-17 조사)
  - 문서에서 만든 PDF 지만 율명이 전용 글꼴의 사설 영역 글자(U+E0xx)로 들어 있다 → GLYPH 표로 율명 변환.
    같은 글꼴의 한자 코드(享 上 不 下 乾 …)는 실제 한자가 아니라 요성 물결·숨표·쉼표 같은 기호 모양 → MARK 표.
  - 가사·줄 번호·행 번호가 없다. 격자는 그려진 선으로만 알 수 있다:
    세로선 쌍 = 율명 칸(폭 ~32–41pt), 그 사이 좁은 띠(~7pt) = 숨표 '<' 자리. 오른쪽 칸이 1줄.
    가로선 = 정간 경계 (12 또는 16정간).
  - 오른쪽 여백(테두리 밖): 큰 제목, '조재수 편곡', 조(예: 산조/정악 6관청). 머리 칸 숫자 = 곡번호(있는 곡만).
"""
import collections
import json
import os
import re
import sys
import unicodedata

import pymupdf

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.dirname(HERE)
ROOT = os.path.dirname(WORK)
REPO = os.path.dirname(os.path.dirname(ROOT))
PDF = os.path.join(REPO, "book", "조재수 대금교본", "조재수 대금교본.pdf")
TSV_DIR = os.path.join(WORK, "tsv")
SRC_DIR = os.path.join(ROOT, "원본")
nfc = lambda s: unicodedata.normalize("NFC", s)

# 전용 글꼴 사설 영역 → 율명 (글리프를 확대해 눈으로 대조해 정함, /tmp/jjs_glyphs.png)
GLYPH = {
    "\ue00d": "淋", "\ue00f": "湳", "\ue010": "潕", "\ue011": "㶐", "\ue012": "潢", "\ue013": "汏", "\ue014": "汰",
    "\ue015": "浹", "\ue016": "㴌", "\ue017": "㳞", "\ue019": "林", "\ue01b": "南", "\ue01c": "無", "\ue01d": "應",
    "\ue01e": "黃", "\ue020": "太", "\ue022": "姑", "\ue023": "仲", "\ue025": "㑣", "\ue027": "㑲", "\ue028": "㒇",
    "\ue029": "㒣",
    # 꾸밈 기호 (앱 악상기호 카탈로그·팔레트에서 가장 비슷한 것)
    "\ue000": "∽", "\ue002": "ㅋ", "\ue003": "ㄱ", "\ue006": "∧", "\ue007": "ʒ", "\ue008": "ʊ", "\ue009": "ㄷ",
    "\ue00a": "ㅅ", "\ue00b": "ㅅ", "\ue01f": "c", "\ue021": "⊏", "\ue026": "·", "\ue041": "h", "\ue043": "ㄴ",
    "\ue044": "է", "\ue049": "И", "\ue04c": "⌝", "\ue051": "○", "\ue052": "ㅎ", "\ue05d": "ㅎ",
}
WAVE = set("享亦亨京交亢亥代仝仟令")              # 요성(물결) 조각 — 이어진 물결 하나를 ∼ 하나로
MARK = {"上": "‹", "不": "△", "下": "Z", "乾": "ﾉ", "乎": "⌝", "丑": "ㅂ", "且": "ㅂ", "丕": "ㅂ",
        "以": "ϟ", "仰": "ϟ", "件": "ϟ", "仲": "ϟ", "亂": "ㄴ", "乘": "∨", "一": "―", "乙": "∧",
        "－": "―", "―": "―", "─": "―", "-": "―", "○": "○", "◎": "○", "．": "·", "˙": "·", "/": "／"}
MUSIC_FONT_PREFIX = "font000000002a6d5"       # 악보 글꼴(쪽마다 이름 끝이 바뀜)


def spans_of(page):
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                for c in s["chars"]:
                    ch = c["c"]
                    if ch.strip():
                        x0, y0, x1, y1 = c["bbox"]
                        out.append(dict(x0=x0, y0=y0, x1=x1, y1=y1, xc=(x0 + x1) / 2, yc=(y0 + y1) / 2,
                                        size=s["size"], font=s["font"], c=ch))
    return out


def grid_of(page):
    V, H = collections.defaultdict(float), collections.defaultdict(float)
    VY = collections.defaultdict(lambda: [1e9, -1e9])
    def addv(x, y0, y1):
        k = round(x, 1); V[k] += abs(y1 - y0); VY[k][0] = min(VY[k][0], y0, y1); VY[k][1] = max(VY[k][1], y0, y1)
    for g in page.get_drawings():
        for it in g["items"]:
            if it[0] == "l":
                a, b = it[1], it[2]
                if abs(a.x - b.x) < 0.6: addv(a.x, a.y, b.y)
                elif abs(a.y - b.y) < 0.6: H[round(a.y, 1)] += abs(a.x - b.x)
            elif it[0] == "re":
                r = it[1]
                if r.width < 1.5: addv(r.x0, r.y0, r.y1)
                elif r.height < 1.5: H[round(r.y0, 1)] += r.width
    vx = sorted(x for x, L in V.items() if L > 150)
    cols = [(a, b) for a, b in zip(vx, vx[1:]) if 25 <= b - a <= 60]
    if not cols:
        return None
    xmin, xmax = cols[0][0], cols[-1][1]
    # 격자 위·아래 = 율명 칸 세로선이 그려진 범위
    y_top = min(VY[a][0] for a, b in cols); y_bot = max(VY[a][1] for a, b in cols)
    hy = sorted({y for y, L in H.items() if y_top - 1.5 <= y <= y_bot + 1.5 and L > 30} | {round(y_top, 1), round(y_bot, 1)})
    merged = []
    for y in hy:
        if not merged or y - merged[-1] > 3: merged.append(y)
    rows = list(zip(merged, merged[1:]))
    return dict(cols=sorted(cols, key=lambda c: -c[0]), rows=rows, xmin=xmin, xmax=xmax)


def stack(items):
    items = sorted(items, key=lambda s: (s["yc"], s["xc"]))
    groups = []
    for s in items:
        if groups and abs(groups[-1][-1]["yc"] - s["yc"]) < 4:
            groups[-1].append(s)
        else:
            groups.append([s])
    return ["".join(x["t"] for x in sorted(g, key=lambda s: s["xc"])) for g in groups]


def parse_page(page):
    S = spans_of(page); G = grid_of(page)
    if not G or len(G["rows"]) < 6:
        return None
    cols, rows = G["cols"], G["rows"]
    # 오른쪽 끝(또는 왼쪽 끝) 칸에 율명 글자가 하나도 없으면 악보 줄이 아니다(제목·정보 칸) → 빼고 여백으로
    def music_in(a, b):
        return sum(1 for s in S if a <= s["xc"] <= b and (s["c"] in GLYPH or s["c"] in ("不", "－")))
    while cols and music_in(*cols[0]) == 0: cols = cols[1:]          # cols 는 오른쪽부터
    while cols and music_in(*cols[-1]) == 0: cols = cols[:-1]
    if not cols:
        return None
    G = dict(G, cols=cols, xmax=cols[0][1], xmin=cols[-1][0])
    top, bot = rows[0][0], rows[-1][1]
    cells = collections.defaultdict(list); notes = collections.defaultdict(list); waves = collections.defaultdict(list)
    for s in S:
        if not (G["xmin"] - 2 <= s["xc"] <= G["xmax"] + 12 and top - 6 <= s["yc"] <= bot + 6):
            continue
        # 칸: 이 글자 x 를 포함하거나 바로 왼쪽에 있는 율명 칸(숨표 띠는 왼쪽 칸 것)
        ci = next((i for i, (a, b) in enumerate(cols) if a - 1 <= s["xc"]), None)
        if ci is None: continue
        ch = s["c"]; music = s["font"].startswith(MUSIC_FONT_PREFIX)
        if ch in WAVE and music:
            waves[ci].append(s); continue
        if ch == "上" and music:                               # 숨표 '<' : 경계 바로 위 정간 끝
            r = min(range(len(rows)), key=lambda k: abs(rows[k][1] - s["yc"]))
            cells[(ci, r)].append(dict(s, t="‹", yc=rows[r][1] - 1)); continue
        r = next((k for k, (a, b) in enumerate(rows) if a - 1 <= s["yc"] <= b + 1), None)
        if r is None: continue
        if music and ch in GLYPH: t = GLYPH[ch]
        elif music and ch in MARK: t = MARK[ch]
        elif ch in MARK: t = MARK[ch]
        elif re.match(r"[0-9A-Za-z()\[\]※,，、=→ⅠⅡⅢ]", ch) or "\uac00" <= ch <= "\ud7a3" or not music:
            notes[(ci, r)].append(ch); continue
        else:
            t = ch                                              # 모르는 글자는 그대로(검수 때 주황 칸)
        cells[(ci, r)].append(dict(s, t=t))
    for ci, ws in waves.items():                                # 이어진 물결 → 시작 칸에 ∼ 하나
        ws.sort(key=lambda s: s["y0"]); run_top = None; last = None
        for w in ws:
            if last is None or w["y0"] - last > 4:
                r = next((k for k, (a, b) in enumerate(rows) if a - 1 <= w["yc"] <= b + 1), None)
                if r is not None:
                    tgt = [x for x in cells[(ci, r)] if x["t"] != "‹"]
                    yc = min(tgt, key=lambda x: abs(x["yc"] - w["yc"]))["yc"] if tgt else w["yc"]
                    cells[(ci, r)].append(dict(w, t="∼", yc=yc, xc=w["xc"] + 100))
            last = w["y1"]
    grid = {}
    for (ci, r), items in cells.items():
        grid[(ci + 1, r + 1)] = ("|".join(stack(items)), "".join(notes.get((ci, r), [])))
    for k, v in notes.items():
        if (k[0] + 1, k[1] + 1) not in grid:
            grid[(k[0] + 1, k[1] + 1)] = ("", "".join(v))
    # 오른쪽 여백 정보 / 머리 칸 숫자
    side = [s for s in S if s["xc"] > G["xmax"] + 4 and s["c"] not in GLYPH and s["c"] not in WAVE and s["c"] not in ("上", "不", "下", "－")]
    # 세로쓰기 열로 묶기 (x 가까운 글자끼리), 열마다 위→아래
    colsx = []
    for s in sorted(side, key=lambda s: s["xc"]):
        if colsx and abs(colsx[-1][-1]["xc"] - s["xc"]) < 9 and abs(colsx[-1][-1]["size"] - s["size"]) < 2:
            colsx[-1].append(s)
        else:
            colsx.append([s])
    def text(col):
        col = sorted(col, key=lambda s: s["yc"]); out = col[0]["c"]
        gaps = [b["y0"] - a["y1"] for a, b in zip(col, col[1:])]
        med = sorted(gaps)[len(gaps) // 2] if gaps else 0
        for g, b in zip(gaps, col[1:]):
            out += (" " if g > med + col[0]["size"] * 0.4 else "") + b["c"]
        return out.strip()
    tcols = [c for c in colsx if c[0]["size"] >= 16 and any("\uac00" <= x["c"] <= "\ud7a3" or x["c"].isalpha() for x in c) and len(c) >= 2]
    tcols = [c for c in tcols if len(c) <= 16]
    tcol = max(tcols, key=lambda c: (c[0]["size"], len(c))) if tcols else None
    big = text(tcol) if tcol else ""
    small = " / ".join(t for t in (text(c) for c in sorted(colsx, key=lambda c: -c[0]["xc"]) if c is not tcol)
                       if t and t.replace(" ", "") not in ("조재수편곡", "편곡조재수"))
    head = "".join(s["c"] for s in S if s["yc"] < top - 4 and s["c"].isdigit() and G["xmin"] < s["xc"] < G["xmax"])
    return dict(lines=len(cols), nrows=len(rows), grid=grid, title=big, info=small, num=head)


def show(pno):
    d = pymupdf.open(PDF); P = parse_page(d[pno - 1])
    if not P: print(f"p{pno}: 격자 없음"); return
    print(f"p{pno} 번호={P['num']!r} 제목={P['title']!r} 정보={P['info']!r} 줄 {P['lines']} 행 {P['nrows']}")
    for r in range(1, P["nrows"] + 1):
        print(f"  {r:>2} " + " | ".join(f"{(P['grid'].get((ln, r), ('', ''))[0] + ('!' + P['grid'].get((ln, r), ('', ''))[1] if P['grid'].get((ln, r), ('', ''))[1] else '')):<11}" for ln in range(P["lines"], 0, -1)))


if __name__ == "__main__" and len(sys.argv) > 2 and sys.argv[1] == "--show":
    for p in sys.argv[2:]:
        show(int(p))


# ───────────────────────────── 전체 추출 ─────────────────────────────
BEAT = {12: "3/4", 16: "4/4", 8: "4/4", 18: "3/4"}      # 그 밖의 정간 수는 '정악'(자유) — 대금율보와 같은 규칙
SECTION = {"간주", "전주", "후렴", "엇모리", "자진모리", "중모리", "중중모리", "굿거리", "휘모리", "세마치", "진양", "진양조", "1절", "2절"}
FIX_TITLE = {"이누야사시대를 초월한 마음": "이누야사 (시대를 초월한 마음)", "서편제 os": "서편제 OST",
             "Danny Boy 아 목동아": "Danny Boy (아, 목동아)", "10 월의 어느 멋진날의": "10월의 어느 멋진날의"}
SCANNED = {13, 15, 49, 50, 80, 83, 132}                  # 글자 층 없는 그림 쪽 → 이미지 판독(따로)


def main():
    d = pymupdf.open(PDF)
    os.makedirs(TSV_DIR, exist_ok=True); os.makedirs(SRC_DIR, exist_ok=True)
    for f in os.listdir(TSV_DIR):
        if f.endswith(".tsv") and not f.startswith("s"): os.remove(os.path.join(TSV_DIR, f))
    songs, cur = [], None
    for pno in range(6, d.page_count + 1):
        if pno in SCANNED:
            cur = None; continue
        P = parse_page(d[pno - 1])
        if not P:
            continue
        t = nfc(P["title"]).strip()
        if t.replace(" ", "") in SECTION and cur is not None:          # 장단·구간 이름은 제목이 아니라 앞 곡의 다음 부분
            P["info"] = (t + (" / " + P["info"] if P["info"] else "")); t = ""
        if t:
            t = FIX_TITLE.get(t, t)
            cur = dict(title=t, info=P["info"], parts=[], first=pno); songs.append(cur)
        elif cur is None:
            print(f"  p{pno}: 제목 없는 쪽인데 앞 곡 없음 — 건너뜀"); continue
        elif P["info"]:
            cur["info"] += (" / " if cur["info"] else "") + P["info"]
        cur["parts"].append(dict(page=pno, **P))
        png = os.path.join(SRC_DIR, f"p{pno:03d}.png")
        if not os.path.exists(png):
            d[pno - 1].get_pixmap(dpi=150).save(png)
    # 한 곡 안에서 정간 수가 바뀌면(장단이 바뀜) 거기서 나눔 → '{제목} ({장단})'
    split = []
    for s in songs:
        groups = [[s["parts"][0]]]
        for p in s["parts"][1:]:
            (groups[-1].append(p) if p["nrows"] == groups[-1][-1]["nrows"] else groups.append([p]))
        for k, g in enumerate(groups):
            if k == 0:
                split.append(dict(s, parts=g)); continue
            lab = next((w for w in SECTION if w in g[0]["info"].replace(" ", "")), "") or str(k + 1)
            split.append(dict(title=f"{s['title']} ({lab})", info=g[0]["info"], parts=g, first=g[0]["page"]))
    songs = split
    # 내용이 똑같은 곡(같은 쪽이 두 번 실린 것)은 하나만
    sig = {}; uniq = []
    for s in songs:
        k = (s["title"], tuple(sorted((tuple(p["grid"].items()) for p in s["parts"]), key=str)))
        if k in sig:
            print(f"  p{s['first']}: '{s['title']}' 이 p{sig[k]} 과 내용이 같음 — 뺌"); continue
        sig[k] = s["first"]; uniq.append(s)
    songs = uniq
    seen = collections.Counter()
    for s in songs:
        seen[s["title"]] += 1
        s["name"] = s["title"] if seen[s["title"]] == 1 else f"{s['title']}-{seen[s['title']]}"
    lst = ["곡ID\t곡명\t쪽\t블록수\t행수\t줄수\t정보\t문제"]
    for n, s in enumerate(songs, 1):
        sid = f"{n:03d}"; probs = []
        R = max(p["nrows"] for p in s["parts"])
        if len({p["nrows"] for p in s["parts"]}) > 1: probs.append(f"쪽마다 정간 수 다름 {sorted({p['nrows'] for p in s['parts']})}")
        beat = BEAT.get(R, "정악")
        names = [f"{sid}_{k + 1}.tsv" for k in range(len(s["parts"]))]
        for k, p in enumerate(s["parts"]):
            meta = [("제목", s["name"]), ("부제", ""), ("키", ""), ("쪽", str(p["page"])), ("박자", beat), ("행수", str(R)),
                    ("줄수", str(p["lines"])), ("원본", f"p{p['page']:03d}.png"), ("이어짐", names[k + 1] if k + 1 < len(names) else ""),
                    ("메모", f"PDF 글자 층(전용 율명 글꼴)에서 자동 추출. PDF {p['page']}쪽." + (f" 쪽 글: {p['info']}" if p["info"] else ""))]
            out = [f"#{a}\t{b}" for a, b in meta] + ["줄\t정간\t율명\t가사\t비고"]
            for ln in range(1, p["lines"] + 1):
                for r in range(1, R + 1):
                    y, n_ = p["grid"].get((ln, r), ("", ""))
                    out.append(f"{ln}\t{r}\t{y}\t\t{n_}")
            open(os.path.join(TSV_DIR, names[k]), "w", encoding="utf-8", newline="\n").write("\n".join(out) + "\n")
        lst.append("\t".join([sid, s["name"], ",".join(str(p["page"]) for p in s["parts"]), str(len(s["parts"])), str(R),
                              ",".join(str(p["lines"]) for p in s["parts"]), s["info"], "; ".join(probs)]))
    open(os.path.join(WORK, "곡목록.tsv"), "w", encoding="utf-8", newline="\n").write("\n".join(lst) + "\n")
    print(f"곡 {len(songs)} / TSV {sum(len(s['parts']) for s in songs)} / 그림 쪽(따로 판독) {sorted(SCANNED)}")
    print("문제", [l.split('\t')[1] + ': ' + l.split('\t')[-1] for l in lst[1:] if l.split('\t')[-1]])


if __name__ == "__main__" and len(sys.argv) == 1:
    main()
