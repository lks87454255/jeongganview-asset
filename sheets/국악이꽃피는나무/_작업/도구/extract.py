#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract.py — 「국악이 꽃피는 나무」 PDF 의 글자 층(텍스트 레이어)에서 정간보를 그대로 뽑아 TSV(판독규격 v2)로 저장
==============================================================================================
이 PDF 는 스캔이 아니라 문서에서 바로 만든 PDF 라 율명·가사가 글자 + 좌표로 들어 있다 → 이미지 판독 없이 정확히 추출.

    python3 extract.py                 # 전체 쪽 → _작업/tsv/*.tsv, _작업/곡목록.tsv, 원본/pNNN.png
    python3 extract.py --show 31       # 한 쪽만 추출해 격자를 화면에 출력 (파일 저장 안 함)

쪽 구조 (조사 결과)
  - 격자 블록: 위쪽 줄 번호 머리(HYwulM 11pt 숫자, 오른쪽 1 → 왼쪽 8·9) + 왼쪽 행 번호(1..12 또는 1..16).
    한 쪽에 블록이 2개인 쪽도 있음(p190: 위 1–8줄, 아래 9–16줄). 두 쪽짜리 곡은 다음 쪽 머리가 9..16.
  - 줄 하나 = 율명 칸(머리 숫자 바로 아래) + 오른쪽 가사 칸. 율명 글꼴 HYhaeseo, 가사는 T10/T14/T17 등.
  - 오른쪽 여백(x > 485): 큰 제목(MHunmin 세로), 곡번호(HYwulM 20pt), 장단·작곡·노래·채보 정보.
  - 한 칸 안: 위→아래로 쌓인 글자는 '|' 로, 같은 높이에 나란한 글자는 붙여 씀(판독규격과 같음).
"""
import collections
import json
import os
import re
import sys
import unicodedata

import pymupdf

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.dirname(HERE)                      # sheets/국악이꽃피는나무/_작업
ROOT = os.path.dirname(WORK)                      # sheets/국악이꽃피는나무
REPO = os.path.dirname(os.path.dirname(ROOT))
PDF = os.path.join(REPO, "book", "국악이 꽃피는 나무", "국악이 꽃피는 나무.pdf")
TSV_DIR = os.path.join(WORK, "tsv")
SRC_DIR = os.path.join(ROOT, "원본")
MARGIN_X = 485          # 이보다 오른쪽은 제목·정보 여백
YUL_FONT = "HYhaeseo"
SYM = {"-": "―", "－": "―", "—": "―", "–": "―", "ㅡ": "―",
       "蕤": "㽔",            # 유빈: 책은 蕤(U+8564), 앱 율명표는 㽔(U+3F54)
       "~": "∼", "ɜ": "ʒ", "▼": "▾", "o": "○", " ": ""}   # 꾸밈 기호 → 앱 악상기호 카탈로그 글자

nfc = lambda s: unicodedata.normalize("NFC", s)


def spans_of(page):
    out = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                t = s["text"]
                if t.strip():
                    x0, y0, x1, y1 = s["bbox"]
                    out.append(dict(x0=x0, y0=y0, x1=x1, y1=y1, xc=(x0 + x1) / 2, yc=(y0 + y1) / 2,
                                    size=s["size"], font=s["font"], t=nfc(t)))
    return out


def is_num(s, lo=10.5, hi=11.5):
    return s["font"] == "HYwulM" and lo <= s["size"] <= hi and s["t"].strip().isdigit()


def find_blocks(S):
    """줄 번호 머리 → 블록. 머리 = 같은 높이(±3)에 숫자 3개 이상, x < MARGIN_X."""
    nums = [s for s in S if is_num(s) and s["x0"] > 85 and s["x1"] < MARGIN_X]
    rows_by_y = collections.defaultdict(list)
    for s in nums:
        rows_by_y[round(s["yc"] / 4)].append(s)
    heads = []
    for k, g in sorted(rows_by_y.items()):
        if len(g) >= 3:
            heads.append(sorted(g, key=lambda s: -s["xc"]))          # 오른쪽부터
    # 같은 머리가 반올림 경계로 둘로 나뉘면 합침
    merged = []
    for h in heads:
        if merged and abs(merged[-1][0]["yc"] - h[0]["yc"]) < 6:
            merged[-1] = sorted(merged[-1] + h, key=lambda s: -s["xc"])
        else:
            merged.append(h)
    labels = [s for s in S if is_num(s) and s["x1"] < 85]
    blocks = []
    for i, h in enumerate(merged):
        top = h[0]["yc"]
        bottom = merged[i + 1][0]["yc"] if i + 1 < len(merged) else 10 ** 6
        rl = sorted((s for s in labels if top < s["yc"] < bottom), key=lambda s: s["yc"])
        # 행 번호는 1부터 연속
        rows, want = [], 1
        for s in rl:
            if int(s["t"]) == want:
                rows.append(s); want += 1
        if not rows:
            continue
        lines = [(int(s["t"]), s["xc"]) for s in h]
        # 못갖춘마디 글자가 머리 숫자를 덮은 경우(p61): 오른쪽에 자리가 있으면 빠진 번호를 간격으로 채움
        lines.sort(key=lambda t: t[0])
        if len(lines) > 1:
            pitch = abs(lines[0][1] - lines[-1][1]) / (lines[-1][0] - lines[0][0])
            while lines[0][0] > 1 and lines[0][1] + pitch < MARGIN_X:
                lines.insert(0, (lines[0][0] - 1, lines[0][1] + pitch))
        blocks.append(dict(lines=lines, rows=[r["yc"] for r in rows], top=top))
    return blocks


def cell_text(items, sep="|"):
    """칸 안 글자들 → 위→아래 줄은 sep, 같은 높이는 왼→오 붙임"""
    items = sorted(items, key=lambda s: (s["yc"], s["xc"]))
    groups = []
    for s in items:
        if groups and abs(groups[-1][-1]["yc"] - s["yc"]) < 3.5:
            groups[-1].append(s)
        else:
            groups.append([s])
    return sep.join("".join(x["t"].strip() for x in sorted(g, key=lambda s: s["xc"])) for g in groups)


def side_info(S, blocks):
    """오른쪽 여백: 곡번호(20pt), 제목(큰 MHunmin), 그 아래 정보 글자들"""
    num = "".join(s["t"].strip() for s in S if s["font"] == "HYwulM" and s["size"] > 18 and s["x0"] > MARGIN_X - 10)
    side = sorted((s for s in S if s["x0"] > MARGIN_X and not is_num(s) and s["font"] != "Gulim"
                   and not (s["font"] == "HYwulM" and s["size"] > 18)), key=lambda s: s["yc"])
    big = [s for s in side if s["size"] >= 14]
    title = "".join(s["t"].strip() for s in big)
    small = "".join(s["t"].strip() for s in side if s["size"] < 14)
    return num, title, small


def parse_page(page):
    S = spans_of(page)
    blocks = find_blocks(S)
    out = []
    for b in blocks:
        xs = sorted(b["lines"], key=lambda t: t[0])                   # (줄번호, x)
        pitch = (abs(xs[0][1] - xs[-1][1]) / (len(xs) - 1)) if len(xs) > 1 else 50
        rows = b["rows"]
        rh = (rows[-1] - rows[0]) / (len(rows) - 1) if len(rows) > 1 else 40
        y_top, y_bot = rows[0] - rh / 2, rows[-1] + rh / 2
        cells = collections.defaultdict(lambda: {"y": [], "g": [], "n": []})
        pick = collections.defaultdict(lambda: {"y": [], "g": []})     # 머리 줄 위치에 적힌 못갖춘마디
        for s in S:
            if b["top"] - 20 < s["yc"] < y_top and not is_num(s) and 85 < s["xc"] < MARGIN_X:
                for ln, hx in xs:
                    dx = s["xc"] - hx
                    if -pitch * 0.3 <= dx < pitch * 0.7:
                        pick[ln]["y" if dx < pitch * 0.3 else "g"].append(s); break
                continue
            if s["xc"] > MARGIN_X or s["xc"] < 85 or not (y_top <= s["yc"] <= y_bot) or is_num(s):
                continue
            # 가장 가까운 줄: 율명 칸 중심 ≈ 머리 x, 가사 칸 ≈ 머리 x + 0.5 pitch
            best = None
            for ln, hx in xs:
                dx = s["xc"] - hx
                if -pitch * 0.3 <= dx < pitch * 0.7:
                    best = (ln, dx); break
            if not best:
                continue
            ln, dx = best
            r = min(range(len(rows)), key=lambda i: abs(rows[i] - s["yc"])) + 1
            kind = "y" if dx < pitch * 0.3 else "g"
            t = s["t"].strip()
            if kind == "y" and re.search(r"[A-Za-z]{2,}|[.]", t) and t not in ("D.S", "D.C"):
                kind = "n"                                              # Fine 같은 글자는 비고로
            if kind == "y" and t in ("D.S", "D.C", "D.S.", "Fine"):
                kind = "n"
            cells[(ln, r)][kind].append(s)
        grid = {}
        for (ln, r), c in cells.items():
            y = cell_text(c["y"])
            y = "".join(SYM.get(ch, ch) for ch in y)
            g = cell_text(c["g"], sep="").replace("-", "")
            n = " ".join(x["t"].strip() for x in c["n"])
            grid[(ln, r)] = (y, g, n)
        for ln, c in pick.items():                                      # 첫 행 비고로 남김
            y = "".join(SYM.get(ch, ch) for ch in cell_text(c["y"])); g = cell_text(c["g"], sep="")
            yy, gg, nn = grid.get((ln, 1), ("", "", ""))
            grid[(ln, 1)] = (yy, gg, (nn + " " if nn else "") + f"못갖춘마디(1행 앞): {y} {g}".strip())
        out.append(dict(lines=[ln for ln, _ in xs], nrows=len(rows), grid=grid))
    num, title, small = side_info(S, blocks)
    return dict(blocks=out, num=num, title=title, info=small)


def show(pno):
    d = pymupdf.open(PDF)
    P = parse_page(d[pno - 1])
    print(f"p{pno} 곡번호={P['num']!r} 제목={P['title']!r}\n  정보={P['info']!r}")
    for bi, b in enumerate(P["blocks"]):
        print(f"  블록{bi + 1}: 줄 {b['lines']} 행 {b['nrows']}")
        for r in range(1, b["nrows"] + 1):
            row = []
            for ln in sorted(b["lines"], reverse=True):
                y, g, n = b["grid"].get((ln, r), ("", "", ""))
                row.append(f"{y}{('/' + g) if g else ''}{('!' + n) if n else ''}")
            print(f"   {r:>2} " + " | ".join(f"{c:<10}" for c in row))


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--show":
        for p in sys.argv[2:]:
            show(int(p))


# ───────────────────────────── 전체 추출 ─────────────────────────────
def read_toc(d):
    """p5–8 목차: (곡번호, 제목, 정렬제목, 정렬번호) 반복. '곡번호' 머리 뒤부터 읽음."""
    toc = {}
    for i in range(4, 8):
        toks = [l.strip() for l in d[i].get_text().splitlines() if l.strip()]
        if "곡번호" in toks:
            toks = toks[toks.index("곡번호"):]
        elif i == 4:
            continue
        j = 0
        while j < len(toks) - 1:
            if toks[j].isdigit() and not toks[j + 1].isdigit() and not toks[j + 1].startswith(("제", "<")):
                n = int(toks[j])
                toc.setdefault(n, nfc(toks[j + 1]))
                j += 4 if j + 3 < len(toks) and toks[j + 3].isdigit() else 2
            else:
                j += 1
    return toc


FIX_TITLE = {"홍도야우지마라": "홍도야 우지마라"}   # 목차에 없는 곡(곡번호가 앞 곡과 겹침)의 띄어쓰기
INFO_KEYS = ("작사", "작곡", "노래", "편곡", "시", "가사")


JANGDAN = ["중중모리", "자진모리", "늦중모리", "중모리", "세마치", "굿거리", "엇모리", "휘모리", "진양조", "진양", "단모리", "동살풀이"]


def parse_info(info):
    """'중모리다음카페국악이꽃피는나무채보송경호' / '작사김창완작곡김창완노래산울림…' → 장단, {작사:…}"""
    info = info.split("다음카페")[0]
    info = re.sub(r"채보.*$", "", info)
    meta = {}
    pat = r"(작사|작곡|노래|편곡)"
    parts = re.split(pat, info)
    head = parts[0].strip()
    for k, v in zip(parts[1::2], parts[2::2]):
        if v.strip():
            meta.setdefault(k, v.strip())
    jangdan = next((j for j in JANGDAN if j in head), "")
    rest = re.sub(r"^\(\d\)|\(\d\)", "", head).replace(jangdan, "").strip(" .")
    if rest:
        meta.setdefault("정보", rest)
    return jangdan, meta


def clean_lyric(g):
    return re.sub(r"[△▽▼●○^~()\[\].,:;|]", "", g).strip()


def split_yul(y):
    """율명 칸에 들어온 한글(간주·후렴 같은 글씨)·숫자·(5) 같은 표시는 비고로 뺌"""
    extra = []
    for m in re.findall(r"\(\d+\)|\d+", y):
        extra.append(m); y = y.replace(m, "", 1)
    note = "".join(ch for ch in y if "\uac00" <= ch <= "\ud7a3") + (" " + " ".join(extra) if extra else "")
    note = note.strip()
    if note:
        y = "|".join(p for p in ("".join(ch for ch in part if not ("\uac00" <= ch <= "\ud7a3")) for part in y.split("|")) if p)
    return y, note


def main():
    d = pymupdf.open(PDF)
    toc = read_toc(d)
    os.makedirs(TSV_DIR, exist_ok=True); os.makedirs(SRC_DIR, exist_ok=True)
    for f in os.listdir(TSV_DIR):
        if f.endswith(".tsv"): os.remove(os.path.join(TSV_DIR, f))
    songs, cur = [], None
    for pno in range(1, d.page_count + 1):
        P = parse_page(d[pno - 1])
        if not P["blocks"]:
            continue
        m = re.match(r"\d+", P["num"].strip())
        num = None
        if m:
            t = m.group(0)
            if len(t) > 3 and len(t) % 2 == 0 and t[:len(t) // 2] == t[len(t) // 2:]:   # 블록 2개 쪽: '119119'
                t = t[:len(t) // 2]
            num = int(t)
        if num is None and cur is None:
            continue                                   # 첫 곡 앞(정간보 보는 법)
        base = re.sub(r"\(\d\)", "", P["title"]).replace(" ", "")
        if len(base) > 1 and len(base) % 2 == 0 and base[:len(base) // 2] == base[len(base) // 2:]:
            base = base[:len(base) // 2]                      # 블록 2개 쪽의 같은 제목 두 번
        other = bool(cur is not None and base and cur.get("base") and base != cur["base"] and num in (None, cur["num"]))
        if other and num == cur["num"]:
            print(f"  p{pno}: 곡번호 {num} 이 같지만 제목이 다름 → 다른 곡 ({base})")
        if num is not None and (cur is None or cur["num"] != num or other):
            jangdan, meta = parse_info(P["info"])
            title = toc.get(num, base) if not other else FIX_TITLE.get(base, base)
            cur = dict(num=num if not other else f"{num}-2", title=title, side=P["title"], base=base,
                       jangdan=jangdan, meta=meta, parts=[])
            songs.append(cur)
        for bi, b in enumerate(P["blocks"]):
            if any(y or g for y, g, n in b["grid"].values()):          # 빈 양식 쪽(p115 등)은 건너뜀
                cur["parts"].append(dict(page=pno, block=bi + 1, **b))
        png = os.path.join(SRC_DIR, f"p{pno:03d}.png")
        if not os.path.exists(png):
            d[pno - 1].get_pixmap(dpi=150).save(png)
    # 제목 겹치면 뒤의 것에 -2, -3
    seen = collections.Counter()
    for s in songs:
        seen[s["title"]] += 1
        s["name"] = s["title"] if seen[s["title"]] == 1 else f"{s['title']}-{seen[s['title']]}"
    rows_tsv = ["곡번호\t곡명\t목차제목\t쪽\t블록수\t행수\t장단\t작사\t작곡\t노래\t문제"]
    for s in songs:
        probs = []
        nrows = {p["nrows"] for p in s["parts"]}
        if len(nrows) > 1: probs.append(f"블록마다 행수 다름 {sorted(nrows)}")
        R = max(nrows); beat = {12: "3/4", 16: "4/4", 20: "정악"}.get(R, "4/4")
        if R not in (12, 16, 20): probs.append(f"행수 {R}")
        names = []
        for k, p in enumerate(s["parts"]):
            fn = (f"{s['num']:03d}" if isinstance(s['num'], int) else f"{int(str(s['num']).split('-')[0]):03d}b") + f"_{k + 1}.tsv"; names.append(fn)
        for k, p in enumerate(s["parts"]):
            base = min(p["lines"]) - 1
            nl = max(p["lines"]) - base
            if nl > 10: probs.append(f"p{p['page']} 블록 {nl}줄 > 10")
            meta = [("제목", s["name"]), ("부제", s["jangdan"]), ("키", ""), ("쪽", str(p["page"])), ("박자", beat),
                    ("행수", str(R)), ("줄수", str(nl)), ("원본", f"p{p['page']:03d}.png"),
                    ("이어짐", names[k + 1] if k + 1 < len(names) else ""),
                    ("메모", f"PDF 글자 층에서 자동 추출(이미지 판독 아님). 곡번호 {s['num']}, PDF {p['page']}쪽 블록{p['block']}, 줄 {min(p['lines'])}–{max(p['lines'])}."
                             + (" " + " ".join(f"{a} {b}" for a, b in s["meta"].items()) if s["meta"] else ""))]
            out = [f"#{a}\t{b}" for a, b in meta] + ["줄\t정간\t율명\t가사\t비고"]
            for ln in range(base + 1, base + nl + 1):
                for r in range(1, R + 1):
                    y, g, n = p["grid"].get((ln, r), ("", "", ""))
                    y, hn = split_yul(y)
                    notes = [x for x in (n, hn) if x and x.strip(" .")]
                    out.append(f"{ln - base}\t{r}\t{y}\t{clean_lyric(g)}\t{'; '.join(notes)}")
            open(os.path.join(TSV_DIR, names[k]), "w", encoding="utf-8", newline="\n").write("\n".join(out) + "\n")
        rows_tsv.append("\t".join([str(s["num"]), s["name"], toc.get(s["num"], ""), ",".join(str(p["page"]) for p in s["parts"]),
                                   str(len(s["parts"])), str(R), s["jangdan"], s["meta"].get("작사", ""), s["meta"].get("작곡", ""),
                                   s["meta"].get("노래", ""), "; ".join(probs)]))
    open(os.path.join(WORK, "곡목록.tsv"), "w", encoding="utf-8", newline="\n").write("\n".join(rows_tsv) + "\n")
    print(f"곡 {len(songs)} / 목차 {len(toc)} / TSV {sum(len(s['parts']) for s in songs)}")
    missing = sorted(set(toc) - {s['num'] for s in songs if isinstance(s['num'], int)}); print("목차에 있는데 못 찾은 곡번호", missing)
    print("문제 있는 곡", [(s.split('\t')[0], s.split('\t')[1], s.split('\t')[-1]) for s in rows_tsv[1:] if s.split('\t')[-1]])


if __name__ == "__main__" and len(sys.argv) == 1:
    main()
