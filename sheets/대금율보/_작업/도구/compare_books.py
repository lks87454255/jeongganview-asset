#!/usr/bin/env python3
"""결정 4: 1권·2권에 같이 있는 곡이 내용까지 같은지 자동 비교.

목차_S0.tsv 에서 두 권 쪽이 다 있는 곡(짝) → 2권 쪽 지도(pages/*.json)의 곡 영역(part 마다)을
1권 해당 인쇄 쪽(±1쪽) 전체에서 템플릿 매칭(여러 배율)한다. 같은 판이면 점수가 높다(0.8 이상).
점수가 낮은 곡만 사람이 눈으로 확인한다 → /tmp/dg_pan/cmp/{곡ID}.png (왼쪽 2권 곡 영역, 오른쪽 1권 쪽).

  python3 compare_books.py            → _작업/권비교.tsv + 요약
  python3 compare_books.py --img      → 낮은 점수 곡 비교 이미지도 만듦
"""
import os
import re
import sys

import cv2
import numpy as np

from common import OFFSET, TMP, WORK, all_maps, nfc, render

DPI = 40
SX = [0.90 + 0.02 * i for i in range(11)]       # 가로 배율 0.90 ~ 1.10
SY = [0.85 + 0.02 * i for i in range(16)]       # 세로 배율 0.85 ~ 1.15 (두 권 줄 간격이 조금 다름)
BLUR = 2.0                                      # 글꼴 차이를 줄이려고 반전 + 흐림 뒤 비교
LOW = 0.6
# 목차 곡명과 지도 곡명이 달라 이름으로 못 잇는 짝 (목차 곡명 → 2권 곡ID)
ALIAS = {"둥대지기": "b2_016_2", "어메징 그레이스": "b2_021_2", "첨밀밀": "b2_030_1", "엽불승": "b2_072_2",
         "엘콘돌파샤": "b2_085_1", "청성곡": "b2_159_1", "엮음 지름시조": "b2_178_1", "한": "b2_187_1"}


def key(s):
    return re.sub(r"[\s()·,.\-]|中", "", nfc(s))


def gray(img):
    a = 255 - np.asarray(img.convert("L"), dtype=np.float32)
    return cv2.GaussianBlur(a, (0, 0), BLUR)


_page_cache = {}


def page_img(book, pdf_page):
    k = (book, pdf_page)
    if k not in _page_cache:
        _page_cache[k] = gray(render(book, pdf_page, DPI))
    return _page_cache[k]


def best_match(tpl, page):
    best = (-1.0, None, None)
    for sx in SX:
        for sy in SY:
            t = cv2.resize(tpl, None, fx=sx, fy=sy, interpolation=cv2.INTER_AREA)
            if t.shape[0] >= page.shape[0] or t.shape[1] >= page.shape[1] or min(t.shape) < 10:
                continue
            r = cv2.matchTemplate(page, t, cv2.TM_CCOEFF_NORMED)
            _, mx, _, loc = cv2.minMaxLoc(r)
            if mx > best[0]:
                best = (mx, (sx, sy), loc)
    return best


def main():
    make_img = "--img" in sys.argv
    s0 = [l.rstrip("\n").split("\t") for l in open(os.path.join(WORK, "목차_S0.tsv"), encoding="utf-8")][1:]
    pairs = [r for r in s0 if len(r) > 3 and r[2].strip() and r[3].strip()]
    # 2권 지도: 인쇄 쪽 → 곡들
    by_page = {}
    for m in all_maps():
        if m["book"] != 2:
            continue
        for s in m.get("songs", []):
            by_page.setdefault(m["printedPage"], []).append((s, m["pdfPage"]))
    out, nomatch = [], []
    for r in pairs:
        title, p2, p1s = r[1], int(r[2]), [int(x) for x in re.findall(r"\d+", r[3])]
        cands = [(s, pdf) for s, pdf in by_page.get(p2, []) if s["id"] == ALIAS.get(title) or key(s.get("title", "")).startswith(key(title)) or key(title).startswith(key(s.get("title", "")))]
        if not cands:
            nomatch.append(r)
            continue
        for s, pdf in cands:
            scores = []
            for part in s["parts"]:
                ppdf = part.get("pdfPage", pdf)
                tpl = gray(render(2, ppdf, DPI, part["box"]))
                best = (-1.0, None, None, None)
                for p1 in p1s:
                    for d in (0, 1, -1):
                        q = p1 + OFFSET[1] + d
                        if q < 1 or q > 262:
                            continue
                        sc, scale, loc = best_match(tpl, page_img(1, q))
                        if sc > best[0]:
                            best = (sc, scale, loc, q)
                        if best[0] >= 0.75:
                            break
                scores.append(best)
            mn = min(b[0] for b in scores)
            out.append((s["id"], s.get("title", ""), p2, r[3], mn, scores))
            if make_img and mn < LOW:
                os.makedirs(os.path.join(TMP, "cmp"), exist_ok=True)
                g = lambda im: np.asarray(im.convert("L"), dtype=np.uint8)
                a = g(render(2, s["parts"][0].get("pdfPage", pdf), 60, s["parts"][0]["box"]))
                pages = [g(render(1, b[3], 60)) for b in scores[:1]] + [g(render(1, p + OFFSET[1], 60)) for p in p1s[:1] if p + OFFSET[1] != scores[0][3]]
                h = max([a.shape[0]] + [p.shape[0] for p in pages])
                pad = lambda x: np.pad(x, ((0, h - x.shape[0]), (0, 8)), constant_values=255)
                cv2.imwrite(os.path.join(TMP, "cmp", f"{s['id']}.png"), np.hstack([pad(a)] + [pad(p) for p in pages]))
    out.sort(key=lambda x: x[4])
    with open(os.path.join(WORK, "권비교.tsv"), "w", encoding="utf-8") as f:
        f.write("곡ID\t곡명\t2권_인쇄쪽\t1권_인쇄쪽\t최저점수\tpart별(점수@1권PDF쪽)\n")
        for sid, t, p2, p1, mn, sc in out:
            f.write(f"{sid}\t{t}\t{p2}\t{p1}\t{mn:.3f}\t" + " ".join(f"{b[0]:.2f}@{b[3]}" for b in sc) + "\n")
    print(f"짝 {len(pairs)} → 비교 {len(out)} 곡, 2권 지도에서 못 찾음 {len(nomatch)}")
    for r in nomatch:
        print("   못 찾음:", r[0], r[1], r[2], r[3])
    for th in (0.9, 0.8, 0.7, 0.6, 0.5):
        print(f"  < {th}: {sum(1 for x in out if x[4] < th)}")


if __name__ == "__main__":
    main()
