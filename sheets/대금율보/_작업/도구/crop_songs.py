#!/usr/bin/env python3
"""S2: 쪽 지도(pages/*.json, confirmed) → 곡별 crop.

  python3 crop_songs.py [곡ID …]      # 없으면 확정된 지도의 모든 곡
  python3 crop_songs.py --force …     # 원본/{곡ID}.png 가 있어도 다시 만듦

산출
  sheets/대금율보/원본/{곡ID}.png      300dpi 회색조. titleBox(있으면) 위에, parts 를 읽는 순서대로 아래로 이어 붙임(오른쪽 맞춤).
                                      검수 때 원본 대조용 — 저장소에 커밋한다.
  /tmp/dg_pan/{곡ID}/                  판독용 조각 (200dpi). part 마다 오른쪽 줄부터 최대 LINES_PER_TILE 줄씩,
                                      높이가 MAX_SIDE 를 넘으면 위/아래로 나눔. 한 변 ≤ 2000px (세션 멈춤 방지).
    t{part}_{k}{T|B}.png             k = 1 이 가장 오른쪽. 조각끼리 약간 겹친다.
    full.png                         곡 전체 축소본 (긴 변 1600px) — 배치 확인용
"""
import os
import sys

from common import ORIG_DIR, TMP, all_maps, render

ORIG_DPI = 300
TILE_DPI = 200
LINES_PER_TILE = 4
MAX_SIDE = 1900
OVERLAP = 0.04  # 조각 겹침 (part 폭/높이 대비)


def stitch(images):
    from PIL import Image
    W = max(i.width for i in images)
    H = sum(i.height for i in images) + 20 * (len(images) - 1)
    out = Image.new("L", (W, H), 255)
    y = 0
    for im in images:
        out.paste(im, (W - im.width, y))  # 오른쪽 맞춤 (1줄이 오른쪽)
        y += im.height + 20
    return out


def tiles_for_part(book, pdf, part, pi, dst):
    x, y, w, h = part["box"]
    lines = max(1, part.get("lines") or 1)
    lw = w / lines
    names = []
    k = 0
    for first in range(0, lines, LINES_PER_TILE):
        k += 1
        n = min(LINES_PER_TILE, lines - first)
        # 오른쪽 끝에서 first 줄 건너뛴 위치부터 n 줄 (가사 칸이 율명 오른쪽이라 오른쪽으로 조금 더)
        right = x + w - first * lw
        left = right - n * lw
        ox = OVERLAP * w
        cx0, cx1 = max(x, left - ox), min(x + w, right + ox * 0.5)
        img = render(book, pdf, TILE_DPI, [cx0, y, cx1 - cx0, h]).convert("L")
        if img.height <= MAX_SIDE:
            img.save(os.path.join(dst, f"t{pi}_{k}.png"))
            names.append(f"t{pi}_{k}.png")
        else:
            half = h / 2
            oy = OVERLAP * h
            for tag, (yy, hh) in (("T", (y, half + oy)), ("B", (y + half - oy, half + oy))):
                im = render(book, pdf, TILE_DPI, [cx0, yy, cx1 - cx0, min(hh, 1 - yy)]).convert("L")
                im.save(os.path.join(dst, f"t{pi}_{k}{tag}.png"))
                names.append(f"t{pi}_{k}{tag}.png")
    return names


def crop_song(m, s, force):
    book, pdf = m["book"], m["pdfPage"]
    sid = s["id"]
    out = os.path.join(ORIG_DIR, f"{sid}.png")
    parts = []
    if s.get("titleBox"):
        parts.append(render(book, pdf, ORIG_DPI, s["titleBox"]).convert("L"))
    for p in s["parts"]:
        b, pg = p["box"], p.get("pdfPage", pdf)  # 다음 쪽으로 이어지는 part 는 pdfPage 를 따로 적는다
        parts.append(render(book, pg, ORIG_DPI, b).convert("L"))
    os.makedirs(ORIG_DIR, exist_ok=True)
    if force or not os.path.exists(out):
        stitch(parts).save(out, optimize=True)
    dst = os.path.join(TMP, sid)
    os.makedirs(dst, exist_ok=True)
    names = []
    for pi, p in enumerate(s["parts"], 1):
        names += tiles_for_part(book, p.get("pdfPage", pdf), p, pi, dst)
    full = stitch(parts)
    full.thumbnail((1600, 1600))
    full.save(os.path.join(dst, "full.png"))
    print(f"{sid} {s.get('title', '')}: 원본/{sid}.png, 조각 {len(names)} → {dst}")


def main(argv):
    force = "--force" in argv
    want = {a for a in argv if not a.startswith("--")}
    n = 0
    for m in all_maps():
        if not m.get("confirmed"):
            continue
        for s in m["songs"]:
            if want and s["id"] not in want:
                continue
            crop_song(m, s, force)
            n += 1
    if want and n < len(want):
        print(f"경고: 확정 지도에서 못 찾은 곡ID 가 있음 (찾음 {n} / 요청 {len(want)})")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
