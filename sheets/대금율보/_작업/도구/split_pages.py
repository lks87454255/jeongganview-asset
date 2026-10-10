#!/usr/bin/env python3
"""S1: 쪽별 곡 영역 지도(_작업/pages/b{권}_{PDF쪽}.json) 초안 만들기 · 확인 이미지 그리기.

  python3 split_pages.py draft b2_013 b2_016 …   # OpenCV 테두리 박스 → 초안 JSON (이미 있으면 건드리지 않음)
  python3 split_pages.py show  b2_013 …          # 지도 박스·곡ID + 0.05 눈금 → /tmp/dg_pan/show/b2_013.png
  python3 split_pages.py grid  b2_013 …          # 눈금만 (지도 만들기 전에 좌표 읽기용)
  python3 split_pages.py check                   # 모든 지도 검사 (목차.tsv 곡ID·형식)
쪽 표기: b2_013 (PDF 쪽) 또는 b2_p11 (인쇄 쪽).

초안은 자동 검출이라 틀리기 쉽다(설계서 1절). show 이미지를 보고 JSON 을 직접 고친 뒤 "confirmed": true 로 둔다.
"""
import json
import os
import sys

from common import (OFFSET, PAGES_DIR, TMP, all_maps, check_box, load_map, page_key, parse_page_arg, render, WORK)

DPI = 100


def detect_boxes(img):
    import cv2
    import numpy as np
    g = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2GRAY)
    bw = cv2.adaptiveThreshold(g, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 25, 15)
    H, W = bw.shape
    v = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, H // 12)))
    h = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (W // 12, 1)))
    grid = cv2.dilate(v | h, np.ones((5, 5), np.uint8))
    cnts, _ = cv2.findContours(grid, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    boxes = []
    for c in cnts:
        x, y, w, hh = cv2.boundingRect(c)
        if w * hh > 0.02 * W * H:
            boxes.append([round(x / W, 3), round(y / H, 3), round(w / W, 3), round(hh / H, 3)])
    # 오른쪽 위 → 왼쪽, 위 → 아래
    boxes.sort(key=lambda b: (round(b[1] * 4), -b[0]))
    return boxes


def draft(book, pdf):
    path = os.path.join(PAGES_DIR, page_key(book, pdf) + ".json")
    if os.path.exists(path):
        print(f"있음, 건너뜀: {path}")
        return
    img = render(book, pdf, DPI)
    printed = pdf - OFFSET[book]
    songs = []
    for i, b in enumerate(detect_boxes(img), 1):
        songs.append({"id": f"b{book}_{printed:03d}_{i}", "title": "", "parts": [{"box": b, "lines": 0, "rows": 0}],
                      "beat": "", "jangdan": "", "tempo": "", "sections": "", "repeat": "", "memo": "자동 검출 초안"})
    m = {"book": book, "pdfPage": pdf, "printedPage": printed, "confirmed": False, "songs": songs}
    os.makedirs(PAGES_DIR, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(m, f, ensure_ascii=False, indent=2)
    print(f"초안 {len(songs)}곡: {path}")


def draw(book, pdf, with_map):
    from PIL import ImageDraw, ImageFont
    img = render(book, pdf, DPI)
    W, H = img.size
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("/System/Library/Fonts/AppleSDGothicNeo.ttc", 14)
    except OSError:
        font = ImageFont.load_default()
    for k in range(1, 20):
        t = k / 20
        col = (0, 160, 255) if k % 2 == 0 else (170, 220, 255)
        d.line([(t * W, 0), (t * W, H)], fill=col, width=1)
        d.line([(0, t * H), (W, t * H)], fill=col, width=1)
        if k % 2 == 0:
            d.text((t * W + 2, 2), f"{t:.1f}", fill=(0, 100, 200), font=font)
            d.text((2, t * H + 2), f"{t:.1f}", fill=(0, 100, 200), font=font)
    m = load_map(book, pdf) if with_map else None
    if m:
        for s in m["songs"]:
            for j, p in enumerate(s["parts"], 1):
                x, y, w, h = p["box"]
                d.rectangle([x * W, y * H, (x + w) * W, (y + h) * H], outline=(255, 0, 0), width=3)
                d.text((x * W + 4, y * H + 4), f"{s['id']}#{j} {s.get('title', '')}", fill=(255, 0, 0), font=font)
    out = os.path.join(TMP, "show", page_key(book, pdf) + ".png")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    img.save(out)
    print(out)


def load_toc_ids():
    ids = {}
    with open(os.path.join(WORK, "목차.tsv"), encoding="utf-8") as f:
        next(f)
        for ln in f:
            c = ln.rstrip("\n").split("\t")
            ids[c[0]] = c[1]
    return ids


def check():
    toc = load_toc_ids()
    errors, seen = [], {}
    for m in all_maps():
        key = page_key(m["book"], m["pdfPage"])
        if m["printedPage"] != m["pdfPage"] - OFFSET[m["book"]]:
            errors.append(f"{key}: printedPage 불일치")
        for s in m["songs"]:
            sid = s["id"]
            if sid in seen:
                errors.append(f"{key}: 곡ID 중복 {sid} ({seen[sid]})")
            seen[sid] = key
            if m.get("confirmed") and sid not in toc:
                errors.append(f"{key}: {sid} 가 목차.tsv 에 없음 (목차.tsv 곡ID 를 맞출 것)")
            if not s.get("parts"):
                errors.append(f"{key}: {sid} parts 없음")
            for j, p in enumerate(s.get("parts", []), 1):
                try:
                    check_box(p["box"], f"{key} {sid}#{j}")
                except ValueError as e:
                    errors.append(str(e))
                if m.get("confirmed") and (p.get("lines", 0) < 1 or p.get("rows", 0) < 1):
                    errors.append(f"{key}: {sid}#{j} lines/rows 없음")
    for e in errors:
        print("오류:", e)
    print(f"지도 {len(all_maps())}쪽, 곡 {len(seen)}, 오류 {len(errors)}")
    return 1 if errors else 0


def main(argv):
    if not argv:
        print(__doc__)
        return 2
    cmd, pages = argv[0], [parse_page_arg(a) for a in argv[1:]]
    if cmd == "check":
        return check()
    for b, p in pages:
        if cmd == "draft":
            draft(b, p)
        elif cmd == "show":
            draw(b, p, True)
        elif cmd == "grid":
            draw(b, p, False)
        else:
            print(__doc__)
            return 2
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
