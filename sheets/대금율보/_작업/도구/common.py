"""대금율보 도구 공통: 경로, PDF 쪽 렌더링, 쪽 지도(JSON) 읽기."""
import json
import os
import unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.dirname(HERE)                      # sheets/대금율보/_작업
ROOT = os.path.dirname(WORK)                      # sheets/대금율보
REPO = os.path.dirname(os.path.dirname(ROOT))     # jeongganview-asset
PAGES_DIR = os.path.join(WORK, "pages")
ORIG_DIR = os.path.join(ROOT, "원본")
TMP = "/tmp/dg_pan"                               # 판독용 조각·확인 이미지 (저장소 밖)

PDF = {b: os.path.join(REPO, "book", "대금율보", f"대금율보 {b}.pdf") for b in (1, 2)}
OFFSET = {1: 6, 2: 2}                             # PDF 쪽 = 인쇄 쪽 + OFFSET


def nfc(s):
    return unicodedata.normalize("NFC", s or "")


def page_key(book, pdf_page):
    return f"b{book}_{pdf_page:03d}"


def parse_page_arg(s):
    """'b2_013' · 'b2:13' · '2:13' → (권, PDF쪽). 'b2_p22'(p = 인쇄 쪽)도 받음."""
    s = s.strip().lower().lstrip("b")
    sep = "_" if "_" in s else ":"
    b, p = s.split(sep, 1)
    if p.startswith("p"):
        return int(b), int(p[1:]) + OFFSET[int(b)]
    return int(b), int(p)


_docs = {}


def doc(book):
    import pymupdf
    if book not in _docs:
        _docs[book] = pymupdf.open(PDF[book])
    return _docs[book]


def render(book, pdf_page, dpi, clip=None):
    """PDF 쪽(1부터)을 PIL 이미지(RGB)로. clip = 쪽 대비 0~1 [x, y, w, h]."""
    import pymupdf
    from PIL import Image
    pg = doc(book)[pdf_page - 1]
    rect = None
    if clip:
        x, y, w, h = clip
        r = pg.rect
        rect = pymupdf.Rect(r.x0 + x * r.width, r.y0 + y * r.height,
                            r.x0 + (x + w) * r.width, r.y0 + (y + h) * r.height)
    pix = pg.get_pixmap(dpi=dpi, clip=rect, colorspace=pymupdf.csRGB)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def load_map(book, pdf_page):
    path = os.path.join(PAGES_DIR, page_key(book, pdf_page) + ".json")
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def all_maps():
    out = []
    if not os.path.isdir(PAGES_DIR):
        return out
    for fn in sorted(os.listdir(PAGES_DIR)):
        if fn.endswith(".json"):
            with open(os.path.join(PAGES_DIR, fn), encoding="utf-8") as f:
                out.append(json.load(f))
    return out


def check_box(box, where):
    if not (isinstance(box, list) and len(box) == 4):
        raise ValueError(f"{where}: box 는 [x, y, w, h]")
    x, y, w, h = box
    if not (0 <= x < 1 and 0 <= y < 1 and 0 < w <= 1 and 0 < h <= 1 and x + w <= 1.0001 and y + h <= 1.0001):
        raise ValueError(f"{where}: box 범위 밖 {box}")
