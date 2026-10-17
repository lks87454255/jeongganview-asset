#!/usr/bin/env python3
"""국악이 꽃피는 나무.pdf 쪽 구조 조사: 글꼴별 글자, 줄 번호 머리(1..8), 행 번호, 제목(세로 큰 글자), 곡번호."""
import pymupdf, collections, json, sys
PDF = "book/국악이 꽃피는 나무/국악이 꽃피는 나무.pdf"
d = pymupdf.open(PDF)
def spans(p):
    out = []
    for b in p.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                if s["text"].strip(): out.append(dict(x0=s["bbox"][0], y0=s["bbox"][1], x1=s["bbox"][2], y1=s["bbox"][3], size=round(s["size"], 1), font=s["font"], t=s["text"]))
    return out
res = []; fonts = collections.Counter(); yul = collections.Counter()
for i, p in enumerate(d):
    S = spans(p)
    hdr = [s for s in S if s["font"] == "HYwulM" and 85 < s["y0"] < 95 and s["t"].strip().isdigit()]
    rows = [s for s in S if s["font"] == "HYwulM" and s["x0"] < 85 and s["t"].strip().isdigit()]
    big = [s for s in S if s["font"] == "MHunmin"]
    num = [s for s in S if s["font"] == "HYwulM" and s["size"] > 18]
    for s in S: fonts[(s["font"], s["size"])] += 1
    for s in S:
        if s["font"] == "HYhaeseo":
            for ch in s["t"].strip(): yul[ch] += 1
    res.append(dict(page=i + 1, lines=len(hdr), rows=max([int(s["t"]) for s in rows] or [0]), title=''.join(s["t"] for s in sorted(big, key=lambda s: s["y0"])), num=''.join(s["t"] for s in num)))
json.dump(res, open('/tmp/gk_survey.json', 'w'), ensure_ascii=False)
print('글꼴', fonts.most_common(20)); print('율명칸 글자', yul.most_common())
c = collections.Counter((r["lines"], r["rows"]) for r in res); print('줄×행', c.most_common())
for r in res[24:40]: print(r)
