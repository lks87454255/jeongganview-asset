#!/usr/bin/env python3
"""남은 일: 지도 없는 쪽, 지도에 있는데 TSV 없는 곡."""
import json, os, sys
W = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
pages = {}
for fn in os.listdir(os.path.join(W, "pages")):
    if fn.endswith(".json"):
        try:
            pages[fn[:-5]] = json.load(open(os.path.join(W, "pages", fn), encoding="utf-8"))
        except ValueError:
            print("깨진 지도", fn)
tsv = {f[:-4] for f in os.listdir(os.path.join(W, "tsv")) if f.endswith(".tsv")}
want2 = [f"b2_{p:03d}" for p in range(13, 241)]
b1 = sorted({int(p) + 6 for l in list(open(os.path.join(W, "목차.tsv"), encoding="utf-8"))[1:]
             for c in [l.rstrip("\n").split("\t")] if c[4] == "1권" for p in c[3].split(",")})
want1 = [f"b1_{p:03d}" for p in b1]
nomap = [k for k in want2 + want1 if k not in pages]
notsv = [(k, s["id"], s.get("title", "")) for k, m in sorted(pages.items()) for s in m.get("songs", []) if s["id"] not in tsv]
print("지도 없는 쪽", len(nomap), " ".join(nomap))
print("TSV 없는 곡", len(notsv))
for k, i, t in notsv:
    print(f"  {k} {i} {t}")
