#!/usr/bin/env python3
"""쪽 지도(pages/*.json) → 목차.tsv 다시 만들기 (S1 이후 곡 목록의 기준은 지도).

목차_S0.tsv(책 목차 판독)와 곡명으로 대조해, 지도에 없는 목차 곡을 '목차에만 있음' 으로 보고한다.
"""
import json, os, re, unicodedata
W = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
def key(s): return re.sub(r"[\s()·,.\-]|中", "", unicodedata.normalize("NFC", s or ""))
songs = []
for fn in sorted(os.listdir(os.path.join(W, "pages"))):
    if not fn.endswith(".json"): continue
    m = json.load(open(os.path.join(W, "pages", fn), encoding="utf-8"))
    for s in m.get("songs", []):
        pages = sorted({m["pdfPage"]} | {p.get("pdfPage", m["pdfPage"]) for p in s.get("parts", [])})
        songs.append((s["id"], s.get("title", ""), m["book"], m["printedPage"], s.get("genre", ""), pages, s.get("memo", "")))
tsv = {f[:-4] for f in os.listdir(os.path.join(W, "tsv")) if f.endswith(".tsv")}
s0 = [l.rstrip("\n").split("\t") for l in open(os.path.join(W, "목차_S0.tsv"), encoding="utf-8")][1:]
have = {key(t) for _, t, *_ in songs}
with open(os.path.join(W, "목차.tsv"), "w", encoding="utf-8") as f:
    f.write("곡ID\t곡명\t2권_인쇄쪽\t1권_인쇄쪽\t채택\t분류\t비고\n")
    for sid, t, b, pp, g, pages, memo in sorted(songs, key=lambda x: (x[2] != 2, x[0])):
        off = 2 if b == 2 else 6
        note = "" if sid in tsv else "TSV 없음"
        f.write(f"{sid}\t{t}\t{pp if b == 2 else ''}\t{pp if b == 1 else ''}\t{b}권\t{g}\t{note}\n")
missing = [r for r in s0 if key(r[1]) not in have]
print(f"지도 곡 {len(songs)} (2권 {sum(1 for s in songs if s[2]==2)}, 1권 {sum(1 for s in songs if s[2]==1)}), TSV {len(tsv)}")
print(f"목차_S0 에만 있는 곡 {len(missing)}:")
for r in missing: print("  ", r[0], r[1], r[2], r[3], r[4])
