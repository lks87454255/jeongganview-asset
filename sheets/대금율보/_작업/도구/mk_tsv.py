#!/usr/bin/env python3
"""판독 TSV 를 짧은 표기로 쓰는 도우미 (S3).

  from mk_tsv import mk
  mk("b2_022_1", "소금장수", "4/4", 16, [ "줄1", "줄2", ... ], book=2, printed=22,
     memo="", jangdan="잦은모리", tempo="", sections="1-3:전주 4-11:노래", repeat="")

줄 문자열 = 정간 수(rows)만큼 쉼표(,)로 나눈 칸. 오른쪽 첫 줄이 줄1, 칸은 위에서부터.
칸 = "율명;가사;비고" (뒤는 생략 가능). 율명 칸 안 위→아래는 | 로 구분. 빈 칸은 "".
쉼표(,)·세미콜론(;)·탭은 칸 안에 쓰지 말 것(비고에도). 칸 수가 rows 와 다르면 오류.
결과: _작업/tsv/{곡ID}.tsv  (#원본 = {곡ID}.png, #출처 = "{book}권 {printed}쪽")
"""
import os

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tsv")


def mk(sid, title, beat, rows, lines, book, printed, memo="", jangdan="", tempo="", sections="", repeat="", sub=""):
    assert beat in ("4/4", "3/4", "2/4", "정악"), beat
    L = [f"#제목\t{title}", f"#부제\t{sub}", "#키\t", "#쪽\t", f"#박자\t{beat}", f"#행수\t{rows}", f"#줄수\t{len(lines)}",
         f"#원본\t{sid}.png", "#이어짐\t", f"#메모\t{memo}", f"#장단\t{jangdan}", f"#빠르기\t{tempo}", f"#구간\t{sections}",
         f"#반복\t{repeat}", f"#출처\t{book}권 {printed}쪽", "줄\t정간\t율명\t가사\t비고"]
    for i, ln in enumerate(lines, 1):
        cells = [c.strip() for c in ln.split(",")]
        if len(cells) != rows:
            raise SystemExit(f"{sid} 줄{i}: 칸 {len(cells)}개 ≠ 행수 {rows}")
        for j, c in enumerate(cells, 1):
            p = (c.split(";") + ["", ""])[:3]
            L.append(f"{i}\t{j}\t{p[0].strip()}\t{p[1].strip()}\t{p[2].strip()}")
    os.makedirs(OUT, exist_ok=True)
    with open(os.path.join(OUT, sid + ".tsv"), "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")
    print(f"{sid} {title}: {len(lines)}줄 × {rows}")
