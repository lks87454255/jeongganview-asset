#!/usr/bin/env python3
"""S0: 목차_2권.tsv + 목차_1권.tsv → 목차.tsv (곡 목록).

규칙(추출설계서 S0)
- 2권 우선, 2권에 없는 곡만 1권 채택.
- 같은 곡 판별: 곡명에서 공백·'中' 을 뺀 이름이 같으면(별칭표 ALIAS 포함) 같은 곡.
- 쪽이 여러 개인 곡(예: 사랑가 47,148)은 버전마다 한 행. 두 권의 버전은 쪽 순서대로 짝짓는다.
- 곡ID = {권}_{인쇄쪽 3자리}_{순번}. 순번은 임시(그 쪽 안 목차 가나다 순) → S1 지도에서 위치 순으로 다시 매긴다.

사용: python3 make_toc.py   (sheets/대금율보/_작업 에서 또는 아무 곳에서)
"""
import os
import sys
from collections import defaultdict

WORK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PAGE_RANGE = {2: (11, 238), 1: (11, 254)}
# 목차 표기가 권마다 다른 같은 곡
ALIAS = {"엽불승": "염불승"}


def key(name):
    k = name.replace(" ", "").replace("中", "")
    return ALIAS.get(k, k)


def load(book):
    path = os.path.join(WORK, f"목차_{book}권.tsv")
    rows, errors = [], []
    with open(path, encoding="utf-8") as f:
        lines = [l.rstrip("\n") for l in f if l.strip() and not l.startswith("#")]
    for n, line in enumerate(lines[1:], 2):  # 첫 줄은 머리글
        cols = line.split("\t")
        name, pages = cols[0].strip(), cols[1].strip()
        note = cols[2].strip() if len(cols) > 2 else ""
        try:
            ps = [int(p) for p in pages.split(",")]
        except ValueError:
            errors.append(f"{book}권 {n}행 쪽 번호 오류: {line!r}")
            continue
        lo, hi = PAGE_RANGE[book]
        for p in ps:
            if not lo <= p <= hi:
                errors.append(f"{book}권 {name}: 쪽 {p} 범위({lo}–{hi}) 밖")
        rows.append((name, sorted(ps), note))
    seen = defaultdict(list)
    for name, _, _ in rows:
        seen[key(name)].append(name)
    for k, v in seen.items():
        if len(v) > 1:
            errors.append(f"{book}권 목차에 같은 곡 중복: {v}")
    return rows, errors


def main():
    b2, e2 = load(2)
    b1, e1 = load(1)
    errors = e2 + e1
    by1 = {key(n): (n, ps, note) for n, ps, note in b1}
    used1 = set()
    out = []  # (권, 쪽, 곡명, 2권쪽, 1권쪽, 채택, 비고)
    for name, ps2, note2 in b2:
        k = key(name)
        ps1, note1 = [], ""
        if k in by1:
            used1.add(k)
            _, ps1, note1 = by1[k]
        for i, p in enumerate(ps2):
            p1 = ps1[i] if i < len(ps1) else ""
            notes = [x for x in (note2, note1 and f"1권: {note1}") if x]
            if len(ps2) > 1:
                notes.insert(0, f"목차 {','.join(map(str, ps2))} → 버전 {i + 1}")
            out.append((2, p, name, p, p1, "2권", "; ".join(notes)))
        # 1권에만 있는 추가 버전
        for i in range(len(ps2), len(ps1)):
            out.append((1, ps1[i], name, "", ps1[i], "1권",
                        f"1권 목차 {','.join(map(str, ps1))} → 버전 {i + 1}, 2권에 없음"))
    for name, ps1, note1 in b1:
        if key(name) in used1:
            continue
        for i, p in enumerate(ps1):
            notes = ["2권에 없음"] + ([note1] if note1 else [])
            if len(ps1) > 1:
                notes.insert(0, f"목차 {','.join(map(str, ps1))} → 버전 {i + 1}")
            out.append((1, p, name, "", p, "1권", "; ".join(notes)))

    # 곡ID: 권·쪽별 임시 순번
    out.sort(key=lambda r: (-r[0], r[1], key(r[2])))
    counter = defaultdict(int)
    lines = ["곡ID\t곡명\t2권_인쇄쪽\t1권_인쇄쪽\t채택\t분류\t비고"]
    for book, page, name, p2, p1, adopt, note in out:
        counter[(book, page)] += 1
        sid = f"b{book}_{page:03d}_{counter[(book, page)]}"
        lines.append(f"{sid}\t{name}\t{p2}\t{p1}\t{adopt}\t\t{note}")

    dst = os.path.join(WORK, "목차.tsv")
    with open(dst, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    n2 = sum(1 for r in out if r[5] == "2권")
    n1 = len(out) - n2
    print(f"2권 목차 {len(b2)}곡, 1권 목차 {len(b1)}곡")
    print(f"목차.tsv {len(out)}행 (2권 채택 {n2}, 1권 전용 {n1}) → {dst}")
    only2 = sorted(n for n, _, _ in b2 if key(n) not in by1)
    print(f"2권에만 있는 곡 {len(only2)}: {' / '.join(only2)}")
    for e in errors:
        print("오류:", e, file=sys.stderr)
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
