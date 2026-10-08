# TSV v2 (판독규격.md) 형식 검사. 판독 직후 한 파일씩 돌려 본다.
# 사용: python3 check_tsv.py <tsv 파일 또는 폴더> ...   (인자 없으면 _작업/tsv 전체)
# 검사: 머리 정보, 박자/행수, 본문 행 수 = 줄수 × 행수, 줄·정간 중복/누락, 율명 칸 팔레트 외 글자
import os, sys, unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
TSV_DIR = os.path.normpath(os.path.join(HERE, "..", "tsv"))

YUL = ("㣴㣕㣖㣣㣨㣡㣸㣩𢓡㣮㣳㣹" "僙㐲㑀俠㑬㑖𠐭㑣侇㑲㒇㒣"
       "黃大太夾姑仲㽔林夷南無應" "潢汏汰浹㴌㳞㶋淋洟湳潕㶐" "㶂𣴘㳲㴺㵈㴢㶙㵉㴣㵜㶃㶝")
MARKS = "―△Иﾉ^⌝ㅋ⌞է⊏⊔·○⁚‹／"
OK = set(YUL) | set(MARKS) | {"?"}
assert len(YUL) == 60
ROWS = {"3/4": {12}, "4/4": {16}, "2/4": {16}, "정악": {20}}
KEYS = {"제목", "부제", "키", "박자", "행수", "줄수", "원본", "이어짐", "메모", "쪽"}


def check(path):
    errs, warns = [], []
    raw = open(path, "rb").read()
    if raw.startswith(b"\xef\xbb\xbf"):
        errs.append("BOM 있음")
    if b"\r" in raw:
        errs.append("CRLF 있음")
    lines = raw.decode("utf-8").split("\n")
    meta, body, hdr = {}, [], False
    for n, ln in enumerate(lines, 1):
        if not ln.strip():
            continue
        if not hdr and ln.startswith("#"):
            k, _, v = ln[1:].partition("\t")
            if k not in KEYS:
                warns.append(f"{n}행 알 수 없는 머리 정보 #{k}")
            meta[k] = v
            continue
        if not hdr:
            if ln.split("\t")[:3] != ["줄", "정간", "율명"]:
                errs.append(f"{n}행 본문 헤더 아님: {ln!r}")
            hdr = True
            continue
        body.append((n, ln.split("\t")))
    beat, rows, jul = meta.get("박자", ""), meta.get("행수", ""), meta.get("줄수", "")
    if beat not in ROWS:
        errs.append(f"#박자 '{beat}' 지원 안 함")
    try:
        rows, jul = int(rows), int(jul)
    except ValueError:
        errs.append(f"#행수/#줄수 숫자 아님 ({rows!r}, {jul!r})")
        return errs, warns, 0
    if beat in ROWS and rows not in ROWS[beat]:
        errs.append(f"#박자 {beat} 인데 #행수 {rows}")
    src = meta.get("원본", "")
    root = os.path.normpath(os.path.join(HERE, "..", "..", "원본"))
    if not src or not os.path.exists(os.path.join(root, unicodedata.normalize("NFD", src))) and not os.path.exists(os.path.join(root, src)):
        errs.append(f"#원본 파일 없음: {src!r}")
    seen = set()
    for n, f in body:
        if len(f) < 2 or len(f) > 5:
            errs.append(f"{n}행 칸 수 {len(f)}")
            continue
        try:
            j, g = int(f[0]), int(f[1])
        except ValueError:
            errs.append(f"{n}행 줄/정간 숫자 아님")
            continue
        if not (1 <= j <= jul and 1 <= g <= rows):
            errs.append(f"{n}행 범위 밖 줄{j} 정간{g}")
        if (j, g) in seen:
            errs.append(f"{n}행 중복 줄{j} 정간{g}")
        seen.add((j, g))
        pitch = f[2] if len(f) > 2 else ""
        for part in pitch.split("|") if pitch else []:
            if part == "":
                errs.append(f"{n}행 율명에 빈 '|' 조각")
            for ch in part:
                if ch not in OK:
                    errs.append(f"{n}행 팔레트 외 '{ch}' U+{ord(ch):04X} (줄{j} 정간{g})")
        if "?" in pitch and not (len(f) > 4 and f[4].strip()):
            warns.append(f"{n}행 '?' 인데 비고 없음")
    missing = [(j, g) for j in range(1, jul + 1) for g in range(1, rows + 1) if (j, g) not in seen]
    if missing:
        errs.append(f"누락 정간 {len(missing)}개: " + ", ".join(f"{j}-{g}" for j, g in missing[:10]))
    if len(body) != jul * rows:
        errs.append(f"본문 {len(body)}행 ≠ 줄수×행수 {jul * rows}")
    return errs, warns, len(body)


def main():
    targets = sys.argv[1:] or [TSV_DIR]
    files = []
    for t in targets:
        if os.path.isdir(t):
            files += sorted(os.path.join(t, f) for f in os.listdir(t) if f.endswith(".tsv"))
        else:
            files.append(t)
    bad = 0
    for p in files:
        errs, warns, n = check(p)
        name = os.path.basename(p)
        print(("OK  " if not errs else "ERR ") + f"{name} ({n}행)")
        for e in errs:
            print("   ✗ " + e)
        for w in warns:
            print("   △ " + w)
        bad += bool(errs)
    print(f"--- {len(files)}개 중 오류 {bad}개")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
