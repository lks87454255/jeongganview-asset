#!/usr/bin/env python3
"""검수 xlsx '악보' 시트: 율명 열 너비를 칸 안 가장 긴 줄에 맞춤 (시트 안 율명 열은 모두 같은 너비).
    python3 fitcols.py <검수 폴더> [--check]
한 칸 '黃\\n黃太\\n△' 이면 가장 긴 줄 '黃太'(2글자) 기준. 한자·한글·△―는 1칸, 영문·작은 꾸밈 기호는 0.6칸.
너비 = 글자 수 × 율명 글자 크기(24pt) + 여유 12pt, 최소 2글자. 행 높이는 이어서 layout_a4.py 로 맞춘다.
"""
import os, re, sys, zipfile, html, unicodedata
D = sys.argv[1]; CHECK = "--check" in sys.argv
FONT_PT, PAD_PT, MIN_CH = 24.0, 12.0, 2.0


def sheet_path(z):
    wb = z.read('xl/workbook.xml').decode(); rels = z.read('xl/_rels/workbook.xml.rels').decode()
    rid = re.search(r'<sheet [^>]*name="악보"[^>]*r:id="([^"]+)"', wb).group(1)
    for rel in re.findall(r'<Relationship [^>]*>', rels):
        if f'Id="{rid}"' in rel:
            t = re.search(r'Target="([^"]+)"', rel).group(1).lstrip('/'); return t if t.startswith('xl/') else 'xl/' + t


def colnum(t):
    k = 0
    for ch in t: k = k * 26 + ord(ch) - 64
    return k


def ch_w(c):
    if unicodedata.east_asian_width(c) in "WF" or c in "△―–○": return 1.0
    return 0.6


def units(pt):                       # pt → 엑셀 열 너비 단위 (px = 7*w + 5, pt = px*0.75)
    return round((pt / 0.75 - 5) / 7, 2)


changed = 0; report = []
for f in sorted(os.listdir(D)):
    if not f.endswith('.xlsx') or f.startswith(('_', '~$', '.~lock')): continue
    if os.path.exists(os.path.join(D, f".~lock.{f}#")) or os.path.exists(os.path.join(D, "~$" + f)):
        print("  열려 있어 건너뜀:", f); continue
    p = os.path.join(D, f); z = zipfile.ZipFile(p); sp = sheet_path(z); x = z.read(sp).decode()
    ss = [html.unescape(''.join(re.findall(r'<t[^>]*>([^<]*)</t>', si))) for si in
          re.findall(r'<si>.*?</si>', z.read('xl/sharedStrings.xml').decode(), re.S)] if 'xl/sharedStrings.xml' in z.namelist() else []
    yul, titlerows = set(), set()
    for a, r1, b, r2 in re.findall(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', x):
        titlerows.add(int(r1)); c1, c2 = colnum(a), colnum(b)
        yul |= {c for c in range(c1, c2 + 1) if (c2 - c) % 2 == 1}
    longest = MIN_CH
    for ref, attrs, inner in re.findall(r'<c r="([A-Z]+\d+)"([^>]*?)(?<!/)>(.*?)</c>', x, re.S):
        col = colnum(re.match(r"[A-Z]+", ref).group(0)); row = int(re.search(r"\d+", ref).group(0))
        if col not in yul or row in titlerows: continue
        v = re.search(r'<v>(\d+)</v>', inner); t = re.findall(r'<t[^>]*>([^<]*)</t>', inner)
        txt = ss[int(v.group(1))] if ('t="s"' in attrs and v) else html.unescape(''.join(t))
        for line in txt.split("\n"):
            longest = max(longest, sum(ch_w(c) for c in line.strip()))
    w = units(longest * FONT_PT + PAD_PT)
    cols = {}
    for el in re.findall(r'<col [^>]*/>', x):
        a = int(re.search(r'min="(\d+)"', el).group(1)); b = int(re.search(r'max="(\d+)"', el).group(1))
        for c in range(a, b + 1): cols[c] = el
    new_cols = []
    for c in sorted(set(cols) | yul):
        el = cols.get(c, f'<col min="{c}" max="{c}" width="8.5" customWidth="1"/>')
        el = re.sub(r'min="\d+"', f'min="{c}"', el); el = re.sub(r'max="\d+"', f'max="{c}"', el)
        if c in yul:
            el = re.sub(r'width="[^"]*"', f'width="{w}"', el)
            if 'customWidth' not in el: el = el.replace('/>', ' customWidth="1"/>')
        new_cols.append(el)
    x2 = re.sub(r'<cols>.*?</cols>', '<cols>' + ''.join(new_cols) + '</cols>', x, 1, flags=re.S)
    if x2 != x:
        changed += 1; report.append((f, longest, w))
        if not CHECK:
            tmp = p + '.tmp'
            with zipfile.ZipFile(tmp, 'w') as zo:
                for i in z.infolist():
                    zo.writestr(i, x2.encode() if i.filename == sp else z.read(i.filename), compress_type=i.compress_type)
            z.close(); os.replace(tmp, p); continue
    z.close()
import collections
print(f"{'(확인만) ' if CHECK else ''}변경 {changed}개 · 가장 긴 줄 글자 수 분포 {dict(sorted(collections.Counter(r[1] for r in report).items()))}")
