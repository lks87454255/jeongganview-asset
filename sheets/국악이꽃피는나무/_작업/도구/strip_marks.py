#!/usr/bin/env python3
"""검수 xlsx '악보' 시트 율명 칸에서 지정한 기호를 지움 (칸 안 줄이 비면 그 줄도 지움).
    python3 strip_marks.py <검수 폴더> <지울 글자들> [--check]
"""
import os, re, sys, zipfile, unicodedata
from xml.sax.saxutils import escape
import html
D, CH = sys.argv[1], set(sys.argv[2]); CHECK = "--check" in sys.argv; OFF = "--offpalette" in sys.argv
import importlib.util
_spec = importlib.util.spec_from_file_location("gsi", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "..", "generate_sheets_index.py"))
_g = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(_g)
KNOWN = _g.KNOWN | {"–", "\n", " "}
U = {'&quot;': '"', '&apos;': "'"}


def sheet_path(z):
    wb = z.read('xl/workbook.xml').decode(); rels = z.read('xl/_rels/workbook.xml.rels').decode()
    rid = re.search(r'<sheet [^>]*name="악보"[^>]*r:id="([^"]+)"', wb).group(1)
    for rel in re.findall(r'<Relationship [^>]*>', rels):
        if f'Id="{rid}"' in rel:
            t = re.search(r'Target="([^"]+)"', rel).group(1).lstrip('/'); return t if t.startswith('xl/') else 'xl/' + t


def clean(t):
    keep = []
    for l in t.split("\n"):
        o = "".join(c for c in l if c not in CH and not (OFF and c not in KNOWN))
        if o.strip() or not l.strip():          # 기호만 있던 줄은 지우고, 원래 빈 줄은 둠
            keep.append(o)
    r = "\n".join(keep)
    return r if r.strip() else ""


tot_files = tot_cells = 0; removed = {}
for f in sorted(os.listdir(D)):
    if not f.endswith('.xlsx') or f.startswith(('_', '~$', '.~lock')): continue
    if os.path.exists(os.path.join(D, f".~lock.{f}#")) or os.path.exists(os.path.join(D, "~$" + f)):
        print("  열려 있어 건너뜀:", f); continue
    p = os.path.join(D, f); z = zipfile.ZipFile(p); sp = sheet_path(z)
    ss = [html.unescape(''.join(re.findall(r'<t[^>]*>([^<]*)</t>', si))) for si in
          re.findall(r'<si>.*?</si>', z.read('xl/sharedStrings.xml').decode(), re.S)] if 'xl/sharedStrings.xml' in z.namelist() else []
    x = z.read(sp).decode(); n = [0]
    def colnum(t):
        k = 0
        for ch in t: k = k * 26 + ord(ch) - 64
        return k
    yulcols = set(); titlerows = set()
    for a, r1, b, r2 in re.findall(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', x):
        titlerows.add(int(r1))
        c1, c2 = colnum(a), colnum(b)
        yulcols |= {c for c in range(c1, c2 + 1) if (c2 - c) % 2 == 1}

    def fix(m):
        attrs, inner = m.group(2), m.group(3)
        if colnum(re.match(r"[A-Z]+", m.group(1)).group(0)) not in yulcols: return m.group(0)
        if int(re.search(r"\d+", m.group(1)).group(0)) in titlerows: return m.group(0)
        v = re.search(r'<v>(\d+)</v>', inner); t = re.findall(r'<t[^>]*>([^<]*)</t>', inner)
        if 't="s"' in attrs and v: old = ss[int(v.group(1))]
        elif 't="inlineStr"' in attrs and t: old = html.unescape(''.join(t))
        else: return m.group(0)
        new = clean(old)
        if new == old: return m.group(0)
        for c in old:
            if c in CH or (OFF and c not in KNOWN): removed[c] = removed.get(c, 0) + 1
        n[0] += 1
        a = re.sub(r'\s*t="[^"]*"', '', attrs)
        if not new: return f'<c r="{m.group(1)}"{a}/>'
        return f'<c r="{m.group(1)}"{a} t="inlineStr"><is><t xml:space="preserve">{escape(new)}</t></is></c>'
    x2 = re.sub(r'<c r="([A-Z]+\d+)"([^>]*?)(?<!/)>(.*?)</c>', fix, x, flags=re.S)
    if n[0]:
        tot_files += 1; tot_cells += n[0]
        if not CHECK:
            tmp = p + '.tmp'
            with zipfile.ZipFile(tmp, 'w') as zo:
                for i in z.infolist():
                    zo.writestr(i, x2.encode() if i.filename == sp else z.read(i.filename), compress_type=i.compress_type)
            z.close(); os.replace(tmp, p); continue
    z.close()
print(f"{'(확인만) ' if CHECK else ''}파일 {tot_files} · 칸 {tot_cells} · 지운 글자 {dict(sorted(removed.items(), key=lambda kv: -kv[1]))}")
