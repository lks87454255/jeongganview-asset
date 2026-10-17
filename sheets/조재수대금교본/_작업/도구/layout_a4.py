#!/usr/bin/env python3
"""악보 시트: 인쇄 시 제목 높이 동일(T pt) · 글자가 선에 닿지 않게(최적 높이+여유) · 페이지 높이 채움. 배율은 시트마다 계산해 고정."""
import zipfile, re, os, sys, glob, json, math, subprocess
D = sys.argv[1]; T = 40.0; PAD = 8.0; SMALL = 6.0
LOPY = "/Applications/LibreOffice.app/Contents/Resources/python"
def sheet_path(z):
    wb = z.read('xl/workbook.xml').decode(); rels = z.read('xl/_rels/workbook.xml.rels').decode()
    rid = re.search(r'<sheet [^>]*name="악보"[^>]*r:id="([^"]+)"', wb).group(1)
    for rel in re.findall(r'<Relationship [^>]*>', rels):
        if f'Id="{rid}"' in rel:
            t = re.search(r'Target="([^"]+)"', rel).group(1).lstrip('/'); return t if t.startswith('xl/') else 'xl/' + t
def colnum(s):
    n = 0
    for ch in s: n = n * 26 + ord(ch) - 64
    return n
info = {}; jobs = []
for f in sorted(glob.glob(os.path.join(D, '*.xlsx'))):
    if os.path.basename(f).startswith('_'): continue
    if os.path.exists(os.path.join(D, '.~lock.' + os.path.basename(f) + '#')): print('  열려 있어 건너뜀:', os.path.basename(f)); continue
    z = zipfile.ZipFile(f); sp = sheet_path(z); x = z.read(sp).decode(); z.close()
    merges = [re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', r).groups() for r in re.findall(r'<mergeCell ref="([^"]+)"', x)]
    titles = sorted(int(a[1]) for a in merges if a[1] == a[3] and colnum(a[2]) > colnum(a[0]))
    c2 = max(colnum(a[2]) for a in merges if int(a[1]) in titles)
    last = max(int(r) for r, b in re.findall(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S) if re.search(r'<v>|<is>|s="', b))
    brks = [int(b) for b in re.findall(r'<brk id="(\d+)"', x) if int(b) < last]
    pages = []
    for i, t in enumerate(titles):
        end = (titles[i + 1] - 1) if i + 1 < len(titles) else last
        while end > t and end + 1 in titles[i + 1:i + 2] and False: pass
        body = list(range(t + 1, end + 1)); pre = list(range((brks[i - 1] + 1) if i else 1, t))
        # 다음 제목 직전 빈 행은 다음 페이지 머리(pre)로 감
        if i + 1 < len(titles):
            nb = brks[i] if i < len(brks) else titles[i + 1] - 1
            body = [r for r in body if r <= nb]
            sep = [r for r in range(nb + 1, titles[i + 1])]
        pages.append({"title": t, "body": body, "pre": pre})
    m = re.search(r'<pageMargins[^>]*>', x).group(0); g = lambda k: float(re.search(r'%s="([^"]+)"' % k, m).group(1))
    info[f] = dict(sp=sp, pages=pages, c2=c2, Wp=595.28 - (g('left') + g('right')) * 72, Hp=(841.89 - (g('top') + g('bottom')) * 72) * 0.985)
    jobs.append({"path": f, "ncols": c2, "rows": [r for p in pages for r in p["body"]]})
json.dump(jobs, open('/tmp/lo_in.json', 'w', encoding='utf-8'), ensure_ascii=False)
subprocess.run([LOPY, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'lo_measure.py'), '/tmp/lo_in.json', '/tmp/lo_out.json'], check=True)
meas = json.load(open('/tmp/lo_out.json', encoding='utf-8'))
for f, I in info.items():
    if f not in meas:
        print('  측정 실패, 건너뜀:', os.path.basename(f)); continue
    M = meas[f]; W = sum(M["cols"]); vals = sorted(M["rows"].values()); med = vals[len(vals) // 2]
    out = [v for v in vals if v > med * 2]        # 메모처럼 유난히 긴 칸 1~2개는 높이 계산에서 뺌 (그 칸은 잘려 보일 수 있음)
    if out and os.environ.get('V'): print('  제외한 긴 칸 높이', out)
    req = max(v for v in vals if v <= med * 2) + PAD     # 시트 안 정간 행은 모두 같은 높이
    s = min(1.0, I["Wp"] / W, *[(I["Hp"] - T) / (len(p["body"]) * req + SMALL * len(p["pre"])) for p in I["pages"]])
    pct = max(10, math.floor(s * 100)); s = pct / 100
    h = min(((I["Hp"] - T) / s - SMALL * len(p["pre"])) / len(p["body"]) for p in I["pages"])
    hts = {}
    for p in I["pages"]:
        for r in p["pre"]: hts[r] = SMALL
        hts[p["title"]] = round(T / s, 2)
        for r in p["body"]: hts[r] = round(h, 2)
    z = zipfile.ZipFile(f); x = z.read(I["sp"]).decode()
    def fix(mm):
        r = int(mm.group(1)); a = mm.group(2)
        if r not in hts: return mm.group(0)
        a = re.sub(r'\s+customHeight="[^"]*"', '', a); a = re.sub(r'\s+ht="[^"]*"', '', a)
        return f'<row r="{r}"{a} ht="{hts[r]}" customHeight="true"'
    x = re.sub(r'<row r="(\d+)"([^>]*?)(?=/?>)', fix, x)
    have = {int(r) for r in re.findall(r'<row r="(\d+)"', x)}
    for r in sorted(hts):
        if r not in have:   # 행 요소가 없는 여백 행은 기본 높이로 인쇄되므로 직접 넣음
            nxt = min([q for q in have if q > r] or [None], key=lambda q: q or 0)
            el = f'<row r="{r}" ht="{hts[r]}" customHeight="true"/>'
            x = x.replace(f'<row r="{nxt}"', el + f'<row r="{nxt}"', 1) if nxt else x.replace('</sheetData>', el + '</sheetData>', 1)
            have.add(r)
    # 내용 아래의 빈 행(셀 없음, 높이만 있음)은 인쇄 범위를 늘려 빈 쪽을 만드므로 지움
    lastp = max(r for p in I['pages'] for r in p['body'])
    x = re.sub(r'<row r="(\d+)"[^>]*/>', lambda mm: '' if int(mm.group(1)) > lastp else mm.group(0), x)
    x = re.sub(r'<pageSetUpPr([^>]*?)\s*fitToPage="[^"]*"', r'<pageSetUpPr\1 fitToPage="0"', x)
    x = re.sub(r'<pageSetup[^>]*/>', f'<pageSetup paperSize="9" orientation="portrait" scale="{pct}"/>', x, 1)
    tmp = f + '.tmp'
    with zipfile.ZipFile(tmp, 'w') as zo:
        for i in z.infolist():
            zo.writestr(i, x.encode() if i.filename == I["sp"] else z.read(i.filename), compress_type=i.compress_type)
    z.close(); os.replace(tmp, f)
    if os.environ.get('V'): print(os.path.basename(f), f'배율 {pct}%', '제목', hts[I["pages"][0]["title"]], '행', h, '필요', round(req, 1), '폭', round(W))
