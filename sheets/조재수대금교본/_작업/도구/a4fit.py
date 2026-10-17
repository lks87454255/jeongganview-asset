#!/usr/bin/env python3
"""검수 xlsx '악보' 시트 페이지 설정: A4 세로, 페이지(행 나누기)마다 한 장에 맞춤(너비 1 · 높이 = 페이지 수)"""
import zipfile, re, os, sys, glob
D = sys.argv[1]
def sheet_path(z, name):
    wb = z.read('xl/workbook.xml').decode(); rels = z.read('xl/_rels/workbook.xml.rels').decode()
    rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % name, wb).group(1)
    for rel in re.findall(r'<Relationship [^>]*>', rels):
        if f'Id="{rid}"' in rel:
            t = re.search(r'Target="([^"]+)"', rel).group(1)
            return t.lstrip('/') if t.startswith('/xl/') or t.startswith('xl/') else 'xl/' + t
n = 0
for f in sorted(glob.glob(os.path.join(D, '*.xlsx'))):
    if os.path.basename(f).startswith('_'): continue
    z = zipfile.ZipFile(f); sp = sheet_path(z, '악보'); x = z.read(sp).decode()
    # 내용이 있는 마지막 행 아래의 페이지 나누기는 빈 페이지만 만들므로 지움
    last = max([int(r) for r, body in re.findall(r'<row r="(\d+)"[^>]*>(.*?)</row>', x, re.S) if re.search(r'<v>|<is>', body)] or [1])
    brks = [b for b in re.findall(r'<brk [^>]*/>', x) if int(re.search(r'id="(\d+)"', b).group(1)) < last]
    if '<rowBreaks' in x:
        x = re.sub(r'<rowBreaks[^>]*>.*?</rowBreaks>|<rowBreaks[^>]*/>',
                   (f'<rowBreaks count="{len(brks)}" manualBreakCount="{len(brks)}">' + ''.join(brks) + '</rowBreaks>') if brks else '', x, 1, flags=re.S)
    pages = len(brks) + 1
    # sheetPr / pageSetUpPr fitToPage
    if re.search(r'<pageSetUpPr[^>]*>', x):
        x = re.sub(r'<pageSetUpPr([^>]*?)\s*fitToPage="[^"]*"', r'<pageSetUpPr\1', x)
        x = re.sub(r'<pageSetUpPr', '<pageSetUpPr fitToPage="1"', x, 1)
    elif re.search(r'<sheetPr[^>]*/>', x):
        x = re.sub(r'<sheetPr([^>]*)/>', r'<sheetPr\1><pageSetUpPr fitToPage="1"/></sheetPr>', x, 1)
    elif '<sheetPr' in x:
        x = re.sub(r'(<sheetPr[^>]*>)', r'\1<pageSetUpPr fitToPage="1"/>', x, 1)   # 순서: tabColor, outlinePr 뒤가 원칙이지만 앞이어도 Excel/LO 모두 읽음
        # 정확한 순서 맞추기: tabColor·outlinePr 다음으로 이동
        m = re.search(r'(<sheetPr[^>]*>)(<pageSetUpPr fitToPage="1"/>)((?:<tabColor[^>]*/>)?(?:<outlinePr[^>]*/>)?)', x)
        if m: x = x[:m.start()] + m.group(1) + m.group(3) + m.group(2) + x[m.end():]
    else:
        x = re.sub(r'(<worksheet[^>]*>)', r'\1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>', x, 1)
    ps = f'<pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="{pages}"/>'
    if '<pageSetup' in x: x = re.sub(r'<pageSetup[^>]*/>', ps, x, 1)
    else: x = re.sub(r'(<pageMargins[^>]*/>)', r'\1' + ps, x, 1)
    tmp = f + '.tmp'
    with zipfile.ZipFile(tmp, 'w') as zo:
        for i in z.infolist():
            d = z.read(i.filename)
            if i.filename == sp: d = x.encode()
            zo.writestr(i, d, compress_type=i.compress_type)
    z.close(); os.replace(tmp, f); n += 1
    if pages > 1: print(os.path.basename(f), pages, '페이지')
print('적용', n)
