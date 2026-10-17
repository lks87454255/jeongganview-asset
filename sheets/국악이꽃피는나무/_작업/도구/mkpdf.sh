#!/bin/bash
# $1 = xlsx 폴더, $2 = 출력 pdf.  악보 시트만, 이름순, 곡 이름 책갈피
set -e; W=$(mktemp -d); mkdir -p $W/y $W/pdf
cp "$1"/*.xlsx $W/; rm -f $W/_목록.xlsx
python3 - "$W" <<'PY'
import zipfile,re,os,glob,sys
W=sys.argv[1]
for f in sorted(glob.glob(W+'/*.xlsx')):
    z=zipfile.ZipFile(f); wb=z.read('xl/workbook.xml').decode()
    idx=re.findall(r'<sheet [^>]*name="([^"]+)"',wb).index('악보')
    def fix(m):
        t=re.sub(r'\s*state="[^"]*"','',m.group(0)); return t if 'name="악보"' in t else t.replace('<sheet ','<sheet state="hidden" ',1)
    wb=re.sub(r'<sheet [^>]*/>',fix,wb)
    if '<workbookView' not in wb: wb=wb.replace('<sheets>',f'<bookViews><workbookView firstSheet="{idx}" activeTab="{idx}"/></bookViews><sheets>',1)
    else: wb=re.sub(r'\s*(activeTab|firstSheet)="\d+"','',wb); wb=re.sub(r'<workbookView',f'<workbookView firstSheet="{idx}" activeTab="{idx}"',wb,1)
    with zipfile.ZipFile(W+'/y/'+os.path.basename(f),'w') as zo:
        for i in z.infolist():
            d=z.read(i.filename)
            if i.filename=='xl/workbook.xml': d=wb.encode()
            elif i.filename.startswith('xl/worksheets/sheet'): d=re.sub(rb'tabSelected="(true|1)"',b'tabSelected="0"',d)
            zo.writestr(i,d,compress_type=i.compress_type)
PY
/Applications/LibreOffice.app/Contents/MacOS/soffice --headless --convert-to pdf --outdir $W/pdf $W/y/*.xlsx >/dev/null 2>&1
[ -x /tmp/pdfvenv/bin/python ] || { python3 -m venv /tmp/pdfvenv >/dev/null && /tmp/pdfvenv/bin/pip install -q "pypdf==5.1.0" >/dev/null 2>&1; }
/tmp/pdfvenv/bin/python - "$W/pdf" "$2" <<'PY'
import glob,os,unicodedata,sys
from pypdf import PdfWriter, PdfReader
fs=sorted(glob.glob(sys.argv[1]+'/*.pdf'), key=lambda p: unicodedata.normalize('NFC',os.path.basename(p))[:-4])
w=PdfWriter()
for f in fs:
    r=PdfReader(f); s=len(w.pages)
    for p in r.pages: w.add_page(p)
    w.add_outline_item(unicodedata.normalize('NFC',os.path.basename(f))[:-4], s)
w.page_mode="/UseOutlines"; w.write(open(sys.argv[2],'wb'))
print('곡',len(fs),'쪽',len(PdfReader(sys.argv[2]).pages))
PY
cp -R $W/pdf /tmp/lastpdf 2>/dev/null || true; rm -rf $W
