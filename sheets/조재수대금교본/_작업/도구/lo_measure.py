# LibreOffice 내장 python: 악보 시트 열 너비(pt)와 행별 최적 높이(pt) 측정. 파일은 저장하지 않음.
import json, os, subprocess, sys, time, unicodedata, uno
from com.sun.star.beans import PropertyValue
SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"; PIPE = "jgv_measure"
def prop(n, v):
    p = PropertyValue(); p.Name, p.Value = n, v; return p
proc = subprocess.Popen([SOFFICE, "--headless", "--invisible", "--norestore", "--nologo", "-env:UserInstallation=file:///tmp/jgv_lo_profile2", f"--accept=pipe,name={PIPE};urp;"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
local = uno.getComponentContext(); res = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(120):
    try:
        desktop = res.resolve(f"uno:pipe,name={PIPE};urp;StarOffice.ComponentContext").ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", None) if False else None
        ctx = res.resolve(f"uno:pipe,name={PIPE};urp;StarOffice.ComponentContext")
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx); break
    except Exception: time.sleep(0.5)
jobs = json.load(open(sys.argv[1], encoding="utf-8")); out = {}
pt = lambda v: round(v / 100 / 25.4 * 72, 2)
try:
    for j in jobs:
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(j["path"])), "_blank", 0, (prop("Hidden", True),))
        try:
            sh = doc.Sheets.getByName(next(n for n in doc.Sheets.ElementNames if unicodedata.normalize("NFC", n) == "악보"))
            cols = [pt(sh.Columns.getByIndex(c).Width) for c in range(j["ncols"])]
            hs = {}
            for r in j["rows"]:
                row = sh.Rows.getByIndex(r - 1); row.OptimalHeight = True; hs[str(r)] = pt(row.Height)
            out[j["path"]] = {"cols": cols, "rows": hs}
        finally: doc.close(True)
finally:
    try: desktop.terminate()
    except Exception: pass
    proc.wait(timeout=30)
json.dump(out, open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False)
