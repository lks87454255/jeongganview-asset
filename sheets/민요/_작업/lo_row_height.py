# -*- coding: utf-8 -*-
"""
lo_row_height.py — LibreOffice '최적 행 높이'(자동 높이) 측정 (xlsx_layout.py 가 LibreOffice 내장 python 으로 실행)

    /Applications/LibreOffice.app/Contents/Resources/python lo_row_height.py in.json out.json

in.json : [{"path": "/tmp/a.xlsx", "sheet": "악보", "rows": [3, 4, ...]}, ...]   (rows: 1부터)
out.json: {"/tmp/a.xlsx": {"3": 75.3, ...}}   (pt) — 파일은 저장하지 않고 닫는다.
"""
import json
import os
import subprocess
import sys
import time
import unicodedata

import uno
from com.sun.star.beans import PropertyValue

SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
PIPE = "jgv_row_height"


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def connect():
    profile = "file:///tmp/jgv_lo_profile"         # 사용자가 띄워 둔 LibreOffice 와 겹치지 않게 별도 프로필
    proc = subprocess.Popen([SOFFICE, "--headless", "--invisible", "--norestore", "--nologo",
                             f"-env:UserInstallation={profile}", f"--accept=pipe,name={PIPE};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(120):
        try:
            ctx = resolver.resolve(f"uno:pipe,name={PIPE};urp;StarOffice.ComponentContext")
            return proc, ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        except Exception:                                   # noqa: BLE001 — 기동 대기
            time.sleep(0.5)
    proc.kill()
    sys.exit("LibreOffice 연결 실패")


def main():
    jobs = json.load(open(sys.argv[1], encoding="utf-8"))
    proc, desktop = connect()
    out = {}
    try:
        for job in jobs:
            url = uno.systemPathToFileUrl(os.path.abspath(job["path"]))
            doc = desktop.loadComponentFromURL(url, "_blank", 0, (prop("Hidden", True),))
            try:
                sheets = doc.Sheets
                name = next(n for n in sheets.ElementNames
                            if unicodedata.normalize("NFC", n).strip() == job["sheet"])
                rows = sheets.getByName(name).Rows
                res = {}
                for r in job["rows"]:
                    row = rows.getByIndex(r - 1)
                    row.OptimalHeight = True
                    res[str(r)] = round(row.Height / 100 / 25.4 * 72, 2)   # 1/100 mm → pt
                out[job["path"]] = res
            finally:
                doc.close(True)
    finally:
        try:
            desktop.terminate()
        except Exception:                                   # noqa: BLE001
            pass
        proc.wait(timeout=30)
    json.dump(out, open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False)


if __name__ == "__main__":
    main()
