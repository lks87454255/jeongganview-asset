# build_all.ps1 실행 후 결과+생성리포트를 로그 파일로 남김 (셸 출력이 깨지거나 crash할 때 대비)
# 사용: python run_build.py [로그파일]   (기본 D:\build_out.txt)
import subprocess, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.normpath(os.path.join(HERE, ".."))                 # sheets/민요/_작업
LOG = sys.argv[1] if len(sys.argv) > 1 else r"D:\build_out.txt"
ps1 = os.path.join(WORK, "build_all.ps1")
rep = os.path.join(WORK, "\uc0dd\uc131\ub9ac\ud3ec\ud2b8.txt")   # 생성리포트.txt

r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1],
                   capture_output=True, text=True, encoding="utf-8", errors="replace")
out = ["RC=" + str(r.returncode), r.stdout or "", r.stderr or "", "--- REPORT ---"]
if os.path.exists(rep):
    out.append(open(rep, encoding="utf-8").read())
open(LOG, "w", encoding="utf-8").write("\n".join(out))
