# 원본 이미지 → 대비강화 full + 4분할(RT/RB/LT/LB) crop 생성 (판독용, 저장소 밖에 출력)
# 사용: python prep_crops.py [출력폴더]   (기본 D:\pan)
# 출력: f{폴더번호}_{순번}_full/RT/RB/LT/LB.png, map.json, map.txt(키→원본파일→tsv명)
# 주의: 한글 파일명은 하드코딩하지 말고 os.listdir 결과만 사용 (코드포인트 오타로 실패한 적 있음)
import os, sys, json, unicodedata, traceback

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, "..", "..", "\uc6d0\ubcf8"))   # sheets/민요/원본
OUT = sys.argv[1] if len(sys.argv) > 1 else r"D:\pan"
LOG = os.path.join(OUT, "prep_log.txt")


def say(s):
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(s + "\n")


os.makedirs(OUT, exist_ok=True)
open(LOG, "w", encoding="utf-8").write("START\n")
try:
    from PIL import Image, ImageEnhance, ImageFilter
    mapping = []
    for folder in sorted(os.listdir(ROOT)):
        fdir = os.path.join(ROOT, folder)
        if not os.path.isdir(fdir):
            continue
        num = folder.rsplit("-", 1)[-1]                      # 민요채보-1 -> "1"
        files = sorted(f for f in os.listdir(fdir) if f.lower().endswith((".png", ".jpg", ".jpeg")))
        for i, fn in enumerate(files, 1):
            key = f"f{num}_{i:02d}"
            with open(os.path.join(fdir, fn), "rb") as fh:
                im = Image.open(fh)
                im.load()
            im = im.convert("L")
            W, H = im.size
            im = ImageEnhance.Contrast(im).enhance(1.6).filter(ImageFilter.SHARPEN)
            im.save(os.path.join(OUT, key + "_full.png"))
            # 4분할(겹침 포함). 각 장은 2000px 미만이어야 read 도구가 받음 → 반드시 1장씩 읽을 것
            ox, oy = int(W * 0.04), int(H * 0.03)
            mx, my = W // 2, H // 2
            boxes = {"RT": (mx - ox, 0, W, my + oy), "RB": (mx - ox, my - oy, W, H),
                     "LT": (0, 0, mx + ox, my + oy), "LB": (0, my - oy, mx + ox, H)}
            for name, box in boxes.items():
                im.crop(box).save(os.path.join(OUT, f"{key}_{name}.png"))
            stem = unicodedata.normalize("NFC", os.path.splitext(fn)[0])
            mapping.append({"key": key, "folder": num, "file": unicodedata.normalize("NFC", fn),
                            "tsv": f"{num}_{stem}.tsv", "orig": [W, H]})
            say(f"done {key} {W}x{H}")
    json.dump(mapping, open(os.path.join(OUT, "map.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    with open(os.path.join(OUT, "map.txt"), "w", encoding="utf-8") as f:
        for m in mapping:
            f.write(f"{m['key']}\t{m['file']}\t{m['tsv']}\t{m['orig'][0]}x{m['orig'][1]}\n")
    say(f"OK total={len(mapping)}")
except Exception as e:
    say("ERR " + repr(e))
    say(traceback.format_exc())
