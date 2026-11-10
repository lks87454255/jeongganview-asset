#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_index.py
=================
로컬 audio/ , book/ 디렉토리를 스캔하여
  - audio-index.json  (음원 카테고리/트랙 목록)
  - book-index.json   (정간보 카테고리/파일 목록)
을 생성합니다.

정간보 PDF ↔ 그리드 JSON 일치 검사도 함께 실행합니다(오류가 있으면 종료 코드 1).
  python3 generate_index.py          # 목록 생성 + 검사
  python3 generate_index.py --check  # 검사만 (파일 변경 없음)

생성 후 git push 하면 앱이 GitHub API 없이
raw.githubusercontent.com 에서 단 1회 요청으로 전체 목록을 가져옵니다.
(GitHub API Rate Limit: 비인증 60회/시간 → 완전 우회)

앱의 URL 인코딩 방식:
  Java: URLEncoder.encode(name, "UTF-8").replace("+", "%20")
  Python: urllib.parse.quote(name, safe="")   ← 완전 동일
"""


import os
import sys
import json
import unicodedata
from datetime import datetime, timezone
from urllib.parse import quote

# ── 레포지토리 설정 ────────────────────────────────────────────
OWNER  = "lks87454255"
REPO   = "jeongganview-asset"
BRANCH = "main"
BASE   = f"https://raw.githubusercontent.com/{OWNER}/{REPO}/{BRANCH}"

# ── 지원 확장자 ─────────────────────────────────────────────────
AUDIO_EXTS = {".mp3", ".wav", ".m4a", ".flac", ".ogg", ".aac", ".opus"}
BOOK_EXTS  = {".pdf"}

# 무시할 파일/디렉토리 패턴
IGNORE = {".DS_Store", ".gitkeep", ".gitignore", "Thumbs.db"}

# 레포에는 두지만 book-index.json 에 넣지 않는 PDF (레포 기준 상대 경로, 그리드 JSON 도 만들지 않음)
# 레포에는 두지만 book-index.json 에 넣지 않는 폴더 (정간보 시트 작업 원본 — 앱 배경 목록·검색에 안 보임)
BOOK_EXCLUDE_DIRS = {"대금율보", "조재수 대금교본", "국악이 꽃피는 나무"}

# 관리자 기기에만 보이는 분류(테스트 중인 디지털 정간보) — book-index.json 의 "adminCategories" 로 나간다
BOOK_ADMIN_ONLY_DIRS: set = set()

BOOK_EXCLUDE: set = {
    "book/대금정악보/가곡.pdf",   # 곡별 파일(book/가곡/가곡-<곡명>)로 보여 준다 — 파일·그리드·싱크 맵은 레포에 두고 카드만 숨김(사용자 결정, 2026-11-08)
}


def nfc(s: str) -> str:
    """한글 조합형 → 완성형(NFC) 정규화 (앱과 동일)"""
    return unicodedata.normalize("NFC", s)


def raw_url(*parts: str) -> str:
    """
    parts 각각을 URL-인코딩하여 raw.githubusercontent.com URL 생성.
    Java의 URLEncoder.encode(s, UTF8).replace("+", "%20") 와 동일.
    """
    encoded = "/".join(quote(nfc(p), safe="") for p in parts)
    return f"{BASE}/{encoded}"


def is_ignored(name: str) -> bool:
    return name in IGNORE or name.startswith(".")


# ── 곡 제목 목록: <PDF 이름>.grid.json 최상위 "songs" ─────────────
# "songs": [{"title": "진양", "page": 1, "grid": 1}, …]
#   page = 그 PDF 안 쪽 번호(1부터), grid = 그 쪽에서 곡이 시작하는 블록 gridId(한 쪽에 곡이 둘 이상이면 필수, 없으면 그 쪽 첫 블록),
#   col·row(선택) = 블록 중간에서 시작할 때. 그리드가 아직 없는 PDF 도 곡 제목만 든 .grid.json(pageData {})을 둔다.
#   (예전 <PDF 이름>.songs.json 은 2026-11-06 폐기 — 남아 있으면 오류)
# book-index.json 에는 앱이 읽는 title·page 만 넣는다(앱 정간보 설정 검색용).
GRID_SUFFIX = ".grid.json"      # 그리드 JSON(<PDF 이름>.grid.json) — 앱 GridManager 가 읽는다
OLD_SONGS_SUFFIX = ".songs.json"


def read_songs(grid_path: str) -> list:
    if not os.path.isfile(grid_path):
        return []
    with open(grid_path, encoding="utf-8") as f:
        data = json.load(f)
    out = []
    for s in data.get("songs") or []:
        title, page = nfc(str(s.get("title", ""))).strip(), s.get("page")
        if title and isinstance(page, int) and page >= 1:
            e = {"title": title, "page": page}
            for k in ("grid", "col", "row"):
                if k in s:
                    e[k] = s[k]
            out.append(e)
    return out


# ── audio-index.json 생성 ──────────────────────────────────────
def scan_audio(audio_dir: str) -> dict:
    categories = []

    cat_names = sorted(
        n for n in os.listdir(audio_dir)
        if os.path.isdir(os.path.join(audio_dir, n)) and not is_ignored(n)
    )

    for cat_name in cat_names:
        cat_path = os.path.join(audio_dir, cat_name)
        tracks = []

        for file_name in sorted(os.listdir(cat_path)):
            if is_ignored(file_name):
                continue
            ext = os.path.splitext(file_name)[1].lower()
            if ext not in AUDIO_EXTS:
                continue

            display_name = os.path.splitext(file_name)[0]  # 확장자 제거
            asset_path   = raw_url("audio", cat_name, file_name)

            tracks.append({
                "displayName": display_name,   # parseAudioCategoriesFromJson 필드
                "fileName":    file_name,
                "assetPath":   asset_path,
            })

        if tracks:
            categories.append({"name": cat_name, "tracks": tracks})
            print(f"  🎵 [{cat_name}]  {len(tracks)}개 트랙")
        else:
            print(f"  ⚠️  [{cat_name}]  음원 파일 없음 — 건너뜀")

    total_tracks = sum(len(c["tracks"]) for c in categories)
    return {
        "version":    1,
        "generated":  datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "owner":      OWNER,
        "repo":       REPO,
        "branch":     BRANCH,
        "categories": categories,
        "_total":     {"categories": len(categories), "tracks": total_tracks},
    }


# ── book-index.json 생성 ──────────────────────────────────────
def pdf_page_count(path: str):
    """PDF 총 쪽 수(앱 정간보 연주 카드에 보임). 못 읽으면 None."""
    try:
        import pymupdf
        with pymupdf.open(path) as doc:
            return doc.page_count
    except Exception:
        return None


def scan_book(book_dir: str) -> dict:
    categories = []
    admin_categories = []

    cat_names = sorted(
        n for n in os.listdir(book_dir)
        if os.path.isdir(os.path.join(book_dir, n)) and not is_ignored(n) and nfc(n) not in BOOK_EXCLUDE_DIRS
    )

    for cat_name in cat_names:
        cat_path = os.path.join(book_dir, cat_name)
        files = []

        for file_name in sorted(os.listdir(cat_path)):
            if is_ignored(file_name):
                continue
            ext = os.path.splitext(file_name)[1].lower()
            if ext not in BOOK_EXTS:
                continue
            if nfc(f"book/{cat_name}/{file_name}") in BOOK_EXCLUDE:
                print(f"  ⏭️  제외(BOOK_EXCLUDE): book/{cat_name}/{file_name}")
                continue

            display_name = os.path.splitext(file_name)[0]  # 확장자 제거
            rel_path     = f"book/{cat_name}/{file_name}"   # 상대 경로
            url_str      = raw_url("book", cat_name, file_name)
            size_bytes   = os.path.getsize(os.path.join(cat_path, file_name))

            files.append({
                "name":      display_name,   # parseBookCategoriesFromJson 필드
                "fileName":  file_name,
                "path":      rel_path,
                "url":       url_str,
                "sizeBytes": size_bytes,
            })
            pages = pdf_page_count(os.path.join(cat_path, file_name))
            if pages:
                files[-1]["pageCount"] = pages
            # 한 PDF 에 곡이 여럿이면 PDF 옆 <이름>.grid.json 의 "songs" 곡 제목·쪽을 넣는다(앱 정간보 설정 검색용)
            songs = read_songs(os.path.join(cat_path, display_name + GRID_SUFFIX))
            if songs:
                files[-1]["songs"] = [{"title": s["title"], "page": s["page"]} for s in songs]
            # PDF 옆 <이름>.sync.json(국립국악원 음원 싱크 맵)이 있으면 앱 카드에 음원 아이콘을 보인다
            sync_path = os.path.join(cat_path, display_name + ".sync.json")
            if os.path.isfile(sync_path):
                files[-1]["hasSync"] = True
                # 음원 전체 길이(초) — 트랙 durationMs 합. 큰 화면 카드에서 음원 아이콘 옆에 보인다.
                try:
                    with open(sync_path, encoding="utf-8") as f:
                        ms = [t.get("durationMs") for t in json.load(f).get("tracks", [])]
                    if ms and all(isinstance(x, int) and x > 0 for x in ms):
                        files[-1]["syncDurationSec"] = round(sum(ms) / 1000)
                except Exception:
                    pass

        if files:
            entry = {"name": cat_name, "files": files}
            # 관리자 전용 분류는 "adminCategories" 에 따로 담는다.
            #   옛 앱은 "categories" 만 읽으므로 이 분류를 보지 못한다(새 앱은 관리자 기기에만 보여 준다).
            (admin_categories if nfc(cat_name) in BOOK_ADMIN_ONLY_DIRS else categories).append(entry)
            total_mb = sum(f["sizeBytes"] for f in files) / 1_048_576
            tag = " (관리자 전용)" if nfc(cat_name) in BOOK_ADMIN_ONLY_DIRS else ""
            print(f"  📚 [{cat_name}]{tag}  {len(files)}개 파일  ({total_mb:.1f} MB)")
        else:
            print(f"  ⚠️  [{cat_name}]  PDF 파일 없음 — 건너뜀")

    total_files = sum(len(c["files"]) for c in categories)
    return {
        "version":    1,
        "generated":  datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "owner":      OWNER,
        "repo":       REPO,
        "branch":     BRANCH,
        "categories": categories,
        "adminCategories": admin_categories,
        "_total":     {"categories": len(categories), "files": total_files},
    }


# ── 정간보 PDF ↔ 그리드 JSON 일치 검사 ─────────────────────────
# 앱은 PDF 주소의 ".pdf" 를 ".grid.json" 으로 바꿔 그리드 JSON 을 찾는다(없으면 예전 이름 ".json").
# 이름이 한 글자라도 다르면(예: 한천수.pdf / 헌천수.grid.json) JSON 이 404 → 그리드가 그려지지 않는다.
#   그리드 JSON 이름은 .sync.json 과 같은 패턴인 <PDF 이름>.grid.json 으로 통일(2026-11-04).
#   그 밖의 .json(예전 이름 <PDF 이름>.json 등)은 오류로 알린다.
#   서비스 파일은 세 가지뿐이다(2026-11-06 결정): <곡>.pdf, <곡>.grid.json(곡 제목 "songs" 포함), <곡>.sync.json. 그 밖의 파일도 오류.
#   오류(ERROR): JSON 짝 PDF 없음 · JSON 파싱/구조 오류 · 파일명 NFC 아님 · 폐기된 .songs.json · 곡 시작 위치 오류
#   경고(WARN) : JSON 안 fileName 이 PDF 이름과 다름
#   정보(INFO) : JSON 없는 PDF·쪽 데이터 없는 그리드 JSON (기본 그리드로 동작 — 허용)
def check_book_pairs(book_dir: str):
    import difflib
    errors, warns, infos = [], [], []
    pdfs, jsons = {}, {}
    for root, _, files in os.walk(book_dir):
        for fn in files:
            if is_ignored(fn):
                continue
            rel = os.path.relpath(os.path.join(root, fn), os.path.dirname(book_dir))
            if nfc(rel).replace(os.sep, "/").split("/")[1] in BOOK_EXCLUDE_DIRS:
                continue                       # 목록에 넣지 않는 폴더는 검사도 안 함
            if fn != nfc(fn):
                errors.append(f"파일명이 NFC(완성형)가 아님 — 앱 URL 과 달라 404: {rel}")
            stem, ext = os.path.splitext(nfc(rel))
            rel_s = nfc(rel).replace(os.sep, "/")
            # BOOK_EXCLUDE PDF 와 그 짝(.grid.json·.sync.json)은 목록에 넣지 않으므로 검사도 안 함
            base = rel_s[:-len(GRID_SUFFIX)] if rel_s.endswith(GRID_SUFFIX) else rel_s[:-len(".sync.json")] if rel_s.endswith(".sync.json") else os.path.splitext(rel_s)[0]
            if rel_s in BOOK_EXCLUDE or base + ".pdf" in BOOK_EXCLUDE or rel_s.split("/")[1] in BOOK_EXCLUDE_DIRS:
                continue
            if nfc(fn).endswith(OLD_SONGS_SUFFIX):
                errors.append(f"폐기된 파일(곡 제목은 {GRID_SUFFIX} 의 \"songs\" 에 둔다): {nfc(rel)}")
                continue
            if nfc(fn).endswith(".sync.json"):      # 음원 싱크 맵(앱 AudioSyncRepository 가 읽는다) — 그리드 JSON 아님
                continue
            if ext.lower() == ".pdf":
                pdfs[stem] = os.path.join(root, fn)
            elif nfc(fn).endswith(GRID_SUFFIX):
                jsons[nfc(rel)[:-len(GRID_SUFFIX)]] = os.path.join(root, fn)
            elif ext.lower() == ".json":
                errors.append(f"그리드 JSON 이름은 <PDF 이름>{GRID_SUFFIX} 이어야 함: {nfc(rel)}")
            else:
                errors.append(f"서비스하지 않는 파일(.pdf · .grid.json · .sync.json 만 둔다): {nfc(rel)}")

    for stem, path in sorted(jsons.items()):
        if stem not in pdfs:
            same_dir = [p for p in pdfs if os.path.dirname(p) == os.path.dirname(stem)]
            hint = difflib.get_close_matches(stem, same_dir, n=1, cutoff=0.6)
            errors.append(f"그리드 JSON 의 짝 PDF 없음: {stem}{GRID_SUFFIX}"
                          + (f"  → 비슷한 PDF: {hint[0]}.pdf (이름 불일치?)" if hint else ""))
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            errors.append(f"JSON 파싱 실패: {stem}{GRID_SUFFIX} — {e}")
            continue
        fname = data.get("fileName", "")
        if fname and nfc(fname) != os.path.basename(stem) + ".pdf":
            warns.append(f"JSON fileName '{fname}' ≠ PDF 이름 '{os.path.basename(stem)}.pdf': {stem}{GRID_SUFFIX}")
        total = data.get("totalPages", 0)
        for pg, blocks in (data.get("pageData") or {}).items():
            if not str(pg).isdigit() or (total and not 1 <= int(pg) <= total):
                errors.append(f"페이지 키 '{pg}' 가 1..totalPages({total}) 밖: {stem}{GRID_SUFFIX}")
            for bi, b in enumerate(blocks or []):
                tag = f"{stem}{GRID_SUFFIX} p{pg} 블록{bi + 1}"
                cols, rows = b.get("cols", 0), b.get("rows", 0)
                if cols < 1 or rows < 1:
                    errors.append(f"cols/rows 가 1 미만 (cols={cols}, rows={rows}): {tag}")
                    continue
                for k in ("x", "y", "width", "height"):
                    if not isinstance(b.get(k), (int, float)):
                        errors.append(f"좌표 '{k}' 없음: {tag}")
                if all(isinstance(b.get(k), (int, float)) for k in ("x", "y", "width", "height")):
                    if b["width"] <= 0 or b["height"] <= 0 or b["x"] < 0 or b["y"] < 0 \
                            or b["x"] + b["width"] > 1.0001 or b["y"] + b["height"] > 1.0001:
                        errors.append(f"좌표가 페이지(0~1) 밖: {tag}")
                for c in b.get("columnConfigs") or []:
                    ci, sr, er = c.get("colIndex", 0), c.get("startRow", 0), c.get("endRow", 0)
                    if not 1 <= ci <= cols or not 1 <= sr <= er <= rows:
                        errors.append(f"열 구성 오류 col{ci} r{sr}~{er} (cols={cols}, rows={rows}): {tag}")

    for stem in sorted(set(pdfs) - set(jsons)):
        infos.append(f"그리드 JSON 없음(기본 그리드로 동작): {stem}.pdf")

    # 곡 제목 목록(그리드 JSON "songs"): 쪽 번호가 PDF 쪽 수 안, grid 는 그 쪽 블록, 블록 곡명은 songs 와 같아야 한다
    for stem, path in sorted(jsons.items()):
        if stem not in pdfs:
            continue
        try:
            songs = read_songs(path)
            with open(path, encoding="utf-8") as f:
                grid = json.load(f)
        except Exception:
            continue                           # 파싱 오류는 위에서 알렸다
        pd = grid.get("pageData") or {}
        if not pd:
            infos.append(f"쪽 데이터 없는 그리드 JSON(곡 제목만, 기본 그리드로 동작): {stem}{GRID_SUFFIX}")
        if not songs:
            continue
        try:
            import pymupdf
            n = pymupdf.open(pdfs[stem]).page_count
        except Exception:
            n = 0
        seen = set()
        for s in songs:
            tag = f"곡 '{s['title']}'"
            if n and not 1 <= s["page"] <= n:
                errors.append(f"{tag} 쪽 {s['page']} 가 1..{n} 밖: {stem}{GRID_SUFFIX}")
            blocks = pd.get(str(s["page"])) or []
            if "grid" in s and not any(b.get("gridId") == s["grid"] for b in blocks):
                errors.append(f"{tag} 시작 블록 grid{s['grid']} 가 {s['page']}쪽에 없음: {stem}{GRID_SUFFIX}")
            key = (s["page"], s.get("grid"))
            if pd and key in seen:          # 그리드가 아직 없으면 블록으로 가를 수 없다(그리드를 만든 뒤 grid 를 적는다)
                errors.append(f"{tag} 시작 위치가 다른 곡과 같음(p{s['page']} grid{s.get('grid')}) — 한 쪽에 곡이 둘이면 grid 를 적는다: {stem}{GRID_SUFFIX}")
            seen.add(key)
        infos.append(f"곡 제목 {len(songs)}개: {stem}.pdf")
        # 블록 곡명: 곡 제목의 원천은 "songs" 하나다. 그리드 블록 title 은 거기서 만든다(앱 상단 곡명).
        #   곡을 더하거나 그리드를 다시 만들면 앱 저장소 scripts/apply_block_titles.py --apply 로 채운다.
        if pd:
            base = os.path.basename(stem)
            want = {s["title"] if s["title"].startswith(base) else f"{base} · {s['title']}" for s in songs}
            got = [b.get("title") for p in pd.values() for b in (p or [])]
            fix = f" → 앱 저장소에서 python3 scripts/apply_block_titles.py --only {base} --apply"
            if any(not t for t in got):
                errors.append(f"블록 곡명(title) 없는 블록 {sum(1 for t in got if not t)}개: {stem}{GRID_SUFFIX}{fix}")
            else:
                if set(got) - want:
                    errors.append(f"블록 곡명이 \"songs\" 에 없음 {sorted(set(got) - want)[:3]}: {stem}{GRID_SUFFIX}{fix}")
                if want - set(got):
                    errors.append(f"\"songs\" 곡이 블록 곡명에 없음 {sorted(want - set(got))[:3]}: {stem}{GRID_SUFFIX}{fix}")
    return errors, warns, infos


def report_book_pairs(book_dir: str) -> int:
    errors, warns, infos = check_book_pairs(book_dir)
    print("🔎 정간보 PDF ↔ 그리드 JSON 검사")
    for m in infos:
        print(f"  ℹ️  {m}")
    for m in warns:
        print(f"  ⚠️  {m}")
    for m in errors:
        print(f"  ❌ {m}")
    if errors:
        print(f"❌ 오류 {len(errors)}건 — 고친 뒤 다시 실행하세요(이 상태로 push 하면 앱에서 그리드가 안 그려집니다).\n")
    else:
        print(f"✅ 오류 없음 (경고 {len(warns)}건)\n")
    return len(errors)


# ── 메인 ──────────────────────────────────────────────────────
def main():
    base_dir   = os.path.dirname(os.path.abspath(__file__))
    audio_dir  = os.path.join(base_dir, "audio")
    book_dir   = os.path.join(base_dir, "book")
    audio_out  = os.path.join(base_dir, "audio-index.json")
    book_out   = os.path.join(base_dir, "book-index.json")

    print(f"📁 레포 경로: {base_dir}")
    print(f"🔗 Base URL : {BASE}\n")

    # ── 검사 전용 모드: python3 generate_index.py --check ──
    if "--check" in sys.argv:
        sys.exit(1 if report_book_pairs(book_dir) else 0)

    # ── audio ──
    if os.path.isdir(audio_dir):
        print("🎵 audio/ 스캔 중...")
        audio_data = scan_audio(audio_dir)
        with open(audio_out, "w", encoding="utf-8") as f:
            json.dump(audio_data, f, ensure_ascii=False, indent=2)
        t = audio_data["_total"]
        print(f"✅ audio-index.json  {t['categories']}개 카테고리 / {t['tracks']}개 트랙\n")
    else:
        print(f"⚠️  audio/ 디렉토리 없음: {audio_dir}\n")

    # ── book ──
    book_errors = 0
    if os.path.isdir(book_dir):
        book_errors = report_book_pairs(book_dir)
        print("📚 book/ 스캔 중...")
        book_data = scan_book(book_dir)
        with open(book_out, "w", encoding="utf-8") as f:
            json.dump(book_data, f, ensure_ascii=False, indent=2)
        t = book_data["_total"]
        print(f"✅ book-index.json   {t['categories']}개 카테고리 / {t['files']}개 파일\n")
    else:
        print(f"⚠️  book/ 디렉토리 없음: {book_dir}\n")

    print("📄 생성된 파일:")
    print(f"   {audio_out}")
    print(f"   {book_out}")
    print()
    print("다음 단계:")
    print("  git add audio-index.json book-index.json")
    print("  git commit -m 'chore: update audio/book index'")
    print("  git push")
    if book_errors:
        print(f"\n❌ 정간보 검사 오류 {book_errors}건 — push 전에 고치세요.")
        sys.exit(1)


if __name__ == "__main__":
    main()

