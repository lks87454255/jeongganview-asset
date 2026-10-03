#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_index.py
=================
로컬 audio/ , book/ , video/ 디렉토리를 스캔하여
  - audio-index.json  (음원 카테고리/트랙 목록)
  - book-index.json   (정간보 카테고리/파일 목록)
  - video-index.json  (동영상 카테고리/파일 목록)
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
VIDEO_EXTS = {".mp4", ".webm", ".mkv", ".avi", ".mov"}

# 무시할 파일/디렉토리 패턴
IGNORE = {".DS_Store", ".gitkeep", ".gitignore", "Thumbs.db"}


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
def scan_book(book_dir: str) -> dict:
    categories = []

    cat_names = sorted(
        n for n in os.listdir(book_dir)
        if os.path.isdir(os.path.join(book_dir, n)) and not is_ignored(n)
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

        if files:
            categories.append({"name": cat_name, "files": files})
            total_mb = sum(f["sizeBytes"] for f in files) / 1_048_576
            print(f"  📚 [{cat_name}]  {len(files)}개 파일  ({total_mb:.1f} MB)")
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
        "_total":     {"categories": len(categories), "files": total_files},
    }


# ── video-index.json 생성 ──────────────────────────────────────
def scan_video(video_dir: str) -> dict:
    categories = []

    cat_names = sorted(
        n for n in os.listdir(video_dir)
        if os.path.isdir(os.path.join(video_dir, n)) and not is_ignored(n)
    )

    for cat_name in cat_names:
        cat_path = os.path.join(video_dir, cat_name)
        videos = []

        for file_name in sorted(os.listdir(cat_path)):
            if is_ignored(file_name):
                continue
            ext = os.path.splitext(file_name)[1].lower()
            if ext not in VIDEO_EXTS:
                continue

            display_name = os.path.splitext(file_name)[0]  # 확장자 제거
            asset_path   = raw_url("video", cat_name, file_name)
            size_bytes   = os.path.getsize(os.path.join(cat_path, file_name))

            videos.append({
                "displayName": display_name,
                "fileName":    file_name,
                "assetPath":   asset_path,
                "sizeBytes":   size_bytes,
            })

        if videos:
            categories.append({"name": cat_name, "videos": videos})
            total_mb = sum(v["sizeBytes"] for v in videos) / 1_048_576
            print(f"  🎬 [{cat_name}]  {len(videos)}개 파일  ({total_mb:.1f} MB)")
        else:
            print(f"  ⚠️  [{cat_name}]  동영상 파일 없음 — 건너뜀")

    total_videos = sum(len(c["videos"]) for c in categories)
    return {
        "version":    1,
        "generated":  datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "owner":      OWNER,
        "repo":       REPO,
        "branch":     BRANCH,
        "categories": categories,
        "_total":     {"categories": len(categories), "videos": total_videos},
    }


# ── 정간보 PDF ↔ 그리드 JSON 일치 검사 ─────────────────────────
# 앱은 PDF 주소의 ".pdf" 를 ".json" 으로 바꿔 그리드 JSON 을 찾는다.
# 이름이 한 글자라도 다르면(예: 한천수.pdf / 헌천수.json) JSON 이 404 → 그리드가 그려지지 않는다.
#   오류(ERROR): JSON 짝 PDF 없음 · JSON 파싱/구조 오류 · 파일명 NFC 아님
#   경고(WARN) : JSON 안 fileName 이 PDF 이름과 다름
#   정보(INFO) : JSON 없는 PDF (기본 그리드로 동작 — 허용)
def check_book_pairs(book_dir: str):
    import difflib
    errors, warns, infos = [], [], []
    pdfs, jsons = {}, {}
    for root, _, files in os.walk(book_dir):
        for fn in files:
            if is_ignored(fn):
                continue
            rel = os.path.relpath(os.path.join(root, fn), os.path.dirname(book_dir))
            if fn != nfc(fn):
                errors.append(f"파일명이 NFC(완성형)가 아님 — 앱 URL 과 달라 404: {rel}")
            stem, ext = os.path.splitext(nfc(rel))
            if ext.lower() == ".pdf":
                pdfs[stem] = os.path.join(root, fn)
            elif ext.lower() == ".json":
                jsons[stem] = os.path.join(root, fn)

    for stem, path in sorted(jsons.items()):
        if stem not in pdfs:
            same_dir = [p for p in pdfs if os.path.dirname(p) == os.path.dirname(stem)]
            hint = difflib.get_close_matches(stem, same_dir, n=1, cutoff=0.6)
            errors.append(f"그리드 JSON 의 짝 PDF 없음: {stem}.json"
                          + (f"  → 비슷한 PDF: {hint[0]}.pdf (이름 불일치?)" if hint else ""))
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            errors.append(f"JSON 파싱 실패: {stem}.json — {e}")
            continue
        fname = data.get("fileName", "")
        if fname and nfc(fname) != os.path.basename(stem) + ".pdf":
            warns.append(f"JSON fileName '{fname}' ≠ PDF 이름 '{os.path.basename(stem)}.pdf': {stem}.json")
        total = data.get("totalPages", 0)
        for pg, blocks in (data.get("pageData") or {}).items():
            if not str(pg).isdigit() or (total and not 1 <= int(pg) <= total):
                errors.append(f"페이지 키 '{pg}' 가 1..totalPages({total}) 밖: {stem}.json")
            for bi, b in enumerate(blocks or []):
                tag = f"{stem}.json p{pg} 블록{bi + 1}"
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
    video_dir  = os.path.join(base_dir, "video")
    audio_out  = os.path.join(base_dir, "audio-index.json")
    book_out   = os.path.join(base_dir, "book-index.json")
    video_out  = os.path.join(base_dir, "video-index.json")

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

    # ── video ──
    if os.path.isdir(video_dir):
        print("🎬 video/ 스캔 중...")
        video_data = scan_video(video_dir)
        with open(video_out, "w", encoding="utf-8") as f:
            json.dump(video_data, f, ensure_ascii=False, indent=2)
        t = video_data["_total"]
        print(f"✅ video-index.json  {t['categories']}개 카테고리 / {t['videos']}개 파일\n")
    else:
        print(f"⚠️  video/ 디렉토리 없음: {video_dir}\n")

    print("📄 생성된 파일:")
    print(f"   {audio_out}")
    print(f"   {book_out}")
    print(f"   {video_out}")
    print()
    print("다음 단계:")
    print("  git add audio-index.json book-index.json video-index.json")
    print("  git commit -m 'chore: update audio/book/video index'")
    print("  git push")
    if book_errors:
        print(f"\n❌ 정간보 검사 오류 {book_errors}건 — push 전에 고치세요.")
        sys.exit(1)


if __name__ == "__main__":
    main()

