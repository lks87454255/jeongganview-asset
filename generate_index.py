#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_index.py
=================
로컬 audio/ 및 book/ 디렉토리를 스캔하여
  - audio-index.json  (음원 카테고리/트랙 목록)
  - book-index.json   (정간보 카테고리/파일 목록)
을 생성합니다.

생성 후 git push 하면 앱이 GitHub API 없이
raw.githubusercontent.com 에서 단 1회 요청으로 전체 목록을 가져옵니다.
(GitHub API Rate Limit: 비인증 60회/시간 → 완전 우회)

앱의 URL 인코딩 방식:
  Java: URLEncoder.encode(name, "UTF-8").replace("+", "%20")
  Python: urllib.parse.quote(name, safe="")   ← 완전 동일
"""


import os
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


# ── 메인 ──────────────────────────────────────────────────────
def main():
    base_dir   = os.path.dirname(os.path.abspath(__file__))
    audio_dir  = os.path.join(base_dir, "audio")
    book_dir   = os.path.join(base_dir, "book")
    audio_out  = os.path.join(base_dir, "audio-index.json")
    book_out   = os.path.join(base_dir, "book-index.json")

    print(f"📁 레포 경로: {base_dir}")
    print(f"🔗 Base URL : {BASE}\n")

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
    if os.path.isdir(book_dir):
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


if __name__ == "__main__":
    main()

