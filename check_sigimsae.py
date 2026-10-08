# -*- coding: utf-8 -*-
"""
check_sigimsae.py  (jeongganview-asset)
=======================================
sigimsae/*.json (악상기호 카탈로그)를 앱과 같은 규칙으로 검사합니다.
앱(JeongGanView)은 원격 파일이 이 규칙을 하나라도 어기면 그 파일을 버리고
마지막 정상 캐시나 앱 내장 사본을 씁니다 → push 전에 반드시 통과시키세요.

    python3 check_sigimsae.py            # sigimsae/ 아래 *.json 전부
    python3 check_sigimsae.py FILE.json  # 지정한 파일만

앱 쪽 규칙: data/sigimsae/SigimsaeCatalogValidator.kt, SigimsaeExpander.kt, SigimsaeModels.kt
오류가 하나라도 있으면 종료 코드 1.
"""
import glob
import json
import os
import sys

# ── 앱 SigimsaeCatalogValidator.SUPPORTED_SCHEMA_VERSION 과 같게 ──
SUPPORTED_SCHEMA_VERSION = 1

# ── 앱 SigimsaeModels.kt 의 enum 과 같게(새 값을 쓰려면 앱부터 배포하고 schemaVersion 을 올린다) ──
CATEGORIES = {"BUHO", "ORNAMENT", "TEMPO", "TECHNIQUE", "JANGGU"}
PLACEMENTS = {"SLOT", "ATTACHED"}
TIMINGS = {"EVEN", "GRACE", "SHORT_LONG"}
EFFECT_TYPES = {
    "CONTINUE", "BREATH", "REST", "STACCATO", "ACCENT", "FERMATA", "LENGTHEN", "SHORTEN", "REPEAT",
    "TEMPO_SLOWER", "TEMPO_FASTER",
    "BEND_UP", "BEND_UP_DOUBLE", "BEND_DOWN", "BEND_DOWN_DOUBLE", "STRONG_ATTACK",
    "VIBRATO", "VIBRATO_DESCEND", "YOSEONG", "YOSEONG_DOUBLE", "TTEOIEO", "SAME_AS_PREVIOUS",
    "JANGGU_STROKE",
}

# ── 앱 util/Yulmyeong.kt NAMES 와 같게(반음 60음, 黃=0 pitch class) ──
YUL = list("㣴㣕㣖㣣㣨㣡㣸㣩") + ["𢓡"] + list("㣮㣳㣹僙㐲㑀俠㑬㑖") + ["𠐭"] + list(
    "㑣侇㑲㒇㒣黃大太夾姑仲㽔林夷南無應潢汏汰浹㴌㳞㶋淋洟湳潕㶐㶂") + ["𣴘"] + list("㳲㴺㵈㴢㶙㵉㴣㵜㶃㶝")
assert len(YUL) == 60, len(YUL)
YUL_INDEX = {n: i for i, n in enumerate(YUL)}
# 대금 기본 평조 5음 黃 太 仲 林 南 (SigimsaeExpander.PYEONGJO_PITCH_CLASSES)
DEGREES = [i for i in range(60) if i % 12 in {0, 2, 5, 7, 9}]

# 앱 data/JangdanModel.kt 의 Instrument enum 이름(장구 appInstrument 검사). 앱에서 악기가 바뀌면 같이 고친다.
# 비워 두면(None) 검사하지 않는다.
INSTRUMENTS = None


class ExpansionError(Exception):
    pass


def step_pitch(base, step):
    if base not in YUL_INDEX:
        raise ExpansionError(f"알 수 없는 율명: {base}")
    idx = YUL_INDEX[base]
    if idx not in DEGREES:
        raise ExpansionError(f"{base} 는 평조 5음(黃太仲林南)이 아님")
    pos = DEGREES.index(idx) + step
    if not 0 <= pos < len(DEGREES):
        raise ExpansionError(f"{base} {step:+d} 음역 밖")
    return YUL[DEGREES[pos]]


def resolve_steps(sym, by_id):
    exp = sym.get("expansion")
    if not exp:
        raise ExpansionError(f"{sym['id']}: 전개 정보 없음")
    compose = exp.get("composeOf") or []
    if not compose:
        return list(exp.get("steps") or [])
    out = []
    for part_id in compose:
        part = by_id.get(part_id)
        if part is None:
            raise ExpansionError(f"{sym['id']}: 구성 기호 없음 {part_id}")
        steps = resolve_steps(part, by_id)
        if out and out[-1] == 0 and steps and steps[0] == 0:
            out += steps[1:]
        else:
            out += steps
    return out


def expand_names(base, sym, by_id):
    steps = resolve_steps(sym, by_id)
    if not steps:
        raise ExpansionError(f"{sym['id']}: 단계열 비어 있음")
    names = [step_pitch(base, s) for s in steps]
    exp = sym["expansion"]
    if exp["timing"] == "GRACE":
        main = exp.get("mainIndex")
        if main is None:
            main = len(names) - 1
        if not 0 <= main < len(names):
            raise ExpansionError(f"{sym['id']}: mainIndex 범위 밖")
    if exp["timing"] == "SHORT_LONG" and len(names) != 2:
        raise ExpansionError(f"{sym['id']}: SHORT_LONG 은 2음이어야 함")
    return names


def has_cycle(start, by_id):
    def visit(sid, path):
        parts = ((by_id.get(sid) or {}).get("expansion") or {}).get("composeOf") or []
        return any(p in path or visit(p, path | {p}) for p in parts)
    return visit(start, {start})


def check(path):
    errors = []
    try:
        data = json.load(open(path, encoding="utf-8"))
    except Exception as e:  # noqa: BLE001
        return [f"JSON 읽기 실패: {e}"]

    ver = data.get("schemaVersion")
    if not isinstance(ver, int) or not 1 <= ver <= SUPPORTED_SCHEMA_VERSION:
        return [f"schemaVersion {ver} 은 지원하지 않음(최대 {SUPPORTED_SCHEMA_VERSION})"]
    src = data.get("source") or {}
    if not isinstance(src.get("title"), str) or not isinstance(src.get("pages"), str):
        errors.append("source.title / source.pages 필요")
    symbols = data.get("symbols")
    if not isinstance(symbols, list) or not symbols:
        return errors + ["symbols 비어 있음"]

    # ── 형식(앱 kotlinx.serialization 이 읽을 수 있는지) ──
    for i, s in enumerate(symbols):
        sid = s.get("id") or f"#{i}"
        for key in ("id", "category", "glyph", "name", "description"):
            if not isinstance(s.get(key), str) or (key == "id" and not s.get(key).strip()):
                errors.append(f"{sid}: {key} 필요(문자열)")
        if not isinstance(s.get("page"), int):
            errors.append(f"{sid}: page 필요(정수)")
        if s.get("category") not in CATEGORIES:
            errors.append(f"{sid}: 앱이 모르는 category {s.get('category')}")
        exp, eff = s.get("expansion"), s.get("effect")
        if (exp is None) == (eff is None):
            errors.append(f"{sid}: expansion 과 effect 중 정확히 하나만 있어야 함")
        if exp is not None:
            if exp.get("placement") not in PLACEMENTS:
                errors.append(f"{sid}: 앱이 모르는 placement {exp.get('placement')}")
            if exp.get("timing") not in TIMINGS:
                errors.append(f"{sid}: 앱이 모르는 timing {exp.get('timing')}")
            if not all(isinstance(v, int) for v in exp.get("steps") or []):
                errors.append(f"{sid}: steps 는 정수 목록")
        if eff is not None and eff.get("type") not in EFFECT_TYPES:
            errors.append(f"{sid}: 앱이 모르는 effect.type {eff.get('type')}")
        ex = s.get("example")
        if ex is not None and (not isinstance(ex.get("base"), str) or not isinstance(ex.get("printed"), list)):
            errors.append(f"{sid}: example 은 base(문자열)·printed(목록) 필요")
        if INSTRUMENTS is not None and s.get("appInstrument") not in (None, *INSTRUMENTS):
            errors.append(f"{sid}: 앱에 없는 악기 {s.get('appInstrument')}")
    if errors:
        return errors

    # ── 의미(앱 SigimsaeCatalogValidator 와 같은 순서) ──
    ids = [s["id"] for s in symbols]
    for dup in sorted({x for x in ids if ids.count(x) > 1}):
        errors.append(f"id 중복: {dup}")
    by_id = {s["id"]: s for s in symbols}
    for s in symbols:
        for part in (s.get("expansion") or {}).get("composeOf") or []:
            if part not in by_id:
                errors.append(f"{s['id']}: composeOf 에 없는 기호 {part}")
    for s in symbols:
        if has_cycle(s["id"], by_id):
            errors.append(f"{s['id']}: composeOf 순환")
    if errors:
        return errors

    for s in symbols:
        ex = s.get("example")
        if ex is None:
            continue
        expected = ex.get("corrected") or ex["printed"]
        try:
            got = expand_names(ex["base"], s, by_id)
        except ExpansionError as e:
            got = str(e)
        if got != expected:
            errors.append(f"{s['id']}: 예시 {ex['base']} 기대 {expected} / 엔진 {got}")
    return errors


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    files = sys.argv[1:] or sorted(glob.glob(os.path.join(here, "sigimsae", "*.json")))
    if not files:
        print("검사할 파일 없음: sigimsae/*.json")
        return 1
    failed = 0
    for f in files:
        errs = check(f)
        name = os.path.relpath(f, here)
        if errs:
            failed += 1
            print(f"ERROR {name}: {len(errs)}건")
            for e in errs:
                print(f"  - {e}")
        else:
            n = len(json.load(open(f, encoding="utf-8"))["symbols"])
            print(f"OK    {name}: 기호 {n}개")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
