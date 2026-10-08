# sigimsae (악상기호)

정간뷰(JeongGanView) 앱 **악상기호** 메뉴의 데이터 정본입니다. 앱 배포 없이 고칠 수 있습니다.

```
sigimsae/악상기호.json                ← 2021 대금정악보 10~15쪽 악상기호 60개 (파일 이름은 앱과 약속된 값 — 바꾸면 앱이 못 찾는다)
check_sigimsae.py                    ← 앱과 같은 규칙으로 검사 (push 전 필수)
```

## 앱이 불러오는 방식

1. 앱 내장 사본(또는 마지막으로 받은 캐시)으로 화면을 바로 띄운다.
2. 이 저장소의 파일을 조건부 요청(ETag)으로 받아, 바뀌었으면 검사한다.
3. 검사를 **통과한 경우에만** 화면에 반영하고 캐시에 저장한다. 하나라도 어기면 버리고 이전 것을 계속 쓴다.

검사 규칙: 앱이 아는 enum 값만 사용 · id 중복 없음 · `expansion` 과 `effect` 중 하나만 · `composeOf` 대상 존재·순환 없음 ·
모든 `example` 을 전개 엔진이 그대로 재현(오기는 `corrected` 기준) · `schemaVersion` ≤ 앱 지원 버전(현재 1).

## 고치는 순서

```
python3 check_sigimsae.py      # OK 가 나와야 함 (오류 있으면 종료 코드 1)
git add sigimsae check_sigimsae.py
git commit -m "chore: update sigimsae"
git push
```

- 설명·이름·예시·glyph 수정은 바로 반영된다(앱 재시작 후 악상기호 화면을 열 때).
- **새 enum 값**(예: 새 `effect.type`)이나 필드 의미 변경은 앱이 먼저 지원해야 한다.
  앱을 배포한 뒤 `schemaVersion` 을 올리고, `check_sigimsae.py` 의 목록·`SUPPORTED_SCHEMA_VERSION` 도 같이 고친다.
  (옛 앱은 새 버전 파일을 버리고 자기 사본을 계속 쓴다)
- 앱 내장 사본(`JeongGanView/app/src/main/assets/sigimsae/`)은 오프라인·첫 실행용이다. 릴리스 전에 이 파일로 덮어써 맞춘다.
