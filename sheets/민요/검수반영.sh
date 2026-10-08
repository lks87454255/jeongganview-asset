#!/usr/bin/env bash
# 검수 xlsx → csv 재생성 → sheets-index.json 갱신 → 커밋 → push
#
# 사용 (어느 폴더에서 실행해도 됨):
#   ./sheets/민요/검수반영.sh              # 바뀐 곡 확인 후 물어보고 반영·커밋·push
#   ./sheets/민요/검수반영.sh --check      # 무엇이 바뀔지만 보기 (파일 변경 없음)
#   ./sheets/민요/검수반영.sh 강강술래 농부가   # 특정 곡만
#
# 주의: _작업/build_all.py 는 실행하지 말 것 (TSV 로 csv·xlsx 를 다시 만들어 검수 내용을 덮어씀)
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"          # sheets/민요
REPO="$(cd "$HERE/../.." && pwd)"               # jeongganview-asset
cd "$REPO"

CHECK=0
SONGS=()
for a in "$@"; do
  if [[ "$a" == "--check" ]]; then CHECK=1; else SONGS+=("$a"); fi
done

# 엑셀/LibreOffice 에서 열려 있으면 저장이 안 끝났을 수 있음
if ls "$HERE/검수/" 2>/dev/null | grep -qE '^(~\$|\.~lock\.)'; then
  echo "⚠️  검수 폴더에 열려 있는 xlsx(잠금 파일)가 있습니다. 저장하고 닫은 뒤 다시 실행하세요."
  ls -a "$HERE/검수/" | grep -E '^(~\$|\.~lock\.)' | sed 's/^/   /'
  exit 1
fi

echo "▶ 원격 최신 받기 (git pull --ff-only)"
git pull --ff-only -q

echo "▶ 바뀔 곡 확인"
python3 "$HERE/_작업/xlsx_to_csv.py" --check ${SONGS[@]+"${SONGS[@]}"}
[[ $CHECK -eq 1 ]] && exit 0

read -r -p "위 곡을 csv 에 반영하고 커밋·push 할까요? [y/N] " ans
[[ "$ans" =~ ^[Yy]$ ]] || { echo "취소했습니다."; exit 0; }

echo "▶ csv 재생성"
python3 "$HERE/_작업/xlsx_to_csv.py" ${SONGS[@]+"${SONGS[@]}"}

# csv·xlsx 가 그대로면 index 도 만들지 않음 (index 는 생성 시각만 바뀌어도 diff 가 생김)
if [[ -z "$(git status --porcelain -- "sheets/민요/검수" "sheets/민요/csv")" ]]; then
  echo "바뀐 내용이 없습니다."
  exit 0
fi

echo "▶ sheets-index.json 갱신"
python3 generate_sheets_index.py >/dev/null

git add -- "sheets/민요/검수" "sheets/민요/csv" sheets-index.json

echo "▶ 커밋할 파일"
git -c core.quotepath=false diff --cached --stat

# 커밋 메시지에 곡 이름 넣기
NAMES=$(git -c core.quotepath=false diff --cached --name-only -- "sheets/민요/csv" \
        | sed -E 's#.*/##; s/\.csv$//' | paste -sd ',' - | sed 's/,/, /g')
git commit -q -m "fix(sheets): 민요 검수 반영 (${NAMES:-xlsx})"
git push -q origin main
echo "✅ push 완료: $(git log --oneline -1)"
