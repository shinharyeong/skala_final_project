#!/bin/zsh
cd -- "$(dirname -- "$0")/.." || exit 1
PYTHON_BIN="parser_ai/.venv/bin/python3"
[ -x "$PYTHON_BIN" ] || PYTHON_BIN="python3"
echo "사용 Python: $PYTHON_BIN"
"$PYTHON_BIN" -B parser_ai/run.py prepare parser/Data/dummy/organizations
result=$?
echo ""
echo "준비 실행 종료 코드: $result (0=완료). API 호출과 성능 비교는 하지 않았습니다."
echo "결과: parser_ai/output/의 새 실행 폴더"
read "?Enter를 누르면 닫습니다."
exit $result
