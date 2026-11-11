#!/bin/zsh
cd -- "$(dirname -- "$0")/.." || exit 1
PYTHON_BIN="parser_ai/.venv/bin/python3"
[ -x "$PYTHON_BIN" ] || PYTHON_BIN="python3"
echo "AI 파싱: 기본은 더미 조직 1개. OPENAI_API_KEY를 parser_ai/.env에 설정하세요."
echo "키를 설정한 경우 OpenAI API 호출 비용이 발생합니다."
echo "사용 Python: $PYTHON_BIN"
"$PYTHON_BIN" -B parser_ai/run.py run parser/Data/dummy/organizations --max-files 1
result=$?
echo ""
echo "AI 파싱 종료 코드: $result (0=완료, 2=설정/추론/검증 확인 필요)."
echo "결과: parser_ai/output/의 새 실행 폴더"
read "?Enter를 누르면 닫습니다."
exit $result
