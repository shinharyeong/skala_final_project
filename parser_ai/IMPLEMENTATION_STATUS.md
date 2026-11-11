# 구현·검증 상태 — 2026-11-11 (갱신)

> **이 문서는 이전 셀별(legacy-cell) 엔진 단계의 기록이며 이후 작업(구조 해석
> 기본 엔진, 실제 API 테스트 결과)을 반영하지 않는다. 현재 상태와 실제 테스트
> 결과는 [README.md](README.md)의 "실제로 실행한 테스트와 검증되지 않은
> 범위"를 기준으로 본다.**

## (과거 기록) 이번 세션에서 한 일

- 기존 `parser/`는 바이트 단위로 수정하지 않음. 모든 변경은 `parser_ai/` 안에서만 수행.
- 아키텍처 변경: 기존에는 LLM이 셀 포인터만 반환하고 Python이
  `normalize_number()`로 숫자를 정규화했다. 이번 세션에서 이를 바꿔,
  **Python은 숫자를 가리지 않고 원본 값·서식·수식 캐시를 그대로 모델에
  전달하며, LLM이 각 재무 후보 셀의 부호·단위·백분율을 직접 해석하고
  정규화된 값을 직접 반환**하도록 `schema.py`/`context.py`/`engine.py`를
  재작성함. Python은 그 값을 재계산·교정하지 않고, 독립 검산과 다르면
  `AI_NORMALIZATION_DISAGREEMENT` 진단만 남긴다.
- 그라운딩 검증 추가: 모델이 주장하는 `source_value_read`가 실제 추출된
  원본 값(또는 수식의 저장 캐시)과 일치하는지 검사하고, 불일치 시 재요청
  (repair) 후에도 실패하면 해당 배치를 명시적으로 실패 처리한다(룰 결과로
  대체하지 않음).
- 행 단위로 묻던 것을 "행 안의 모든 숫자/수식/지정 값열 셀" 단위로 세분화하여,
  셀 커버리지(누락·중복) 검증을 추가함.
- `parser_ai/.venv`에 Python 3.11 전용 가상환경 생성, 패키지를 editable
  설치(외부 의존성 없음).
- `parser_ai/.env.example`에서 실수로 들어가 있던 **실제 작동 가능한 OpenAI
  키를 발견하여 제거함(2회, 파일이 중간에 복원됨)**. 이 키는 이미 파일에
  노출되었으므로 사용자가 OpenAI 대시보드에서 즉시 폐기해야 함.
- 테스트(`tests/test_parser_ai.py`)를 새 아키텍처에 맞게 전면 재작성: 모델
  정규화값을 Python이 교정하지 않는지, 숫자가 실제로 모델 입력에 포함되는지,
  그라운딩 실패·셀 커버리지·수식 캐시 누락·0/대시/비율 구분·기존 출력 계약
  필드 호환성·API 실패 시 룰 대체가 없는지를 검증. 26개 전부 통과(mock).
- 기존 `parser/` 자체 테스트 40개(37 통과 + 외부 데이터 의존 3개 skip) — 무변경 확인.
- 실제 더미 조직 17개(153시트, 435,965 content cell)로 `prepare` 실행 —
  원본 해시 보존, 독립 원본 대조, 후보 보존 모두 통과(API 미호출).
- 스크립트 기반 mock 모델로 더미 조직 1개(14,132개 숫자/수식 후보)를 대상으로
  `run` 전체 경로(레이아웃→행/셀 추론→그라운딩→compile_records→출처 검증)를
  review on/off 양쪽 다 실행하여 통과 확인. 이 mock은 실제 LLM 품질을
  검증하지 않으며, 그라운딩·커버리지 로직이 실제 크기의 워크북에서도
  동작하는지만 확인한 것.
- OpenAI 공식 문서를 재확인: `gpt-6-astra` 모델 존재, `reasoning.effort`
  단계(low/medium/high/xhigh/max, `none` 미지원), Responses API
  `text.format` 구조화 출력(json_schema, strict), `store=false` 파라미터
  모두 현재 문서와 일치함을 확인 — 클라이언트 코드 변경 불필요.

## 아직 안 된 것 / 명시적 범위 제외

- 실제 OpenAI API 호출 0회. API 키 미설정(사용자가 `parser_ai/.env`에 입력
  예정). **실제 LLM 파싱 품질·정확도는 전혀 검증되지 않음.**
- 성능 비교·정답표 평가·마스터 집계·HTML 비교 화면은 이번 범위에서 제외.
- 소스 GitHub push/커밋/PR 없음. `.git` 저장소 자체가 없음(ZIP 스냅샷).

## 다음 단계 (사용자가 할 일)

1. `parser_ai/.env.example`을 `parser_ai/.env`로 복사(없으면 생성됨)하고
   `OPENAI_API_KEY=` 뒤에 본인 키를 입력한다.
   **주의: 과거 `.env.example`에 노출되었던 키는 이미 유출된 것으로 간주하고
   OpenAI 대시보드에서 폐기(Revoke)한 뒤 새 키를 써야 한다.**
2. `parser_ai/run_ai_parser.command`를 실행하거나
   `python3 -B parser_ai/run.py run parser/Data/dummy/organizations --max-files 1`
   로 조직 1개에 대해 실제 API 파싱을 시작한다.
3. 결과는 `parser_ai/output/` 아래 새 폴더에 저장된다. `inference.json`의
   호출 수·토큰 사용량을 확인한 뒤 17개 전체 실행 여부를 판단한다(비용 발생).
4. 이후 새로운 확인용 데이터로 비교할 때는 동일한 정답 기준과 표준 출력
   계약을 사용하고, 입력 경로만 바꿔 실행한다.
