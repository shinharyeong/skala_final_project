# 구현·검증 상태 — 2026-11-11 (과거 기록)

> **이 문서의 아래 섹션들은 이전 셀별(legacy-cell) 엔진 단계의 기록이며 이후
> 작업(구조 해석 기본 엔진, 실제 API 테스트 결과)을 반영하지 않는다. 가장
> 최신 작업 기록은 "rectangle 검증 실패 원인 분석·시트별 독립 처리" 섹션과
> [README.md](README.md)의 "실제로 실행한 테스트와 검증되지 않은 범위"를
> 기준으로 본다.**

## rectangle 검증 실패 원인 분석·시트별 독립 처리·검증 메시지 구체화

**배경**: `--engine structure`로 `01_F10_사업부1_재무제표.xlsx`를 실제
실행했을 때(output: `parser_ai/output/20261111-235354-17cc6bdc`),
`매출_세부_IFRS15後` 시트에서 `validate_plan` 검증이
`"invalid table/header rectangle"`로 실패했고, repair 재시도도 똑같이
실패했다. 이 실패가 `infer_structure()`의 메인 루프를 즉시 `break`시켜
나머지 8개 시트는 LLM 호출 자체를 시도조차 못 했다(`inference.json`에
이벤트가 `sheet_structure_v1`/`sheet_structure_v1_repair` 2건뿐임을 확인).

### 1. 원인 파악 (실제 캐시 데이터로 재현)

캐시 파일 `cache/12335627f16ca6f423ba95cf0bd4d7f6bb59b1c86cb4895c4a3c6152f3a44dee.json`
(1차 응답)과 `cache/525cf40f82b6f6f61ad20bcda037031d9c876bb15ed076e14db4265f581976cf.json`
(repair 응답) 모두 동일하게 `start_row=3, header_start_row=3,
header_end_row=4`를 선언했다. 즉 **`start_row`가 `header_end_row + 1`(=5)이
아니라 `header_start_row`와 같은 값으로 선언됨** — 올바른 `start_row`는 5여야
했다.

실제 시트 구조를 직접 추출해 대조:
- 1행: 제목/안내 텍스트("작성" 등), 숫자 없음.
- **2행**: `A2`/`F2`에 안내 텍스트가 있지만, `H2:BJ2`에 걸쳐 52개의
  숫자(비율/예시성 수식 값)가 있고 `numeric_span=[8,62]` — 이는 **실제
  데이터 첫 행(5행)과 정확히 같은 숫자 분포(`numeric_span=[8,62]`,
  `numeric_count=52`)**다.
- 3행: 실제 컬럼 헤더 레이블(조직명/관리분류/.../사업부코드/'25년 실적/
  '26년 계획 등), 숫자 없음. A3:A4~F3:F4가 2행에 걸친 병합 헤더셀.
- 4행: 월 하위 헤더(1월~12월 등 반복), 숫자 없음.
- **5행부터 실제 데이터** 시작(`A5="사업부1"`, `B5="음성기본료"` 등).

즉 올바른 선언은 `header_start_row=3, header_end_row=4, start_row=5`여야
하는데 모델은 `start_row=3`(헤더 시작 행과 동일)으로 잘못 선언했다. 기여
요인으로 보이는 두 가지: (a) A~F열의 2행짜리 병합 헤더셀이 헤더가 "어디서
끝나는지"에 대한 판단을 흐렸을 가능성, (b) 2행이 안내/예시성 숫자 행임에도
실제 데이터 행(5행)과 **숫자 분포가 완전히 동일**해서, 모델이 숫자 분포만으로
"데이터가 어디서 시작하는지" 판단했을 경우 혼동을 유발했을 가능성. 두 가설
모두 완전히 배타적이지 않고 함께 작용했을 수 있다.

기존 `validate_plan`은 이 네 가지 서로 다른 조건(행 역전/헤더 역전/헤더-
데이터 경계/열 역전)을 전부 `"invalid table/header rectangle"`이라는 단일
문자열로 묶어서 반환했기 때문에, repair 요청에 "무엇이 왜 잘못됐는지"에 대한
구체적 신호가 전혀 전달되지 않았다 — repair가 **1차와 완전히 동일한 실수를
반복한 것**이 이를 뒷받침한다.

### 2. 바꾼 파일과 함수

| 파일 | 함수/항목 | 바뀐 내용 |
|---|---|---|
| `src/excel_parser_ai/structured.py` | `validate_plan` | rectangle 검증 하나의 generic 메시지를 4개의 구체적·숫자 포함 메시지로 분리(`start_row vs end_row`, `header_start_row vs header_end_row`, **`start_row vs header_end_row`**(병합 헤더·숫자행 혼동 안내 포함), `start_col vs end_col`) |
| 〃 | `PROMPT` | `start_row`는 반드시 `header_end_row + 1`이며 병합 헤더의 앵커 행도 헤더로 취급해야 함, 헤더 바로 위/안의 숫자처럼 보이는 행(안내·예시 행)이 있어도 숫자 분포만으로 데이터 시작 행을 판단하지 말 것이라는 지침 추가 |
| 〃 | `infer_structure` | 메인 루프의 `break # preserve completed plans/caches; no further requests`를 제거. 한 시트가 최종 실패하면 `failures`에 기록만 하고 **다음 시트는 계속 시도**(`continue`). 단, 그 시점에 `client.call_count >= client.max_calls`(파일 전체 호출 상한 소진)이면 더 시도해도 의미가 없으므로 그때만 `break` |

### 3. 전후 비교

| 항목 | 이전 | 이후 |
|---|---|---|
| 한 시트의 rectangle 검증 실패 시 나머지 시트 | 메인 루프가 즉시 `break` — 나머지 시트는 호출 자체를 시도하지 않음 | 해당 시트만 `failures`에 기록하고 나머지 시트는 계속 시도. `max_calls` 예산은 전체 파일 기준으로 그대로 적용(예산 소진 시에만 멈춤) |
| rectangle 검증 실패 메시지 | `"invalid table/header rectangle"` 한 줄(어느 조건이 왜 위반됐는지 불명) | 조건별로 분리된 구체적 메시지(예: `"table start_row (3) must be greater than header_end_row (4): ... 병합 헤더의 앵커 행도 헤더다 ..."`) — repair 재시도가 무엇을 고쳐야 하는지 알 수 있음 |
| PROMPT의 `start_row` 정의 | 암묵적(명시적 수식 없음) | `start_row = header_end_row + 1`을 명시, 병합 헤더·숫자처럼 보이는 안내 행에 대한 지침 추가 |

### 4. 테스트

`tests/test_structured.py`에 2개 범주, 5개 테스트 추가:
- `PerSheetIndependentFailureTests`(2개) — 한 시트가 실패해도 다음 시트는
  시도되는지(`test_one_sheets_failure_does_not_block_the_next_sheets_attempt`),
  호출 예산이 소진되면 그 이후 시트는 시도조차 안 하는지
  (`test_budget_exhaustion_stops_remaining_sheets_without_attempting_them`).
- `RectangleValidationMessageTests`(3개) — 실제 실패와 동일한 형태
  (`start_row == header_start_row`, 다중 행 헤더)가 generic 메시지 대신
  구체적 메시지를 내는지, `start_row`를 `header_end_row + 1`로 고치면
  검증을 통과하는지, 행/열 역전도 각각 구체적 메시지로 분리되는지.

전체 87개 통과(기존 82개 + 신규 5개, 모두 mock·실제 캐시 데이터 기반, API
미호출). 기존 `parser/` 테스트 40개(37 통과 + 3 skip)도 무변경 재확인.

**실제 캐시 데이터로 직접 재현**: 위 두 캐시 파일에서 읽은 실제 LLM
응답(`start_row=3`)을 실제 추출된 `매출_세부_IFRS15後` 시트에 대해
`validate_plan`으로 재검증한 결과, 새 메시지
`"table start_row (3) must be greater than header_end_row (4): ..."`가
정확히 출력됨을 확인(API 호출 없음, 기존 캐시 파일과 로컬 추출만 사용).

---

## 구조 해석 엔진 보강 — 시트 범위 제한·블록 분할·타임아웃 축소 재시도

**배경**: `--engine structure`(기본 엔진)로 실제 실행했을 때 `답안지_IFRS15後`
시트(수식 10,425개, 요청 크기 46.6KB)에서 240초 타임아웃이 연속 2회(같은
요청을 그대로 재시도) 모두 실패했다. 원인은 (a) 자동계산 시트까지 구조
해석 대상에 포함돼 있었고, (b) 시트 하나를 통째로 1콜로 보내 큰 시트가
분할되지 않았고, (c) 타임아웃이 나도 같은 크기 요청을 그대로 재시도해
똑같이 실패하는 구조였기 때문이다.

### 바꾼 파일과 함수

| 파일 | 함수/항목 | 바뀐 내용 |
|---|---|---|
| `src/excel_parser_ai/structured.py` | `_formula_ratio` (신규) | 시트의 `formula_count / content_cell_count`를 셀에서 직접 계산 |
| 〃 | `infer_structure` | 비율이 `FORMULA_SKIP_RATIO`(0.3) 이상이면 해당 시트를 LLM 호출 없이 스킵하고 `skipped_sheets`에 사유·비율 기록 |
| 〃 | `_recompute_sum_crosscheck` (신규) | 스킵된 시트의 `SUM(범위)` 수식만(단일 연속 범위, 동일/다른 시트) 원본 저장값 합계와 캐시된 계산값을 대조. 레코드 값은 바꾸지 않고 진단만 남김 |
| 〃 | `apply_plans` | `formula_crosschecks` 인자 추가 — 대조 결과를 `FORMULA_RECOMPUTE_MATCH/MISMATCH/...` 진단으로만 첨부 |
| 〃 | `structure_context` | `row_range`/`header_row_range` 인자 추가 — 큰 시트를 행 범위 블록으로 쪼갰을 때 블록별 맥락과 공통 헤더 맥락을 구성 |
| 〃 | `validate_plan` | `row_range` 인자 추가 — 블록 하나의 선언 범위·커버리지만 검사 |
| 〃 | `_row_blocks`, `_header_row_range`, `_merge_block_plans` (신규) | 데이터 행을 100행 단위 블록으로 나누고, 블록별 응답을 시트 하나의 계획으로 병합 |
| 〃 | `_resolve_sheet_plan`, `_plan_from_blocks` (신규) | 큰 시트는 처음부터 블록으로, 작은 시트는 전체 1콜 시도 후 타임아웃 시 블록으로 축소 재시도 |
| `src/excel_parser_ai/client.py` | `ModelError` | `connection_error` 플래그 추가(타임아웃/연결 오류와 HTTP 오류를 구분) |
| 〃 | `ResponsesClient.ask`/`_request` | 호출별 `max_connection_attempts` 오버라이드 추가(기본 동작은 그대로, 명시적으로 넘길 때만 적용) |
| `src/excel_parser_ai/engine.py` | `infer` | `max_connection_attempts`를 선택적으로 `client.ask`에 전달(값이 없으면 기존 테스트 더블과 100% 호환) |
| `src/excel_parser_ai/pipeline.py` | `run` | `summary.json`의 파일별 항목에 `structure_interpretation_skipped_sheets`, `structure_timeout_fallbacks` 추가 |

### 전후 비교 (실제 `01_F10_사업부1_재무제표.xlsx` 파일, mock 호출로 측정)

| 항목 | 이전 | 이후 |
|---|---|---|
| `답안지_IFRS15後` 구조 해석 호출 | 1회(실패, 240초×2) | **0회** (자동계산 시트로 스킵, 수식 비율 0.955) |
| `예산코드`(데이터 행 314개) 구조 해석 호출 | 1회(시트 전체 요청) | **4회**(100행 단위 블록) |
| 작은 시트(`매출_세부`, `OpEx_세부`) | 1회 전체 요청 | 1회 전체 요청(변화 없음 — 150행 이하는 그대로 1콜) |
| 타임아웃 시 재시도 | 같은 요청 그대로 2회 | 전체 요청 1회 시도 후 타임아웃이면 블록 단위로 축소해 재시도(같은 요청 반복 안 함) |
| 자동계산 시트의 수식 캐시 | 전혀 다루지 않음 | `SUM(범위)` 수식 2,152건을 원본 합계와 대조(전부 일치 확인, mock 아님 — 실제 파일의 실제 캐시값) |

### mock 테스트

`tests/test_structured.py`에 3개 범주, 6개 테스트 추가(아래 3가지를 모두
커버):
- `LargeSheetBlockSplittingTests` — 큰 시트가 여러 블록 요청으로 나뉘는지, 작은 시트는 그대로 1콜인지.
- `FormulaHeavySheetSkipTests` — 수식 비율이 높은 시트가 LLM 호출 없이 스킵되는지, `SUM` 수식의 일치/불일치가 진단으로만 남고 레코드 값은 바뀌지 않는지.
- `TimeoutFallbackTests` — 타임아웃(연결 오류)이 같은 요청을 반복하지 않고 블록으로 축소되는지, HTTP 오류는 이 축소 로직을 타지 않는지.

전체 82개 통과(기존 76개 + 신규 6개, 모두 mock). 기존 `parser/` 테스트
40개(37 통과 + 3 skip)도 무변경 재확인.

실제 파일(`01_F10_사업부1_재무제표.xlsx`)에 대해 스크립트 기반 mock
클라이언트로(API 호출 없음) `infer_structure`를 직접 실행해 확인: 시트 9개
전부 계획이 생성됨(스킵 포함), 실패 0건, `답안지_IFRS15後`는 스킵되고
그 수식 2,152건의 재계산 대조가 전부 일치. **이 결과는 로직이 실제 크기의
워크북에서 멈추지 않고 동작함을 보인 것이며, 실제 LLM 구조 해석의 성공·
정확도를 의미하지 않는다** — 실제 API 테스트는 수행하지 않았다.

---

*(아래부터는 이전 셀별 엔진 단계의 과거 기록)*

## 이 단계에서 한 일

- 기존 `parser/`는 바이트 단위로 수정하지 않음. 모든 변경은 `parser_ai/` 안에서만 수행.
- 아키텍처 변경: 기존에는 LLM이 셀 포인터만 반환하고 Python이
  `normalize_number()`로 숫자를 정규화했다. 이 단계에서 이를 바꿔,
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

## 이 단계에서 아직 안 된 것

- 이 단계 시점에는 실제 OpenAI API 호출 0회. **실제 LLM 파싱 품질·정확도는
  이 단계에서는 검증되지 않음.**
- 성능 비교·정답표 평가·마스터 집계·HTML 비교 화면은 범위에서 제외.
- 소스 GitHub push/커밋/PR 없음(이 단계 시점 기준).

## 이후 진행 상황

이 문서 이후에 실제로 진행된 OpenAI API 테스트 결과, 구조 해석 기본 엔진
도입, 테스트 현황은 모두 [README.md](README.md)에 정리되어 있다.
