# 실행 중 겪은 문제와 조치

`--engine structure`(기본 엔진) 실행 중 실제로 겪은 문제를 증상/원인/조치
순서로 정리한다. 각 항목은 가능하면 날짜와 관련 커밋·실행(output) ID를
남긴다. [IMPLEMENTATION_STATUS.md](../../IMPLEMENTATION_STATUS.md)에 더 긴
설명(파일/함수별 변경, 전후 비교, 테스트 목록)이 있는 항목은 그쪽을 함께
참고.

## a. 큰 시트 통째로 요청해서 240초 타임아웃

**날짜**: 2026-11-11. 관련 실행: `output/20261111-225706-440465b6`.

**증상**: `01_F10_사업부1_재무제표.xlsx`의 `답안지_IFRS15後` 시트(수식
10,425개, 요청 크기 약 46.6KB)를 구조 해석할 때 240초 타임아웃이 발생했고,
재시도도 똑같은 요청으로 다시 보내 또 240초 타임아웃으로 실패했다
(`structure_failures`에 `오류 종류=TimeoutError`로 기록).

**원인**: 시트 하나를 행 수와 무관하게 통째로 1번의 구조 해석 호출로
보냈다. 데이터 행이 많은 시트는 요청이 커져 응답 시간이 늘어나고, 타임아웃이
나도 재시도가 같은 크기의 요청을 그대로 반복했다.

**조치**: `structured.py`에 `LARGE_SHEET_ROW_THRESHOLD`(150행)를 넘는
시트는 `ROW_BLOCK_SIZE`(100행) 단위 블록으로 나눠 여러 번 요청하도록 바꿈
(`_row_blocks`, `_resolve_sheet_plan`, `_plan_from_blocks`). 블록마다 헤더
행 범위(`header_row_range`)를 공통으로 포함해 헤더/단위 맥락이 끊기지
않게 함. 자세한 내용은 IMPLEMENTATION_STATUS.md의 "구조 해석 엔진 보강"
섹션.

## b. 자동계산 시트(수식 비율 높음)까지 LLM 호출 대상에 포함

**날짜**: 2026-11-11. 관련 실행: `output/20261111-235354-17cc6bdc`,
`output/20261112-001725-0e53c13f`(수정 후 실제 확인).

**증상**: `답안지_IFRS15後`처럼 거의 모든 셀이 수식인 "자동계산" 시트도
다른 시트와 똑같이 구조 해석 LLM 호출 대상이었다. 이런 시트는 수식 자체가
이미 계산 결과를 담고 있어 LLM이 구조를 해석해도 얻는 정보가 적은데, 가장
크고 느린 요청이 되어 (a)의 타임아웃을 유발하는 주범이었다.

**원인**: 시트를 LLM 호출 대상에 포함할지 판단하는 기준이 없었다 — 수식
비율과 무관하게 숫자·수식 후보가 있으면 전부 호출했다.

**조치**: `_formula_ratio(sheet) = formula_count / content_cell_count`를
계산해 `FORMULA_SKIP_RATIO`(0.3) 이상이면 "자동계산 시트"로 분류하고 LLM
호출 자체를 생략. 대신 그 시트의 `SUM(범위)` 수식만 원본 셀 합계와 캐시된
계산값을 재계산 대조(`_recompute_sum_crosscheck`)해서 `FORMULA_RECOMPUTE_*`
진단으로만 남김(레코드 값은 바꾸지 않음). `summary.json`의
`structure_interpretation_skipped_sheets`에 어떤 시트를 왜(비율 얼마) 스킵
했는지 기록. **실제 확인**: `output/20261112-001725-0e53c13f`(API 호출,
실제 run)에서 `답안지_IFRS15後`가 `formula_ratio: 0.9551`로 스킵되고 LLM
호출 없이 처리됨을 확인.

## c. 시트 1개 실패 시 전체 파일의 나머지 시트가 시도조차 안 됨

**날짜**: 2026-11-11~12. 관련 실행: `output/20261111-235354-17cc6bdc`
(문제 재현) → 수정 후 `output/20261112-001725-0e53c13f`(실제 확인).

**증상**: `매출_세부_IFRS15後` 시트 하나의 구조 해석이 최종 실패하자(원인은
d 참고), 워크북에 시트가 9개 있었는데도 `inference.json`에 호출 이벤트가
`sheet_structure_v1`/`sheet_structure_v1_repair` 2건뿐이었다 — 나머지
7개 시트는 LLM 호출 자체를 시도하지 않았다.

**원인**: `infer_structure()`의 메인 루프가 한 시트의 구조 해석이
실패하면(`except Exception`) 즉시 `break`해서 루프를 끝내도록 되어 있었다.
"이미 성공한 계획·캐시는 보존하고 더 이상 호출하지 않는다"는 비용 안전
설계였지만, 한 시트의 실패가 다른 시트와 무관한 경우까지 전체를 멈추는
것은 비효율적이었다.

**조치**: 시트별 실패를 `failures`에 기록만 하고 다음 시트로 계속
진행하도록(`continue`) 바꿈. 단 그 시점에 `client.call_count >=
client.max_calls`(파일 전체 호출 상한 소진)이면 남은 시트도 호출할 수 없는
것이 확정이므로 그때만 `break`. 파일 전체 호출 예산은 그대로 지킨다.

## d. "매출_세부_IFRS15後"에서 표 범위(rectangle) 오판

**날짜**: 2026-11-11~12. 관련 실행: `output/20261111-235354-17cc6bdc`
(실패) → `output/20261112-001725-0e53c13f`(수정 후 실제 성공 확인).

**증상**: `validate_plan` 검증이 `"invalid table/header rectangle"`로
실패했고, repair 재시도도 완전히 동일한 응답을 내 똑같이 실패했다.

**원인**: 실제 캐시된 LLM 응답(1차·repair 모두 동일)을 확인한 결과,
모델이 `start_row`를 `header_end_row + 1`(올바른 값, 5)이 아니라
`header_start_row`와 같은 값(3)으로 선언했다. 이 시트는 A~F열에 2행짜리
병합 헤더(3~4행)가 있고, 헤더 바로 위인 2행에 안내/예시성 숫자가 있는데
그 숫자 분포(`numeric_span=[8,62]`, `numeric_count=52`)가 실제 첫 데이터
행(5행)과 완전히 같았다 — 병합 헤더와 이 숫자 분포 일치가 모델의 판단을
흐렸을 가능성이 있다. 기존 `validate_plan`은 이 조건 위반을 generic한
한 문장(`"invalid table/header rectangle"`)으로만 반환해, repair 요청에
무엇이 왜 잘못됐는지 신호를 주지 못했다 — 그래서 repair도 똑같은 실수를
반복했다.

**조치**: `validate_plan`의 rectangle 검증을 조건별로 분리된 구체적
메시지로 바꾸고(예: `"table start_row (3) must be greater than
header_end_row (4): ..."`), `PROMPT`에 `start_row = header_end_row + 1`을
명시하고 병합 헤더·숫자처럼 보이는 안내 행에 속지 말라는 지침을 추가함.
**실제 확인**: `output/20261112-001725-0e53c13f`(API 호출, 실제 run)에서
`매출_세부_IFRS15後`가 더 이상 `failures`에 나타나지 않음(성공). 자세한
원인 분석은 IMPLEMENTATION_STATUS.md의 "rectangle 검증 실패 원인 분석"
섹션.

## e. `.env.example` 키 노출

**날짜**: 2026-11-11, 작업 초반(구조 해석 엔진 도입 이전).

**증상**: `parser_ai/.env.example`에 placeholder가 아닌 실제 사용 가능한
형태의 OpenAI API 키 값이 들어가 있었다. 1차 발견 후 제거했으나 파일이
다시 복원되면서 동일한 키가 재발견되어 2차로 제거했다.

**원인**: `.env.example`은 설계상 커밋·공유를 전제로 하는 템플릿 파일인데,
실제 키 값이 그 안에 그대로 남아 있었다. 이 사건이 있었던 시점에는 이
작업 폴더에 `.git` 저장소 자체가 없었고(ZIP 스냅샷), 이 시점까지 커밋·
푸시·원격 공유가 전혀 없었다 — 따라서 이 키가 GitHub 등 외부로 실제
유출(공개 노출)되었다는 증거는 없다.

**조치**: 두 차례 모두 발견 즉시 값을 제거하고 빈 템플릿
(`OPENAI_API_KEY=`)으로 되돌림. 이후 어떤 코드·문서·대화 로그에도 실제 키
값을 기록하지 않음(길이·접두사 종류 등 메타데이터만 필요 시 언급). 이후
모든 커밋·푸시 전에 `sk-`/`Bearer`/`svcacct` 패턴 grep 스캔을 수행했고,
GitHub `ai_parser` 브랜치의 `.env.example`은 빈 값만 포함함(별도 확인됨).
더 요약된 기록은 [IMPLEMENTATION_STATUS.md](../../IMPLEMENTATION_STATUS.md)의
"보안 사고 기록" 섹션에도 있다.
