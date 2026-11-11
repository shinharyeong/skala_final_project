# LLM 구조 해석 기반 엑셀 파서

비교 대상은 **룰베이스 파서 vs LLM 구조 해석 기반 파서**다. LLM은 워크북의 표·헤더·항목·계층·기간·시나리오·단위·예외를 해석해 선언형 범위 규칙을 반환한다. Python은 그 규칙을 **현재 파일의 원본 셀 전체**에 적용하고 `Decimal` 기반으로 숫자·부호·단위를 정규화한다. 일반 숫자 셀마다 LLM이 금액과 메타데이터를 반복 출력하지 않는다.

기존 `parser/`의 원본 추출기·숫자 정규화기·출처 검증기를 재사용한다. 룰베이스의 의미 매핑 결과나 별도 정답 마스터는 LLM에 보내지 않는다. `parser/`와 원본 Excel은 수정하지 않는다. 백엔드·프론트·DB·Gap 계산·공식 마스터 발급은 범위 밖이다.

## 처리 단계

1. XLSX의 모든 원본 셀, 수식 원문, 마지막 저장 계산값, 서식, 병합, 숨김, 메모, 출처를 보존하고 원본과 대조한다.
2. 숫자·수식 후보가 있는 시트마다 구조 입력을 만든다. 모든 행을 스캔하되 실제 숫자값은 자료형·열 범위·스타일별 대표 사례만 모델에 보낸다. 라벨과 구조 변화, 병합, 서식, 반복 헤더 단서를 포함한다. 큰 참조 시트는 대표 행과 변화 지점을 보내며, 생략된 행은 자동 확정하지 않는다.
3. LLM이 엄격한 JSON 스키마로 표·헤더 범위, 조직과 항목 위치, 기간 열 범위, 행 역할, 원본 단위와 적용 범위, 제외·예외 영역을 선언한다. 정규화 방식은 `amount_to_won`·`ratio_fraction`·`quantity_native`·`unresolved` 중 하나다. Python 코드나 임의 표현식은 받지 않으며 `eval`/`exec`하지 않는다.
4. 존재하지 않는 근거 셀, 표·행·기간·단위 범위의 중복과 빈틈을 검사한다. 유효한 규칙만 현재 파일의 후보 셀에 적용한다. 규칙 밖과 모호한 영역은 `unmapped` 또는 `partial`로 남긴다.
5. 표준 숫자 변환은 Python의 `Decimal` 정규화기가 수행한다. 수식은 새로 계산하지 않는다. 수식 캐시가 없거나 오류면 정규화값을 만들지 않는다.

기본 구조 엔진은 숫자·수식 후보가 없는 시트를 로컬에서 보존하고, 후보가 있는 시트마다 구조 요청을 한 번 보낸다. 유효하지 않은 응답은 근거/범위 오류를 붙여 한 번 복구 요청할 수 있다. `--max-exception-calls`는 여러 예외 영역을 묶은 별도 요청의 상한이며 기본 0이다. 미해결 항목만을 이유로 반복 호출하지 않는다. HTTP 520을 포함한 일시 서버 오류의 재시도는 최대 4회, 연결 오류는 최대 2회로 제한된다.

## 실제로 실행한 테스트와 검증되지 않은 범위

**로컬 mock 테스트:** `PYTHONPATH=parser_ai/src python3 -B -m unittest discover -s parser_ai/tests -v`로
76개 전부 통과했다(구조 엔진 14개, 셀별 엔진과 공통 API/캐시/연결 오류 처리 62개).
그라운딩·좌표 커버리지·구조 핑거프린트 재사용·예외 처리 상한·API 실패 시 캐시 보존 등을
가짜 응답으로 검증한다. **mock 통과는 실제 LLM 해석의 정확도나 성공을 보증하지 않는다.**

**실제 OpenAI API 테스트(가장 최근 실행):** 더미 조직 1개 파일(`01_F10_사업부1_재무제표.xlsx`)을
기본 엔진(`structure`)으로 실제 실행한 결과, **첫 번째 시트(`답안지_IFRS15後`)의 구조 해석 요청이
240초 타임아웃을 연속 2회(연결 재시도 한도) 기록하며 실패**했다. 성공한 구조 해석은 0개이며,
해당 파일은 `llm_inference_completed=false`로 종료됐다(`summary.json`의
`error_type: "PartialStructureFailure"` 참고). **원본 추출(raw-cells/workbooks, 원본 바이트 무변경
검증)은 성공했으나, 구조 해석의 성공 여부·정확도·룰베이스 파서 대비 성능 우열은 이번 실행에서
전혀 검증되지 않았다.** 더 작은 입력이나 다른 timeout 설정으로의 재시도는 이번 작업 범위에서
수행하지 않았다.

별도로 진행했던 이전 셀별 엔진(`--engine legacy-cell`) 실제 실행에서는 레이아웃 해석 9개 시트는
모두 성공했고, 행 배치 152개 중 32개를 처리한 뒤 서버 측 일시 오류(HTTP 520)로 중단됐다(성공한
32개 배치의 레코드는 보존됨). 원본 대조 결과 모델이 읽은 원본값 자체는 정확했으나, 같은 표의
이웃한 셀들 사이에서 단위 환산(원/억원)을 일관되게 적용하지 못한 사례를 발견했다 — Python은
이 불일치를 교정하지 않고 진단으로만 남겼다. 이 역시 정확도 검증이 아니라 관찰 결과다.

## 동일 양식 재사용

`parser_ai/cache/structure_templates/`에 구조 계획과 지문을 저장한다. 지문에는 모든 시트의 셀 좌표·자료형·서식·병합·숨김·문자 라벨이 포함되고 실제 금액은 포함되지 않는다. 프롬프트·스키마 계약도 함께 검사한다. 선언된 조직·연도·시나리오 근거 셀은 파일마다 다시 읽도록 동적 위치로 취급한다. 지문이나 계약이 일치하지 않으면 재사용하지 않고 해당 파일을 새로 해석한다. 파일명 유사성은 판단에 쓰지 않는다. 새 파일의 조직명·금액·연도·시나리오는 언제나 새 파일에서 읽는다. 큰 시트에서 구조 입력에 싣지 못한 행은 재사용 여부와 관계없이 미해결이다.

이 지문은 보수적이다. 조직명·연도가 선언된 근거 셀 이외의 임의 위치에서도 달라지면 재사용이 거부될 수 있다. 이는 다른 구조를 잘못 재사용하는 것보다 안전하다. 저장된 과거 셀별 캐시는 삭제하지 않으며 새 구조 계획 캐시와 키가 다르다.

## 환경과 명령

Python 3.11 이상과 표준 라이브러리만 필요하다. API 키는 `parser_ai/.env`의 `OPENAI_API_KEY`에서 읽는다. 키와 Authorization 헤더는 출력하지 않는다. `.env`, `.venv`, 입력 데이터, 캐시, 실행 산출물은 Git에서 제외한다.

```bash
# API 없이 원본 추출만 확인
python3 -B parser_ai/run.py prepare parser/Data/dummy/organizations --max-files 1

# 첫 번째 파일의 구조 해석 실제 API 시험 (현재 작업에서는 실행하지 않음)
parser_ai/.venv/bin/python -B parser_ai/run.py run \
  parser/Data/dummy/organizations/01_F10_* \
  --engine structure --review off --max-calls 12 --max-exception-calls 0

# 원본 출처 연결 CSV
python3 -B parser_ai/audit.py parser_ai/output/<실행 ID>

# Mock/로컬 테스트
PYTHONPATH=parser_ai/src python3 -B -m unittest discover -s parser_ai/tests -v
```

첫 번째 조직 파일을 로컬에서 추출해 센 결과, 9개 시트 중 숫자·수식 후보가 있는 시트는 4개다. 유효한 응답 기준 구조 요청 **4회**가 예상되며, 근거/범위 복구 요청이 필요한 경우 최대 4회가 더 발생할 수 있다. 예외 호출은 위 명령에서 0회로 제한했다. `--max-calls 12`는 비용 상한액이 아니라 논리 요청 상한이다. 캐시 일치 시 API 요청은 줄어든다.

기존 152개 행 배치 방식은 구조 외에도 매 숫자 셀의 메타데이터와 정규화값을 반복 생성했다. 첫 파일의 이전 레이아웃 계획으로 로컬에서 요청 맥락을 재구성하면 행 배치 입력 약 6,121,010바이트(152건), 레이아웃 입력 약 436,711바이트(9건)다. 새 구조 입력은 후보 시트 순서대로 37,154·17,744·20,166·17,324바이트, 합계 **92,388바이트(4건)**다. 두 값 모두 공백 없는 로컬 JSON 맥락 크기이며 새 구조 입력은 이전 전체 맥락보다 약 **98.6% 작다**. 시스템 프롬프트·스키마·API 포장 크기와 실제 토큰 사용량은 제외한 비교다. 출력도 셀별 정규화값 대신 표·행·기간 범위를 반환하도록 설계했지만 감소량, 속도, 정확도는 실제 API 실행 전 확정하지 않는다.

## 결과 해석

기존 출력 계약 1.0.0의 `raw-cells`, `standard-records`, `workbooks`, `diagnostics`, `validation`, `summary`, `ai-semantics`, `inference` 등을 `parser_ai/output/<실행 ID>/`에 저장한다. `mapping.rule_ids`에 `LLM_STRUCTURE`, `PYTHON_RULE_APPLICATION` 또는 `LLM_EXCEPTION_REGION`을 남기고, `value.rule_ids`에 `PYTHON_DECIMAL_NORMALIZATION`을 남긴다. `ai-semantics`에는 선언된 구조 계획과 재사용 시트가 기록된다.

`summary.raw_extraction_complete`는 원본 추출과 무변경 검증, `llm_inference_completed`는 관련 시트의 구조 요청 완료, `validation_passed`는 자동 출처·구조 검사 통과를 뜻한다. `semantic_interpretation_complete`와 독립 정답 정확도는 별개다. `mapped`/`confirmed`만으로 정답이라고 단정하지 않는다. 수식의 저장 계산값은 최신성을 보증하지 않는다. `source-audit.csv`는 원본 시트·셀·값과 해석한 조직·항목·기간·시나리오·금액·단위·진단을 나란히 보여준다.

`inference.json`의 `logical_attempted_calls`는 캐시 제외 시도한 논리 요청 수, `http_post_attempts`는 재전송을 포함한 HTTP POST 시도 수다. 기존 호환 필드 `api_calls`는 이벤트로 기록된 논리 요청 수다. 응답이 없는 요청의 토큰 사용량은 알 수 없으며 0으로 단정하지 않는다.

## 이전 셀별 실험

`--engine legacy-cell`은 이전 결과 재현용이다. 이 모드만 `--batch-size`, `--max-batch-cells`, `--max-batch-bytes`, `--max-new-row-batches`를 사용한다. 이전 셀별 모드의 성공 캐시와 실행 결과는 보존되어 있다. 새 비교와 이후 실행의 기본값은 `structure`다.

## 더 자세한 문서

- [실행 중 겪은 문제와 조치](docs/development/troubleshooting.md)
- [구현·검증 상태와 변경 기록](IMPLEMENTATION_STATUS.md)
