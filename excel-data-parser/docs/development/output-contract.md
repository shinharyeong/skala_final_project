# 로컬 파서 출력 계약 1.0.0

## 범위와 책임

현재 계약은 로컬 Python 파서의 결과 파일과 메모리 반환값을 정의합니다. DB 제품·테이블·웹 API와 독립적입니다. 데이터 소비자는 이 계약의 출처·타입·확인 상태를 기준으로 결과를 해석합니다.

기본 실행은 읽기 전용이며, 저장을 요청한 작업에서 `--write-output`을 명시하면 `output/` 아래 새 실행 폴더를 만듭니다. 원본과 결과 파일은 저장소에 포함하지 않습니다. 기존 원본·결과를 덮어쓰지 않습니다.

## 실행별 파일

| 파일 | 단위와 내용 |
| --- | --- |
| `raw-cells.jsonl.gz` | 행마다 원본 내용 셀 또는 메모가 있는 셀 1개. 모든 시트를 포함 |
| `standard-records.jsonl.gz` | 행마다 숫자·수식 또는 확인된 재무 필드의 표준 레코드 1개 |
| `workbooks.jsonl.gz` | 행마다 파일 식별·시트 구조·빈 구간·스타일·공유 문자열·패키지 메타데이터 |
| `diagnostics.jsonl.gz` | 파일 처리·의미 미확정 등의 출처 연결 진단 |
| `summary.json` | 파일·시트별 원본 추출·표준 매핑·정규화 상태와 개수 |
| `validation.json` | 독립 OOXML 대조, 출처 연결, 원본 해시, 원본 마스터 비교 |
| `normalization-observations.json` | 실제 적용 rule_id별 개수와 원본·정규화·출처 예시 |
| `gap-relationships.json` | 영업이익·Contingency·제외 영업이익의 원본 연결과 미확정 이유 |

JSON은 UTF-8, JSONL은 줄마다 독립 객체입니다. gzip은 중복된 출처·헤더·스타일 구조의 크기를 줄이기 위한 압축이며, 내용을 생략하지 않습니다. 빈 입력 영역 전체를 레코드로 만들지 않습니다. 새 실행 폴더만 생성하고 기존 결과는 덮어쓰지 않습니다.

## 원본 식별과 복원

`workbooks.file.file_id`는 원본 파일 바이트의 SHA-256입니다. `source_id`는 출처 경로와 SHA-256을 결합한 실행 내 파일 발생 식별자입니다. 동일 바이트의 다른 경로를 구분합니다. 셀 `record_id`는 `source_id/sheet_id/address`입니다.

| 원본 셀 필드 | 타입 | 의미 |
| --- | --- | --- |
| `record_id`, `file_id`, `sheet_id`, `sheet`, `address` | string | 파일·시트·셀 연결 |
| `row`, `column_index` / `column` | integer / string | 1부터 시작하는 행·열 번호와 열 문자 |
| `value` | string / boolean / null | 저장값. 숫자는 원본 십진 문자열을 유지 |
| `data_type` | string | number/string/boolean/error/date/blank |
| `storage_type`, `raw_lexeme`, `xml_attributes` | string/null, string/null, object | 원본 OOXML 타입·v 내용·셀 속성 |
| `formula` | object / null | 수식 원문·속성·공유 기준·캐시 상태 |
| `cached_value` | string / boolean / null | 수식 셀의 저장된 계산 결과. 재계산 결과가 아님 |
| `style_id`, `number_format` | integer/null, string/null | 스타일 표 참조와 숫자 표시 형식 |
| `row_hidden`, `column_hidden`, `sheet_hidden` | boolean | 숨김 상태 |
| `merge` | object / null | 병합 범위와 기준 셀 |
| `comments` | array | 원본 메모 텍스트·작성자·XML |
| `has_content` | boolean | XML에 v/f/is 내용 요소가 있는지 여부 |

공유 문자열의 숫자 같은 텍스트는 `data_type=string`입니다. XML 숫자의 원문은 `raw_lexeme`와 `value`에 보존합니다. 날짜 일련번호는 날짜 서식이 확인되면 `data_type=date`로 구분하지만 원본 체계·값을 바꾸지 않습니다. ISO 날짜도 문자열로 유지합니다.

`workbooks.sheets[].blank_spans`는 행, 시작 열, 끝 열, 셀 속성으로 동일 스타일의 연속 빈 XML 셀을 나타냅니다. `cells`의 원본 주소와 이 구간을 합치면 원본 XML의 c 주소 목록을 복원할 수 있습니다. 병합 헤더를 해석해도 빈 원본 셀에 금액을 복제하지 않습니다. 메모만 있는 셀은 별도 원본 레코드이며 `has_content=false`입니다.

스타일 전체 XML에는 글꼴·색상·채움·정렬·숫자 형식·조건부 차등 서식이 남습니다. 시트의 `structure_xml`, `conditional_formatting`에는 범위·조건·아이콘 등의 원문이 남고 `evaluated=false`입니다. 그려진 표시값이나 조건 충족 여부를 새로 계산했다고 주장하지 않습니다. 등록된 모든 기본 숫자 서식을 렌더링하지는 않으며, 모르는 형식 ID는 원본 스타일 표에 보존합니다.

## 수식과 캐시

- `formula.text`는 원본 수식 앞에 `=`를 붙인 표현입니다. 수식은 계산하지 않습니다.
- 공유 수식의 종속 셀은 원문이 없을 수 있습니다. 원본 속성 및 `shared_base`로 기준 셀과 수식을 연결하며 상대 참조를 임의 전개하지 않습니다.
- `cached_status=stored_unverified`는 파일에 저장된 결과가 있다는 뜻입니다. 최신 계산 여부나 수식 전체의 정확성을 보장하지 않습니다.
- 캐시가 없거나 빈 v이면 `missing`, 오류 타입이면 `error`입니다. 정규화 값을 0으로 채우지 않습니다.
- 원본 외부 참조와 정의 이름·계산 설정·관계는 보존합니다. 외부 파일에 접속하거나 계산 엔진을 새로 만들지 않습니다.

## 표준 레코드

| 필드 | 타입 | 의미 |
| --- | --- | --- |
| `contract_version`, `record_id` | string | 출력 계약 버전과 표준 레코드 식별자 |
| `source` | object | raw_record_id, file_id, source_id, path, sheet, cell, workbook_role |
| `organization` | object | 코드·원래 조직명·범위·출처. 불명확하면 null/unknown |
| `item` | object | 원본 항목 코드·이름·헤더 경로·근거·GAP metric_code |
| `dimensions` | object | 원본에 있는 팀·과제·계정 등 추가 구분 |
| `period` | object | 연도·월·분기·month/annual/unknown 집계 범위·시나리오·헤더 출처 |
| `role` | object | 상세·소계·합계·비율·비재무 등 역할과 집계 가능 여부 |
| `value` | object | 원문·Decimal 정규화 값·단위·통화·배율·규칙·확인 상태 |
| `submission` | object | 제출 월·버전 원문·월별 회차·개정번호·확인 상태 |
| `mapping` | object | mapped/partial/unmapped, 적용 규칙과 미확정 사유 |
| `formula`, `raw_value`, `data_type`, `style_reference` | object/null, any, string, object | 원본 값·수식·스타일과 연결 |

`mapping.status=mapped`는 확인된 재무 의미를 연결했다는 뜻입니다. 제출본의 승인이나 마스터 포함을 뜻하지 않습니다. `value.status`는 confirmed/unresolved/not_numeric/missing으로 별도 관리합니다. 표준 항목으로 못 해석한 숫자·수식도 출처 레코드와 원본 추출 결과를 남깁니다. 제목·헤더·일반 텍스트는 원본 추출에서 보존하며 모두 재무 레코드로 만들지는 않습니다.

`source.workbook_role=consolidated_reference`는 원본 전체 집계 참고 파일입니다. 개별 조직 파일과 함께 합산하지 않습니다. 원본 항목 코드가 없는 경우 내부 GAP 식별자는 `item.metric_code`로 분리하며 코드가 원본에 있었다고 가장하지 않습니다.

현재 모든 레코드는 안전한 자동 집계 정책이 미정이므로 `role.aggregatable=false`로 제공합니다. 상세·소계·합계, 월별·연간·비율, 원본 참고 마스터와 조직 제출본, 여러 제출 버전을 일괄 합산하지 않습니다.

## 숫자와 단위 타입

내부 정규화는 `decimal.Decimal`입니다. `normalized_value`, `multiplier`는 JSON에서 **십진 문자열**로 직렬화합니다. 합성 예로 천원 값 `100`은 원문을 유지하고 정규화 원화 문자열 `100000`로 연결합니다. API/DB 소비자는 float로 바꾸지 말고 decimal 또는 DB의 정밀 십진 타입으로 읽어야 합니다.

원본 숫자의 부호·0은 저장값이 기준입니다. 빨강·괄호·삼각형·아이콘으로 부호를 다시 붙이지 않습니다. 숫자로 저장된 비율은 % 배율을 재적용하지 않습니다. 명시 단위와 서식 표시 배율은 구분하며, 화면 표시용 쉼표 배율을 실제 값에 곱하지 않습니다. 원본이 이미 가진 Excel 정밀도보다 더 정확한 값을 복원했다고 주장하지 않습니다.

문자열 숫자 규칙과 실제 적용 근거는 [정규화 규칙](normalization-rules.md)을 따릅니다. 명확하지 않은 기호·로케일·단위 충돌은 null 값과 unresolved 상태로 남깁니다.

## 메타데이터와 버전

`--metadata`는 다음 구조의 JSON을 읽습니다. 아래는 형식 예시이며 해당 파일을 자동 생성하지 않습니다.

```json
{
  "defaults": {"submission_month": "2026-11", "submission_version": "V1.01"},
  "files": {"원본파일명.xlsx": {"organization_code": "명시할 코드"}}
}
```

파일별 키는 원본 출처 경로를 우선하고 파일명을 보조로 사용합니다. 같은 이름의 다른 경로는 전체 출처 경로로 구분합니다. 조직 메타데이터는 원본에서 조직을 확인하지 못했을 때만 사용하며, 원본에서 확인한 조직을 덮어쓰지 않습니다.

제출 월은 `period.month`와 별개입니다. 버전 원문은 문자열로 유지하고 명확한 V1/V1.01 형태일 때만 회차·개정번호를 정수로 분해합니다. 원본에 없는 제출 월·버전·승인 상태는 만들지 않습니다. 조직 제출본 버전과 공식 마스터 버전은 다른 개념이며 이번 파서는 버전을 발급하거나 승인·마스터 발행을 하지 않습니다. 파일 목록이 전 사업부 승인 완료 목록이라는 가정도 하지 않습니다.

## 성공과 한계

`raw_extraction_complete`는 전체 입력 파일·시트·XML 내용과 희소 구조의 독립 대조 및 원본 해시 보존을 뜻합니다. `validation_passed`는 출처 연결과 적용 가능한 마스터 비교도 포함합니다. `semantic_interpretation_complete=false`와 `formula_recalculation_performed=false`를 별도로 보고합니다.

의미 매핑 단계가 실패해도 이미 추출한 원본은 보존하고 매핑 실패를 별도 진단합니다. 원본 추출 실패 파일은 누락을 숨기지 않고 failures에 기록합니다. 모든 값의 업무 의미 확정·계산 검증 완료와 전체 원본 추출 완료를 혼동하지 않습니다.
