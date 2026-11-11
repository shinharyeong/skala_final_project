# 현재 구현 함수 명세서

코드 docstring을 기준으로 실제 함수·메서드 61개의 명세를 정리합니다. 공개·비공개 모두 포함하며, 각 함수 변경 시 이 문서도 같은 변경에서 갱신합니다. 자동 생성되는 dataclass 생성자는 클래스의 Attributes와 생성 규칙 설명을 따릅니다.

## 전체 사용 흐름

```text
CLI main() → pipeline.run_parser()
  ├── paths.ProjectPaths.load() / collect_inputs()
  ├── extraction.extract_workbook() → ooxml 읽기 보조 함수
  ├── mapping.map_workbook() → normalization.normalize_number()
  ├── validation.validate_extraction() / validate_provenance()
  ├── validation.compare_master()
  └── output_dir를 명시한 경우에만 JSON/JSONL 저장
```

외부 Python 코드는 extract_workbook과 map_workbook을 순서대로 호출해 메모리에서 처리할 수 있습니다. 결과 파일 생성은 pipeline에서 분리합니다. Python 실행에는 -B 또는 PYTHONDONTWRITEBYTECODE=1을 사용합니다. 정확한 실행·입출력 형식은 [README](../../README.md)와 [출력 계약](output-contract.md)을 따릅니다.

## 함수별 명세

## __main__.py

### main

정의: [__main__.py](../../src/excel_parser/__main__.py)

서명: `main() -> int`

정적 호출 관계: 외부 호출 또는 동적 호출 가능. 전체 실행 흐름과 호출 코드를 함께 확인합니다..

```text
CLI 조회 또는 파싱·검증을 실행하고 종료 상태를 반환한다.

Args:
    없음. sys.argv에서 --root와 paths/files/parse 명령을 읽는다.
    parse의 inputs는 파일·폴더 목록이며 --write-output은 결과 저장 요청이다.

Returns:
    정상 완료 또는 의미 미확정이 있는 원본 추출 완료는 0.
    parse의 파일 처리·독립 검증 실패는 1. 프로세스 종료는 호출자가 처리한다.

Raises:
    SystemExit: 도움말은 0, 인자 오류는 2, 포착한 OSError/ValueError는 1로
        종료한다. 도움말·인자 오류는 설정 파일을 읽기 전에 처리한다.
    RuntimeError: 경로 해석 등에서 발생하면 이 함수가 포착하지 않고 전파한다.

Side Effects:
    조회는 메타데이터만, parse는 원본 Excel 내용을 읽는다.
    --write-output이 있을 때만 output 하위에 새 실행 결과를 생성한다.
    Data 내부와 기존 출력에는 쓰지 않는다. 콘솔 출력이 있다.
    Python 바이트코드 캐시는 호출 시 -B 또는 환경변수로 별도 차단해야 한다.
```

## extraction.py

### file_fingerprint

정의: [extraction.py](../../src/excel_parser/extraction.py)

서명: `file_fingerprint(path: Path) -> str`

정적 호출 관계: `extraction.extract_workbook`, `pipeline.run_parser`.

```text
원본 바이트의 SHA-256 식별자를 계산한다.

Args:
    path: 읽을 파일 경로.
Returns:
    SHA-256 16진수 문자열.
Raises:
    OSError: 파일 읽기에 실패한 경우.
Side Effects:
    파일을 청크 단위로 읽기만 한다.
```

### _cell_value

정의: [extraction.py](../../src/excel_parser/extraction.py)

서명: `_cell_value(node: ET.Element, shared: list[str], number_format: str | None) -> tuple[object, str, str | None]`

정적 호출 관계: `extraction._extract_sheet`.

```text
셀 저장 타입에 따라 원본 값을 손실 없이 읽는다.

Args:
    node: 원본 c 요소.
    shared: 공유 문자열 표.
    number_format: 숫자 표시 형식.
Returns:
    해석한 값, 자료형, v 요소 원문 튜플. 숫자는 문자열로 유지한다.
Raises:
    ValueError: 공유 문자열 인덱스가 유효하지 않은 경우.
Side Effects:
    없음. 수식을 계산하거나 표시 서식으로 부호를 바꾸지 않는다.
```

### _formula

정의: [extraction.py](../../src/excel_parser/extraction.py)

서명: `_formula(node: ET.Element, value: object, kind: str, shared: dict) -> dict | None`

정적 호출 관계: `extraction._extract_sheet`.

```text
수식과 저장 캐시 상태를 분리하여 보존한다.

Args:
    node: 원본 셀 요소.
    value: 저장 캐시의 해석 값.
    kind: 캐시 자료형.
    shared: 같은 시트의 공유 수식 원본 표.
Returns:
    수식 원문·속성·캐시 상태 또는 수식이 없으면 None.
Raises:
    없음.
Side Effects:
    없음. 공유 수식의 상대 참조를 임의 전개하거나 재계산하지 않는다.
```

### _blank_span

정의: [extraction.py](../../src/excel_parser/extraction.py)

서명: `_blank_span(spans: list[dict], node: ET.Element, row: int, column: int) -> None`

정적 호출 관계: `extraction._extract_sheet`.

```text
동일 스타일의 연속 빈 XML 셀을 구조 구간으로 압축한다.

Args:
    spans: 시트의 빈 셀 구간 목록.
    node: 내용이 없는 원본 c 요소.
    row: 1부터 시작하는 행 번호.
    column: 1부터 시작하는 열 번호.
Returns:
    없음. 전달받은 목록에 구간을 추가하거나 마지막 구간을 확장한다.
Raises:
    없음.
Side Effects:
    메모리 목록만 변경한다. 빈 셀에 값은 채우지 않는다.
```

### _extract_sheet

정의: [extraction.py](../../src/excel_parser/extraction.py)

서명: `_extract_sheet(archive: ZipFile, info: dict, file: dict, shared: list[str], styles: dict) -> dict`

정적 호출 관계: `extraction.extract_workbook`.

```text
한 시트의 내용 셀·숨김·병합·서식·메모와 빈 구조를 읽는다.

Args:
    archive: 읽기 전용 ZIP 패키지.
    info: name, sheet_id, state, part를 포함한 시트 정보.
    file: 원본 파일 식별 정보.
    shared: 공유 문자열 표.
    styles: 읽어 둔 셀 스타일 표.
Returns:
    셀 목록과 구조 정보를 포함한 시트 사전.
Raises:
    ValueError: 셀 주소·공유 문자열 참조가 유효하지 않은 경우.
    ET.ParseError: 시트 XML이 손상된 경우.
Side Effects:
    시트와 메모 XML을 읽기만 한다.
```

### extract_workbook

정의: [extraction.py](../../src/excel_parser/extraction.py)

서명: `extract_workbook(path: Path, root: Path | None=None) -> dict`

정적 호출 관계: `pipeline.run_parser`.

```text
한 xlsx/xlsm 파일의 모든 시트를 읽기 전용으로 추출한다.

Args:
    path: 원본 OOXML 파일 경로.
    root: 출처 경로를 상대 표현할 프로젝트 루트. 없으면 절대 경로를 사용한다.
Returns:
    file, sheets, styles, package_metadata, diagnostics를 가진 추출 결과.
Raises:
    ValueError: 지원하지 않는 형식이나 잘못된 관계·셀 구조인 경우.
    OSError: 원본 파일을 읽을 수 없는 경우.
    BadZipFile: ZIP 패키지가 손상된 경우.
    ET.ParseError: XML 파트가 손상된 경우.
Side Effects:
    원본을 읽기만 한다. Excel 실행·재계산·파일 생성은 하지 않는다.
```

## mapping.py

### _text

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_text(cell: dict | None) -> str`

정적 호출 관계: `mapping._header_at`, `mapping._period`, `mapping._layout`, `mapping._organization_registry`, `mapping._workbook_organization`, `mapping._unit`, `mapping._statement_items`, `mapping._row_context`.

```text
셀의 저장값을 검색용 문자열로 반환한다.

Args:
    cell: 추출 셀 또는 None. 수식은 이미 추출된 캐시값을 사용한다.
Returns:
    값이 없으면 빈 문자열, 그 외에는 저장값의 문자열 표현.
Raises:
    없음. 추출 계약을 만족하는 입력을 전제로 한다.
Side Effects:
    없음. 입력을 변경하거나 수식을 계산하지 않는다.
```

### _ref

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_ref(cell: dict | None) -> dict | None`

정적 호출 관계: `mapping._period`, `mapping._organization_registry`, `mapping._workbook_organization`, `mapping._unit`, `mapping._statement_items`, `mapping._row_context`, `mapping.map_workbook`.

```text
해석 근거 셀과 원본 레코드의 연결 정보를 만든다.

Args:
    cell: 근거 셀 또는 None.
Returns:
    시트·셀·원본 레코드 참조 또는 None.
Raises:
    없음. 추출 계약을 만족하는 입력을 전제로 한다.
Side Effects:
    없음.
```

### _index

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_index(sheet: dict) -> dict`

정적 호출 관계: `mapping.map_workbook`.

```text
행·열별 셀과 주소별 셀의 메모리 색인을 만든다.

Args:
    sheet: extraction의 시트 구조.
Returns:
    rows와 addresses 색인을 포함하는 사전.
Raises:
    KeyError: 필수 행·열·주소가 누락된 잘못된 추출 자료.
    ValueError: 병합 범위 주소가 잘못된 경우.
Side Effects:
    없음. 원본 셀 사전은 읽기 참조로만 사용한다.
```

### _header_at

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_header_at(index: dict, row: int, column: int) -> dict | None`

정적 호출 관계: `mapping._layout`.

```text
병합 헤더를 읽을 때만 기준 셀과 상속 위치를 연결한다.

Args:
    index: 시트 색인과 병합 범위 목록.
    row: 조회할 1부터 시작하는 헤더 행.
    column: 조회할 1부터 시작하는 헤더 열.
Returns:
    직접 셀 또는 상속 근거가 붙은 기준 셀 사본. 없으면 None.
Raises:
    없음. 추출 계약의 유효한 병합 범위를 전제로 한다.
Side Effects:
    없음. 원본 빈 셀에 금액을 복제하지 않고 헤더 조회에만 사용한다.
```

### _scenario

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_scenario(label: str) -> str`

정적 호출 관계: `mapping._period`, `mapping._layout`.

```text
명시된 한국어 값 구분을 계약의 시나리오로 해석한다.

Args:
    label: 원본 헤더 문자열.
Returns:
    actual, plan, estimate, mixed, unknown 중 하나.
Raises:
    없음.
Side Effects:
    없음. 월이나 파일 이름으로 시나리오를 추측하지 않는다.
```

### _year

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_year(label: str) -> int | None`

정적 호출 관계: `mapping._period`, `mapping._layout`.

```text
명시된 연도 또는 이 템플릿의 작은따옴표 연도를 읽는다.

Args:
    label: 2027년 또는 '27년 같은 원본 연도 헤더.
Returns:
    4자리 연도. 명시 연도가 없으면 None.
Raises:
    없음.
Side Effects:
    없음.
```

### _period

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_period(cells: list[dict], *, annual: bool=False) -> dict`

정적 호출 관계: `mapping._layout`, `mapping.map_workbook`.

```text
근거 헤더에서 대상 연월과 기간 집계 방식을 해석한다.

Args:
    cells: 같은 값 열을 설명하는 원본 헤더 셀 목록.
    annual: 단일 복합 헤더의 연간 집계를 명시하는 템플릿 규칙.
Returns:
    기간·시나리오·원본 헤더·근거 셀 사전.
Raises:
    없음. 추출 계약을 만족하는 입력을 전제로 한다.
Side Effects:
    없음. 제출 월과 대상 월을 결합하지 않는다.
```

### _layout

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_layout(sheet: dict, index: dict) -> dict`

정적 호출 관계: `mapping.map_workbook`.

```text
파일명이 아닌 실제 헤더 조합으로 알려진 재무 표를 식별한다.

Args:
    sheet: 시트 구조.
    index: _index 결과.
Returns:
    표 종류, 헤더 행, 값 열별 기간 및 헤더명별 열 색인.
Raises:
    없음. 추출 계약을 만족하는 입력을 전제로 한다.
Side Effects:
    없음. 빈 셀에 원본 값이나 금액을 복제하지 않는다.
```

### _periods_with_cross_evidence

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_periods_with_cross_evidence(layouts: list[dict]) -> None`

정적 호출 관계: `mapping.map_workbook`.

```text
같은 워크북의 명시 월별 시나리오로 혼합 기간의 의미를 보완한다.

Args:
    layouts: 표 레이아웃 목록. 시트별 기간 사전이 갱신된다.
Returns:
    없음.
Raises:
    없음. 내부 생성 레이아웃을 전제로 한다.
Side Effects:
    전달된 레이아웃의 기간만 갱신한다. 원본 workbook은 변경하지 않는다.
```

### _organization_registry

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_organization_registry(indices: list[dict]) -> dict[str, dict]`

정적 호출 관계: `mapping.map_workbook`.

```text
워크북 내 사업부 코드표에서 코드와 이름의 확인 근거를 읽는다.

Args:
    indices: 각 시트의 메모리 색인.
Returns:
    사업부 코드별 이름·근거. 충돌 이름은 확정하지 않는다.
Raises:
    없음. 추출 계약을 만족하는 입력을 전제로 한다.
Side Effects:
    없음. 다른 워크북이나 마스터 파일에 접근하지 않는다.
```

### _workbook_organization

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_workbook_organization(layouts: list[dict], indices: list[dict]) -> dict`

정적 호출 관계: `mapping.map_workbook`.

```text
상세 행의 조직 식별자를 모아 단일 조직과 여러 조직 범위를 구분한다.

Args:
    layouts: 시트별 레이아웃.
    indices: 같은 순서의 시트별 색인.
Returns:
    단일 코드·원래 이름 또는 multi_organization 범위와 원본 근거.
Raises:
    없음. 내부 생성 자료를 전제로 한다.
Side Effects:
    없음. 파일 이름의 조직을 확정값으로 사용하지 않는다.
```

### _metric

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_metric(label: str) -> str | None`

정적 호출 관계: `mapping._statement_items`, `mapping._row_context`, `mapping.map_workbook`.

```text
실제 항목명에서 GAP 관련 지표를 식별한다.

Args:
    label: 표의 항목명. 파일명이나 임의 추정값은 전달하지 않는다.
Returns:
    OP, CONTINGENCY, OP_EX_CONTINGENCY 또는 None.
Raises:
    없음.
Side Effects:
    없음. 지표 값을 계산하거나 조정하지 않는다.
```

### _unit

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_unit(index: dict) -> tuple[str | None, list[dict]]`

정적 호출 관계: `mapping.map_workbook`.

```text
제목 또는 단위 설명에 명시된 기본 금액 단위를 찾는다.

Args:
    index: 현재 시트 색인.
Returns:
    원·천원·백만원·억원 중 하나와 근거 셀 목록. 없으면 None.
Raises:
    없음.
Side Effects:
    없음. 표시 형식의 쉼표를 저장값의 단위로 추측하지 않는다.
```

### _submission

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_submission(metadata: dict) -> dict`

정적 호출 관계: `mapping.map_workbook`.

```text
사용자가 전달한 제출 정보만 보존하고 명확한 버전 표기를 분리한다.

Args:
    metadata: submission_month, submission_version 등의 선택 메타데이터.
Returns:
    버전 원문, 월별 회차, 개정번호, 제출 월과 확인 상태.
Raises:
    ValueError: 버전 원문이 문자열이 아니거나 제출 월 형식이 잘못된 경우.
Side Effects:
    없음. 버전을 발급하거나 승인 여부를 판정하지 않는다.
```

### _statement_items

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_statement_items(index: dict, layout: dict) -> dict[int, dict]`

정적 호출 관계: `mapping.map_workbook`.

```text
답안지의 들여쓰기 열을 계층 항목 경로로 보존한다.

Args:
    index: 답안지 셀 색인.
    layout: financial_statement 레이아웃.
Returns:
    행별 항목명, 계층 경로와 원본 근거.
Raises:
    없음. 내부 생성 자료를 전제로 한다.
Side Effects:
    없음. 원본 빈 셀에는 값을 쓰지 않는다.
```

### _row_context

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_row_context(cell: dict, index: dict, layout: dict, common_org: dict, registry: dict, statement_items: dict) -> dict`

정적 호출 관계: `mapping.map_workbook`.

```text
값 셀에 대응하는 조직·항목·역할을 실제 헤더에서 가져온다.

Args:
    cell: 값 또는 수식의 원본 셀.
    index: 현재 시트 색인.
    layout: 확인된 표 레이아웃.
    common_org: 워크북 상세 행에서 확인한 조직 범위.
    registry: 같은 워크북의 조직 코드표.
    statement_items: 답안지의 행별 항목.
Returns:
    organization, item, dimensions, role, financial을 담은 사전.
Raises:
    없음. 내부 생성 자료를 전제로 한다.
Side Effects:
    없음. 원본이나 공통 조직 사전을 변경하지 않는다.
```

### _gap_relationships

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `_gap_relationships(records: list[dict]) -> list[dict]`

정적 호출 관계: `mapping.map_workbook`.

```text
동일 조직·기간에서 GAP 지표의 원본과 명시 제외값을 연결한다.

Args:
    records: 현재 워크북의 표준 레코드.
Returns:
    값 계산 없이 원본 지표 참조와 관계 확인 상태를 담은 목록.
Raises:
    없음. 내부 생성 표준 레코드를 전제로 한다.
Side Effects:
    없음. 영업이익에 Contingency를 더하거나 빼지 않는다.
```

### map_workbook

정의: [mapping.py](../../src/excel_parser/mapping.py)

서명: `map_workbook(workbook: dict, metadata: dict | None=None) -> dict`

정적 호출 관계: `pipeline.run_parser`.

```text
추출 셀 전체 후보를 출처가 연결된 표준 레코드로 변환한다.

Args:
    workbook: extract_workbook의 원본 보존 결과. 입력은 변경하지 않는다.
    metadata: 선택 조직명·코드, 제출 월·버전. 파일명에서 추측하지 않는다.
Returns:
    records, diagnostics, summary, gap_relationships 사전. Decimal은 저장
    계층에서 문자열로 직렬화한다. 숫자·수식 후보는 미매핑이어도 남는다.
Raises:
    ValueError: 잘못된 제출 메타데이터. 정규화 계약 오류도 전파한다.
    KeyError: 추출 계약 필수 필드가 없는 입력.
Side Effects:
    없음. 파일 접근·수식 계산·DB 작업·결과 저장을 하지 않는다.
```

## normalization.py

### inspect_number_format

정의: [normalization.py](../../src/excel_parser/normalization.py)

서명: `inspect_number_format(number_format: str | None) -> dict`

정적 호출 관계: `mapping.map_workbook`, `normalization._base_result`.

```text
숫자 표시 형식의 구간·조건·색·표시 배율을 값 변경 없이 요약한다.

Args:
    number_format: 원본 Excel 표시 형식. 없으면 None.

Returns:
    원문, 구간별 메타데이터, 백분율·날짜 표시 여부를 담은 사전.

Raises:
    없음. 알 수 없는 형식은 원문을 보존하며 완전한 렌더링을 시도하지 않는다.

Side Effects:
    없음. 파일을 읽거나 쓰지 않는다.
```

### _base_result

정의: [normalization.py](../../src/excel_parser/normalization.py)

서명: `_base_result(value: object, value_kind: str | None, unit: str | None, number_format: str | None) -> dict`

정적 호출 관계: `normalization.normalize_number`.

```text
정규화 상태를 담을 독립된 결과 객체를 만든다.

Args:
    value: 원본 저장값 또는 문자열.
    value_kind: 호출자가 확인한 필드 종류.
    unit: 호출자가 확인한 원본 단위.
    number_format: 원본 표시 형식.

Returns:
    정규화값이 없는 초기 결과 사전.

Raises:
    없음.

Side Effects:
    없음. 입력 객체를 변경하지 않는다.
```

### _stop

정의: [normalization.py](../../src/excel_parser/normalization.py)

서명: `_stop(result: dict, status: str, code: str, message: str) -> dict`

정적 호출 관계: `normalization.normalize_number`.

```text
값을 확정하지 못한 이유와 상태를 결과에 기록한다.

Args:
    result: 현재 호출에서 생성한 결과 사전.
    status: missing, not_numeric, unresolved 중 하나.
    code: 기계 판별에 사용하는 규칙·진단 코드.
    message: 사람이 읽을 수 있는 미확정 사유.

Returns:
    상태와 진단이 추가된 같은 결과 객체.

Raises:
    없음.

Side Effects:
    해당 결과 사전만 변경하며 원본 입력이나 파일은 변경하지 않는다.
```

### _unit_info

정의: [normalization.py](../../src/excel_parser/normalization.py)

서명: `_unit_info(unit: str | None) -> tuple[str | None, Decimal, str | None] | None`

정적 호출 관계: `normalization.normalize_number`.

```text
등록된 단위의 통화·배율·값 종류를 찾는다.

Args:
    unit: 단위 문자열. None은 배율이 지정되지 않았음을 뜻한다.

Returns:
    통화, Decimal 배율, 값 종류의 튜플. 미등록 단위이면 None.

Raises:
    없음.

Side Effects:
    없음.
```

### _remove_affixes

정의: [normalization.py](../../src/excel_parser/normalization.py)

서명: `_remove_affixes(text: str) -> tuple[str, list[str], list[str], str | None]`

정적 호출 관계: `normalization.normalize_number`.

```text
숫자 바깥의 등록된 통화·단위 표기만 분리한다.

Args:
    text: 유니코드와 공백을 정규화한 전체 문자열.

Returns:
    숫자 부분, 단위 목록, 통화 기호 목록, 오류 코드의 튜플.

Raises:
    없음.

Side Effects:
    없음.
```

### _parse_decimal_token

정의: [normalization.py](../../src/excel_parser/normalization.py)

서명: `_parse_decimal_token(token: str, locale: str | None) -> Decimal | None`

정적 호출 관계: `normalization.normalize_number`.

```text
부호 없는 숫자 토큰을 로케일과 전체 구분자 패턴으로 해석한다.

Args:
    token: 부호와 단위가 제거된 숫자 문자열. 지수부의 부호는 허용한다.
    locale: 등록된 로케일 또는 None. None에서는 모호한 단일 구분자를 거부한다.

Returns:
    유효한 유한 Decimal. 구분자나 전체 형태가 모호하면 None.

Raises:
    없음. Decimal 변환 실패는 None으로 반환한다.

Side Effects:
    없음.
```

### normalize_number

정의: [normalization.py](../../src/excel_parser/normalization.py)

서명: `normalize_number(value: object, *, data_type: str, value_kind: str | None=None, unit: str | None=None, number_format: str | None=None, locale: str | None=None, symbol_rules: dict | None=None) -> dict`

정적 호출 관계: `mapping.map_workbook`.

```text
원본 자료형을 우선하여 숫자·문자열 부호·단위를 한 번만 정규화한다.

Args:
    value: 원본 저장값. 숫자 셀의 XML 숫자 문자열도 허용한다.
    data_type: number/string/boolean/error/date/blank 등 원본 자료형.
    value_kind: amount/ratio/quantity/number 또는 보호할 code/date/text 등.
    unit: 원본 제목·헤더 등에서 확인한 단위. 표시 배율과 별개이다.
    number_format: 원본 숫자 표시 형식. 저장 숫자의 부호를 바꾸지 않는다.
    locale: 문자열 숫자의 확인된 로케일. 미지정 모호 표기는 거부한다.
    symbol_rules: 확인된 기호별 부호 규칙. 값은 1/-1 또는 sign을 가진 사전.

Returns:
    원본·Decimal 정규화값·상태·단위·배율·규칙·진단을 담은 사전.
    confirmed만 정규화값을 가진다. JSON 직렬화 시 Decimal은 문자열로 쓴다.

Raises:
    없음. 지원하지 않는 입력·표기는 미확정 진단으로 반환한다.

Side Effects:
    없음. 파일·DB·콘솔을 사용하거나 원본 객체를 변경하지 않는다.
```

## ooxml.py

### local_name

정의: [ooxml.py](../../src/excel_parser/ooxml.py)

서명: `local_name(tag: str) -> str`

정적 호출 관계: `extraction._extract_sheet`, `extraction.extract_workbook`, `ooxml.xml_text`, `ooxml.read_comments`.

```text
네임스페이스를 제외한 XML 이름을 반환한다.

Args:
    tag: ElementTree 태그 또는 속성명.
Returns:
    마지막 닫는 중괄호 뒤의 이름.
Raises:
    없음. 문자열 입력을 전제로 한다.
Side Effects:
    없음.
```

### xml_text

정의: [ooxml.py](../../src/excel_parser/ooxml.py)

서명: `xml_text(element: ET.Element | None) -> str`

정적 호출 관계: `extraction._cell_value`, `extraction.extract_workbook`, `ooxml.read_comments`.

```text
리치 텍스트의 실제 표시 텍스트를 연결한다.

Args:
    element: 공유 문자열·inline 문자열·메모 요소 또는 None.
Returns:
    발음 안내 rPh를 제외한 t 요소의 연결 문자열.
Raises:
    없음.
Side Effects:
    없음. 전달받은 트리를 변경하지 않는다.
```

### cell_position

정의: [ooxml.py](../../src/excel_parser/ooxml.py)

서명: `cell_position(address: str) -> tuple[int, int, str]`

정적 호출 관계: `extraction._extract_sheet`, `ooxml.range_bounds`.

```text
A1 주소를 1부터 시작하는 행·열 번호로 해석한다.

Args:
    address: 달러 기호를 포함할 수 있는 셀 주소.
Returns:
    행 번호, 열 번호, 열 문자 튜플.
Raises:
    ValueError: 유효한 단일 셀 주소가 아닌 경우.
Side Effects:
    없음.
```

### range_bounds

정의: [ooxml.py](../../src/excel_parser/ooxml.py)

서명: `range_bounds(reference: str) -> tuple[int, int, int, int]`

정적 호출 관계: `extraction._extract_sheet`, `mapping._index`.

```text
직사각형 A1 범위의 경계를 반환한다.

Args:
    reference: 단일 주소 또는 시작:끝 주소.
Returns:
    시작 행·열, 끝 행·열 번호.
Raises:
    ValueError: 주소가 잘못된 경우.
Side Effects:
    없음.
```

### relationships

정의: [ooxml.py](../../src/excel_parser/ooxml.py)

서명: `relationships(archive: ZipFile, part: str) -> dict[str, dict]`

정적 호출 관계: `extraction._extract_sheet`, `extraction.extract_workbook`, `ooxml.read_comments`.

```text
패키지 내부 관계를 대상 경로와 함께 반환한다.

Args:
    archive: 읽기 전용 ZIP 패키지.
    part: 소유 XML 파트. 패키지 루트는 빈 문자열.
Returns:
    관계 ID를 키로 하는 속성과 해석된 part 경로.
Raises:
    ValueError: 내부 관계가 패키지 밖을 가리키는 경우.
    ET.ParseError: 관계 XML이 손상된 경우.
Side Effects:
    관계 XML만 읽는다. 외부 관계는 접속하지 않는다.
```

### read_styles

정의: [ooxml.py](../../src/excel_parser/ooxml.py)

서명: `read_styles(archive: ZipFile, part: str | None) -> dict`

정적 호출 관계: `extraction.extract_workbook`.

```text
셀 서식과 원본 스타일 XML을 보존한다.

Args:
    archive: 읽기 전용 ZIP 패키지.
    part: styles 파트 경로. 없으면 None.
Returns:
    custom_formats, cell_formats, 원본 XML을 담은 사전.
Raises:
    ET.ParseError: 스타일 XML이 손상된 경우.
Side Effects:
    스타일 XML만 읽으며 서식을 적용하거나 값을 변경하지 않는다.
```

### is_date_format

정의: [ooxml.py](../../src/excel_parser/ooxml.py)

서명: `is_date_format(number_format: str | None) -> bool`

정적 호출 관계: `extraction._cell_value`.

```text
명시적인 날짜·시간 서식을 보수적으로 탐지한다.

Args:
    number_format: Excel 표시 형식 또는 None.
Returns:
    리터럴 제거 뒤 날짜·시간 토큰이 확인되면 True.
Raises:
    없음.
Side Effects:
    없음. 일련번호를 날짜나 금액으로 환산하지 않는다.
```

### read_comments

정의: [ooxml.py](../../src/excel_parser/ooxml.py)

서명: `read_comments(archive: ZipFile, sheet_part: str) -> dict[str, list[dict]]`

정적 호출 관계: `extraction._extract_sheet`.

```text
시트에 연결된 일반·스레드 메모를 셀 주소별로 보존한다.

Args:
    archive: 읽기 전용 ZIP 패키지.
    sheet_part: 시트 XML 경로.
Returns:
    셀 주소별 텍스트·작성자·원본 XML 목록.
Raises:
    ET.ParseError: 메모 XML이 손상된 경우.
Side Effects:
    연결된 메모 파트를 읽는다. 외부 서비스에 접근하지 않는다.
```

## paths.py

### _paths_overlap

정의: [paths.py](../../src/excel_parser/paths.py)

서명: `_paths_overlap(first: Path, second: Path) -> bool`

정적 호출 관계: `paths.ProjectPaths.load`.

```text
두 경로가 같거나 서로를 포함하는지 문자열 기준으로 비교한다.

Args:
    first: 비교할 첫 번째 경로. 호출자가 해석을 마친 경로를 전달한다.
    second: 비교할 두 번째 경로. 첫 번째 경로와 같은 기준을 사용한다.

Returns:
    대소문자와 유니코드 조합 차이를 무시했을 때 겹치면 True.

Raises:
    없음. Path 타입의 정상 입력을 기준으로 한다.

Side Effects:
    없음. 파일시스템에 접근하지 않으며 심볼릭 링크도 직접 해석하지 않는다.
```

### _find_project_root

정의: [paths.py](../../src/excel_parser/paths.py)

서명: `_find_project_root() -> Path`

정적 호출 관계: `paths.ProjectPaths.load`.

```text
설정 파일과 패키지 정의가 함께 있는 프로젝트 루트를 찾는다.

Args:
    없음. 현재 작업 위치와 이 소스 파일의 위치를 탐색 기준으로 사용한다.

Returns:
    pyproject.toml과 config/paths.toml을 포함하는 루트의 절대 경로.

Raises:
    FileNotFoundError: 모든 후보에서 프로젝트 루트를 찾지 못한 경우.
    OSError: 현재 위치 확인이나 경로 해석에서 파일시스템 오류가 전파된 경우.
    RuntimeError: 실행 환경에 따라 링크 순환 등 경로 해석이 실패한 경우.

Side Effects:
    후보 경로의 파일 존재 여부를 확인한다. 내용 읽기와 파일 쓰기는 없다.
```

### ProjectPaths.load

정의: [paths.py](../../src/excel_parser/paths.py)

서명: `ProjectPaths.load(cls, root: str | Path | None=None) -> 'ProjectPaths'`

정적 호출 관계: `__main__.main`, `pipeline.run_parser`.

```text
설정 파일을 읽어 서로 겹치지 않는 프로젝트 내부 경로를 반환한다.

Args:
    root: 프로젝트 루트. None이면 자동 탐색한다. 상대 경로를 주면
        현재 작업 위치를 기준으로 해석하며 '~'는 사용자 홈으로 확장한다.

Returns:
    root, input_dir, output_dir, logs_dir를 담은 ProjectPaths 객체.
    설정된 폴더의 실제 존재 여부나 접근 권한은 보장하지 않는다.

Raises:
    FileNotFoundError: 프로젝트 루트나 경로 설정 파일을 찾지 못한 경우.
    tomllib.TOMLDecodeError: 설정 파일이 올바른 TOML 형식이 아닌 경우.
    ValueError: 필수 설정이 없거나 경로가 프로젝트 밖 또는 서로 겹친 경우.
    OSError: 설정 읽기나 경로 해석에서 파일시스템 오류가 전파된 경우.
    RuntimeError: 홈 경로나 심볼릭 링크를 실행 환경에서 해석할 수 없는 경우.

Side Effects:
    config/paths.toml을 읽고 경로 정보를 조회한다.
    원본 파일 내용을 읽거나 파일·디렉토리를 생성하지 않는다.
```

### ProjectPaths.list_input_files

정의: [paths.py](../../src/excel_parser/paths.py)

서명: `ProjectPaths.list_input_files(self, *, recursive: bool=False) -> list[Path]`

정적 호출 관계: 외부 호출 또는 동적 호출 가능. 전체 실행 흐름과 호출 코드를 함께 확인합니다..

```text
입력 디렉토리에서 대상 엑셀 파일의 경로를 정렬해 반환한다.

Args:
    recursive: True이면 하위 폴더도 탐색한다. 기본값 False는 직속 항목만
        확인한다. self.input_dir를 기준으로 파일 목록을 읽는다.

Returns:
    대소문자와 관계없이 .xlsx 또는 .xlsb인 파일의 경로 목록.
    '~$'로 시작하는 파일은 제외한다. 전체 경로의 casefold() 결과로
    정렬하고 같으면 원래 경로 문자열로 비교한다. 대상이 없으면 빈 목록.

Raises:
    FileNotFoundError: 입력 경로가 없거나 디렉토리로 확인되지 않는 경우.
    OSError: 디렉토리 탐색 중 파일시스템 접근 오류가 전파된 경우.

Side Effects:
    디렉토리 항목과 파일 정보를 조회한다. 파일 내용 읽기와 쓰기는 없다.
    파일을 가리키는 심볼릭 링크도 포함할 수 있으며, 반환 경로의 대상이
    입력 디렉토리 안에 있는지까지 검증하지는 않는다.
```

## pipeline.py

### json_value

정의: [pipeline.py](../../src/excel_parser/pipeline.py)

서명: `json_value(value: object) -> str`

정적 호출 관계: 외부 호출 또는 동적 호출 가능. 전체 실행 흐름과 호출 코드를 함께 확인합니다..

```text
Decimal을 정밀도 손실 없는 JSON 문자열 값으로 직렬화한다.

Args:
    value: json 모듈이 기본 처리할 수 없는 객체.
Returns:
    Decimal의 고정 소수점 문자열.
Raises:
    TypeError: 등록하지 않은 타입인 경우.
    ValueError: 유한하지 않은 Decimal인 경우.
Side Effects:
    없음.
```

### collect_inputs

정의: [pipeline.py](../../src/excel_parser/pipeline.py)

서명: `collect_inputs(inputs: list[Path], *, recursive: bool=False) -> list[Path]`

정적 호출 관계: `pipeline.run_parser`.

```text
명시한 파일·폴더에서 Excel 입력을 빠짐없이 수집한다.

Args:
    inputs: 단일·복수 파일 또는 디렉토리 경로.
    recursive: 디렉토리 하위도 탐색할지 여부.
Returns:
    중복 실제 경로를 제거한 정렬 목록. 미지원 Excel 형식도 진단용으로 포함한다.
Raises:
    FileNotFoundError: 입력이 없거나 파일·디렉토리가 아닌 경우.
    ValueError: 탐색 결과 Excel 파일이 없는 경우.
Side Effects:
    경로 정보만 조회한다. Excel '~$' 잠금 파일은 제외한다.
```

### _write_json

정의: [pipeline.py](../../src/excel_parser/pipeline.py)

서명: `_write_json(stream: TextIO, value: object) -> None`

정적 호출 관계: `pipeline.run_parser`.

```text
열린 결과 스트림에 JSON 레코드 한 줄을 기록한다.

Args:
    stream: 호출자가 명시적으로 연 출력 스트림.
    value: JSON으로 직렬화할 레코드.
Returns:
    없음.
Raises:
    TypeError: 계약 밖의 객체가 포함된 경우.
    OSError: 결과 기록 실패.
Side Effects:
    전달받은 스트림에 한 줄을 쓴다. 새 경로를 결정하거나 열지 않는다.
```

### _file_metadata

정의: [pipeline.py](../../src/excel_parser/pipeline.py)

서명: `_file_metadata(metadata: dict, file: dict) -> dict`

정적 호출 관계: `pipeline.run_parser`.

```text
실행 기본값과 파일별 명시 메타데이터를 결합한다.

Args:
    metadata: defaults와 files를 포함하는 실행 메타데이터.
    file: 파일명·출처 경로·식별자.
Returns:
    해당 파일에 적용할 평탄한 메타데이터 사전.
Raises:
    ValueError: 메타데이터 구조가 사전이 아닌 경우.
Side Effects:
    없음. 원본 파일과 입력 사전을 변경하지 않는다.
```

### _record_observations

정의: [pipeline.py](../../src/excel_parser/pipeline.py)

서명: `_record_observations(records: list[dict], observations: dict) -> None`

정적 호출 관계: `pipeline.run_parser`.

```text
실제로 적용한 정규화 규칙의 횟수와 출처 예시를 집계한다.

Args:
    records: 한 파일의 표준 레코드.
    observations: 누적할 규칙별 사전.
Returns:
    없음. observations를 메모리에서 갱신한다.
Raises:
    없음. 매핑 계약을 충족한 입력을 전제로 한다.
Side Effects:
    전달받은 누적 사전만 변경한다.
```

### run_parser

정의: [pipeline.py](../../src/excel_parser/pipeline.py)

서명: `run_parser(inputs: list[Path], *, root: Path, recursive: bool=False, output_dir: Path | None=None, metadata: dict | None=None) -> dict`

정적 호출 관계: `__main__.main`.

```text
입력 전체를 순차 처리하고 원본 보존·독립 검증 결과를 반환한다.

Args:
    inputs: 원본 파일 또는 디렉토리 목록.
    root: 프로젝트 루트.
    recursive: 하위 디렉토리 탐색 여부.
    output_dir: 명시적으로 요청된 새 결과 디렉토리. None이면 저장하지 않는다.
    metadata: defaults와 파일별 files 메타데이터. 누락 정보는 추정하지 않는다.
Returns:
    실행 요약. 파일별 실패·미확정·검증 통과 여부를 별도 필드로 담는다.
Raises:
    ValueError: 입력·메타데이터·출력 경로가 허용된 범위를 벗어난 경우.
    OSError: 원본 탐색 또는 결과 기록에 실패한 경우.
Side Effects:
    원본을 읽기만 한다. output_dir 지정 시에만 새로운 결과 파일을 생성한다.
    기존 결과를 덮어쓰지 않으며 Data 내부에는 쓰지 않는다.
```

## validation.py

### _index_sheets

정의: [validation.py](../../src/excel_parser/validation.py)

서명: `_index_sheets(archive: ZipFile) -> list[dict[str, str]]`

정적 호출 관계: `validation.validate_extraction`, `validation._read_numeric_sources`, `validation.compare_master`.

```text
통합문서 관계를 직접 읽어 모든 시트의 이름과 XML 위치를 얻는다.

Args:
    archive: 읽기용으로 열린 OOXML ZIP.
Returns:
    문서 순서의 name, state, part 사전 목록.
Raises:
    KeyError: 필수 ZIP 부분이나 관계가 없는 경우.
    ET.ParseError: 원본 XML이 손상된 경우.
Side Effects:
    통합문서와 관계 XML만 읽는다. 파일을 생성하지 않는다.
```

### _read_strings

정의: [validation.py](../../src/excel_parser/validation.py)

서명: `_read_strings(archive: ZipFile) -> list[str]`

정적 호출 관계: `validation.validate_extraction`, `validation._read_numeric_sources`.

```text
공유 문자열의 실제 텍스트를 원본 순서대로 읽는다.

Args:
    archive: 읽기용 OOXML ZIP.
Returns:
    공유 문자열 목록. 부분이 없으면 빈 목록.
Raises:
    ET.ParseError: 문자열 XML이 손상된 경우.
Side Effects:
    공유 문자열 XML을 메모리로 읽으며 쓰지 않는다.
```

### _decode_cell

정의: [validation.py](../../src/excel_parser/validation.py)

서명: `_decode_cell(cell: ET.Element, strings: list[str]) -> dict[str, Any]`

정적 호출 관계: `validation._scan_sheet`.

```text
개별 원본 셀의 저장값·수식·공유 문자열을 독립 해석한다.

Args:
    cell: 원본 XML의 c 요소.
    strings: 원본 공유 문자열 목록.
Returns:
    비교에 필요한 원본 값과 수식·속성의 사전.
Raises:
    IndexError: 공유 문자열 번호가 유효하지 않은 경우.
    ValueError: 공유 문자열 번호를 정수로 읽을 수 없는 경우.
Side Effects:
    없음. 제공된 XML 요소만 조회한다.
```

### _scan_sheet

정의: [validation.py](../../src/excel_parser/validation.py)

서명: `_scan_sheet(archive: ZipFile, part: str, strings: list[str]) -> dict[str, Any]`

정적 호출 관계: `validation.validate_extraction`, `validation._read_numeric_sources`.

```text
한 시트 XML을 순회해 주소·값·행열·병합의 독립 근거를 만든다.

Args:
    archive: 읽기용 OOXML ZIP.
    part: 검사할 시트 XML 부분 경로.
    strings: 공유 문자열 목록.
Returns:
    전체 XML 주소 집합, 값 셀, 행열 속성, 병합 범위와 메모 목록.
Raises:
    KeyError: 지정된 ZIP 부분이 없는 경우.
    ET.ParseError: 원본 XML이 손상된 경우.
Side Effects:
    시트와 연결된 메모 XML을 읽으며 원본은 수정하지 않는다.
```

### _column_label

정의: [validation.py](../../src/excel_parser/validation.py)

서명: `_column_label(index: int) -> str`

정적 호출 관계: `validation._preserved_addresses`.

```text
1부터 시작하는 열 번호를 A1 표기의 열 이름으로 변환한다.

Args:
    index: 양의 정수 열 번호.
Returns:
    A, Z, AA 등의 열 이름.
Raises:
    ValueError: 열 번호가 1보다 작은 경우.
Side Effects:
    없음. 메모리에서만 처리한다.
```

### _preserved_addresses

정의: [validation.py](../../src/excel_parser/validation.py)

서명: `_preserved_addresses(sheet: dict[str, Any]) -> set[str]`

정적 호출 관계: `validation.validate_extraction`.

```text
추출 셀과 압축 빈 구간에서 보존된 주소를 복원한다.

Args:
    sheet: 생산 파서의 한 시트 결과.
Returns:
    명시적으로 보존된 모든 셀 주소 집합.
Raises:
    KeyError: 필수 주소나 빈 구간 필드가 누락된 경우.
    ValueError: 잘못된 열 번호가 전달된 경우.
Side Effects:
    없음. 메모리에서 주소 집합만 만든다.
```

### _cell_issues

정의: [validation.py](../../src/excel_parser/validation.py)

서명: `_cell_issues(expected: dict[str, Any], actual: dict[str, Any]) -> list[str]`

정적 호출 관계: `validation.validate_extraction`.

```text
원본 셀과 추출 셀의 값·수식·스타일 불일치를 분류한다.

Args:
    expected: 독립 XML 점검으로 읽은 셀.
    actual: 생산 파서에서 반환한 셀.
Returns:
    불일치한 필드명의 목록. 일치하면 빈 목록.
Raises:
    없음. 누락 필드는 불일치로 처리한다.
Side Effects:
    없음. 원본과 결과 사전을 변경하지 않는다.
```

### validate_extraction

정의: [validation.py](../../src/excel_parser/validation.py)

서명: `validate_extraction(path: Path, workbook: dict[str, Any]) -> dict[str, Any]`

정적 호출 관계: `pipeline.run_parser`.

```text
원본 OOXML과 추출 결과의 전체 시트·셀·구조를 독립 비교한다.

Args:
    path: 읽기 전용 원본 xlsx 경로.
    workbook: extract_workbook이 반환한 통합문서 결과.
Returns:
    passed, 전체 개수, 시트별 개수 및 위치가 포함된 issues 사전.
Raises:
    OSError: 원본을 읽을 수 없는 경우.
    KeyError: 필수 OOXML 부분이 없는 경우.
    ET.ParseError: 원본 XML이 손상된 경우.
Side Effects:
    원본 ZIP을 다시 읽는다. 생산 파서의 reader나 결과값을 정답으로 쓰지
    않으며 파일을 수정·생성하지 않는다.
```

### _read_numeric_sources

정의: [validation.py](../../src/excel_parser/validation.py)

서명: `_read_numeric_sources(path: Path, required: set[str]) -> dict[str, dict[str, Any]]`

정적 호출 관계: `validation.compare_master`.

```text
마스터 비교에 필요한 원본 시트를 생산 파서 없이 읽는다.

Args:
    path: 원본 xlsx 경로.
    required: 비교에 필요한 시트명 집합.
Returns:
    시트명별 주소·원본 셀 사전.
Raises:
    OSError: 원본을 읽을 수 없는 경우.
    ET.ParseError: 원본 XML이 손상된 경우.
Side Effects:
    지정 원본의 필요한 시트만 읽으며 산출물을 만들지 않는다.
```

### validate_provenance

정의: [validation.py](../../src/excel_parser/validation.py)

서명: `validate_provenance(workbook: dict[str, Any], mapped: dict[str, Any]) -> dict[str, Any]`

정적 호출 관계: `pipeline.run_parser`.

```text
표준 레코드가 같은 원본의 숫자·수식과 정확히 연결되는지 확인한다.

Args:
    workbook: 독립 원본 대조를 수행한 추출 결과.
    mapped: map_workbook의 표준 레코드와 GAP 관계 결과.
Returns:
    passed, record_count, required_candidate_count, 참조 검사 수와 issues.
    숫자 및 모든 자료형의 수식은 필수 후보이며, 일반 날짜·불리언은
    필수 금액 후보로 간주하지 않는다.
Raises:
    KeyError: 원본 계약의 record_id 등 필수 필드가 없는 경우.
Side Effects:
    없음. 메모리에서 전체 후보와 출처·원문·보조 근거를 비교한다.
```

### compare_master

정의: [validation.py](../../src/excel_parser/validation.py)

서명: `compare_master(paths: list[Path]) -> dict[str, Any]`

정적 호출 관계: `pipeline.run_parser`.

```text
실제 마스터 수식과 조직 원본 합계를 제한된 비교 범위에서 대조한다.

Args:
    paths: 이번 입력의 원본 파일 목록. 파일명으로 조직을 확정하지 않는다.
Returns:
    status, comparison_count, 비교 근거와 issues를 포함한 사전.
    사업부별집계 시트가 없으면 not_applicable로 반환한다.
Raises:
    OSError: 원본 파일을 읽을 수 없는 경우.
    ET.ParseError: 원본 XML이 손상된 경우.
Side Effects:
    원본의 집계 수식·캐시와 조직별 상세 셀을 독립 조회한다. 공식 마스터를
    발행하거나 일반 수식 계산 엔진을 실행하지 않으며 원본을 변경하지 않는다.
```
