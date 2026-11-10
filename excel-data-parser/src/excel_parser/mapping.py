"""원본 셀을 보존하면서 확인된 표 구조를 DB 독립 레코드로 해석한다.

원본 파일을 읽거나 결과를 저장하지 않는다. 원본 추출 결과와 같은 워크북의
헤더·코드표만 사용하며, 미확정 의미나 Contingency 조정값을 만들지 않는다.
"""

import re
from collections import Counter, defaultdict

from .normalization import inspect_number_format, normalize_number
from .ooxml import range_bounds


def _text(cell: dict | None) -> str:
    """셀의 저장값을 검색용 문자열로 반환한다.

    Args:
        cell: 추출 셀 또는 None. 수식은 이미 추출된 캐시값을 사용한다.
    Returns:
        값이 없으면 빈 문자열, 그 외에는 저장값의 문자열 표현.
    Raises:
        없음. 추출 계약을 만족하는 입력을 전제로 한다.
    Side Effects:
        없음. 입력을 변경하거나 수식을 계산하지 않는다.
    """
    value = cell.get("value") if cell else None
    return "" if value is None else str(value)


def _ref(cell: dict | None) -> dict | None:
    """해석 근거 셀과 원본 레코드의 연결 정보를 만든다.

    Args:
        cell: 근거 셀 또는 None.
    Returns:
        시트·셀·원본 레코드 참조 또는 None.
    Raises:
        없음. 추출 계약을 만족하는 입력을 전제로 한다.
    Side Effects:
        없음.
    """
    if cell is None:
        return None
    result = {
        "sheet": cell.get("sheet"),
        "cell": cell.get("address"),
        "raw_record_id": cell.get("record_id"),
    }
    if cell.get("_header_inheritance"):
        result["inheritance"] = cell["_header_inheritance"]
    return result


def _index(sheet: dict) -> dict:
    """행·열별 셀과 주소별 셀의 메모리 색인을 만든다.

    Args:
        sheet: extraction의 시트 구조.
    Returns:
        rows와 addresses 색인을 포함하는 사전.
    Raises:
        KeyError: 필수 행·열·주소가 누락된 잘못된 추출 자료.
        ValueError: 병합 범위 주소가 잘못된 경우.
    Side Effects:
        없음. 원본 셀 사전은 읽기 참조로만 사용한다.
    """
    rows: dict[int, dict[int, dict]] = defaultdict(dict)
    addresses = {}
    for cell in sheet["cells"]:
        rows[cell["row"]][cell["column_index"]] = cell
        addresses[cell["address"]] = cell
    return {
        "rows": rows,
        "addresses": addresses,
        "merges": [
            (merge, range_bounds(merge["range"])) for merge in sheet.get("merges", [])
        ],
    }


def _header_at(index: dict, row: int, column: int) -> dict | None:
    """병합 헤더를 읽을 때만 기준 셀과 상속 위치를 연결한다.

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
    """
    direct = index["rows"].get(row, {}).get(column)
    if _text(direct):
        return direct
    for merge, (row1, column1, row2, column2) in index["merges"]:
        if row1 <= row <= row2 and column1 <= column <= column2:
            anchor = index["addresses"].get(merge["anchor"])
            if anchor:
                return {
                    **anchor,
                    "_header_inheritance": {
                        "range": merge["range"],
                        "anchor": merge["anchor"],
                        "target_row": row,
                        "target_column": column,
                    },
                }
    return direct


def _scenario(label: str) -> str:
    """명시된 한국어 값 구분을 계약의 시나리오로 해석한다.

    Args:
        label: 원본 헤더 문자열.
    Returns:
        actual, plan, estimate, mixed, unknown 중 하나.
    Raises:
        없음.
    Side Effects:
        없음. 월이나 파일 이름으로 시나리오를 추측하지 않는다.
    """
    if "실적" in label and "추정" in label:
        return "mixed"
    for word, scenario in (("계획", "plan"), ("추정", "estimate"), ("실적", "actual")):
        if word in label:
            return scenario
    return "unknown"


def _year(label: str) -> int | None:
    """명시된 연도 또는 이 템플릿의 작은따옴표 연도를 읽는다.

    Args:
        label: 2027년 또는 '27년 같은 원본 연도 헤더.
    Returns:
        4자리 연도. 명시 연도가 없으면 None.
    Raises:
        없음.
    Side Effects:
        없음.
    """
    full = re.search(r"(?<!\d)(20\d{2})(?:년|\s|$)", label)
    short = re.search(r"['’](\d{2})년", label)
    return int(full.group(1)) if full else 2000 + int(short.group(1)) if short else None


def _period(cells: list[dict], *, annual: bool = False) -> dict:
    """근거 헤더에서 대상 연월과 기간 집계 방식을 해석한다.

    Args:
        cells: 같은 값 열을 설명하는 원본 헤더 셀 목록.
        annual: 단일 복합 헤더의 연간 집계를 명시하는 템플릿 규칙.
    Returns:
        기간·시나리오·원본 헤더·근거 셀 사전.
    Raises:
        없음. 추출 계약을 만족하는 입력을 전제로 한다.
    Side Effects:
        없음. 제출 월과 대상 월을 결합하지 않는다.
    """
    labels = [_text(cell) for cell in cells if _text(cell)]
    joined = " ".join(labels)
    month = next(
        (
            int(match.group(1))
            for label in labels
            if (match := re.fullmatch(r"\s*(\d{1,2})월\s*", label))
        ),
        None,
    )
    quarter = next(
        (
            int(match.group(1))
            for label in labels
            if (match := re.fullmatch(r"\s*([1-4])Q\s*", label))
        ),
        None,
    )
    aggregation = (
        "month" if month else "annual" if annual or "연간" in joined else "unknown"
    )
    return {
        "year": next((year for label in labels if (year := _year(label))), None),
        "month": month,
        "quarter": quarter,
        "aggregation": aggregation,
        "scenario": _scenario(joined),
        "raw_headers": labels,
        "provenance": [_ref(cell) for cell in cells],
    }


def _layout(sheet: dict, index: dict) -> dict:
    """파일명이 아닌 실제 헤더 조합으로 알려진 재무 표를 식별한다.

    Args:
        sheet: 시트 구조.
        index: _index 결과.
    Returns:
        표 종류, 헤더 행, 값 열별 기간 및 헤더명별 열 색인.
    Raises:
        없음. 추출 계약을 만족하는 입력을 전제로 한다.
    Side Effects:
        없음. 빈 셀에 원본 값이나 금액을 복제하지 않는다.
    """
    rows = index["rows"]
    layout = {"kind": "unknown", "header_row": None, "periods": {}, "columns": {}}
    for row_number, row in sorted(rows.items()):
        labels = {
            _text(cell).strip(): column for column, cell in row.items() if _text(cell)
        }
        if {"조직명", "관리분류", "중분류", "소분류", "사업부코드"} <= labels.keys():
            kind, group_row, month_row = "revenue_detail", row_number, row_number + 1
        elif {"대분류", "중분류", "소분류", "조직명", "예산Code"} <= labels.keys():
            kind, group_row, month_row = "expense_detail", row_number - 1, row_number
        elif {"사업부코드", "사업부명"} <= labels.keys() and any(
            _year(label) and "매출" in label for label in labels
        ):
            periods = {
                column: _period([cell], annual=True)
                for column, cell in row.items()
                if _year(_text(cell)) is not None
            }
            return {
                "kind": "organization_summary",
                "header_row": row_number,
                "periods": periods,
                "columns": labels,
            }
        else:
            continue
        groups = [
            (column, cell)
            for column, cell in sorted(rows.get(group_row, {}).items())
            if _year(_text(cell)) and _scenario(_text(cell)) != "unknown"
        ]
        periods = {}
        for column, cell in sorted(rows.get(month_row, {}).items()):
            if not (
                re.fullmatch(r"\d{1,2}월", _text(cell).strip()) or "연간" in _text(cell)
            ):
                continue
            candidates = [
                (start, group) for start, group in groups if 0 <= column - start <= 12
            ]
            if candidates:
                periods[column] = _period([candidates[-1][1], cell])
        return {
            "kind": kind,
            "header_row": month_row,
            "group_row": group_row,
            "periods": periods,
            "columns": labels,
        }
    for row_number, row in sorted(rows.items()):
        year_columns = [
            column
            for column, cell in row.items()
            if re.fullmatch(r"20\d{2}년", _text(cell))
        ]
        if not year_columns:
            continue
        year_columns = sorted(set(year_columns) | set(rows.get(row_number + 3, {})))
        periods = {}
        for column in year_columns:
            if _year(_text(_header_at(index, row_number, column))) is None:
                continue
            cells = [
                _header_at(index, row_number + offset, column) for offset in range(4)
            ]
            cells = [cell for cell in cells if cell is not None]
            period = _period(cells)
            if period["scenario"] != "unknown" and period["aggregation"] != "unknown":
                periods[column] = period
        if periods:
            return {
                "kind": "financial_statement",
                "header_row": row_number + 3,
                "periods": periods,
                "columns": {},
            }
    return layout


def _periods_with_cross_evidence(layouts: list[dict]) -> None:
    """같은 워크북의 명시 월별 시나리오로 혼합 기간의 의미를 보완한다.

    Args:
        layouts: 표 레이아웃 목록. 시트별 기간 사전이 갱신된다.
    Returns:
        없음.
    Raises:
        없음. 내부 생성 레이아웃을 전제로 한다.
    Side Effects:
        전달된 레이아웃의 기간만 갱신한다. 원본 workbook은 변경하지 않는다.
    """
    evidence: dict[tuple[int, int], list[dict]] = defaultdict(list)
    for layout in layouts:
        if layout["kind"] != "financial_statement":
            continue
        for period in layout["periods"].values():
            if period["month"] and period["scenario"] in {"actual", "estimate"}:
                evidence[(period["year"], period["month"])].append(period)
    for layout in layouts:
        for period in layout["periods"].values():
            if period["scenario"] == "mixed" and period["month"]:
                matches = evidence.get((period["year"], period["month"]), [])
                scenarios = {match["scenario"] for match in matches}
                if len(scenarios) == 1:
                    period["scenario"] = next(iter(scenarios))
                    period["provenance"] = (
                        period["provenance"] + matches[0]["provenance"]
                    )
                    period["rule_id"] = "PERIOD_CROSS_SHEET_MONTH_SCENARIO"
            if (
                layout["kind"] == "financial_statement"
                and period["aggregation"] == "annual"
            ):
                if period["scenario"] != "actual":
                    continue
                monthly = [
                    p
                    for p in layout["periods"].values()
                    if p["year"] == period["year"]
                    and p["month"]
                    and p["scenario"] in {"actual", "estimate"}
                ]
                if {p["scenario"] for p in monthly} == {"actual", "estimate"}:
                    period["scenario"] = "mixed"
                    period["rule_id"] = "PERIOD_ANNUAL_MIXED_COMPONENTS"
                    period["provenance"] = period["provenance"] + [
                        ref for p in monthly for ref in p["provenance"]
                    ]


def _organization_registry(indices: list[dict]) -> dict[str, dict]:
    """워크북 내 사업부 코드표에서 코드와 이름의 확인 근거를 읽는다.

    Args:
        indices: 각 시트의 메모리 색인.
    Returns:
        사업부 코드별 이름·근거. 충돌 이름은 확정하지 않는다.
    Raises:
        없음. 추출 계약을 만족하는 입력을 전제로 한다.
    Side Effects:
        없음. 다른 워크북이나 마스터 파일에 접근하지 않는다.
    """
    registry: dict[str, dict] = {}
    for index in indices:
        for row_number, row in index["rows"].items():
            labels = {_text(cell).strip(): column for column, cell in row.items()}
            if not {"사업부코드", "사업부명", "팀코드"} <= labels.keys():
                continue
            for number, data in index["rows"].items():
                if number <= row_number:
                    continue
                code_cell, name_cell = (
                    data.get(labels["사업부코드"]),
                    data.get(labels["사업부명"]),
                )
                code, name = _text(code_cell), _text(name_cell)
                if code and name:
                    previous = registry.get(code)
                    conflict = previous is not None and previous["name"] != name
                    registry[code] = {
                        "name": None if conflict else name,
                        "provenance": [_ref(code_cell), _ref(name_cell)],
                    }
    return registry


def _workbook_organization(layouts: list[dict], indices: list[dict]) -> dict:
    """상세 행의 조직 식별자를 모아 단일 조직과 여러 조직 범위를 구분한다.

    Args:
        layouts: 시트별 레이아웃.
        indices: 같은 순서의 시트별 색인.
    Returns:
        단일 코드·원래 이름 또는 multi_organization 범위와 원본 근거.
    Raises:
        없음. 내부 생성 자료를 전제로 한다.
    Side Effects:
        없음. 파일 이름의 조직을 확정값으로 사용하지 않는다.
    """
    organizations = {}
    for layout, index in zip(layouts, indices):
        if layout["kind"] not in {"revenue_detail", "expense_detail"}:
            continue
        columns = layout["columns"]
        for row_number, row in index["rows"].items():
            if row_number <= layout["header_row"]:
                continue
            code_cell = row.get(columns.get("사업부코드"))
            name_cell = row.get(columns.get("조직명"))
            code = _text(code_cell)
            if code:
                organizations[code] = {
                    "code": code,
                    "name": _text(name_cell) or None,
                    "original_name": _text(name_cell) or None,
                    "scope": "organization",
                    "provenance": [_ref(code_cell), _ref(name_cell)],
                }
    if len(organizations) == 1:
        return next(iter(organizations.values()))
    return {
        "code": None,
        "name": None,
        "original_name": None,
        "scope": "multi_organization" if organizations else "unknown",
        "provenance": [
            ref for value in organizations.values() for ref in value["provenance"]
        ],
    }


def _metric(label: str) -> str | None:
    """실제 항목명에서 GAP 관련 지표를 식별한다.

    Args:
        label: 표의 항목명. 파일명이나 임의 추정값은 전달하지 않는다.
    Returns:
        OP, CONTINGENCY, OP_EX_CONTINGENCY 또는 None.
    Raises:
        없음.
    Side Effects:
        없음. 지표 값을 계산하거나 조정하지 않는다.
    """
    compact = re.sub(r"\s+", "", label).casefold()
    contingency = "contingency" in compact or "컨틴전시" in compact
    if "영업이익률" in compact:
        return None
    if "영업이익" in compact:
        return "OP_EX_CONTINGENCY" if contingency and "제외" in compact else "OP"
    if contingency:
        return "CONTINGENCY"
    return None


def _unit(index: dict) -> tuple[str | None, list[dict]]:
    """제목 또는 단위 설명에 명시된 기본 금액 단위를 찾는다.

    Args:
        index: 현재 시트 색인.
    Returns:
        원·천원·백만원·억원 중 하나와 근거 셀 목록. 없으면 None.
    Raises:
        없음.
    Side Effects:
        없음. 표시 형식의 쉼표를 저장값의 단위로 추측하지 않는다.
    """
    for row_number, row in sorted(index["rows"].items()):
        if row_number > 5:
            break
        for cell in row.values():
            label = _text(cell)
            if "단위" in label or ("(" in label and "원" in label):
                match = re.search(r"억원|백만원|천원|(?<![가-힣])원(?![가-힣])", label)
                if match:
                    return match.group(), [_ref(cell)]
    return None, []


def _submission(metadata: dict) -> dict:
    """사용자가 전달한 제출 정보만 보존하고 명확한 버전 표기를 분리한다.

    Args:
        metadata: submission_month, submission_version 등의 선택 메타데이터.
    Returns:
        버전 원문, 월별 회차, 개정번호, 제출 월과 확인 상태.
    Raises:
        ValueError: 버전 원문이 문자열이 아니거나 제출 월 형식이 잘못된 경우.
    Side Effects:
        없음. 버전을 발급하거나 승인 여부를 판정하지 않는다.
    """
    version = metadata.get("submission_version")
    month = metadata.get("submission_month")
    if version is not None and not isinstance(version, str):
        raise ValueError("submission_version은 소수가 아닌 원문 문자열이어야 합니다.")
    if month is not None and (
        not isinstance(month, str) or not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", month)
    ):
        raise ValueError("submission_month는 YYYY-MM 문자열이어야 합니다.")
    for field in ("submission_cycle", "revision"):
        value = metadata.get(field)
        if value is not None and (type(value) is not int or value < 0):
            raise ValueError(f"{field}는 0 이상의 정수여야 합니다.")
    match = re.fullmatch(r"V(\d+)(?:\.(\d+))?", version) if version else None
    return {
        "version_raw": version,
        "cycle": int(match.group(1)) if match else metadata.get("submission_cycle"),
        "revision": int(match.group(2))
        if match and match.group(2)
        else metadata.get("revision"),
        "submission_month": month,
        "status": "provided" if version or month else "unknown",
    }


def _statement_items(index: dict, layout: dict) -> dict[int, dict]:
    """답안지의 들여쓰기 열을 계층 항목 경로로 보존한다.

    Args:
        index: 답안지 셀 색인.
        layout: financial_statement 레이아웃.
    Returns:
        행별 항목명, 계층 경로와 원본 근거.
    Raises:
        없음. 내부 생성 자료를 전제로 한다.
    Side Effects:
        없음. 원본 빈 셀에는 값을 쓰지 않는다.
    """
    hierarchy: dict[int, dict] = {}
    items = {}
    first_value_column = min(layout["periods"])
    for row_number, row in sorted(index["rows"].items()):
        if row_number <= layout["header_row"]:
            continue
        labels = [
            (column, cell)
            for column, cell in sorted(row.items())
            if column < first_value_column
            and cell.get("data_type") == "string"
            and _text(cell)
        ]
        if not labels:
            continue
        first_column = labels[0][0]
        hierarchy = {
            column: cell for column, cell in hierarchy.items() if column < first_column
        }
        hierarchy.update(labels)
        path = []
        for cell in hierarchy.values():
            if not path or path[-1] != _text(cell):
                path.append(_text(cell))
        items[row_number] = {
            "code": _metric(_text(labels[-1][1])),
            "name": _text(labels[-1][1]),
            "header_path": path,
            "provenance": [_ref(cell) for cell in hierarchy.values()],
        }
    return items


def _row_context(
    cell: dict,
    index: dict,
    layout: dict,
    common_org: dict,
    registry: dict,
    statement_items: dict,
) -> dict:
    """값 셀에 대응하는 조직·항목·역할을 실제 헤더에서 가져온다.

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
    """
    row = index["rows"][cell["row"]]
    columns = layout["columns"]
    kind = layout["kind"]
    financial = cell["column_index"] in layout["periods"]
    organization = dict(common_org)
    dimensions = {}
    item = {"code": None, "name": None, "header_path": [], "provenance": []}
    role = "unknown"
    if kind in {"revenue_detail", "expense_detail"}:
        data_row = cell["row"] > layout["header_row"]
        code_cell, name_cell = (
            row.get(columns.get("사업부코드")),
            row.get(columns.get("조직명")),
        )
        if data_row and (_text(code_cell) or _text(name_cell)):
            code = _text(code_cell) or None
            lookup = registry.get(code, {})
            organization = {
                "code": code,
                "name": _text(name_cell) or lookup.get("name"),
                "original_name": _text(name_cell) or None,
                "scope": "organization",
                "provenance": [
                    ref for ref in [_ref(code_cell), _ref(name_cell)] if ref
                ],
            }
            names = (
                ["관리분류", "중분류", "소분류"]
                if kind == "revenue_detail"
                else ["대분류", "중분류", "소분류"]
            )
            labels = [row.get(columns.get(name)) for name in names]
            path = [_text(label) for label in labels if _text(label)]
            name_cell = (
                row.get(columns.get("예산명칭"))
                if kind == "expense_detail"
                else labels[-1]
            )
            code_cell = row.get(columns.get("예산Code"))
            item = {
                "code": _text(code_cell) or _metric(_text(name_cell)),
                "name": _text(name_cell) or (path[-1] if path else None),
                "header_path": path,
                "provenance": [_ref(label) for label in labels if label],
            }
            for name in (
                "관리분류",
                "중분류",
                "소분류",
                "본부명",
                "팀Code",
                "팀명",
                "실행과제Code",
                "실행과제명",
                "예산Code",
                "실적집계기준",
                "비용구분",
                "담당자",
                "거래내역 및 산출근거",
            ):
                source = row.get(columns.get(name))
                if source:
                    dimensions[name] = {
                        "value": source.get("value"),
                        "source": _ref(source),
                    }
            role = "detail" if financial else "reference"
        elif (
            financial
            and cell.get("formula")
            and "SUBTOTAL(" in (cell["formula"].get("text") or "").upper()
        ):
            item = {
                "code": None,
                "name": "매출" if kind == "revenue_detail" else "영업비용",
                "header_path": [],
                "provenance": [],
            }
            role = "subtotal"
        else:
            financial = False
    elif (
        kind == "financial_statement"
        and financial
        and cell["row"] > layout["header_row"]
    ):
        item = statement_items.get(cell["row"], item)
        role = (
            "total"
            if item["code"] or item["name"] in {"매출", "영업비용"}
            else "subtotal"
        )
    elif (
        kind == "organization_summary"
        and financial
        and cell["row"] > layout["header_row"]
    ):
        code_cell, name_cell = (
            row.get(columns.get("사업부코드")),
            row.get(columns.get("사업부명")),
        )
        organization = {
            "code": _text(code_cell) or None,
            "name": _text(name_cell) or None,
            "original_name": _text(name_cell) or None,
            "scope": "organization" if _text(code_cell) else "multi_organization",
            "provenance": [ref for ref in [_ref(code_cell), _ref(name_cell)] if ref],
        }
        header = index["rows"][layout["header_row"]][cell["column_index"]]
        name = re.sub(r"^20\d{2}\s+(?:실적/추정|실적|계획|추정)\s+", "", _text(header))
        item = {
            "code": _metric(name),
            "name": name,
            "header_path": [name],
            "provenance": [_ref(header)],
        }
        role = "total"
    else:
        financial = False
        labels = [
            other
            for column, other in sorted(row.items())
            if kind == "unknown"
            and column < cell["column_index"]
            and other.get("data_type") == "string"
            and _text(other)
        ]
        if labels:
            label = labels[-1]
            item = {
                "code": _metric(_text(label)),
                "name": _text(label),
                "header_path": [_text(label)],
                "provenance": [_ref(label)],
            }
            financial = item["code"] is not None
        role = "reference" if kind == "unknown" else "unknown"
    return {
        "organization": organization,
        "item": item,
        "dimensions": dimensions,
        "role": role,
        "financial": financial,
    }


def _gap_relationships(records: list[dict]) -> list[dict]:
    """동일 조직·기간에서 GAP 지표의 원본과 명시 제외값을 연결한다.

    Args:
        records: 현재 워크북의 표준 레코드.
    Returns:
        값 계산 없이 원본 지표 참조와 관계 확인 상태를 담은 목록.
    Raises:
        없음. 내부 생성 표준 레코드를 전제로 한다.
    Side Effects:
        없음. 영업이익에 Contingency를 더하거나 빼지 않는다.
    """
    groups: dict[tuple, dict[str, list[dict]]] = {}
    for record in records:
        metric = record["item"].get("metric_code")
        if metric not in {"OP", "CONTINGENCY", "OP_EX_CONTINGENCY"}:
            continue
        period = record["period"]
        key = (
            record["source"]["sheet"],
            record["organization"]["code"],
            period["year"],
            period["month"],
            period["aggregation"],
            period["scenario"],
        )
        groups.setdefault(key, {}).setdefault(metric, []).append(record)
    relationships = []
    for group in groups.values():
        explicit = group.get("OP_EX_CONTINGENCY", [])
        relationships.append(
            {
                "status": "explicit_excluded_value" if explicit else "unresolved",
                "preferred_record_ids": [record["record_id"] for record in explicit],
                "metric_records": {
                    metric: [record["record_id"] for record in values]
                    for metric, values in group.items()
                },
                "formulas": [
                    {"source": record["source"], "formula": record["formula"]}
                    for values in group.values()
                    for record in values
                    if record["formula"]
                ],
                "calculation_performed": False,
                "reason": "원본의 명시적 제외값 사용; 추가 조정 없음"
                if explicit
                else "Contingency 원본 및 반영 방향이 확인되지 않아 제외 영업이익을 계산하지 않음",
                "rule_id": "GAP_PREFER_EXPLICIT_NO_AUTOMATIC_ADJUSTMENT",
            }
        )
    return relationships


def map_workbook(workbook: dict, metadata: dict | None = None) -> dict:
    """추출 셀 전체 후보를 출처가 연결된 표준 레코드로 변환한다.

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
    """
    metadata = metadata or {}
    submission = _submission(metadata)
    sheets = workbook["sheets"]
    indices = [_index(sheet) for sheet in sheets]
    layouts = [_layout(sheet, index) for sheet, index in zip(sheets, indices)]
    _periods_with_cross_evidence(layouts)
    registry = _organization_registry(indices)
    common_org = _workbook_organization(layouts, indices)
    workbook_role = (
        "consolidated_reference"
        if common_org["scope"] == "multi_organization"
        and any(layout["kind"] == "organization_summary" for layout in layouts)
        else "organization_workbook"
        if common_org["scope"] == "organization"
        else "unknown"
    )
    if common_org["scope"] == "unknown" and (
        metadata.get("organization_code") or metadata.get("organization_name")
    ):
        common_org = {
            "code": metadata.get("organization_code"),
            "name": metadata.get("organization_name"),
            "original_name": None,
            "scope": "organization",
            "provenance": [{"metadata": "organization"}],
        }
    records, diagnostics, sheet_summaries = [], [], []
    ratio_formats = {}
    file = workbook["file"]
    for sheet, index, layout in zip(sheets, indices, layouts):
        sheet_records = []
        unit, unit_sources = _unit(index)
        statement_items = (
            _statement_items(index, layout)
            if layout["kind"] == "financial_statement"
            else {}
        )
        for cell in sheet["cells"]:
            if not cell.get("has_content", True):
                continue
            context = _row_context(
                cell, index, layout, common_org, registry, statement_items
            )
            context["item"]["metric_code"] = _metric(context["item"]["name"] or "")
            candidate = (
                cell.get("data_type") == "number" or cell.get("formula") is not None
            )
            candidate = candidate or (
                context["financial"] and cell.get("value") is not None
            )
            if not candidate:
                continue
            financial = context["financial"]
            period = (
                dict(layout["periods"].get(cell["column_index"], _period([])))
                if financial
                else _period([])
            )
            source_unit = unit if financial else None
            formula = cell.get("formula")
            rules = ["MAP_" + layout["kind"].upper()]
            unit_evidence = list(unit_sources) if financial else []
            if context["role"] == "subtotal" and layout["kind"].endswith("_detail"):
                expression = (formula or {}).get("text") or ""
                if unit == "천원" and re.search(r"/\s*100000(?:\.0*)?\s*$", expression):
                    source_unit = "억원"
                    unit_evidence.append(_ref(cell))
                    rules.append("UNIT_SUBTOTAL_THOUSAND_TO_HUNDRED_MILLION")
                else:
                    source_unit = None
                    rules.append("UNIT_SUBTOTAL_CONVERSION_UNRESOLVED")
            value_kind = (
                "amount"
                if financial
                else "identifier"
                if layout["kind"] in {"expense_detail", "revenue_detail"}
                else "number"
            )
            number_format = cell.get("number_format")
            if number_format not in ratio_formats:
                ratio_formats[number_format] = inspect_number_format(number_format)[
                    "percent_display"
                ]
            if financial and ratio_formats[number_format]:
                value_kind, source_unit = "ratio", None
                context["role"] = "ratio"
            normalization = normalize_number(
                cell.get("value"),
                data_type=cell.get("data_type", "blank"),
                value_kind=value_kind,
                unit=source_unit,
                number_format=cell.get("number_format"),
                locale=metadata.get("locale"),
            )
            # 표시 형식 전체는 원본 스타일에 이미 보존되므로 값마다 반복하지 않는다.
            normalization.pop("format_analysis", None)
            if formula and formula.get("cached_status") in {"missing", "error"}:
                normalization["normalized_value"] = None
                normalization["status"] = "unresolved"
                rules.append("FORMULA_CACHE_UNAVAILABLE")
            reasons = []
            if not financial:
                reasons.append("재무 값 영역 또는 지표 의미 미확정")
            if context["organization"]["scope"] == "unknown":
                reasons.append("조직 미확정")
            if not context["item"]["name"]:
                reasons.append("항목 미확정")
            if period["year"] is None or period["aggregation"] == "unknown":
                reasons.append("대상 기간 미확정")
            if period["scenario"] == "unknown":
                reasons.append("계획·추정·실적 미확정")
            if financial and value_kind == "amount" and source_unit is None:
                reasons.append("금액 단위 미확정")
            if normalization["status"] != "confirmed":
                reasons.append("값 정규화 미확정 또는 비숫자")
            status = "mapped" if not reasons else "partial" if financial else "unmapped"
            source = {
                "raw_record_id": cell["record_id"],
                "file_id": file["file_id"],
                "source_id": file.get("source_id"),
                "path": file["path"],
                "sheet": sheet["name"],
                "cell": cell["address"],
                "workbook_role": workbook_role,
            }
            record = {
                "record_id": "standard:" + cell["record_id"],
                "source": source,
                "organization": context["organization"],
                "item": context["item"],
                "dimensions": context["dimensions"],
                "period": period,
                "role": {
                    "kind": context["role"],
                    "aggregatable": False,
                    "reason": "집계 시 조직·버전·기간·상세/합계 중복 제거 정책 확정 필요",
                },
                "value": normalization,
                "unit_provenance": unit_evidence,
                "submission": dict(submission),
                "mapping": {
                    "status": status,
                    "rule_ids": rules
                    + ([period["rule_id"]] if period.get("rule_id") else []),
                    "reasons": reasons,
                },
                "formula": formula,
                "raw_value": cell.get("value"),
                "data_type": cell.get("data_type"),
            }
            record["style_reference"] = {
                "style_id": cell.get("style_id"),
                "number_format": cell.get("number_format"),
                "raw_record_id": cell["record_id"],
            }
            sheet_records.append(record)
        counts = Counter(record["mapping"]["status"] for record in sheet_records)
        sheet_summaries.append(
            {
                "sheet": sheet["name"],
                "layout": layout["kind"],
                "records": len(sheet_records),
                "mapped": counts["mapped"],
                "partial": counts["partial"],
                "unmapped": counts["unmapped"],
            }
        )
        records.extend(sheet_records)
    if submission["status"] == "unknown":
        diagnostics.append(
            {
                "code": "SUBMISSION_METADATA_UNKNOWN",
                "severity": "info",
                "message": "제출 월과 버전은 파일에서 확정하지 못했으며 자동 발급하지 않았습니다.",
            }
        )
    gap = _gap_relationships(records)
    if not any(
        record["item"].get("metric_code") == "CONTINGENCY" for record in records
    ):
        diagnostics.append(
            {
                "code": "CONTINGENCY_NOT_IDENTIFIED",
                "severity": "info",
                "message": "원본 항목에서 Contingency를 식별하지 못해 제외 영업이익을 계산하지 않았습니다.",
            }
        )
    counts = Counter(record["mapping"]["status"] for record in records)
    return {
        "records": records,
        "diagnostics": diagnostics,
        "summary": {
            "records": len(records),
            "mapped": counts["mapped"],
            "partial": counts["partial"],
            "unmapped": counts["unmapped"],
            "workbook_role": workbook_role,
            "sheets": sheet_summaries,
        },
        "gap_relationships": gap,
    }
