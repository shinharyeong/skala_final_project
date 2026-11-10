"""합성 재무 표와 선택적 외부 데이터로 의미 매핑을 검증한다.

원본은 읽기 전용이며 테스트 파일·폴더·캐시를 생성하지 않는다.
"""

import re
import unittest
from copy import deepcopy
from decimal import Decimal

from helpers import PROJECT_ROOT, external_workbooks

from excel_parser.extraction import extract_workbook
from excel_parser.mapping import map_workbook
from excel_parser.validation import validate_provenance


def memory_workbook(
    rows: list[tuple[str, object, str]], *, missing_formula: bool = False
) -> dict:
    """파일 없이 경계 검증용 추출 계약을 메모리에 만든다.

    Args:
        rows: 셀 주소, 저장값, 추출 자료형 목록.
        missing_formula: B5에 캐시 없는 수식을 추가할지 여부.
    Returns:
        하나의 시트를 가진 추출 구조 사전.
    Raises:
        ValueError: 검증 코드의 셀 주소가 잘못된 경우.
    Side Effects:
        없음. 디스크에 샘플 워크북을 만들지 않는다.
    """
    cells = []
    for address, value, data_type in rows:
        match = re.fullmatch(r"([A-Z]+)(\d+)", address)
        if match is None:
            raise ValueError(address)
        column = 0
        for letter in match.group(1):
            column = column * 26 + ord(letter) - 64
        cells.append(
            {
                "record_id": "memory/1/" + address,
                "sheet": "검증",
                "sheet_id": "1",
                "address": address,
                "row": int(match.group(2)),
                "column": match.group(1),
                "column_index": column,
                "value": value,
                "data_type": data_type,
                "formula": None,
                "number_format": "General",
                "has_content": True,
            }
        )
    if missing_formula:
        cells.append(
            {
                "record_id": "memory/1/B5",
                "sheet": "검증",
                "sheet_id": "1",
                "address": "B5",
                "row": 5,
                "column": "B",
                "column_index": 2,
                "value": None,
                "data_type": "number",
                "number_format": "General",
                "has_content": True,
                "formula": {"text": "=B2-B3", "cached_status": "missing"},
            }
        )
    return {
        "file": {
            "file_id": "memory",
            "source_id": "memory-source",
            "path": "memory.xlsx",
        },
        "sheets": [{"name": "검증", "sheet_id": "1", "cells": cells, "merges": []}],
    }


class MappingTests(unittest.TestCase):
    """원본 데이터와 독립적인 합성 표 및 경계 사례를 확인한다."""

    @classmethod
    def setUpClass(cls) -> None:
        """천원 상세와 억원 소계, 월별 근거 표를 합성해 준비한다.

        Args:
            없음. 아래의 독립적인 합성 리터럴만 사용한다.
        Returns:
            없음. raw, result, records 클래스 속성을 준비한다.
        Raises:
            Exception: 합성 계약의 매핑에 실패한 경우.
        Side Effects:
            메모리 사전만 생성하며 외부 파일을 읽지 않는다.
        """
        cls.raw = memory_workbook(
            [
                ("A1", "단위: 천원", "string"),
                ("F2", "0.001", "number"),
                ("A3", "조직명", "string"),
                ("B3", "관리분류", "string"),
                ("C3", "중분류", "string"),
                ("D3", "소분류", "string"),
                ("E3", "사업부코드", "string"),
                ("F3", "2030 실적/추정", "string"),
                ("I3", "예산Code", "string"),
                ("F4", "연간", "string"),
                ("G4", "1월", "string"),
                ("H4", "2월", "string"),
                ("A5", "가상 조직", "string"),
                ("B5", "매출", "string"),
                ("C5", "상품", "string"),
                ("D5", "서비스", "string"),
                ("E5", "DEMO_ORG", "string"),
                ("F5", "100", "number"),
                ("G5", "40", "number"),
                ("H5", "60", "number"),
                ("I5", "12345", "number"),
            ]
        )
        by_address = {cell["address"]: cell for cell in cls.raw["sheets"][0]["cells"]}
        by_address["F2"]["formula"] = {
            "text": "=SUBTOTAL(9,F5:F5)/100000",
            "cached_status": "stored_unverified",
        }
        by_address["A5"]["formula"] = {
            "text": '=CONCAT("가상 ","조직")',
            "cached_status": "stored_unverified",
        }
        evidence = memory_workbook(
            [
                ("A1", "단위: 천원", "string"),
                ("B2", "2030년", "string"),
                ("B3", "실적", "string"),
                ("C3", "실적", "string"),
                ("D3", "추정", "string"),
                ("B4", "연간", "string"),
                ("C4", "1Q", "string"),
                ("D4", "1Q", "string"),
                ("B5", "연간", "string"),
                ("C5", "1월", "string"),
                ("D5", "2월", "string"),
                ("A6", "영업이익", "string"),
                ("B6", "100", "number"),
                ("C6", "40", "number"),
                ("D6", "60", "number"),
            ]
        )["sheets"][0]
        evidence.update(
            name="기간 근거", sheet_id="2", merges=[{"range": "B2:D2", "anchor": "B2"}]
        )
        for cell in evidence["cells"]:
            cell.update(
                sheet="기간 근거", sheet_id="2", record_id="memory/2/" + cell["address"]
            )
        cls.raw["sheets"].append(evidence)
        cls.result = map_workbook(cls.raw)
        cls.records = {
            (record["source"]["sheet"], record["source"]["cell"]): record
            for record in cls.result["records"]
        }

    def test_all_numeric_and_formula_sources_preserved(self) -> None:
        """모든 합성 숫자·수식 후보의 출처와 원문을 확인한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 후보가 빠지거나 원문·수식이 바뀐 경우.
        Side Effects:
            메모리 결과만 조회한다.
        """
        originals = {
            cell["record_id"]: cell
            for sheet in self.raw["sheets"]
            for cell in sheet["cells"]
        }
        expected = {
            key
            for key, cell in originals.items()
            if cell["data_type"] == "number" or cell["formula"] is not None
        }
        actual = {
            record["source"]["raw_record_id"] for record in self.result["records"]
        }
        self.assertTrue(expected <= actual)
        self.assertEqual(len(actual), len(self.result["records"]))
        for record in self.result["records"]:
            original = originals[record["source"]["raw_record_id"]]
            self.assertEqual(record["raw_value"], original["value"])
            self.assertEqual(record["formula"], original["formula"])

    def test_detail_amount_and_organization(self) -> None:
        """합성 상세 40천원과 명시 조직 코드·기간을 독립 기대값과 대조한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 금액 환산이나 명시 조직·기간이 다른 경우.
        Side Effects:
            메모리 결과만 조회한다.
        """
        record = self.records[("검증", "G5")]
        self.assertEqual(record["raw_value"], "40")
        self.assertEqual(record["value"]["normalized_value"], Decimal("40000"))
        self.assertEqual(record["organization"]["code"], "DEMO_ORG")
        self.assertEqual(record["period"]["year"], 2030)
        self.assertEqual(record["period"]["month"], 1)
        self.assertEqual(record["role"]["kind"], "detail")

    def test_subtotal_has_different_unit_than_detail(self) -> None:
        """천원 합계를 /100000으로 나눈 합성 억원 소계를 원으로 환산한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 소계에 상세 단위가 중복 적용된 경우.
        Side Effects:
            메모리 결과만 조회한다.
        """
        record = self.records[("검증", "F2")]
        self.assertEqual(record["value"]["source_unit"], "억원")
        self.assertEqual(record["value"]["normalized_value"], Decimal("100000"))
        self.assertEqual(record["role"]["kind"], "subtotal")
        self.assertIn("/100000", record["formula"]["text"])

    def test_mixed_year_and_month_scenario(self) -> None:
        """명시된 합성 월별 헤더로 혼합 연간과 실적·추정 월을 구분한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 명시된 시나리오와 교차 시트 근거가 다른 경우.
        Side Effects:
            메모리 결과만 조회한다.
        """
        self.assertEqual(
            self.records[("기간 근거", "B6")]["period"]["scenario"], "mixed"
        )
        self.assertEqual(self.records[("검증", "G5")]["period"]["scenario"], "actual")
        estimated = self.records[("검증", "H5")]["period"]
        self.assertEqual(estimated["scenario"], "estimate")
        self.assertTrue(
            any(ref["sheet"] == "기간 근거" for ref in estimated["provenance"])
        )

    def test_identifier_and_formula_text_are_not_amounts(self) -> None:
        """합성 숫자 코드와 조직명 수식의 문자열 캐시를 금액과 구분한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 식별자 값이 금액으로 정규화된 경우.
        Side Effects:
            메모리 결과만 조회한다.
        """
        for address in ("I5", "A5"):
            record = self.records[("검증", address)]
            self.assertIsNone(record["value"]["normalized_value"])
            self.assertEqual(record["value"]["value_kind"], "identifier")

    def test_missing_submission_is_not_inferred(self) -> None:
        """대상 연도만으로 제출 월이나 버전을 만들어내지 않는다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 명시되지 않은 제출 정보가 생성된 경우.
        Side Effects:
            메모리 결과만 조회한다.
        """
        submission = self.result["records"][0]["submission"]
        self.assertEqual(submission["status"], "unknown")
        self.assertIsNone(submission["version_raw"])
        self.assertIsNone(submission["submission_month"])

    def test_version_string_keeps_original_precision(self) -> None:
        """V1.01 원문을 보존하고 소수형 버전 입력을 거부한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 버전 원문이나 회차·개정 분리가 틀린 경우.
        Side Effects:
            없음. 메모리 워크북만 사용한다.
        """
        raw = memory_workbook([("B2", "1", "number")])
        record = map_workbook(
            raw, {"submission_version": "V1.01", "submission_month": "2026-11"}
        )["records"][0]
        self.assertEqual(record["submission"]["version_raw"], "V1.01")
        self.assertEqual(record["submission"]["cycle"], 1)
        self.assertEqual(record["submission"]["revision"], 1)
        with self.assertRaises(ValueError):
            map_workbook(raw, {"submission_version": 1.01})

    def test_explicit_excluded_profit_is_not_adjusted(self) -> None:
        """원본의 명시 제외 영업이익을 우선 연결하고 추가 계산을 막는다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 명시 제외값 110이 다시 조정되거나 원본이 변경된 경우.
        Side Effects:
            없음. 메모리 워크북만 사용한다.
        """
        raw = memory_workbook(
            [
                ("A1", "단위: 원", "string"),
                ("A2", "영업이익", "string"),
                ("B2", "100", "number"),
                ("A3", "Contingency", "string"),
                ("B3", "10", "number"),
                ("A4", "Contingency 제외 영업이익", "string"),
                ("B4", "110", "number"),
            ]
        )
        before = deepcopy(raw)
        result = map_workbook(raw)
        excluded = next(
            record
            for record in result["records"]
            if record["item"]["code"] == "OP_EX_CONTINGENCY"
        )
        self.assertEqual(excluded["value"]["normalized_value"], Decimal("110"))
        relation = result["gap_relationships"][0]
        self.assertEqual(relation["status"], "explicit_excluded_value")
        self.assertEqual(relation["preferred_record_ids"], [excluded["record_id"]])
        self.assertFalse(relation["calculation_performed"])
        self.assertEqual(raw, before)

    def test_missing_formula_cache_remains_unresolved(self) -> None:
        """수식 캐시가 없으면 수식과 출처를 남기고 0으로 대체하지 않는다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 캐시가 없는 값이 확정되거나 누락된 경우.
        Side Effects:
            없음.
        """
        raw = memory_workbook([("A5", "영업이익", "string")], missing_formula=True)
        record = map_workbook(raw)["records"][0]
        self.assertIsNone(record["value"]["normalized_value"])
        self.assertEqual(record["value"]["status"], "unresolved")
        self.assertEqual(record["formula"]["text"], "=B2-B3")

    def test_merged_headers_keep_anchor_provenance(self) -> None:
        """병합 헤더의 기준 셀을 연결하고 원본 빈 헤더를 생성하지 않는다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 상속 헤더 출처가 없거나 원본이 변경된 경우.
        Side Effects:
            없음. 메모리 워크북만 사용한다.
        """
        raw = memory_workbook(
            [
                ("A1", "단위: 원", "string"),
                ("B2", "2027년", "string"),
                ("B3", "계획", "string"),
                ("B4", "1Q", "string"),
                ("B5", "1월", "string"),
                ("C5", "2월", "string"),
                ("A6", "영업이익", "string"),
                ("B6", "10", "number"),
                ("C6", "20", "number"),
            ]
        )
        raw["sheets"][0]["merges"] = [
            {"range": f"B{row}:C{row}", "anchor": f"B{row}"} for row in (2, 3, 4)
        ]
        before = deepcopy(raw)
        result = map_workbook(raw)
        record = next(
            record for record in result["records"] if record["source"]["cell"] == "C6"
        )
        self.assertEqual(record["period"]["year"], 2027)
        self.assertEqual(record["period"]["month"], 2)
        inherited = [
            ref for ref in record["period"]["provenance"] if ref.get("inheritance")
        ]
        self.assertEqual(inherited[0]["cell"], "B2")
        self.assertEqual(inherited[0]["inheritance"]["target_column"], 3)
        self.assertEqual(raw, before)

    def test_recognized_table_header_is_not_a_metric_value(self) -> None:
        """복합 헤더의 영업이익 문구 때문에 옆 헤더를 금액으로 읽지 않는다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 제목만 있는 표에 표준 값 레코드가 만들어진 경우.
        Side Effects:
            없음. 메모리 워크북만 사용한다.
        """
        raw = memory_workbook(
            [
                ("A5", "사업부코드", "string"),
                ("B5", "사업부명", "string"),
                ("C5", "2025 실적 매출", "string"),
                ("D5", "2025 실적 영업비용", "string"),
                ("E5", "2025 실적 영업이익", "string"),
                ("F5", "2026 계획 매출", "string"),
            ]
        )
        self.assertEqual(map_workbook(raw)["records"], [])


class ExternalMappingIntegrationTests(unittest.TestCase):
    """환경변수로 명시한 별도 원본에서 매핑 출처 계약을 확인한다."""

    def test_external_records_preserve_sources(self) -> None:
        """외부 입력 전체의 후보 포함·원문·수식·근거 연결을 검사한다.

        Args:
            없음. EXCEL_PARSER_TEST_DATA를 사용한다.
        Returns:
            없음.
        Raises:
            unittest.SkipTest: 외부 데이터 경로가 지정되지 않은 경우.
            AssertionError: 출처 계약 검증이 실패하거나 지정 경로가 잘못된 경우.
        Side Effects:
            지정한 외부 원본을 읽기만 하며 파일을 생성하지 않는다.
        """
        for path in external_workbooks():
            with self.subTest(path=path.name):
                raw = extract_workbook(path, root=PROJECT_ROOT)
                result = validate_provenance(raw, map_workbook(raw))
                self.assertTrue(result["passed"], result["issues"][:3])


if __name__ == "__main__":
    unittest.main()
