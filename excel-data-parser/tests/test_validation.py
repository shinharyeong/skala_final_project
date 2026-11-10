"""독립 검증기의 누락·변조 검출과 선택적 외부 데이터 비교를 확인한다."""

import unittest
from copy import deepcopy
from pathlib import Path
from unittest.mock import patch

from helpers import (
    PROJECT_ROOT,
    external_workbooks,
    extract_memory,
    make_package,
    open_memory_archive,
)

from excel_parser.mapping import map_workbook
from excel_parser.validation import (
    compare_master,
    validate_extraction,
    validate_provenance,
)


class IndependentValidationTests(unittest.TestCase):
    """독립 기대값을 가진 합성 패키지로 검증기의 오류 검출을 확인한다."""

    @classmethod
    def setUpClass(cls) -> None:
        """원본 파일 없이 합성 OOXML과 추출·매핑 결과를 메모리에 준비한다.

        Args:
            없음.
        Returns:
            없음. 클래스 속성에 준비된 결과를 둔다.
        Raises:
            Exception: 합성 패키지 해석에 실패한 경우.
        Side Effects:
            메모리만 사용하며 외부 데이터를 읽지 않는다.
        """
        cls.package = make_package()
        cls.workbook = extract_memory(cls.package)
        cls.mapped = map_workbook(cls.workbook)
        cls.source = PROJECT_ROOT / cls.workbook["file"]["path"]

    def _validate(self, workbook: dict) -> dict:
        """독립 검증기에 생산 결과와 별도의 원본 ZIP 바이트를 전달한다.

        Args:
            workbook: 검사할 추출 결과.
        Returns:
            원본 XML과의 검증 보고서.
        Raises:
            Exception: 원본 ZIP 또는 계약 해석에 실패한 경우.
        Side Effects:
            메모리 ZIP만 읽으며 디스크 파일을 만들지 않는다.
        """
        with patch(
            "excel_parser.validation.ZipFile",
            side_effect=lambda _: open_memory_archive(self.package),
        ):
            return validate_extraction(self.source, workbook)

    def test_unmodified_package_matches_independent_xml(self) -> None:
        """직접 작성한 XML의 다섯 내용 셀과 한 빈 구조 셀을 대조한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 독립 XML 개수나 값이 일치하지 않는 경우.
        Side Effects:
            메모리 패키지만 읽는다.
        """
        report = self._validate(self.workbook)
        self.assertTrue(report["passed"], report["issues"])
        self.assertEqual(report["source_xml_cells"], 6)
        self.assertEqual(report["source_content_cells"], 5)
        self.assertEqual(report["source_formula_cells"], 1)

    def test_standard_records_preserve_all_required_sources(self) -> None:
        """숫자·수식 후보와 보조 근거가 합성 원본을 가리키는지 확인한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 출처 또는 원문이 다르게 연결된 경우.
        Side Effects:
            메모리 결과만 조회한다.
        """
        report = validate_provenance(self.workbook, self.mapped)
        self.assertTrue(report["passed"], report["issues"])
        self.assertEqual(report["required_candidate_count"], 2)
        self.assertEqual(report["covered_candidate_count"], 2)
        self.assertGreater(report["evidence_reference_count"], 0)

    def test_missing_candidate_and_wrong_source_are_detected(self) -> None:
        """표준 후보 누락과 다른 파일 출처로의 변조를 검출한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 원본 후보 누락이나 잘못된 출처를 허용한 경우.
        Side Effects:
            결과를 메모리에서만 복사·변경한다.
        """
        records = deepcopy(self.mapped["records"])
        del records[0]
        records[0]["source"]["file_id"] = "unrelated-file"
        report = validate_provenance(self.workbook, {**self.mapped, "records": records})
        codes = {issue["code"] for issue in report["issues"]}
        self.assertIn("UNMAPPED_NUMERIC_OR_FORMULA_SOURCE", codes)
        self.assertIn("PROVENANCE_VALUE_MISMATCH", codes)

    def test_changed_original_value_and_evidence_are_detected(self) -> None:
        """원문과 존재하지 않는 헤더 근거의 변조를 검출한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 변조된 값이나 근거를 허용한 경우.
        Side Effects:
            메모리 사전만 복사·변경한다.
        """
        records = deepcopy(self.mapped["records"])
        records[0]["raw_value"] = "altered"
        records[0]["unit_provenance"] = [{"raw_record_id": "missing-header"}]
        report = validate_provenance(self.workbook, {**self.mapped, "records": records})
        codes = {issue["code"] for issue in report["issues"]}
        self.assertIn("PROVENANCE_VALUE_MISMATCH", codes)
        self.assertIn("UNKNOWN_EVIDENCE_REFERENCE", codes)

    def test_missing_value_cell_is_detected(self) -> None:
        """한 내용 셀을 삭제하면 독립 원본 주소 대조에서 누락을 찾는다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 삭제한 주소를 누락으로 보고하지 않은 경우.
        Side Effects:
            메모리 결과만 변경한다.
        """
        workbook = deepcopy(self.workbook)
        removed = workbook["sheets"][0]["cells"].pop(0)["address"]
        result = self._validate(workbook)
        self.assertTrue(
            any(
                issue["code"] == "MISSING_CELLS" and removed in issue["cells"]
                for issue in result["issues"]
            )
        )

    def test_changed_cached_value_is_detected(self) -> None:
        """수식의 저장 캐시만 바꿔도 독립 원문 대조에서 검출한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 변조된 수식 캐시를 허용한 경우.
        Side Effects:
            메모리 결과만 변경한다.
        """
        workbook = deepcopy(self.workbook)
        changed = next(
            cell for cell in workbook["sheets"][0]["cells"] if cell["formula"]
        )
        changed["cached_value"] = "999"
        result = self._validate(workbook)
        self.assertTrue(
            any(
                issue["code"] == "CELL_MISMATCH" and "cached_value" in issue["fields"]
                for issue in result["issues"]
            )
        )

    def test_blank_structure_omission_is_detected(self) -> None:
        """빈 구조 구간을 삭제해도 주소 보존 검증이 실패하는지 확인한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 빈 구조 누락을 찾지 못한 경우.
        Side Effects:
            메모리 결과만 변경한다.
        """
        workbook = deepcopy(self.workbook)
        workbook["sheets"][0]["blank_spans"] = []
        result = self._validate(workbook)
        self.assertFalse(result["passed"])
        self.assertTrue(
            any(issue["code"] == "MISSING_CELLS" for issue in result["issues"])
        )

    def test_master_is_optional_and_duplicate_inputs_are_rejected(self) -> None:
        """비교표 부재를 구분하고 동일 입력 경로를 재합산하지 않는다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 적용 불가 또는 중복 처리 상태가 틀린 경우.
        Side Effects:
            메모리 ZIP만 읽으며 파일을 생성하지 않는다.
        """
        with patch(
            "excel_parser.validation.ZipFile",
            side_effect=lambda _: open_memory_archive(self.package),
        ):
            self.assertEqual(compare_master([self.source])["status"], "not_applicable")
        report = compare_master([self.source, self.source])
        self.assertEqual(report["status"], "failed")
        self.assertEqual(report["issues"][0]["code"], "DUPLICATE_INPUT_PATH")

    def test_synthetic_master_uses_independent_source_values(self) -> None:
        """합성 상세 100+200 및 비용 20+30의 비교표를 독립 계산한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 합성 원본의 합계·차감값이 독립 기대값과 다른 경우.
        Side Effects:
            두 패키지를 BytesIO로만 만들고 읽는다.
        """
        namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        revenue = f"""<worksheet xmlns="{namespace}"><sheetData>
<row r="2"><c r="A2" t="inlineStr"><is><t>DEMO</t></is></c><c r="B2"><v>100</v></c></row>
<row r="3"><c r="A3" t="inlineStr"><is><t>DEMO</t></is></c><c r="B3"><v>200</v></c></row>
</sheetData></worksheet>"""
        expense = revenue.replace("<v>100</v>", "<v>20</v>").replace(
            "<v>200</v>", "<v>30</v>"
        )
        summary = f"""<worksheet xmlns="{namespace}"><sheetData><row r="2">
<c r="A2" t="inlineStr"><is><t>DEMO</t></is></c>
<c r="B2"><f>SUMIFS(상세!$B$2:$B$3,상세!$A$2:$A$3,$A2)/1000</f><v>0.3</v></c>
<c r="C2"><f>SUMIFS(비용!$B$2:$B$3,비용!$A$2:$A$3,$A2)/1000</f><v>0.05</v></c>
<c r="D2"><f>B2-C2</f><v>0.25</v></c>
</row></sheetData></worksheet>"""
        source, reference = (
            Path("synthetic-source.xlsx"),
            Path("synthetic-reference.xlsx"),
        )
        packages = {
            source: make_package({"상세": revenue, "비용": expense}),
            reference: make_package(
                {"상세": revenue, "비용": expense, "사업부별집계": summary}
            ),
        }
        with patch(
            "excel_parser.validation.ZipFile",
            side_effect=lambda path: open_memory_archive(packages[path]),
        ):
            result = compare_master([source, reference])
        self.assertEqual(result["status"], "passed", result["issues"])
        self.assertEqual(result["comparison_count"], 3)
        values = {
            item["cell"]: item["independent_value"] for item in result["comparisons"]
        }
        self.assertEqual(values, {"B2": "0.3", "C2": "0.05", "D2": "0.25"})


class ExternalMasterIntegrationTests(unittest.TestCase):
    """명시적으로 제공한 외부 데이터에서만 비교 가능한 집계를 검사한다."""

    def test_external_source_totals_match_reference(self) -> None:
        """별도 입력에 비교표가 있으면 원본 합계와 대조한다.

        Args:
            없음. EXCEL_PARSER_TEST_DATA를 사용한다.
        Returns:
            없음.
        Raises:
            unittest.SkipTest: 외부 데이터나 비교 가능한 표가 없는 경우.
            AssertionError: 외부 원본과 비교표의 값이 다른 경우.
        Side Effects:
            지정된 외부 원본만 읽으며 산출물은 만들지 않는다.
        """
        result = compare_master(external_workbooks())
        if result["status"] == "not_applicable":
            self.skipTest(result.get("reason", "비교 가능한 집계표가 없습니다."))
        self.assertEqual(result["status"], "passed", result["issues"])
        self.assertGreater(result["comparison_count"], 0)


if __name__ == "__main__":
    unittest.main()
