"""원본 OOXML 및 메모리 경계 사례로 셀·수식·구조 보존을 검증한다."""

import hashlib
import unittest
from io import BytesIO
from zipfile import ZipFile

from helpers import PROJECT_ROOT, external_workbooks

from excel_parser.extraction import _extract_sheet, extract_workbook
from excel_parser.ooxml import read_styles
from excel_parser.validation import validate_extraction

_SHEET_XML = """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1:M4"/><cols><col min="1" max="1" hidden="1"/><col min="13" max="13" hidden="1"/></cols>
<sheetData>
<row r="1"><c r="A1" s="1"><v>-100</v></c><c r="B1" s="2"><v>100</v></c>
<c r="C1" s="3"><v>0</v></c><c r="D1" s="4"><v>0.125</v></c>
<c r="E1"><f>SUM(A1:B1)</f><v/></c><c r="F1" t="e"><f>1/0</f><v>#DIV/0!</v></c>
<c r="G1"><f t="shared" si="0" ref="G1:G2">A1+1</f><v>2</v></c>
<c r="H1" s="5"><v>45292</v></c><c r="I1" t="b"><v>1</v></c>
<c r="J1" t="inlineStr"><is><t>00123</t></is></c><c r="K1" s="1"/></row>
<row r="2" hidden="true"><c r="A2"><v>4</v></c><c r="B2" t="d"><v>2026-11-10</v></c>
<c r="C2" t="e"><v>#N/A</v></c><c r="D2" t="str"><f>IF(A1=0,&quot;&quot;,&quot;x&quot;)</f><v/></c>
<c r="G2"><f t="shared" si="0"/><v>3</v></c><c r="L2"><f>[Book.xlsx]S!A1</f></c></row>
<row r="3"><c r="A3" s="1"/></row><row r="4" hidden="1"/>
</sheetData><mergeCells count="1"><mergeCell ref="J1:K1"/></mergeCells>
<conditionalFormatting sqref="B1"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>
</worksheet>"""
_STYLES_XML = """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="4"><numFmt numFmtId="164" formatCode="0;0"/><numFmt numFmtId="165" formatCode="[Red](0)"/><numFmt numFmtId="166" formatCode="0;-0;&quot;-&quot;"/><numFmt numFmtId="167" formatCode="0.0%"/></numFmts>
<cellXfs count="6"><xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="165"/><xf numFmtId="166"/><xf numFmtId="167"/><xf numFmtId="14"/></cellXfs>
</styleSheet>"""


def _memory_sheet() -> dict:
    """실제 저장값과 서식이 다른 OOXML 경계 사례를 메모리에서 만든다.

    Args:
        없음.
    Returns:
        메모리 ZIP에서 추출한 한 시트 결과.
    Raises:
        Exception: XML 구성 또는 추출이 실패한 경우 테스트를 실패시킨다.
    Side Effects:
        BytesIO만 사용하며 샘플 파일·임시 폴더를 생성하지 않는다.
    """
    stream = BytesIO()
    with ZipFile(stream, "w") as archive:
        archive.writestr("xl/worksheets/sheet1.xml", _SHEET_XML)
        archive.writestr("xl/styles.xml", _STYLES_XML)
        archive.writestr(
            "xl/worksheets/_rels/sheet1.xml.rels",
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>""",
        )
        archive.writestr(
            "xl/comments1.xml",
            """<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Fixture</author></authors><commentList><comment ref="A3" authorId="0"><text><t>내용 없는 셀의 메모</t></text></comment><comment ref="M4" authorId="0"><text><t>숨김 위치의 메모</t></text></comment></commentList></comments>""",
        )
    stream.seek(0)
    with ZipFile(stream) as archive:
        styles = read_styles(archive, "xl/styles.xml")
        return _extract_sheet(
            archive,
            {
                "name": "경계",
                "sheet_id": "1",
                "state": "hidden",
                "part": "xl/worksheets/sheet1.xml",
            },
            {"source_id": "memory", "file_id": "memory"},
            [],
            styles,
        )


class ExtractionBoundaryTests(unittest.TestCase):
    """저장 숫자·캐시 상태·숨김·병합을 파일 생성 없이 검증한다."""

    def test_stored_values_ignore_visual_signs(self) -> None:
        """색상·괄호·대시·퍼센트 서식이 원본 저장값을 바꾸지 않음을 확인한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 독립 OOXML 리터럴과 값이 다른 경우.
        Side Effects:
            메모리에서만 동작한다.
        """
        cells = {cell["address"]: cell for cell in _memory_sheet()["cells"]}
        for address, value in {
            "A1": "-100",
            "B1": "100",
            "C1": "0",
            "D1": "0.125",
        }.items():
            with self.subTest(address=address):
                self.assertEqual(cells[address]["value"], value)
                self.assertEqual(cells[address]["raw_lexeme"], value)
        self.assertEqual(cells["H1"]["data_type"], "date")
        self.assertEqual(cells["H1"]["value"], "45292")
        self.assertIs(cells["I1"]["value"], True)
        self.assertEqual(cells["J1"]["value"], "00123")

    def test_formula_cache_and_shared_formula_provenance(self) -> None:
        """수식의 캐시 누락·오류·공유 기준·외부 참조가 구분됨을 확인한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 수식 상태나 원문이 사라진 경우.
        Side Effects:
            메모리에서만 동작한다. 수식을 계산하지 않는다.
        """
        cells = {cell["address"]: cell for cell in _memory_sheet()["cells"]}
        self.assertEqual(cells["E1"]["formula"]["cached_status"], "missing")
        self.assertIsNone(cells["E1"]["cached_value"])
        self.assertEqual(cells["D2"]["formula"]["cached_status"], "missing")
        self.assertEqual(cells["F1"]["formula"]["cached_status"], "error")
        self.assertEqual(cells["F1"]["cached_value"], "#DIV/0!")
        self.assertEqual(cells["G2"]["formula"]["shared_base"]["address"], "G1")
        self.assertIsNone(cells["G2"]["formula"]["raw_text"])
        self.assertFalse(cells["G2"]["formula"]["recalculated"])
        self.assertTrue(cells["L2"]["formula"]["external_reference"])

    def test_sparse_merges_hidden_metadata_and_comments(self) -> None:
        """병합 빈 셀에 값을 복제하지 않고 숨김·메모·아이콘 정보를 보존한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 구조·메모 또는 숨김 상태가 보존되지 않은 경우.
        Side Effects:
            메모리에서만 동작한다.
        """
        sheet = _memory_sheet()
        cells = {cell["address"]: cell for cell in sheet["cells"]}
        self.assertNotIn("K1", cells)
        self.assertTrue(
            any(
                span["row"] == 1 and span["start_column"] <= 11 <= span["end_column"]
                for span in sheet["blank_spans"]
            )
        )
        self.assertEqual(cells["J1"]["merge"], {"range": "J1:K1", "anchor": "J1"})
        self.assertTrue(cells["A2"]["row_hidden"])
        self.assertTrue(cells["A2"]["column_hidden"])
        self.assertTrue(cells["A2"]["sheet_hidden"])
        self.assertEqual(cells["A3"]["comments"][0]["text"], "내용 없는 셀의 메모")
        self.assertFalse(cells["A3"]["has_content"])
        self.assertTrue(cells["M4"]["annotation_only"])
        self.assertTrue(cells["M4"]["row_hidden"])
        self.assertTrue(cells["M4"]["column_hidden"])
        self.assertIn("3Arrows", sheet["conditional_formatting"][0]["raw_xml"])
        self.assertFalse(sheet["conditional_formatting"][0]["evaluated"])


class ActualExtractionTests(unittest.TestCase):
    """명시적으로 제공된 외부 워크북 전체를 독립 OOXML 점검과 대조한다."""

    def test_all_sheets_cells_formulas_and_notes_are_preserved(self) -> None:
        """외부 입력 전체의 추출 결과와 실행 전후 원본 해시를 점검한다.

        Args:
            없음. EXCEL_PARSER_TEST_DATA의 파일 목록을 사용한다.
        Returns:
            없음.
        Raises:
            unittest.SkipTest: 외부 데이터 경로를 지정하지 않은 경우.
            AssertionError: 독립 OOXML 비교 또는 원본 SHA-256이 다른 경우.
            Exception: 원본 읽기·추출·독립 검증 오류를 그대로 전파한다.
        Side Effects:
            명시된 xlsx/xlsm 파일을 순차적으로 읽는다. 원본을 수정하거나
            출력 파일·임시 파일·캐시를 만들지 않는다.
        """
        for path in external_workbooks():
            with self.subTest(workbook=path.name):
                with path.open("rb") as stream:
                    before = hashlib.file_digest(stream, "sha256").hexdigest()
                try:
                    workbook = extract_workbook(path, root=PROJECT_ROOT)
                    report = validate_extraction(path, workbook)
                    self.assertTrue(report["passed"], report["issues"][:3])
                    self.assertEqual(workbook["file"]["file_id"], before)
                finally:
                    with path.open("rb") as stream:
                        after = hashlib.file_digest(stream, "sha256").hexdigest()
                    self.assertEqual(after, before, "검증 중 원본 내용이 바뀌었습니다.")


if __name__ == "__main__":
    unittest.main()
