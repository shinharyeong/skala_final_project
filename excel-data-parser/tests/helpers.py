"""별도 관리 데이터의 선택적 입력과 합성 OOXML 메모리 fixture를 제공한다."""

import hashlib
import os
import unittest
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from xml.sax.saxutils import quoteattr
from zipfile import ZipFile

from excel_parser.extraction import extract_workbook

PROJECT_ROOT = Path(__file__).resolve().parents[1]
_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_DEFAULT_SHEET = f'''<worksheet xmlns="{_NS}"><dimension ref="A1:C3"/>
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>단위: 원</t></is></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>영업이익</t></is></c><c r="B2"><v>100</v></c></row>
<row r="3"><c r="A3" t="inlineStr"><is><t>비교항목</t></is></c><c r="B3"><f>B2*2</f><v>200</v></c><c r="C3" s="0"/></row>
</sheetData></worksheet>'''


def external_workbooks() -> list[Path]:
    """환경변수로 명시한 별도 데이터 폴더의 OOXML 파일을 찾는다.

    Args:
        없음. EXCEL_PARSER_TEST_DATA 환경변수를 사용한다.
    Returns:
        해당 폴더와 하위 폴더의 xlsx/xlsm 경로 목록.
    Raises:
        unittest.SkipTest: 환경변수를 지정하지 않은 경우.
        AssertionError: 명시한 폴더가 없거나 지원 파일을 찾지 못한 경우.
    Side Effects:
        명시한 외부 폴더의 경로 정보만 읽고 파일을 생성하지 않는다.
    """
    configured = os.environ.get("EXCEL_PARSER_TEST_DATA")
    if not configured:
        raise unittest.SkipTest(
            "별도 데이터 검증은 EXCEL_PARSER_TEST_DATA 지정 시 실행합니다."
        )
    directory = Path(configured).expanduser().resolve()
    if not directory.is_dir():
        raise AssertionError(
            "EXCEL_PARSER_TEST_DATA는 존재하는 데이터 폴더여야 합니다."
        )
    files = sorted(
        path
        for path in directory.rglob("*")
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xlsm"}
        and not path.name.startswith("~$")
    )
    if not files:
        raise AssertionError("별도 데이터 폴더에 xlsx/xlsm 파일이 없습니다.")
    return files


def make_package(sheets: dict[str, str] | None = None) -> bytes:
    """직접 정의한 XML로 최소 OOXML 패키지를 메모리에서 만든다.

    Args:
        sheets: 시트 이름과 XML 본문. 생략하면 합성 값 100/200을 사용한다.
    Returns:
        디스크 파일이 아닌 ZIP 바이트.
    Raises:
        OSError: 메모리 ZIP 작성이 실패한 경우.
    Side Effects:
        BytesIO만 사용하며 파일·폴더를 생성하지 않는다.
    """
    sheets = sheets or {"검증표": _DEFAULT_SHEET}
    entries = []
    relations = []
    stream = BytesIO()
    with ZipFile(stream, "w") as archive:
        for index, (name, document) in enumerate(sheets.items(), start=1):
            entries.append(
                f'<sheet name={quoteattr(name)} sheetId="{index}" r:id="rId{index}"/>'
            )
            relations.append(
                f'<Relationship Id="rId{index}" Type="{_REL}/worksheet" Target="worksheets/sheet{index}.xml"/>'
            )
            archive.writestr(f"xl/worksheets/sheet{index}.xml", document)
        archive.writestr(
            "xl/workbook.xml",
            f'<workbook xmlns="{_NS}" xmlns:r="{_REL}"><sheets>{"".join(entries)}</sheets></workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(relations)
            + "</Relationships>",
        )
    return stream.getvalue()


def open_memory_archive(package: bytes) -> ZipFile:
    """합성 OOXML 바이트의 독립 읽기 스트림을 연다.

    Args:
        package: make_package가 만든 ZIP 바이트.
    Returns:
        호출자가 닫아야 하는 ZipFile 객체.
    Raises:
        BadZipFile: 잘못된 ZIP 바이트인 경우.
    Side Effects:
        메모리 스트림만 열며 디스크에는 접근하지 않는다.
    """
    return ZipFile(BytesIO(package))


def extract_memory(package: bytes, *, name: str = "synthetic.xlsx") -> dict:
    """파일 경계만 대체해 실제 추출기를 합성 ZIP 바이트로 실행한다.

    Args:
        package: 메모리 OOXML ZIP 바이트.
        name: 출처 식별용 가상 파일명. 실제 파일을 만들지 않는다.
    Returns:
        extract_workbook의 실제 추출 결과.
    Raises:
        Exception: 패키지 해석 오류를 그대로 전파한다.
    Side Effects:
        파일 해시·ZIP 열기·크기 조회를 mock으로 대체하고 메모리만 사용한다.
    """
    source = PROJECT_ROOT / "__memory__" / name
    with (
        patch(
            "excel_parser.extraction.file_fingerprint",
            return_value=hashlib.sha256(package).hexdigest(),
        ),
        patch(
            "excel_parser.extraction.ZipFile",
            side_effect=lambda _: open_memory_archive(package),
        ),
        patch.object(Path, "stat", return_value=SimpleNamespace(st_size=len(package))),
    ):
        return extract_workbook(source, root=PROJECT_ROOT)
