"""OOXML 패키지의 주소·관계·공유 문자열·서식을 읽기 전용으로 해석한다."""

import posixpath
import re
from pathlib import PurePosixPath
from xml.etree import ElementTree as ET
from zipfile import ZipFile

BUILTIN_FORMATS = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}


def local_name(tag: str) -> str:
    """네임스페이스를 제외한 XML 이름을 반환한다.

    Args:
        tag: ElementTree 태그 또는 속성명.
    Returns:
        마지막 닫는 중괄호 뒤의 이름.
    Raises:
        없음. 문자열 입력을 전제로 한다.
    Side Effects:
        없음.
    """
    return tag.rsplit("}", 1)[-1]


def xml_text(element: ET.Element | None) -> str:
    """리치 텍스트의 실제 표시 텍스트를 연결한다.

    Args:
        element: 공유 문자열·inline 문자열·메모 요소 또는 None.
    Returns:
        발음 안내 rPh를 제외한 t 요소의 연결 문자열.
    Raises:
        없음.
    Side Effects:
        없음. 전달받은 트리를 변경하지 않는다.
    """
    if element is None:
        return ""
    if local_name(element.tag) == "t":
        return element.text or ""
    return "".join(
        xml_text(child) for child in element if local_name(child.tag) != "rPh"
    )


def cell_position(address: str) -> tuple[int, int, str]:
    """A1 주소를 1부터 시작하는 행·열 번호로 해석한다.

    Args:
        address: 달러 기호를 포함할 수 있는 셀 주소.
    Returns:
        행 번호, 열 번호, 열 문자 튜플.
    Raises:
        ValueError: 유효한 단일 셀 주소가 아닌 경우.
    Side Effects:
        없음.
    """
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?([1-9]\d*)", address)
    if not match:
        raise ValueError(f"셀 주소를 해석할 수 없습니다: {address}")
    column = match.group(1).upper()
    index = 0
    for character in column:
        index = index * 26 + ord(character) - ord("A") + 1
    return int(match.group(2)), index, column


def range_bounds(reference: str) -> tuple[int, int, int, int]:
    """직사각형 A1 범위의 경계를 반환한다.

    Args:
        reference: 단일 주소 또는 시작:끝 주소.
    Returns:
        시작 행·열, 끝 행·열 번호.
    Raises:
        ValueError: 주소가 잘못된 경우.
    Side Effects:
        없음.
    """
    start, _, end = reference.partition(":")
    row1, col1, _ = cell_position(start)
    row2, col2, _ = cell_position(end or start)
    return row1, col1, row2, col2


def relationships(archive: ZipFile, part: str) -> dict[str, dict]:
    """패키지 내부 관계를 대상 경로와 함께 반환한다.

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
    """
    path = PurePosixPath(part)
    rel_path = (
        str(path.parent / "_rels" / (path.name + ".rels")) if part else "_rels/.rels"
    )
    if rel_path not in archive.namelist():
        return {}
    result = {}
    for element in ET.fromstring(archive.read(rel_path)):
        relationship = dict(element.attrib)
        if relationship.get("TargetMode") != "External":
            target = relationship["Target"]
            resolved = posixpath.normpath(
                target.lstrip("/")
                if target.startswith("/")
                else posixpath.join(str(path.parent), target)
            )
            if resolved.startswith("../") or resolved == "..":
                raise ValueError(f"잘못된 패키지 관계: {target}")
            relationship["part"] = resolved
        result[relationship["Id"]] = relationship
    return result


def read_styles(archive: ZipFile, part: str | None) -> dict:
    """셀 서식과 원본 스타일 XML을 보존한다.

    Args:
        archive: 읽기 전용 ZIP 패키지.
        part: styles 파트 경로. 없으면 None.
    Returns:
        custom_formats, cell_formats, 원본 XML을 담은 사전.
    Raises:
        ET.ParseError: 스타일 XML이 손상된 경우.
    Side Effects:
        스타일 XML만 읽으며 서식을 적용하거나 값을 변경하지 않는다.
    """
    if part is None:
        return {
            "part": None,
            "custom_formats": {},
            "cell_formats": [{"style_id": 0, "number_format": "General"}],
            "raw_xml": None,
        }
    raw = archive.read(part)
    root = ET.fromstring(raw)
    custom = {
        int(node.attrib["numFmtId"]): node.attrib["formatCode"]
        for node in root.findall("./{*}numFmts/{*}numFmt")
    }
    formats = []
    for index, node in enumerate(root.findall("./{*}cellXfs/{*}xf")):
        number_id = int(node.get("numFmtId", "0"))
        formats.append(
            {
                "style_id": index,
                "num_fmt_id": number_id,
                "number_format": custom.get(number_id, BUILTIN_FORMATS.get(number_id)),
                "attributes": dict(node.attrib),
                "raw_xml": ET.tostring(node, encoding="unicode"),
            }
        )
    return {
        "part": part,
        "custom_formats": custom,
        "cell_formats": formats,
        "raw_xml": raw.decode("utf-8-sig"),
    }


def is_date_format(number_format: str | None) -> bool:
    """명시적인 날짜·시간 서식을 보수적으로 탐지한다.

    Args:
        number_format: Excel 표시 형식 또는 None.
    Returns:
        리터럴 제거 뒤 날짜·시간 토큰이 확인되면 True.
    Raises:
        없음.
    Side Effects:
        없음. 일련번호를 날짜나 금액으로 환산하지 않는다.
    """
    if not number_format:
        return False
    cleaned = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", number_format).lower()
    return bool(re.search(r"[ydhs]|m{2,}", cleaned)) and "general" not in cleaned


def read_comments(archive: ZipFile, sheet_part: str) -> dict[str, list[dict]]:
    """시트에 연결된 일반·스레드 메모를 셀 주소별로 보존한다.

    Args:
        archive: 읽기 전용 ZIP 패키지.
        sheet_part: 시트 XML 경로.
    Returns:
        셀 주소별 텍스트·작성자·원본 XML 목록.
    Raises:
        ET.ParseError: 메모 XML이 손상된 경우.
    Side Effects:
        연결된 메모 파트를 읽는다. 외부 서비스에 접근하지 않는다.
    """
    result = {}
    for relationship in relationships(archive, sheet_part).values():
        if (
            "comment" not in relationship.get("Type", "").lower()
            or "part" not in relationship
        ):
            continue
        part = relationship["part"]
        root = ET.fromstring(archive.read(part))
        authors = [node.text or "" for node in root.findall("./{*}authors/{*}author")]
        for node in root.iter():
            if local_name(node.tag) not in {"comment", "threadedComment"}:
                continue
            author_index = int(node.get("authorId", "-1"))
            text_node = node.find("{*}text")
            text = (
                xml_text(text_node)
                if text_node is not None and len(text_node)
                else (text_node.text or "" if text_node is not None else "")
            )
            entry = {
                "part": part,
                "text": text,
                "author": authors[author_index]
                if 0 <= author_index < len(authors)
                else node.get("personId"),
                "attributes": dict(node.attrib),
                "raw_xml": ET.tostring(node, encoding="unicode"),
            }
            result.setdefault(node.attrib["ref"], []).append(entry)
    return result
