"""원본 OOXML 값을 재계산·수정 없이 추출하고 빈 구조는 희소 구간으로 보존한다."""

import hashlib
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from .ooxml import (
    cell_position,
    is_date_format,
    local_name,
    range_bounds,
    read_comments,
    read_styles,
    relationships,
    xml_text,
)


def file_fingerprint(path: Path) -> str:
    """원본 바이트의 SHA-256 식별자를 계산한다.

    Args:
        path: 읽을 파일 경로.
    Returns:
        SHA-256 16진수 문자열.
    Raises:
        OSError: 파일 읽기에 실패한 경우.
    Side Effects:
        파일을 청크 단위로 읽기만 한다.
    """
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _cell_value(
    node: ET.Element, shared: list[str], number_format: str | None
) -> tuple[object, str, str | None]:
    """셀 저장 타입에 따라 원본 값을 손실 없이 읽는다.

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
    """
    stored = node.find("{*}v")
    raw = stored.text if stored is not None else None
    kind = node.get("t", "n")
    if kind == "s":
        try:
            return shared[int(raw)], "string", raw
        except (ValueError, IndexError, TypeError) as error:
            raise ValueError(f"공유 문자열 참조 오류: {node.get('r')}") from error
    if kind == "inlineStr":
        return xml_text(node.find("{*}is")), "string", raw
    if kind == "b":
        return raw == "1" if raw is not None else None, "boolean", raw
    if kind == "e":
        return raw, "error", raw
    if kind == "d":
        return raw, "date", raw
    if kind == "str":
        return raw if raw is not None else "", "string", raw
    if raw is None:
        return None, "blank", raw
    return raw, "date" if is_date_format(number_format) else "number", raw


def _formula(node: ET.Element, value: object, kind: str, shared: dict) -> dict | None:
    """수식과 저장 캐시 상태를 분리하여 보존한다.

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
    """
    formula = node.find("{*}f")
    if formula is None:
        return None
    raw = formula.text
    base = shared.get(formula.get("si")) if formula.get("t") == "shared" else None
    expression = raw or (base or {}).get("raw_text", "")
    stored = node.find("{*}v")
    status = (
        "error"
        if kind == "error"
        else "missing"
        if value is None or stored is None or stored.text is None
        else "stored_unverified"
    )
    return {
        "text": "=" + raw if raw else None,
        "raw_text": raw,
        "attributes": dict(formula.attrib),
        "shared_base": base,
        "cached_status": status,
        "external_reference": bool(re.search(r"\[[^\]]+\].*!", expression)),
        "recalculated": False,
    }


def _blank_span(spans: list[dict], node: ET.Element, row: int, column: int) -> None:
    """동일 스타일의 연속 빈 XML 셀을 구조 구간으로 압축한다.

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
    """
    attributes = {key: value for key, value in node.attrib.items() if key != "r"}
    if (
        spans
        and spans[-1]["row"] == row
        and spans[-1]["end_column"] + 1 == column
        and spans[-1]["attributes"] == attributes
    ):
        spans[-1]["end_column"] = column
    else:
        spans.append(
            {
                "row": row,
                "start_column": column,
                "end_column": column,
                "attributes": attributes,
            }
        )


def _extract_sheet(
    archive: ZipFile, info: dict, file: dict, shared: list[str], styles: dict
) -> dict:
    """한 시트의 내용 셀·숨김·병합·서식·메모와 빈 구조를 읽는다.

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
    """
    root = ET.fromstring(archive.read(info["part"]))
    comments = read_comments(archive, info["part"])
    columns = [dict(node.attrib) for node in root.findall("./{*}cols/{*}col")]
    merges = [
        {"range": node.attrib["ref"], "anchor": node.attrib["ref"].split(":")[0]}
        for node in root.findall("./{*}mergeCells/{*}mergeCell")
    ]
    merge_bounds = [(range_bounds(merge["range"]), merge) for merge in merges]
    shared_formulas = {}
    for node in root.findall("./{*}sheetData/{*}row/{*}c"):
        formula = node.find("{*}f")
        if formula is not None and formula.get("t") == "shared" and formula.text:
            shared_formulas[formula.get("si")] = {
                "address": node.attrib["r"],
                "raw_text": formula.text,
                "attributes": dict(formula.attrib),
            }
    dimension = root.find("{*}dimension")
    sheet_format = root.find("{*}sheetFormatPr")
    default_hidden = sheet_format is not None and sheet_format.get("zeroHeight") in {
        "1",
        "true",
    }
    sheet = {
        **info,
        "dimension": dimension.get("ref") if dimension is not None else None,
        "rows": [],
        "columns": columns,
        "merges": merges,
        "blank_spans": [],
        "cells": [],
        "conditional_formatting": [],
        "relationships": relationships(archive, info["part"]),
        "structure_xml": [],
        "counts": {"content_cells": 0, "xml_cells": 0, "blank_cells": 0, "formulas": 0},
    }
    for child in root:
        if local_name(child.tag) != "sheetData":
            raw = ET.tostring(child, encoding="unicode")
            sheet["structure_xml"].append(raw)
            if "conditionalFormatting" in raw:
                sheet["conditional_formatting"].append(
                    {"range": child.get("sqref"), "raw_xml": raw, "evaluated": False}
                )
    seen = set()
    for row_node in root.findall("./{*}sheetData/{*}row"):
        sheet["rows"].append(dict(row_node.attrib))
        row_hidden = row_node.get("hidden", "true" if default_hidden else "false") in {
            "1",
            "true",
        }
        for node in row_node.findall("{*}c"):
            address = node.attrib["r"]
            seen.add(address)
            row, column_index, column = cell_position(address)
            sheet["counts"]["xml_cells"] += 1
            has_content = any(
                node.find("{*}" + tag) is not None for tag in ("v", "f", "is")
            )
            if not has_content and address not in comments and len(node) == 0:
                _blank_span(sheet["blank_spans"], node, row, column_index)
                sheet["counts"]["blank_cells"] += 1
                continue
            column_info = next(
                (
                    entry
                    for entry in columns
                    if int(entry["min"]) <= column_index <= int(entry["max"])
                ),
                {},
            )
            style_id = int(
                node.get("s", row_node.get("s", column_info.get("style", "0")))
            )
            style = (
                styles["cell_formats"][style_id]
                if style_id < len(styles["cell_formats"])
                else {}
            )
            number_format = style.get("number_format")
            value, kind, raw = _cell_value(node, shared, number_format)
            formula = _formula(node, value, kind, shared_formulas)
            merge = next(
                (
                    entry
                    for (r1, c1, r2, c2), entry in merge_bounds
                    if r1 <= row <= r2 and c1 <= column_index <= c2
                ),
                None,
            )
            cell = {
                "record_id": f"{file['source_id']}/{info['sheet_id']}/{address}",
                "file_id": file["file_id"],
                "sheet_id": info["sheet_id"],
                "sheet": info["name"],
                "address": address,
                "row": row,
                "column": column,
                "column_index": column_index,
                "value": value,
                "data_type": kind,
                "raw_lexeme": raw,
                "storage_type": node.get("t", "n"),
                "formula": formula,
                "cached_value": value if formula is not None else None,
                "style_id": style_id,
                "number_format": number_format,
                "row_hidden": row_hidden,
                "column_hidden": column_info.get("hidden") in {"1", "true"},
                "sheet_hidden": info["state"] != "visible",
                "merge": merge,
                "comments": comments.get(address, []),
                "has_content": has_content,
                "xml_attributes": dict(node.attrib),
            }
            inline = node.find("{*}is")
            if inline is not None:
                cell["inline_string_xml"] = ET.tostring(inline, encoding="unicode")
            extra = [
                ET.tostring(child, encoding="unicode")
                for child in node
                if local_name(child.tag) not in {"v", "f", "is"}
            ]
            if extra:
                cell["extra_xml"] = extra
            sheet["cells"].append(cell)
            sheet["counts"]["content_cells"] += int(has_content)
            sheet["counts"]["formulas"] += int(formula is not None)
    for address, entries in comments.items():
        if address in seen:
            continue
        row, column_index, column = cell_position(address)
        row_info = next(
            (entry for entry in sheet["rows"] if int(entry["r"]) == row), {}
        )
        column_info = next(
            (
                entry
                for entry in columns
                if int(entry["min"]) <= column_index <= int(entry["max"])
            ),
            {},
        )
        merge = next(
            (
                entry
                for (r1, c1, r2, c2), entry in merge_bounds
                if r1 <= row <= r2 and c1 <= column_index <= c2
            ),
            None,
        )
        sheet["cells"].append(
            {
                "record_id": f"{file['source_id']}/{info['sheet_id']}/{address}",
                "file_id": file["file_id"],
                "sheet_id": info["sheet_id"],
                "sheet": info["name"],
                "address": address,
                "row": row,
                "column": column,
                "column_index": column_index,
                "value": None,
                "data_type": "blank",
                "raw_lexeme": None,
                "storage_type": None,
                "formula": None,
                "cached_value": None,
                "style_id": None,
                "number_format": None,
                "row_hidden": row_info.get(
                    "hidden", "true" if default_hidden else "false"
                )
                in {"1", "true"},
                "column_hidden": column_info.get("hidden") in {"1", "true"},
                "sheet_hidden": info["state"] != "visible",
                "merge": merge,
                "comments": entries,
                "has_content": False,
                "xml_attributes": {},
                "annotation_only": True,
            }
        )
    return sheet


def extract_workbook(path: Path, root: Path | None = None) -> dict:
    """한 xlsx/xlsm 파일의 모든 시트를 읽기 전용으로 추출한다.

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
    """
    path = Path(path).resolve()
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError(
            f"원본 추출 미지원 형식: {path.suffix}. xlsx/xlsm을 사용하세요."
        )
    fingerprint = file_fingerprint(path)
    display_path = (
        str(path.relative_to(root.resolve()))
        if root is not None and path.is_relative_to(root.resolve())
        else str(path)
    )
    file = {
        "file_id": fingerprint,
        "source_id": hashlib.sha256(
            (display_path + "\0" + fingerprint).encode()
        ).hexdigest()[:24],
        "path": display_path,
        "name": path.name,
        "size_bytes": path.stat().st_size,
        "extension": path.suffix.lower(),
    }
    result = {
        "file": file,
        "sheets": [],
        "styles": {},
        "package_metadata": {},
        "diagnostics": [],
    }
    with ZipFile(path) as archive:
        root_relations = relationships(archive, "")
        book_part = next(
            (
                entry["part"]
                for entry in root_relations.values()
                if entry.get("Type", "").endswith("/officeDocument")
            ),
            "xl/workbook.xml",
        )
        workbook = ET.fromstring(archive.read(book_part))
        rels = relationships(archive, book_part)
        styles_part = next(
            (
                entry.get("part")
                for entry in rels.values()
                if entry.get("Type", "").endswith("/styles")
            ),
            None,
        )
        strings_part = next(
            (
                entry.get("part")
                for entry in rels.values()
                if entry.get("Type", "").endswith("/sharedStrings")
            ),
            None,
        )
        shared_xml = archive.read(strings_part) if strings_part else None
        shared = (
            [xml_text(node) for node in ET.fromstring(shared_xml)] if shared_xml else []
        )
        result["styles"] = read_styles(archive, styles_part)
        result["package_metadata"] = {
            "workbook_part": book_part,
            "workbook_xml": archive.read(book_part).decode("utf-8-sig"),
            "relationships": rels,
            "root_relationships": root_relations,
            "shared_strings_part": strings_part,
            "shared_strings_xml": shared_xml.decode("utf-8-sig")
            if shared_xml
            else None,
            "parts": [
                {"name": item.filename, "size_bytes": item.file_size, "crc": item.CRC}
                for item in archive.infolist()
            ],
            "supplementary_xml": {},
        }
        for name in archive.namelist():
            if name.endswith(".xml") and any(
                part in name
                for part in (
                    "/externalLinks/",
                    "/drawings/",
                    "/theme/",
                    "/persons/",
                    "docProps/",
                )
            ):
                result["package_metadata"]["supplementary_xml"][name] = archive.read(
                    name
                ).decode("utf-8-sig")
        for sheet in workbook.findall("./{*}sheets/{*}sheet"):
            relationship_id = next(
                value for key, value in sheet.attrib.items() if local_name(key) == "id"
            )
            relation = rels[relationship_id]
            info = {
                "sheet_id": sheet.attrib["sheetId"],
                "name": sheet.attrib["name"],
                "state": sheet.get("state", "visible"),
                "part": relation["part"],
                "relationship_type": relation["Type"],
            }
            if not relation["Type"].endswith("/worksheet"):
                result["sheets"].append(
                    {
                        **info,
                        "cells": [],
                        "counts": {
                            "content_cells": 0,
                            "xml_cells": 0,
                            "blank_cells": 0,
                            "formulas": 0,
                        },
                        "raw_xml": archive.read(relation["part"]).decode("utf-8-sig"),
                    }
                )
                result["diagnostics"].append(
                    {
                        "code": "NON_WORKSHEET_PART",
                        "sheet": info["name"],
                        "message": "워크시트가 아닌 시트는 원본 XML로 보존했습니다.",
                    }
                )
                continue
            result["sheets"].append(
                _extract_sheet(archive, info, file, shared, result["styles"])
            )
    return result
