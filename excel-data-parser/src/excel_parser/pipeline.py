"""전체 입력의 추출·매핑·독립 검증과 명시적으로 요청한 결과 저장을 연결한다."""

import gzip
import json
from collections import Counter
from contextlib import ExitStack
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import TextIO

from .extraction import extract_workbook, file_fingerprint
from .mapping import map_workbook
from .paths import ProjectPaths
from .validation import compare_master, validate_extraction, validate_provenance

CONTRACT_VERSION = "1.0.0"
EXCEL_EXTENSIONS = frozenset({".xlsx", ".xlsm", ".xlsb", ".xls"})


def json_value(value: object) -> str:
    """Decimal을 정밀도 손실 없는 JSON 문자열 값으로 직렬화한다.

    Args:
        value: json 모듈이 기본 처리할 수 없는 객체.
    Returns:
        Decimal의 고정 소수점 문자열.
    Raises:
        TypeError: 등록하지 않은 타입인 경우.
        ValueError: 유한하지 않은 Decimal인 경우.
    Side Effects:
        없음.
    """
    if isinstance(value, Decimal):
        if not value.is_finite():
            raise ValueError("유한하지 않은 숫자는 JSON 결과로 저장할 수 없습니다.")
        return format(value, "f")
    raise TypeError(f"직렬화 미지원 타입: {type(value).__name__}")


def collect_inputs(inputs: list[Path], *, recursive: bool = False) -> list[Path]:
    """명시한 파일·폴더에서 Excel 입력을 빠짐없이 수집한다.

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
    """
    result = set()
    for supplied in inputs:
        path = Path(supplied).expanduser().resolve()
        if path.is_file():
            result.add(path)
        elif path.is_dir():
            candidates = path.rglob("*") if recursive else path.iterdir()
            result.update(
                candidate.resolve()
                for candidate in candidates
                if candidate.is_file()
                and candidate.suffix.lower() in EXCEL_EXTENSIONS
                and not candidate.name.startswith("~$")
            )
        else:
            raise FileNotFoundError(f"입력 경로를 찾을 수 없습니다: {supplied}")
    if not result:
        raise ValueError("지정한 입력에서 Excel 파일을 찾지 못했습니다.")
    return sorted(result, key=lambda path: (str(path).casefold(), str(path)))


def _write_json(stream: TextIO, value: object) -> None:
    """열린 결과 스트림에 JSON 레코드 한 줄을 기록한다.

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
    """
    stream.write(
        json.dumps(
            value,
            ensure_ascii=False,
            default=json_value,
            allow_nan=False,
            separators=(",", ":"),
        )
        + "\n"
    )


def _file_metadata(metadata: dict, file: dict) -> dict:
    """실행 기본값과 파일별 명시 메타데이터를 결합한다.

    Args:
        metadata: defaults와 files를 포함하는 실행 메타데이터.
        file: 파일명·출처 경로·식별자.
    Returns:
        해당 파일에 적용할 평탄한 메타데이터 사전.
    Raises:
        ValueError: 메타데이터 구조가 사전이 아닌 경우.
    Side Effects:
        없음. 원본 파일과 입력 사전을 변경하지 않는다.
    """
    defaults = metadata.get("defaults", {})
    files = metadata.get("files", {})
    if not isinstance(defaults, dict) or not isinstance(files, dict):
        raise ValueError("메타데이터의 defaults와 files는 객체여야 합니다.")
    specific = files.get(file["path"], files.get(file["name"], {}))
    if not isinstance(specific, dict):
        raise ValueError("파일별 메타데이터는 객체여야 합니다.")
    return {**defaults, **specific}


def _record_observations(records: list[dict], observations: dict) -> None:
    """실제로 적용한 정규화 규칙의 횟수와 출처 예시를 집계한다.

    Args:
        records: 한 파일의 표준 레코드.
        observations: 누적할 규칙별 사전.
    Returns:
        없음. observations를 메모리에서 갱신한다.
    Raises:
        없음. 매핑 계약을 충족한 입력을 전제로 한다.
    Side Effects:
        전달받은 누적 사전만 변경한다.
    """
    for record in records:
        value = record.get("value", {})
        for rule in value.get("rule_ids", []):
            entry = observations.setdefault(rule, {"count": 0, "examples": []})
            entry["count"] += 1
            if len(entry["examples"]) < 3:
                entry["examples"].append(
                    {
                        "source": record["source"],
                        "original_value": value.get("original_value"),
                        "normalized_value": value.get("normalized_value"),
                        "unit": value.get("source_unit"),
                        "number_format": record.get("style_reference", {}).get(
                            "number_format"
                        ),
                    }
                )


def run_parser(
    inputs: list[Path],
    *,
    root: Path,
    recursive: bool = False,
    output_dir: Path | None = None,
    metadata: dict | None = None,
) -> dict:
    """입력 전체를 순차 처리하고 원본 보존·독립 검증 결과를 반환한다.

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
    """
    paths = ProjectPaths.load(root)
    files = collect_inputs(inputs, recursive=recursive)
    metadata = {} if metadata is None else metadata
    if not isinstance(metadata, dict):
        raise ValueError("실행 메타데이터는 JSON 객체여야 합니다.")
    streams = {}
    summary = {
        "contract_version": CONTRACT_VERSION,
        "started_at_utc": datetime.now(timezone.utc).isoformat(),
        "input_file_count": len(files),
        "processed_file_count": 0,
        "sheet_count": 0,
        "content_cell_count": 0,
        "raw_record_count": 0,
        "formula_count": 0,
        "standard_record_count": 0,
        "mapping_status_counts": {},
        "normalization_status_counts": {},
        "files": [],
        "failures": [],
        "originals_unchanged": True,
        "raw_extraction_complete": False,
        "semantic_interpretation_complete": False,
        "formula_recalculation_performed": False,
        "business_master_issued": False,
        "output_dir": None,
    }
    validations = []
    provenance_checks = []
    observations = {}
    gap_relationships = []
    mapping_counts = Counter()
    normalization_counts = Counter()
    before = {str(path): file_fingerprint(path) for path in files}
    with ExitStack() as stack:
        if output_dir is not None:
            output_dir = Path(output_dir).resolve()
            protected = (paths.root / "Data").resolve()
            if (
                output_dir == paths.output_dir
                or not output_dir.is_relative_to(paths.output_dir)
                or output_dir.is_relative_to(protected)
                or protected.is_relative_to(output_dir)
            ):
                raise ValueError(
                    "출력은 설정된 output 아래의 새로운 실행 하위 폴더여야 합니다."
                )
            if output_dir.exists():
                raise FileExistsError(f"기존 결과를 덮어쓰지 않습니다: {output_dir}")
            output_dir.mkdir(parents=True, exist_ok=False)
            summary["output_dir"] = str(output_dir.relative_to(paths.root))
            for name in ("raw-cells", "standard-records", "workbooks", "diagnostics"):
                streams[name] = stack.enter_context(
                    gzip.open(output_dir / f"{name}.jsonl.gz", "wt", encoding="utf-8")
                )
        for path in files:
            try:
                workbook = extract_workbook(path, root=paths.root)
            except Exception as error:
                failure = {
                    "path": str(path),
                    "stage": "extraction",
                    "code": "EXTRACTION_FAILED",
                    "error_type": type(error).__name__,
                    "message": str(error),
                }
                summary["failures"].append(failure)
                if streams:
                    _write_json(streams["diagnostics"], failure)
                continue
            try:
                validation = validate_extraction(path, workbook)
            except Exception as error:
                validation = {
                    "passed": False,
                    "issues": [
                        {"code": "RAW_VALIDATION_FAILED", "message": str(error)}
                    ],
                }
            try:
                mapped = map_workbook(
                    workbook, metadata=_file_metadata(metadata, workbook["file"])
                )
            except Exception as error:
                failure = {
                    "path": str(path),
                    "stage": "mapping",
                    "code": "MAPPING_FAILED",
                    "error_type": type(error).__name__,
                    "message": str(error),
                }
                summary["failures"].append(failure)
                mapped = {
                    "records": [],
                    "diagnostics": [failure],
                    "summary": {"status": "failed", "records": 0},
                    "gap_relationships": [],
                }
            try:
                provenance = validate_provenance(workbook, mapped)
            except Exception as error:
                provenance = {
                    "passed": False,
                    "issues": [
                        {"code": "PROVENANCE_VALIDATION_FAILED", "message": str(error)}
                    ],
                }
            provenance_checks.append({"file": workbook["file"], **provenance})
            validations.append({"file": workbook["file"], **validation})
            sheet_summaries = []
            for sheet in workbook["sheets"]:
                counts = sheet["counts"]
                per_sheet_records = [
                    record
                    for record in mapped["records"]
                    if record["source"]["sheet"] == sheet["name"]
                ]
                per_sheet = {
                    "name": sheet["name"],
                    "state": sheet["state"],
                    **counts,
                    "raw_records": len(sheet["cells"]),
                    "standard_records": len(per_sheet_records),
                    "mapping_status_counts": dict(
                        Counter(
                            record["mapping"]["status"] for record in per_sheet_records
                        )
                    ),
                    "normalization_status_counts": dict(
                        Counter(
                            record["value"]["status"] for record in per_sheet_records
                        )
                    ),
                    "merges": len(sheet.get("merges", [])),
                    "comments": sum(
                        len(cell.get("comments", [])) for cell in sheet["cells"]
                    ),
                    "conditional_formatting": len(
                        sheet.get("conditional_formatting", [])
                    ),
                }
                sheet_summaries.append(per_sheet)
                summary["sheet_count"] += 1
                summary["content_cell_count"] += counts["content_cells"]
                summary["raw_record_count"] += len(sheet["cells"])
                summary["formula_count"] += counts["formulas"]
                if streams:
                    for cell in sheet["cells"]:
                        _write_json(streams["raw-cells"], cell)
            mapping_counts.update(
                record["mapping"]["status"] for record in mapped["records"]
            )
            normalization_counts.update(
                record["value"]["status"] for record in mapped["records"]
            )
            _record_observations(mapped["records"], observations)
            gap_relationships.append(
                {
                    "file": workbook["file"],
                    "relationships": mapped.get("gap_relationships", []),
                }
            )
            summary["processed_file_count"] += 1
            summary["standard_record_count"] += len(mapped["records"])
            summary["files"].append(
                {
                    "file": workbook["file"],
                    "sheets": sheet_summaries,
                    "mapping_summary": mapped["summary"],
                    "raw_validation_passed": validation["passed"],
                    "provenance_validation_passed": provenance["passed"],
                }
            )
            if streams:
                for record in mapped["records"]:
                    _write_json(
                        streams["standard-records"],
                        {"contract_version": CONTRACT_VERSION, **record},
                    )
                for diagnostic in [*workbook["diagnostics"], *mapped["diagnostics"]]:
                    _write_json(
                        streams["diagnostics"], {"file": workbook["file"], **diagnostic}
                    )
                structure = {
                    key: value for key, value in workbook.items() if key != "sheets"
                }
                structure["sheets"] = [
                    {key: value for key, value in sheet.items() if key != "cells"}
                    for sheet in workbook["sheets"]
                ]
                _write_json(
                    streams["workbooks"],
                    {"contract_version": CONTRACT_VERSION, **structure},
                )
        try:
            master_validation = compare_master(
                [path for path in files if path.suffix.lower() in {".xlsx", ".xlsm"}]
            )
        except Exception as error:
            master_validation = {
                "status": "failed",
                "comparison_count": 0,
                "issues": [{"code": "MASTER_CHECK_FAILED", "message": str(error)}],
            }
        source_checks = [
            {
                "path": str(path),
                "sha256_before": before[str(path)],
                "sha256_after": file_fingerprint(path),
            }
            for path in files
        ]
        summary["originals_unchanged"] = all(
            check["sha256_before"] == check["sha256_after"] for check in source_checks
        )
        summary["mapping_status_counts"] = dict(mapping_counts)
        summary["normalization_status_counts"] = dict(normalization_counts)
        summary["raw_extraction_complete"] = (
            summary["processed_file_count"] == summary["input_file_count"]
            and all(item["passed"] for item in validations)
            and summary["originals_unchanged"]
        )
        summary["validation_passed"] = (
            summary["raw_extraction_complete"]
            and not summary["failures"]
            and all(item["passed"] for item in provenance_checks)
            and master_validation["status"] != "failed"
        )
        summary["completed_at_utc"] = datetime.now(timezone.utc).isoformat()
        validation_report = {
            "raw_extraction": validations,
            "provenance": provenance_checks,
            "master_comparison": master_validation,
            "source_integrity": source_checks,
            "formula_recalculation_performed": False,
        }
        if output_dir is not None:
            for name, value in (
                ("summary", summary),
                ("validation", validation_report),
                ("normalization-observations", observations),
                ("gap-relationships", gap_relationships),
            ):
                with (output_dir / f"{name}.json").open(
                    "x", encoding="utf-8"
                ) as stream:
                    json.dump(
                        value,
                        stream,
                        ensure_ascii=False,
                        indent=2,
                        default=json_value,
                        allow_nan=False,
                    )
                    stream.write("\n")
    return summary
