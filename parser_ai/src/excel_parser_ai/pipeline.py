"""Read-only parser pipeline. No aggregation, scoring, approval or DB writes."""
import gzip
import json
import time
from collections import Counter
from contextlib import ExitStack
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path

from .context import workbook_context
from .engine import compile_records, infer_workbook, unresolved_records
from .structured import infer_structure
from .shared import REPO, extract_workbook, file_fingerprint, validate_extraction, validate_provenance


def encode(value):
    if isinstance(value, Decimal) and value.is_finite():
        return format(value, "f")
    raise TypeError("Unsupported JSON type: " + type(value).__name__)


def dump(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=encode, allow_nan=False) + "\n", encoding="utf-8")


def collect_inputs(inputs, recursive=False):
    result = set()
    for supplied in inputs:
        p = Path(supplied).expanduser().resolve()
        if p.is_file():
            candidates = [p]
        elif p.is_dir():
            candidates = p.rglob("*") if recursive else p.iterdir()
        else:
            raise FileNotFoundError(str(p))
        result.update(c.resolve() for c in candidates if c.is_file() and c.suffix.lower() in {".xlsx", ".xlsm", ".xls", ".xlsb"} and not c.name.startswith(("~$", "._")))
    if not result:
        raise ValueError("입력 Excel 파일이 없습니다")
    return sorted(result)


def run(inputs, output, *, client=None, prepare=False, metadata=None, recursive=False,
        review=True, batch_size=12, max_batch_output_cells=None, max_batch_context_bytes=None,
        concurrency=2, max_files=None, max_new_row_batches=None, progress=None,
        structure_mode=False, max_exception_calls=0):
    files = collect_inputs(inputs, recursive)
    if max_files:
        files = files[:max_files]
    output = Path(output).resolve()
    allowed = (REPO / "parser_ai" / "output").resolve()
    if output == allowed or not output.is_relative_to(allowed):
        raise ValueError("출력은 parser_ai/output 아래 새 실행 폴더만 허용합니다")
    for p in files:
        if output == p or p.is_relative_to(output):
            raise ValueError("입력과 출력 경로가 겹칩니다")
    if output.exists():
        raise FileExistsError("기존 결과를 덮어쓰지 않습니다: " + str(output))
    output.mkdir(parents=True)
    metadata = metadata or {}
    if not isinstance(metadata.get("defaults", {}), dict) or not isinstance(metadata.get("files", {}), dict):
        raise ValueError("metadata.defaults와 metadata.files는 객체여야 합니다")
    started = time.monotonic()
    summary = {"contract_version": "1.0.0", "engine": "llm_structure_parser" if structure_mode else "llm_semantic_parser", "mode": "prepare" if prepare else "run",
               "started_at_utc": datetime.now(timezone.utc).isoformat(), "input_file_count": len(files),
               "processed_file_count": 0, "sheet_count": 0, "content_cell_count": 0, "raw_record_count": 0,
               "formula_count": 0, "standard_record_count": 0, "mapping_status_counts": {}, "normalization_status_counts": {},
               "files": [], "failures": [], "originals_unchanged": True, "raw_extraction_complete": False,
               "semantic_interpretation_complete": False, "llm_inference_completed": False,
               "formula_recalculation_performed": False, "business_master_issued": False,
               "performance_comparison_performed": False, "output_dir": str(output)}
    checks = {"raw_extraction": [], "provenance": [], "source_integrity": [],
              "master_comparison": {"status": "not_performed", "reason": "이번 범위는 AI 파서만. 별도 정답표 대조·성능 비교 없음"},
              "formula_recalculation_performed": False}
    observations, gaps = {}, []
    mapping_counts, value_counts = Counter(), Counter()
    inferred_count = 0
    with ExitStack() as stack:
        streams = {n: stack.enter_context(gzip.open(output / (n + ".jsonl.gz"), "wt", encoding="utf-8"))
                   for n in ("raw-cells", "standard-records", "workbooks", "diagnostics", "ai-semantics", "ai-contexts")}

        def write(name, value):
            streams[name].write(json.dumps(value, ensure_ascii=False, default=encode, separators=(",", ":"), allow_nan=False) + "\n")

        for number, path in enumerate(files, 1):
            if progress:
                progress(f"[{number}/{len(files)}] {path.name}: 원본 추출" + ("·준비" if prepare else " → LLM 의미 해석"))
            before = file_fingerprint(path)
            try:
                workbook = extract_workbook(path)
                raw_check = validate_extraction(path, workbook)
            except Exception as error:
                failure = {"path": str(path), "stage": "extraction", "error_type": type(error).__name__, "message": str(error)}
                summary["failures"].append(failure)
                write("diagnostics", failure)
                checks["source_integrity"].append({"path": str(path), "sha256_before": before, "sha256_after": file_fingerprint(path)})
                continue
            checks["raw_extraction"].append({"file": workbook["file"], **raw_check})
            # Persist raw source before any model call, so failures cannot erase it.
            for sheet in workbook["sheets"]:
                for cell in sheet["cells"]:
                    write("raw-cells", cell)
                summary["sheet_count"] += 1
                summary["content_cell_count"] += sheet["counts"]["content_cells"]
                summary["raw_record_count"] += len(sheet["cells"])
                summary["formula_count"] += sheet["counts"]["formulas"]
            structure = {k: v for k, v in workbook.items() if k != "sheets"}
            structure["sheets"] = [{k: v for k, v in s.items() if k != "cells"} for s in workbook["sheets"]]
            write("workbooks", {"contract_version": "1.0.0", **structure})
            inferred = False
            row_batches = {}
            semantic = None
            try:
                write("ai-contexts", {"file": workbook["file"], "context": workbook_context(workbook)})
                if prepare:
                    mapped = unresolved_records(workbook)
                else:
                    if client is None:
                        raise ValueError("run requires an OpenAI client")
                    specific = metadata.get("files", {}).get(str(path), metadata.get("files", {}).get(path.name, {}))
                    if not isinstance(specific, dict):
                        raise ValueError("파일별 metadata는 객체여야 합니다")
                    effective_metadata = {**metadata.get("defaults", {}), **specific}
                    if structure_mode:
                        mapped, semantic = infer_structure(workbook, client, review=review,
                                                           max_exception_calls=max_exception_calls,
                                                           metadata=effective_metadata)
                    else:
                        semantic = infer_workbook(workbook, client, review=review, batch_size=batch_size,
                                                  max_output_cells_per_batch=max_batch_output_cells,
                                                  max_context_bytes_per_batch=max_batch_context_bytes,
                                                  concurrency=concurrency, max_new_row_batches=max_new_row_batches)
                    # Even when some row batches failed, infer_workbook still
                    # returns normally with whatever succeeded (see
                    # _run_pool_partial) — compile_records already treats any
                    # row/cell the model never reached as AI_NOT_CLASSIFIED,
                    # so the batches that DID succeed are not thrown away.
                    if not structure_mode:
                        mapped = compile_records(workbook, semantic, effective_metadata)
                    write("ai-semantics", {"file": workbook["file"], **semantic})
                    row_batches = semantic.get("row_batches", {})
                    fully_completed = (len(semantic.get("structure_plans", [])) == len(workbook["sheets"])) if structure_mode else (
                        row_batches.get("failed", 0) == 0
                        and not row_batches.get("stopped_by_new_batch_limit", False)
                        and row_batches.get("completed", 0) == row_batches.get("planned", 0))
                    inferred = fully_completed
                    if inferred:
                        inferred_count += 1
                    elif structure_mode:
                        failure = {"path": str(path), "stage": "ai_structure", "error_type": "PartialStructureFailure",
                                   "message": "구조 해석 일부 실패; 성공한 시트 계획과 캐시는 보존함",
                                   "structure_failures": semantic.get("failures", [])}
                        summary["failures"].append(failure)
                        write("diagnostics", failure)
                    elif not structure_mode and row_batches.get("failed", 0) > 0:
                        failure = {"path": str(path), "stage": "ai_semantics", "error_type": "PartialRowBatchFailure",
                                  "message": (f"{row_batches.get('completed', 0)}/{row_batches.get('planned', 0)} row batch(es) "
                                             f"completed; stopped after {row_batches.get('failed', 0)} failure(s). "
                                             "이미 성공한 배치의 결과는 보존되어 저장됨(해당 레코드는 mapped/partial 상태)."),
                                  "row_batches": row_batches}
                        summary["failures"].append(failure)
                        write("diagnostics", failure)
                    elif not structure_mode:
                        # max_new_row_batches intentionally stopped this file
                        # early for a small, cheap test — not a failure.
                        write("diagnostics", {"file": workbook["file"], "code": "AI_UNCERTAINTY", "severity": "info",
                                              "message": (f"테스트 목적(max_new_row_batches)으로 신규 행 배치 "
                                                         f"{row_batches.get('completed', 0)}개 처리 후 의도적으로 중단함. "
                                                         "파일 전체는 미완료이며 llm_inference_completed=false."),
                                              "row_batches": row_batches})
            except Exception as error:
                failure = {"path": str(path), "stage": "ai_semantics", "error_type": type(error).__name__, "message": str(error)}
                summary["failures"].append(failure)
                write("diagnostics", failure)
                mapped = unresolved_records(workbook)
            provenance = validate_provenance(workbook, mapped)
            checks["provenance"].append({"file": workbook["file"], **provenance})
            for record in mapped["records"]:
                write("standard-records", {"contract_version": "1.0.0", **record})
                for rule in record["value"]["rule_ids"]:
                    observation = observations.setdefault(rule, {"count": 0, "examples": []})
                    observation["count"] += 1
                    if len(observation["examples"]) < 3:
                        observation["examples"].append({"source": record["source"], "original_value": record["raw_value"], "normalized_value": record["value"]["normalized_value"]})
            for diagnostic in [*workbook["diagnostics"], *mapped["diagnostics"]]:
                write("diagnostics", {"file": workbook["file"], **diagnostic})
            gaps.append({"file": workbook["file"], "relationships": mapped["gap_relationships"]})
            mapping_counts.update(r["mapping"]["status"] for r in mapped["records"])
            value_counts.update(r["value"]["status"] for r in mapped["records"])
            summary["standard_record_count"] += len(mapped["records"])
            summary["processed_file_count"] += 1
            summary["files"].append({"file": workbook["file"], "llm_inference_completed": inferred,
                                     "mapping_summary": mapped["summary"], "raw_validation_passed": raw_check["passed"],
                                     "provenance_validation_passed": provenance["passed"], "row_batches": row_batches,
                                     "structure_sheets": len(semantic.get("structure_plans", [])) if semantic and structure_mode else None,
                                     "structure_interpretation_skipped_sheets": semantic.get("skipped_sheets", []) if semantic and structure_mode else [],
                                     "structure_timeout_fallbacks": semantic.get("timeout_fallbacks", []) if semantic and structure_mode else []})
            checks["source_integrity"].append({"path": str(path), "sha256_before": before, "sha256_after": file_fingerprint(path)})
    summary["originals_unchanged"] = all(c["sha256_before"] == c["sha256_after"] for c in checks["source_integrity"])
    summary["raw_extraction_complete"] = summary["processed_file_count"] == len(files) and all(c["passed"] for c in checks["raw_extraction"]) and summary["originals_unchanged"]
    summary["llm_inference_completed"] = not prepare and inferred_count == len(files) and not summary["failures"]
    summary["validation_passed"] = summary["raw_extraction_complete"] and not summary["failures"] and all(c["passed"] for c in checks["provenance"])
    summary["mapping_status_counts"] = dict(mapping_counts)
    summary["normalization_status_counts"] = dict(value_counts)
    summary["elapsed_seconds"] = round(time.monotonic() - started, 3)
    summary["completed_at_utc"] = datetime.now(timezone.utc).isoformat()
    summary["limitations"] = ["mapped는 독립 정답셋에 대한 정확도 점수가 아닙니다", "수식 캐시의 최신성을 보증하지 않습니다", "prepare는 LLM을 실행하지 않습니다",
                              "새 데이터의 룰베이스 비교와 정답 대조는 수행하지 않았습니다",
                              ("LLM이 선언한 구조를 Python이 원본 셀에 적용하고 Decimal로 정규화합니다. 구조 밖/충돌/예외는 미해결입니다"
                               if structure_mode else "value.normalized_value는 모델이 직접 반환한 값이며 Python 검산은 진단 전용입니다")]
    events = client.events if client else []
    # A failed call has no usage data (we never received a parseable
    # response to read it from) — failed_calls/events still records that
    # the attempt happened, with whatever elapsed time and error type is
    # knowable, rather than only ever showing successful usage.
    inference = {"model": client.model if client else None, "reasoning_effort": client.effort if client else None,
                 "semantic_review_enabled": review, "api_calls": sum(not e["cached"] for e in events),
                 "logical_attempted_calls": getattr(client, "call_count", None) if client else 0,
                 "http_post_attempts": getattr(client, "http_post_attempts", None) if client else 0,
                 "cache_hits": sum(e["cached"] for e in events), "failed_calls": sum(e.get("failed", False) for e in events),
                 "events": events,
                 "input_tokens": sum(e["usage"].get("input_tokens", 0) for e in events),
                 "output_tokens": sum(e["usage"].get("output_tokens", 0) for e in events),
                 "financial_amounts_generated_by_model": not structure_mode,
                 "python_normalization_fallback_used": False,
                 "python_decimal_normalization_used": structure_mode,
                 "python_independent_crosscheck_is_diagnostic_only": not structure_mode,
                 "baseline_semantic_mapping_used": False,
                 "external_reference_workbook_used": False}
    for name, data in [("summary", summary), ("validation", checks), ("normalization-observations", observations),
                       ("gap-relationships", gaps), ("inference", inference)]:
        dump(output / (name + ".json"), data)
    return summary
