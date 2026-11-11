"""Ground LLM-returned values in immutable source cells; never rewrite them.

The model reads real numbers, formulas and cached values and returns its own
normalized decimal directly for every financial candidate cell. Python's job
here is strictly: (1) validate that every claimed source value actually
matches what was extracted (grounding), (2) validate that every required
row/cell got exactly one answer (coverage), (3) validate the response is
well-formed JSON honoring the schema, and (4) deterministically expand the
accepted response into standard records, assigning technical IDs and source
metadata. Python never substitutes its own normalization for the model's
value; a disagreement is recorded as a diagnostic only.
"""
import re
from collections import Counter
from concurrent.futures import CancelledError, ThreadPoolExecutor, as_completed
from decimal import Decimal, InvalidOperation

from .context import (
    LAYOUT_PROMPT, ROWS_PROMPT, data_rows, plan_row_batches, rows_context,
    sheet_layout_context, value_columns,
)
from .schema import LAYOUT_SCHEMA, ROWS_SCHEMA, SHEET_LAYOUT_SCHEMA
from .shared import cell_index, column_number, normalize_number, ref


def schema_errors(value, schema, path="$"):
    """Local validation also applies to cached responses and test transports."""
    if "anyOf" in schema:
        return [] if any(not schema_errors(value, option, path) for option in schema["anyOf"]) else [path + ": no valid alternative"]
    kinds = schema.get("type", [])
    if isinstance(kinds, str):
        kinds = [kinds]
    valid = {"object": isinstance(value, dict), "array": isinstance(value, list),
             "string": isinstance(value, str), "integer": type(value) is int,
             "null": value is None, "boolean": type(value) is bool}
    if kinds and not any(valid.get(k, False) for k in kinds):
        return [path + ": wrong type"]
    if "enum" in schema and value not in schema["enum"]:
        return [path + ": invalid enum"]
    errors = []
    if isinstance(value, dict):
        properties = schema.get("properties", {})
        errors += [path + ": missing " + k for k in schema.get("required", []) if k not in value]
        if schema.get("additionalProperties") is False:
            errors += [path + ": extra " + k for k in value.keys() - properties.keys()]
        for k in value.keys() & properties.keys():
            errors += schema_errors(value[k], properties[k], path + "." + k)
    elif isinstance(value, list):
        for i, entry in enumerate(value):
            errors += schema_errors(entry, schema["items"], f"{path}[{i}]")
    return errors


def references(value):
    if isinstance(value, dict):
        if set(value) == {"sheet", "cell"}:
            yield value
        else:
            for item in value.values():
                yield from references(item)
    elif isinstance(value, list):
        for item in value:
            yield from references(item)


def pointer_errors(value, index):
    return [f"Unknown source pointer {p['sheet']}!{p['cell']}" for p in references(value) if (p["sheet"], p["cell"]) not in index]


def _sheet_plan_content_errors(sheet):
    """Per-sheet-plan content checks, shared by the whole-workbook validator
    (layout_errors, used as a post-aggregation sanity check) and the
    per-sheet-call validator (sheet_plan_errors, used for each individual
    layout API call)."""
    errors = []
    if sheet["header_end_row"] < 0:
        errors.append("Negative header_end_row")
    columns = [p["column"] for p in sheet["periods"]]
    if len(set(columns)) != len(columns):
        errors.append(sheet["sheet"] + ": duplicate period column")
    for col in [*columns, sheet["organization_code_column"], sheet["organization_name_column"], *(d["column"] for d in sheet["dimensions"])]:
        if col is not None:
            try:
                if column_number(col) > 16384:
                    raise ValueError("Beyond Excel column limit")
            except ValueError:
                errors.append(f"Invalid column {col!r}")
    if sheet["unit"] is not None and not sheet["unit_evidence"]:
        errors.append(sheet["sheet"] + ": unit requires source evidence")
    for period in sheet["periods"]:
        if period["year"] is not None and not 1900 <= period["year"] <= 2200:
            errors.append("Invalid target year")
        if period["month"] is not None and not 1 <= period["month"] <= 12:
            errors.append("Invalid target month")
        if period["quarter"] is not None and not 1 <= period["quarter"] <= 4:
            errors.append("Invalid quarter")
        if period["aggregation"] == "month" and period["month"] is None:
            errors.append("Monthly column requires month")
        if period["aggregation"] == "annual" and period["month"] is not None:
            errors.append("Annual column cannot carry a month")
        if not period["evidence"]:
            errors.append("Period requires source evidence; use explicit unknowns")
    return errors


def layout_errors(plan, workbook):
    """Whole-workbook coverage + content check. Used as a cheap, local
    post-aggregation sanity pass after collecting all per-sheet layout
    calls (see infer_workbook) — it makes no API call itself."""
    errors = schema_errors(plan, LAYOUT_SCHEMA)
    if errors:
        return errors
    index = cell_index(workbook)
    errors += pointer_errors(plan, index)
    actual = [s["name"] for s in workbook["sheets"]]
    supplied = [s["sheet"] for s in plan["sheets"]]
    if Counter(actual) != Counter(supplied):
        errors.append("Return exactly one plan for every source sheet")
    for sheet in plan["sheets"]:
        errors += _sheet_plan_content_errors(sheet)
    return errors


def sheet_plan_errors(result, workbook, sheet_name):
    """Validator for one per-sheet layout API call (see sheet_layout_context
    and the SHEET_LAYOUT_SCHEMA wrapper)."""
    errors = schema_errors(result, SHEET_LAYOUT_SCHEMA)
    if errors:
        return errors
    plan = result["sheet_plan"]
    if plan["sheet"] != sheet_name:
        errors.append(f"sheet_plan.sheet must be {sheet_name!r}, not {plan['sheet']!r}")
        return errors
    index = cell_index(workbook)
    errors += pointer_errors(plan, index)
    errors += _sheet_plan_content_errors(plan)
    return errors


def _same_source_value(claimed, actual):
    """Loose numeric-equivalence grounding check.

    Accepts harmless formatting differences (e.g. "100" vs "100.0") but
    rejects a claimed value that is not actually derivable from the source.
    """
    if claimed is None or actual is None:
        return claimed == actual
    if claimed.strip() == actual.strip():
        return True
    try:
        return Decimal(claimed) == Decimal(actual)
    except InvalidOperation:
        return False


def _ground_truth_value(real):
    formula = real.get("formula")
    if formula:
        if formula.get("cached_status") == "stored_unverified":
            return str(real["cached_value"]) if real.get("cached_value") is not None else None
        return None  # missing/error cache: nothing reliable to ground against
    return str(real["value"]) if real.get("value") is not None else None


def cell_errors(cell, real, sheet, address):
    errors = []
    if cell["unit"] is not None and not cell["unit_evidence"]:
        errors.append(f"{sheet}!{address}: cell unit override requires evidence")
    formula = real.get("formula")
    cache_unusable = bool(formula) and formula.get("cached_status") in {"missing", "error"}
    if cache_unusable and (cell["status"] == "confirmed" or cell["normalized_value"] is not None):
        errors.append(f"{sheet}!{address}: formula cache missing/error cannot be confirmed or carry a normalized_value; cache is not a recalculation source")
    if cell["status"] == "confirmed":
        if cell["value_kind"] in {"amount", "ratio", "quantity", "number"}:
            if cell["normalized_value"] is None:
                errors.append(f"{sheet}!{address}: confirmed numeric cell requires normalized_value")
            else:
                try:
                    Decimal(cell["normalized_value"])
                except InvalidOperation:
                    errors.append(f"{sheet}!{address}: normalized_value must be a plain decimal string")
            if cell["value_kind"] in {"amount", "quantity"} and cell["unit"] is None:
                errors.append(f"{sheet}!{address}: confirmed amount/quantity requires a unit")
    elif cell["normalized_value"] is not None:
        errors.append(f"{sheet}!{address}: only a confirmed cell may carry normalized_value")
    if not cache_unusable and not _same_source_value(cell["source_value_read"], _ground_truth_value(real)):
        errors.append(f"{sheet}!{address}: source_value_read does not match the real source value")
    return errors


def row_errors(result, workbook, sheet, plan, required):
    errors = schema_errors(result, ROWS_SCHEMA)
    if errors:
        return errors
    index = cell_index(workbook)
    errors += pointer_errors(result, index)
    if Counter(r["row"] for r in result["rows"]) != Counter(required):
        errors.append("Return exactly one row descriptor for every required row, without omissions or duplicates")
    sheet_obj = next(s for s in workbook["sheets"] if s["name"] == sheet)
    by_row = {}
    for c in sheet_obj["cells"]:
        by_row.setdefault(c["row"], []).append(c)
    for row in result["rows"]:
        for name in ("item_name", "item_code"):
            p = row[name]
            if p and (p["sheet"] != sheet or int(re.search(r"\d+$", p["cell"]).group()) != row["row"]):
                errors.append("Row item name/code must point to the same source sheet and row")
        required_columns = value_columns(by_row.get(row["row"], []), plan)
        returned_columns = [c["column"] for c in row["cells"]]
        if Counter(returned_columns) != Counter(required_columns):
            errors.append(f"Row {row['row']}: return exactly one cell descriptor for every required value column, without omissions or duplicates")
            continue
        for cell in row["cells"]:
            address = f"{cell['column']}{row['row']}"
            real = index.get((sheet, address))
            if real is None:
                errors.append(f"Unknown value cell {sheet}!{address}")
                continue
            errors += cell_errors(cell, real, sheet, address)
    return errors


def infer(client, task, prompt, context, schema, validator, *, review=True, max_connection_attempts=None):
    # max_connection_attempts is only ever forwarded when a caller explicitly
    # passes it (e.g. structured.py's whole-sheet-then-blocks timeout
    # fallback) — omitted entirely otherwise, so existing test doubles with
    # a plain ask(task, prompt, context, schema) signature are unaffected.
    ask_kwargs = {"max_connection_attempts": max_connection_attempts} if max_connection_attempts is not None else {}
    proposal = client.ask(task, prompt, context, schema, **ask_kwargs)
    errors = validator(proposal)
    if errors:
        proposal = client.ask(task + "_repair", prompt + "\nRepair the validation errors using the original source. Do not hide missing rows.",
                              {"source": context, "proposal": proposal, "validation_errors": errors[:80]}, schema, **ask_kwargs)
        errors = validator(proposal)
    if errors:
        raise ValueError("LLM output failed grounding/coverage validation: " + "; ".join(errors[:12]))
    if review:
        reviewed = client.ask(task + "_review", prompt + "\nAct as an independent semantic reviewer. Read the source first. Check year/scenario, month vs annual, units, codes vs amounts, subtotal/detail, sign/percent handling and source evidence for every cell. Return the full corrected object. Ground every change. Do not assume the proposal is correct.",
                              {"source": context, "proposal_to_audit": proposal}, schema, **ask_kwargs)
        errors = validator(reviewed)
        if errors:
            # A broken critic response must never replace a grounded proposal or
            # be accepted silently. Stop this workbook for explicit reprocessing.
            raise ValueError("Semantic review failed validation: " + "; ".join(errors[:12]))
        proposal = reviewed
    return proposal


def _run_pool(fn, items, concurrency):
    """Run fn over items on a thread pool and return the results as a list.

    On any exception — including KeyboardInterrupt — shuts the pool down
    without waiting for already-queued-but-not-started work, instead of the
    `with ThreadPoolExecutor() as pool:` default, which always blocks for
    wait=True on exit regardless of why the block was left. This keeps a
    Ctrl+C (or a genuine failure) from hanging behind pending, never-useful
    batches; a thread already mid-flight on a blocking network call cannot
    be force-stopped by Python either way, but nothing further is started.
    """
    pool = ThreadPoolExecutor(max_workers=max(1, concurrency))
    try:
        results = list(pool.map(fn, items))
    except BaseException:
        pool.shutdown(wait=False, cancel_futures=True)
        raise
    pool.shutdown(wait=True)
    return results


def _run_pool_partial(fn, items, concurrency):
    """Run fn over items on a thread pool; a failure stops further pending
    work but never discards results already produced.

    As soon as one job raises, every not-yet-started job is cancelled (no
    new API calls are issued for this workbook after a failure) — but any
    job already in flight is left to finish, and whatever already
    succeeded is kept and returned rather than thrown away. A worker
    thread already blocked on a network call cannot be forcibly cancelled
    by Python either way; that limitation is inherent, not something this
    function can close.

    Returns (results, failures) where failures is a list of (item, error)
    for jobs that raised (CancelledError from a job we ourselves cancelled
    after a prior failure is not counted as a new failure).
    """
    results, failures = [], []
    pool = ThreadPoolExecutor(max_workers=max(1, concurrency))
    futures = {pool.submit(fn, item): item for item in items}
    stop_requested = False
    try:
        for future in as_completed(futures):
            try:
                results.append(future.result())
            except CancelledError:
                continue
            except BaseException as error:
                failures.append((futures[future], error))
                if not stop_requested:
                    stop_requested = True
                    for other in futures:
                        if not other.done():
                            other.cancel()
                if isinstance(error, KeyboardInterrupt):
                    raise
    finally:
        pool.shutdown(wait=False, cancel_futures=True)
    return results, failures


def infer_workbook(workbook, client, *, review=True, batch_size=12,
                   max_output_cells_per_batch=None, max_context_bytes_per_batch=None, concurrency=2,
                   max_new_row_batches=None):
    # One smaller layout call per sheet instead of one large call for the
    # whole workbook: a real dummy workbook's combined layout context ran
    # ~480KB, ~90% of it numeric values the layout stage never needs (it
    # only locates structure; real amounts are sent later, per cell).
    def layout_for_sheet(sheet):
        context = sheet_layout_context(workbook, sheet["name"])
        result = infer(client, "workbook_layout_sheet", LAYOUT_PROMPT, context, SHEET_LAYOUT_SCHEMA,
                       lambda p: sheet_plan_errors(p, workbook, sheet["name"]), review=review)
        return result["sheet_plan"], result["uncertainties"]

    layout_results = _run_pool(layout_for_sheet, workbook["sheets"], concurrency)
    layout = {"sheets": [plan for plan, _ in layout_results],
              "uncertainties": [u for _, uncertainties in layout_results for u in uncertainties]}
    errors = layout_errors(layout, workbook)  # local sanity re-check only; no extra API call
    if errors:
        raise ValueError("Aggregated layout plan failed validation: " + "; ".join(errors[:12]))

    plans = {p["sheet"]: p for p in layout["sheets"]}
    batch_kwargs = {"max_rows": batch_size}
    if max_output_cells_per_batch is not None:
        batch_kwargs["max_output_cells"] = max_output_cells_per_batch
    if max_context_bytes_per_batch is not None:
        batch_kwargs["max_context_bytes"] = max_context_bytes_per_batch
    jobs = []
    for sheet in workbook["sheets"]:
        plan = plans[sheet["name"]]
        numbers = data_rows(sheet, plan)
        for batch in plan_row_batches(sheet, plan, numbers, **batch_kwargs):
            jobs.append((sheet, plan, batch))

    def classify(job):
        sheet, plan, numbers = job
        context = rows_context(workbook, sheet, plan, numbers)
        result = infer(client, "row_semantics", ROWS_PROMPT, context, ROWS_SCHEMA,
                       lambda p: row_errors(p, workbook, sheet["name"], plan, numbers), review=review)
        return sheet["name"], result

    rows = {s["name"]: {} for s in workbook["sheets"]}
    uncertainties = list(layout["uncertainties"])
    # A failed batch stops pending ones (no further API calls for this
    # workbook) but never discards batches that already succeeded — those
    # rows are still real mapped/partial records; only the rows in failed
    # or never-attempted batches fall back to AI_NOT_CLASSIFIED.
    stopped_by_new_batch_limit = False
    if max_new_row_batches is None:
        completed, failures = _run_pool_partial(classify, jobs, concurrency)
    else:
        # Deterministic, sequential path for testing a fix on a small,
        # exact number of genuinely NEW (non-cached) row batches, reusing
        # whatever layout/row calls are already cached for free. This is
        # not a race-prone "best effort" like the concurrent path above:
        # jobs run one at a time in this one thread, so counting
        # client.call_count before/after each job is exact.
        completed, failures = [], []
        new_batches_done = 0
        for job in jobs:
            if new_batches_done >= max_new_row_batches:
                stopped_by_new_batch_limit = True
                break
            before = client.call_count
            try:
                completed.append(classify(job))
            except Exception as error:
                failures.append((job, error))
                break
            if client.call_count > before:
                new_batches_done += 1
            # else: this job was fully served from cache — does not count
            # against the new-batch limit, and processing continues.
    for sheet_name, result in completed:
        rows[sheet_name].update({r["row"]: r for r in result["rows"]})
        uncertainties.extend(result["uncertainties"])
    row_batches = {
        "planned": len(jobs), "completed": len(completed), "failed": len(failures),
        "stopped_early": bool(failures) and len(completed) + len(failures) < len(jobs),
        "stopped_by_new_batch_limit": stopped_by_new_batch_limit,
        "failures": [{"sheet": job[0]["name"], "rows": job[2], "error_type": type(error).__name__,
                      "message": str(error)} for job, error in failures],
    }
    return {"layout": layout, "rows": rows, "uncertainties": uncertainties, "row_batches": row_batches}


def submission_metadata(metadata):
    version = metadata.get("submission_version")
    month = metadata.get("submission_month")
    if version is not None and not isinstance(version, str):
        raise ValueError("submission_version must be the original string, never a float")
    if month is not None and (not isinstance(month, str) or not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", month)):
        raise ValueError("submission_month must be YYYY-MM")
    old = re.fullmatch(r"V(\d+)(?:\.(\d+))?", version or "")
    new = re.fullmatch(r"(\d{4})-(0[1-9]|1[0-2])-(\d+)", version or "")
    # Explicit backend metadata only. Neither target period nor version text
    # is used to manufacture a submission_month if it was not supplied.
    return {"version_raw": version, "cycle": int(old[1]) if old else metadata.get("submission_cycle"),
            "revision": int(old[2]) if old and old[2] else int(new[3]) if new else metadata.get("revision"),
            "submission_month": month, "status": "provided" if version or month else "unknown"}


def compile_records(workbook, semantic, metadata=None):
    """Deterministic expansion of the model's own cell-level interpretation.

    Python assigns record IDs and source metadata; it does not recompute,
    correct or invent any normalized value. A cell the model never answered
    for (inference failure, prepare-only run) is preserved as an unmapped,
    unresolved candidate rather than silently interpreted by Python.
    """
    metadata = metadata or {}
    submission = submission_metadata(metadata)
    index = cell_index(workbook)
    plans = {p["sheet"]: p for p in semantic["layout"]["sheets"]}
    records = []
    diagnostics = [{"code": "AI_UNCERTAINTY", "severity": "warning", "message": u} for u in semantic["uncertainties"]]

    def get(p):
        return index[(p["sheet"], p["cell"])] if p else None

    def pointer_ref(p):
        return ref(get(p)) if p else None

    def text(p):
        c = get(p)
        return str(c["value"]) if c and c.get("value") is not None else None

    file = workbook["file"]
    for sheet in workbook["sheets"]:
        plan = plans[sheet["name"]]
        periods = {p["column"]: p for p in plan["periods"]}
        row_plans = semantic["rows"][sheet["name"]]
        cells_by_row_col = {(row_number, c["column"]): c for row_number, row in row_plans.items() for c in row.get("cells", [])}
        for cell in sheet["cells"]:
            if not cell.get("has_content", True):
                continue
            row = row_plans.get(cell["row"])
            cell_desc = cells_by_row_col.get((cell["row"], cell["column"]))
            candidate = cell["data_type"] == "number" or cell.get("formula") is not None or cell_desc is not None
            if not candidate:
                continue
            reasons = []
            org_code_cell = index.get((sheet["name"], f"{plan['organization_code_column']}{cell['row']}")) if plan["organization_code_column"] else None
            org_name_cell = index.get((sheet["name"], f"{plan['organization_name_column']}{cell['row']}")) if plan["organization_name_column"] else None
            org_code_cell = org_code_cell if org_code_cell and org_code_cell.get("value") is not None else get(plan["fallback_organization_code"])
            org_name_cell = org_name_cell if org_name_cell and org_name_cell.get("value") is not None else get(plan["fallback_organization_name"])
            org_code = str(org_code_cell["value"]) if org_code_cell and org_code_cell.get("value") is not None else metadata.get("organization_code")
            org_name = str(org_name_cell["value"]) if org_name_cell and org_name_cell.get("value") is not None else metadata.get("organization_name")
            org_refs = [ref(c) for c in [org_code_cell, org_name_cell] if c]
            organization = {"code": org_code, "name": org_name, "original_name": str(org_name_cell["value"]) if org_name_cell and org_name_cell.get("value") is not None else None,
                            "scope": "organization" if org_code or org_name else "unknown", "provenance": org_refs}
            item = {"code": text(row["item_code"]) if row else None, "name": text(row["item_name"]) if row else None,
                    "header_path": [text(p) for p in row["header_path"]] if row else [],
                    "provenance": [pointer_ref(p) for p in ([*row["header_path"], row["item_name"], row["item_code"]] if row else []) if p],
                    "metric_code": row["metric_code"] if row else None}
            period = periods.get(cell["column"])
            effective_period = ({k: period[k] for k in ("year", "month", "quarter", "aggregation", "scenario")} if period else
                                {"year": None, "month": None, "quarter": None, "aggregation": "unknown", "scenario": "unknown"})
            evidence = period["evidence"] if period else []
            effective_period.update(raw_headers=[text(p) for p in evidence], provenance=[pointer_ref(p) for p in evidence])
            role = row["role"] if row else "reference"
            financial = cell_desc is not None and cell_desc["value_kind"] in {"amount", "ratio", "quantity"} and role not in {"reference", "unknown"}
            if cell_desc is None:
                value = {"original_value": cell.get("value"), "normalized_value": None, "status": "unresolved",
                         "value_kind": None, "currency": None, "source_unit": None, "multiplier": None,
                         "rule_ids": ["AI_NOT_CLASSIFIED"], "diagnostics": ["LLM이 이 후보 셀을 분류하지 않음; Python은 의미를 추정하지 않음"]}
                unit_evidence = []
                reasons.append("LLM이 분류하지 않은 누락 후보; 원본만 보존")
            else:
                normalized = Decimal(cell_desc["normalized_value"]) if cell_desc["normalized_value"] is not None else None
                try:
                    original_numeric = Decimal(cell_desc["source_value_read"]) if cell_desc["source_value_read"] is not None else None
                except InvalidOperation:
                    original_numeric = None
                implied_multiplier = normalized / original_numeric if normalized is not None and original_numeric not in (None, Decimal(0)) else None
                value = {"original_value": cell_desc["source_value_read"], "normalized_value": normalized,
                         "status": cell_desc["status"], "value_kind": cell_desc["value_kind"], "currency": None,
                         "source_unit": cell_desc["unit"], "multiplier": implied_multiplier,
                         "rule_ids": ["AI_CELL_DIRECT_INTERPRETATION"], "diagnostics": []}
                unit_evidence = cell_desc["unit_evidence"]
                if cell_desc["status"] != "confirmed":
                    reasons.append("값 정규화 미확정 또는 비숫자")
                # Independent cross-check for observation only. The model's
                # own value is kept either way; a mismatch is never corrected.
                if cell_desc["status"] == "confirmed" and cell_desc["value_kind"] in {"amount", "ratio", "quantity", "number"} and normalized is not None:
                    check = normalize_number(cell.get("value"), data_type=cell["data_type"], value_kind=cell_desc["value_kind"],
                                             unit=cell_desc["unit"], number_format=cell.get("number_format"))
                    if check["status"] == "confirmed" and check["normalized_value"] is not None and check["normalized_value"] != normalized:
                        diagnostics.append({"code": "AI_NORMALIZATION_DISAGREEMENT", "severity": "info",
                                            "message": "모델 정규화값과 Python 독립 검산값이 다름; 모델 반환값을 그대로 유지함",
                                            "record_id": "standard:" + cell["record_id"],
                                            "model_normalized_value": format(normalized, "f"),
                                            "python_crosscheck_value": format(check["normalized_value"], "f")})
            if cell.get("formula") and cell["formula"].get("cached_status") in {"missing", "error"}:
                reasons.append("수식 저장 계산값 없음 또는 오류")
            if not financial:
                reasons.append("재무 값 영역 또는 의미 미확정; 원본 후보 보존")
            if organization["scope"] == "unknown":
                reasons.append("조직 미확정")
            if not item["name"]:
                reasons.append("항목 미확정")
            if effective_period["year"] is None or effective_period["aggregation"] == "unknown":
                reasons.append("대상 기간 미확정")
            if effective_period["scenario"] == "unknown":
                reasons.append("계획·추정·실적 미확정")
            if financial and value["value_kind"] in {"amount", "quantity"} and value["source_unit"] is None:
                reasons.append("금액 단위 미확정")
            dimensions = {}
            if row:
                for dimension in plan["dimensions"]:
                    original = index.get((sheet["name"], f"{dimension['column']}{cell['row']}"))
                    if original:
                        dimensions[dimension["name"]] = {"value": original.get("value"), "source": ref(original)}
            records.append({
                "record_id": "standard:" + cell["record_id"],
                "source": {"raw_record_id": cell["record_id"], "file_id": file["file_id"], "source_id": file["source_id"],
                           "path": file["path"], "sheet": sheet["name"], "cell": cell["address"], "workbook_role": "unknown"},
                "organization": organization, "item": item, "dimensions": dimensions, "period": effective_period,
                "role": {"kind": role, "aggregatable": False, "reason": "집계 정책과 제출본 선택은 이번 AI 파서 범위 밖"},
                "value": value, "unit_provenance": [pointer_ref(p) for p in unit_evidence],
                "submission": dict(submission),
                "mapping": {"status": "mapped" if not reasons else "partial" if financial else "unmapped",
                            "rule_ids": ["AI_LAYOUT", "AI_ROW_SEMANTICS", "AI_CELL_DIRECT_INTERPRETATION"] if cell_desc is not None else ["AI_NOT_CLASSIFIED"],
                            "reasons": reasons},
                "formula": cell.get("formula"), "raw_value": cell.get("value"), "data_type": cell["data_type"],
                "style_reference": {"style_id": cell.get("style_id"), "number_format": cell.get("number_format"), "raw_record_id": cell["record_id"]},
            })
    # A workbook's organization scope comes from its own actual detail rows.
    orgs = {r["organization"]["code"] for r in records if r["role"]["kind"] == "detail" and r["organization"]["code"]}
    has_summary = any(p["kind"] == "organization_summary" for p in plans.values())
    workbook_role = "consolidated_reference" if len(orgs) > 1 and has_summary else "organization_workbook" if len(orgs) == 1 else "unknown"
    for record in records:
        record["source"]["workbook_role"] = workbook_role
    gap = []
    for record in records:
        metric = record["item"]["metric_code"]
        if metric:
            gap.append({"status": "explicit_excluded_value" if metric == "OP_EX_CONTINGENCY" else "unresolved",
                        "preferred_record_ids": [record["record_id"]] if metric == "OP_EX_CONTINGENCY" else [],
                        "metric_records": {metric: [record["record_id"]]},
                        "formulas": [{"source": record["source"], "formula": record["formula"]}] if record["formula"] else [],
                        "calculation_performed": False, "reason": "원본 지표만 보존; Gap·Contingency 조정은 계산하지 않음",
                        "rule_id": "AI_SOURCE_METRIC_ONLY"})
    return {"records": records, "diagnostics": diagnostics,
            "summary": {"records": len(records), "workbook_role": workbook_role, **dict(Counter(r["mapping"]["status"] for r in records))},
            "gap_relationships": gap}


def unresolved_records(workbook):
    """Inference failure still preserves numeric/formula candidates, without a rule fallback."""
    plans = []
    for sheet in workbook["sheets"]:
        plans.append({"sheet": sheet["name"], "kind": "unknown", "header_end_row": 0,
                      "organization_code_column": None, "organization_name_column": None,
                      "fallback_organization_code": None, "fallback_organization_name": None,
                      "unit": None, "unit_evidence": [], "periods": [], "dimensions": [], "notes": []})
    return compile_records(workbook, {"layout": {"sheets": plans}, "rows": {s["sheet"]: {} for s in plans}, "uncertainties": ["LLM 의미 해석 미완료"]})
