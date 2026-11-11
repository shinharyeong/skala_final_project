"""Financial integrity, model failure, grounding, coverage and API contract tests.

These tests use mocked model responses. They verify the plumbing (schema,
grounding, coverage, no silent Python correction, contract compatibility).
They do not and cannot verify real LLM parsing quality/accuracy.
"""
import contextlib
import copy
import gzip
import io
import json
import tempfile
import time
import unittest
import urllib.error
from decimal import Decimal
from email.message import Message
from pathlib import Path
from unittest.mock import patch

from excel_parser_ai.client import ModelError, ResponsesClient
from excel_parser_ai.context import (
    compact_cell, plan_row_batches, rows_context, sheet_layout_context, value_columns, workbook_context,
)
from excel_parser_ai.engine import (
    _run_pool, _run_pool_partial, cell_errors, compile_records, infer, infer_workbook, layout_errors,
    row_errors, schema_errors, sheet_plan_errors, submission_metadata, unresolved_records,
)
from excel_parser_ai.schema import LAYOUT_SCHEMA, ROWS_SCHEMA, SHEET_LAYOUT_SCHEMA
from excel_parser_ai.shared import validate_provenance
from excel_parser_ai.pipeline import run


def fixture(value="100", *, data_type="number", formula=None, cached_value=None,
            status="confirmed", value_kind="amount", normalized="100000", unit="천원",
            source_value_read="__default__"):
    file = {"file_id": "hash", "source_id": "source", "path": "/fixture.xlsx", "name": "fixture.xlsx"}
    values = {"A1": ("Unit: 천원", "string"), "C2": ("2026년 계획", "string"),
              "C3": ("1월", "string"), "A4": ("F01", "string"), "B4": ("일반광고비", "string"),
              "C4": (value, data_type)}
    cells = []
    for address, (v, kind) in values.items():
        col, row = address[0], int(address[1:])
        cell_formula = formula if address == "C4" else None
        cells.append({"record_id": "source/1/" + address, "file_id": "hash", "sheet_id": "1", "sheet": "Sheet",
                      "address": address, "row": row, "column": col, "column_index": ord(col) - 64,
                      "value": v, "data_type": kind, "formula": cell_formula,
                      "cached_value": cached_value if address == "C4" else None,
                      "has_content": True, "style_id": 0, "number_format": "#,##0;[Red](#,##0)"})
    workbook = {"file": file, "sheets": [{"name": "Sheet", "cells": cells,
                 "counts": {"content_cells": len(cells), "formulas": int(formula is not None)}}], "diagnostics": []}
    r = lambda a: {"sheet": "Sheet", "cell": a}
    sheet = {"sheet": "Sheet", "kind": "expense_detail", "header_end_row": 3,
             "organization_code_column": "A", "organization_name_column": None,
             "fallback_organization_code": None, "fallback_organization_name": None,
             "unit": "천원", "unit_evidence": [r("A1")],
             "periods": [{"column": "C", "year": 2026, "month": 1, "quarter": None,
                          "aggregation": "month", "scenario": "plan", "evidence": [r("C2"), r("C3")]}],
             "dimensions": [], "notes": []}
    if source_value_read == "__default__":
        source_value_read = cached_value if formula else value
    cell_desc = {"column": "C", "status": status, "value_kind": value_kind,
                 "source_value_read": source_value_read, "normalized_value": normalized,
                 "unit": unit, "unit_evidence": [r("A1")] if unit else [], "reason": "source value"}
    row = {"row": 4, "role": "detail", "item_name": r("B4"), "item_code": None,
           "header_path": [r("B4")], "metric_code": None, "cells": [cell_desc], "reason": "source label"}
    semantic = {"layout": {"sheets": [sheet], "uncertainties": []}, "rows": {"Sheet": {4: row}}, "uncertainties": []}
    return workbook, semantic


def financial(mapped):
    return next(r for r in mapped["records"] if r["source"]["cell"] == "C4")


class FinancialIntegrityTests(unittest.TestCase):
    def test_model_normalized_value_is_passed_through_unmodified(self):
        """The model returns the normalized value directly; Python only copies it."""
        w, s = fixture(value="123456789012345678901234567890.125", normalized="123456789012345678901234567890125.000")
        record = financial(compile_records(w, s))
        self.assertEqual(str(record["value"]["normalized_value"]), "123456789012345678901234567890125.000")
        self.assertEqual(record["value"]["original_value"], "123456789012345678901234567890.125")
        self.assertEqual(record["raw_value"], "123456789012345678901234567890.125")
        self.assertFalse(record["role"]["aggregatable"])
        self.assertTrue(validate_provenance(w, compile_records(w, s))["passed"])

    def test_wrong_model_normalization_is_kept_and_only_flagged_as_a_diagnostic(self):
        """Python must never silently correct a model value it disagrees with."""
        # Real stored value 100, confirmed unit 천원 => correct conversion is 100000.
        # The model instead (wrongly) forgets the unit multiplier and returns 100.
        w, s = fixture(value="100", normalized="100")
        mapped = compile_records(w, s)
        record = financial(mapped)
        self.assertEqual(record["value"]["normalized_value"], Decimal("100"))  # kept as the model said
        disagreement = [d for d in mapped["diagnostics"] if d["code"] == "AI_NORMALIZATION_DISAGREEMENT"]
        self.assertEqual(len(disagreement), 1)
        self.assertEqual(disagreement[0]["model_normalized_value"], "100")
        self.assertEqual(disagreement[0]["python_crosscheck_value"], "100000")

    def test_visual_red_parentheses_do_not_flip_stored_positive(self):
        w, s = fixture(value="100", normalized="100000")
        self.assertEqual(financial(compile_records(w, s))["value"]["normalized_value"], Decimal("100000"))
        w, s = fixture(value="-100", normalized="-100000")
        self.assertEqual(financial(compile_records(w, s))["value"]["normalized_value"], Decimal("-100000"))

    def test_zero_is_not_missing_and_dash_is_not_zero(self):
        w, s = fixture(value="0", normalized="0")
        self.assertEqual(financial(compile_records(w, s))["value"]["normalized_value"], Decimal("0"))
        w, s = fixture(value="-", data_type="string", status="missing", value_kind="text", normalized=None, unit=None)
        r = financial(compile_records(w, s))
        self.assertIsNone(r["value"]["normalized_value"])
        self.assertEqual(r["value"]["status"], "missing")

    def test_missing_formula_cache_preserves_formula_and_stays_unresolved(self):
        formula = {"text": "=SUM(C5:C8)", "cached_status": "missing"}
        w, s = fixture(value=None, data_type="blank", formula=formula, cached_value=None,
                       status="unresolved", value_kind="amount", normalized=None, unit=None,
                       source_value_read=None)
        r = financial(compile_records(w, s))
        self.assertIsNone(r["value"]["normalized_value"])
        self.assertEqual(r["formula"], formula)
        self.assertEqual(r["mapping"]["status"], "partial")
        # A model that claims confirmed+a value despite a missing cache must fail validation.
        bad_cell = {"column": "C", "status": "confirmed", "value_kind": "amount",
                    "source_value_read": "999", "normalized_value": "999000", "unit": "천원",
                    "unit_evidence": [{"sheet": "Sheet", "cell": "A1"}], "reason": "x"}
        real = next(c for c in w["sheets"][0]["cells"] if c["address"] == "C4")
        self.assertTrue(cell_errors(bad_cell, real, "Sheet", "C4"))

    def test_ratio_storage_is_not_scaled_twice(self):
        w, s = fixture(value="0.125", value_kind="ratio", normalized="0.125", unit=None)
        r = financial(compile_records(w, s))
        self.assertEqual(r["value"]["normalized_value"], Decimal("0.125"))
        self.assertEqual(r["value"]["value_kind"], "ratio")

    def test_unclassified_candidate_is_not_interpreted_by_python(self):
        """If the LLM never answers for a cell, Python leaves it unresolved
        rather than quietly computing a value for it itself."""
        w, _ = fixture()
        mapped = unresolved_records(w)
        self.assertTrue(validate_provenance(w, mapped)["passed"])
        self.assertTrue(all(r["mapping"]["status"] == "unmapped" for r in mapped["records"]))
        record = financial(mapped)
        self.assertIsNone(record["value"]["normalized_value"])
        self.assertEqual(record["value"]["rule_ids"], ["AI_NOT_CLASSIFIED"])

    def test_no_rule_based_mapping_is_used_as_a_fallback(self):
        import excel_parser_ai.engine as engine_module
        self.assertFalse(hasattr(engine_module, "map_workbook"))
        w, s = fixture()
        mapped = compile_records(w, s)
        for record in mapped["records"]:
            for rule in record["mapping"]["rule_ids"]:
                self.assertTrue(rule.startswith("AI_"))

    def test_target_year_does_not_become_submission_date(self):
        w, s = fixture()
        self.assertIsNone(financial(compile_records(w, s))["submission"]["submission_month"])
        self.assertEqual(submission_metadata({"submission_version": "2026-11-02"})["revision"], 2)
        self.assertIsNone(submission_metadata({"submission_version": "2026-11-02"})["submission_month"])
        with self.assertRaises(ValueError):
            submission_metadata({"submission_version": 1.01})


class OutputContractCompatibilityTests(unittest.TestCase):
    def test_standard_record_has_existing_contract_fields_and_types(self):
        w, s = fixture()
        record = financial(compile_records(w, s))
        for key in ("record_id", "source", "organization", "item", "dimensions", "period",
                    "role", "value", "unit_provenance", "submission", "mapping",
                    "formula", "raw_value", "data_type", "style_reference"):
            self.assertIn(key, record)
        self.assertIsInstance(record["record_id"], str)
        self.assertIsInstance(record["value"]["normalized_value"], Decimal)
        self.assertFalse(record["role"]["aggregatable"])
        self.assertIn(record["mapping"]["status"], {"mapped", "partial", "unmapped"})
        self.assertIn(record["value"]["status"], {"confirmed", "unresolved", "not_numeric", "missing"})


class ContextGroundingTests(unittest.TestCase):
    def test_numbers_are_not_masked_in_model_context(self):
        """The core architecture requirement: real numbers must reach the model."""
        w, _ = fixture(value="123456")
        cell = next(c for c in w["sheets"][0]["cells"] if c["address"] == "C4")
        compact = compact_cell(cell)
        self.assertEqual(compact["value"], "123456")
        self.assertNotIn("masked", str(compact).lower())
        context = workbook_context(w)
        serialized = json.dumps(context, ensure_ascii=False)
        self.assertIn("123456", serialized)

    def test_formula_cache_value_is_exposed_to_the_model(self):
        formula = {"text": "=C1+C2", "cached_status": "stored_unverified"}
        w, _ = fixture(value="42", formula=formula, cached_value="42")
        cell = next(c for c in w["sheets"][0]["cells"] if c["address"] == "C4")
        compact = compact_cell(cell)
        self.assertEqual(compact["cached_value"], "42")

    def test_value_columns_cover_numeric_formula_and_configured_period_cells(self):
        w, s = fixture()
        sheet = w["sheets"][0]
        plan = s["layout"]["sheets"][0]
        row_cells = [c for c in sheet["cells"] if c["row"] == 4]
        self.assertEqual(value_columns(row_cells, plan), ["C"])


class ModelGroundingTests(unittest.TestCase):
    def test_layout_rejects_fabricated_evidence_and_duplicate_columns(self):
        w, s = fixture()
        p = copy.deepcopy(s["layout"])
        p["sheets"][0]["periods"][0]["evidence"][0]["cell"] = "Z999"
        self.assertTrue(layout_errors(p, w))
        p = copy.deepcopy(s["layout"])
        p["sheets"][0]["periods"].append(p["sheets"][0]["periods"][0])
        self.assertTrue(layout_errors(p, w))

    def test_rows_reject_missing_duplicate_and_wrong_row_pointers(self):
        w, s = fixture()
        plan = s["layout"]["sheets"][0]
        row = s["rows"]["Sheet"][4]
        self.assertTrue(row_errors({"rows": [], "uncertainties": []}, w, "Sheet", plan, [4]))
        self.assertTrue(row_errors({"rows": [row, row], "uncertainties": []}, w, "Sheet", plan, [4]))
        bad = copy.deepcopy(row)
        bad["item_name"]["cell"] = "C2"
        self.assertTrue(row_errors({"rows": [bad], "uncertainties": []}, w, "Sheet", plan, [4]))

    def test_rows_reject_missing_and_duplicate_cell_columns(self):
        w, s = fixture()
        plan = s["layout"]["sheets"][0]
        row = s["rows"]["Sheet"][4]
        no_cells = copy.deepcopy(row)
        no_cells["cells"] = []
        self.assertTrue(row_errors({"rows": [no_cells], "uncertainties": []}, w, "Sheet", plan, [4]))
        dup = copy.deepcopy(row)
        dup["cells"] = dup["cells"] * 2
        self.assertTrue(row_errors({"rows": [dup], "uncertainties": []}, w, "Sheet", plan, [4]))

    def test_cell_grounding_rejects_a_value_not_present_in_the_source(self):
        w, s = fixture(value="100")
        plan = s["layout"]["sheets"][0]
        row = s["rows"]["Sheet"][4]
        hallucinated = copy.deepcopy(row)
        hallucinated["cells"][0]["source_value_read"] = "999999"  # not the real stored value
        self.assertTrue(row_errors({"rows": [hallucinated], "uncertainties": []}, w, "Sheet", plan, [4]))

    def test_cell_requires_unit_evidence_and_valid_decimal_string(self):
        w, s = fixture()
        plan = s["layout"]["sheets"][0]
        row = s["rows"]["Sheet"][4]
        no_evidence = copy.deepcopy(row)
        no_evidence["cells"][0]["unit_evidence"] = []
        self.assertTrue(row_errors({"rows": [no_evidence], "uncertainties": []}, w, "Sheet", plan, [4]))
        bad_decimal = copy.deepcopy(row)
        bad_decimal["cells"][0]["normalized_value"] = "100,000"  # not a plain decimal string
        self.assertTrue(row_errors({"rows": [bad_decimal], "uncertainties": []}, w, "Sheet", plan, [4]))

    def test_schema_forbids_extra_undeclared_fields(self):
        _, s = fixture()
        row = copy.deepcopy(s["rows"]["Sheet"][4])
        row["amount"] = 999
        self.assertTrue(schema_errors({"rows": [row], "uncertainties": []}, ROWS_SCHEMA))

    def test_missing_rows_are_repaired_before_acceptance(self):
        w, s = fixture()
        plan = s["layout"]["sheets"][0]
        good = {"rows": [s["rows"]["Sheet"][4]], "uncertainties": []}
        class Client:
            def __init__(self):
                self.calls = 0
            def ask(self, *args):
                self.calls += 1
                return {"rows": [], "uncertainties": []} if self.calls == 1 else good
        c = Client()
        result = infer(c, "rows", "", {}, ROWS_SCHEMA, lambda p: row_errors(p, w, "Sheet", plan, [4]), review=False)
        self.assertEqual(result, good)
        self.assertEqual(c.calls, 2)

    def test_invalid_review_cannot_silently_override(self):
        w, s = fixture()
        plan = s["layout"]["sheets"][0]
        good = {"rows": [s["rows"]["Sheet"][4]], "uncertainties": []}
        class Client:
            def ask(self, task, *args):
                return {"rows": [], "uncertainties": []} if task.endswith("review") else good
        with self.assertRaises(ValueError):
            infer(Client(), "rows", "", {}, ROWS_SCHEMA, lambda p: row_errors(p, w, "Sheet", plan, [4]))


class APIContractTests(unittest.TestCase):
    def test_structured_output_cache_and_no_model_downgrade(self):
        with tempfile.TemporaryDirectory() as tmp:
            calls = []
            def transport(body):
                calls.append(body)
                return {"status": "completed", "model": "test-model", "usage": {"input_tokens": 2, "output_tokens": 3},
                        "output": [{"content": [{"type": "output_text", "text": '{"sheets":[],"uncertainties":[]}'}]}]}
            c = ResponsesClient(tmp, model="test-model", effort="high", transport=transport)
            c.ask("layout", "source only", {}, LAYOUT_SCHEMA)
            c.ask("layout", "source only", {}, LAYOUT_SCHEMA)
            self.assertEqual(len(calls), 1)
            self.assertFalse(calls[0]["store"])
            self.assertEqual(calls[0]["model"], "test-model")
            self.assertTrue(calls[0]["text"]["format"]["strict"])
            self.assertNotIn("temperature", calls[0])
            self.assertTrue(c.events[1]["cached"])

    def test_refusal_and_truncation_are_explicit_failures(self):
        responses = [{"status": "incomplete"}, {"status": "completed", "output": [{"content": [{"type": "refusal"}]}]}]
        for response in responses:
            with tempfile.TemporaryDirectory() as tmp:
                c = ResponsesClient(tmp, transport=lambda body: response)
                with self.assertRaises(ModelError):
                    c.ask("layout", "", {}, LAYOUT_SCHEMA)

    def test_key_is_never_needed_for_mock_or_read_from_source(self):
        with tempfile.TemporaryDirectory() as tmp:
            c = ResponsesClient(tmp)
            with patch.dict("os.environ", {}, clear=True):
                with self.assertRaisesRegex(ModelError, "OPENAI_API_KEY"):
                    c._request({})


class HTTPErrorDiagnosticsTests(unittest.TestCase):
    """A real HTTP error (e.g. 403) must surface code/type/message/request id
    without ever leaking the API key, the Authorization header, or retrying
    a non-transient failure, and without silently switching models."""

    def _http_error(self, status, error_body, *, request_id="req_abc123xyz"):
        headers = Message()
        if request_id:
            headers.add_header("x-request-id", request_id)
        body = json.dumps({"error": error_body}).encode()
        return urllib.error.HTTPError("https://api.openai.com/v1/responses", status, "error", headers, io.BytesIO(body))

    def test_403_detail_has_no_secrets_and_surfaces_code_type_message_request_id(self):
        http_error = self._http_error(403, {"code": "insufficient_permissions", "type": "permission_error",
                                            "message": "You do not have access to this model or project."})
        with tempfile.TemporaryDirectory() as tmp:
            with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-should-never-appear-in-any-error-text"}):
                c = ResponsesClient(tmp, model="gpt-6-astra")
                with patch("excel_parser_ai.client.urllib.request.urlopen", side_effect=http_error):
                    with self.assertRaises(ModelError) as ctx:
                        c._request({"model": "gpt-6-astra"})
        message = str(ctx.exception)
        for expected in ("403", "insufficient_permissions", "permission_error",
                          "You do not have access to this model or project.", "req_abc123xyz", "gpt-6-astra"):
            self.assertIn(expected, message)
        for secret in ("sk-should-never-appear-in-any-error-text", "Authorization", "Bearer"):
            self.assertNotIn(secret, message)

    def test_403_is_not_retried_and_does_not_change_model(self):
        http_error = self._http_error(403, {"code": "insufficient_permissions", "type": "permission_error", "message": "forbidden"})
        calls = []
        def fake_urlopen(request, timeout=None):
            calls.append(request.get_header("Authorization"))  # never asserted/printed; just counts attempts
            raise http_error
        with tempfile.TemporaryDirectory() as tmp:
            with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}):
                c = ResponsesClient(tmp, model="gpt-6-astra")
                with patch("excel_parser_ai.client.urllib.request.urlopen", side_effect=fake_urlopen):
                    with self.assertRaises(ModelError):
                        c._request({"model": "gpt-6-astra"})
        self.assertEqual(len(calls), 1)  # no retry for a non-transient 403
        self.assertEqual(c.model, "gpt-6-astra")  # model is never swapped on failure

    def test_520_is_retried_with_existing_bounded_http_policy(self):
        attempts = {"n": 0}
        def fake_urlopen(request, timeout=None):
            attempts["n"] += 1
            raise self._http_error(520, {"code": None, "type": "server_error", "message": "transient"})
        with tempfile.TemporaryDirectory() as tmp:
            with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}):
                c = ResponsesClient(tmp, model="gpt-5.4")
                with patch("excel_parser_ai.client.urllib.request.urlopen", side_effect=fake_urlopen), \
                     patch("time.sleep", return_value=None):
                    with self.assertRaises(ModelError):
                        c._request({"model": "gpt-5.4"})
        self.assertEqual(attempts["n"], 4)
        self.assertEqual(c.http_post_attempts, 4)

    def test_malformed_error_body_still_yields_a_safe_diagnostic(self):
        headers = Message()
        headers.add_header("x-request-id", "req_zzz")
        http_error = urllib.error.HTTPError("https://api.openai.com/v1/responses", 500, "error", headers, io.BytesIO(b"not json"))
        with tempfile.TemporaryDirectory() as tmp:
            with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}):
                c = ResponsesClient(tmp, model="gpt-6-astra")
                # Non-transient after exhausting retries: force attempt to be the last one by
                # using a 500 code (retryable) but confirm it still raises with a safe detail
                # once retries are exhausted, rather than hanging or leaking anything.
                with patch("excel_parser_ai.client.urllib.request.urlopen", side_effect=http_error), \
                     patch("time.sleep", return_value=None):
                    with self.assertRaises(ModelError) as ctx:
                        c._request({"model": "gpt-6-astra"})
        self.assertIn("req_zzz", str(ctx.exception))
        self.assertIn("500", str(ctx.exception))


class PipelineIntegrationTests(unittest.TestCase):
    def execute(self, tmp, *, fail=False):
        w, s = fixture(value="-100", normalized="-100000")
        root = Path(tmp)
        source = root / "input.xlsx"
        source.write_bytes(b"read-only test input; reader boundary is patched")
        def transport(body):
            if fail:
                raise ModelError("simulated API failure")
            task = body["text"]["format"]["name"]
            if task.startswith("workbook_layout"):
                data = {"sheet_plan": s["layout"]["sheets"][0], "uncertainties": s["layout"]["uncertainties"]}
            else:
                data = {"rows": [s["rows"]["Sheet"][4]], "uncertainties": []}
            return {"status": "completed", "output": [{"content": [{"type": "output_text", "text": json.dumps(data)}]}]}
        client = ResponsesClient(root / "cache", model="mock-only", transport=transport)
        with patch("excel_parser_ai.pipeline.REPO", root), patch("excel_parser_ai.pipeline.extract_workbook", return_value=w), patch("excel_parser_ai.pipeline.validate_extraction", return_value={"passed": True}):
            summary = run([source], root / "parser_ai/output/run", client=client)
        return summary, root / "parser_ai/output/run"

    def test_full_mocked_model_pipeline_emits_existing_contract_without_scoring(self):
        with tempfile.TemporaryDirectory() as tmp:
            summary, output = self.execute(tmp)
            self.assertTrue(summary["llm_inference_completed"])
            self.assertTrue(summary["validation_passed"])
            self.assertFalse(summary["performance_comparison_performed"])
            self.assertFalse(summary["business_master_issued"])
            with gzip.open(output / "standard-records.jsonl.gz", "rt") as f:
                records = [json.loads(line) for line in f]
            self.assertEqual(financial({"records": records})["value"]["normalized_value"], "-100000")
            self.assertEqual(records[0]["contract_version"], "1.0.0")
            self.assertEqual(json.loads((output / "validation.json").read_text())["master_comparison"]["status"], "not_performed")
            inference = json.loads((output / "inference.json").read_text())
            self.assertTrue(inference["financial_amounts_generated_by_model"])
            self.assertFalse(inference["python_normalization_fallback_used"])

    def test_api_failure_preserves_raw_and_does_not_fake_inference_success(self):
        with tempfile.TemporaryDirectory() as tmp:
            summary, output = self.execute(tmp, fail=True)
            self.assertFalse(summary["llm_inference_completed"])
            self.assertFalse(summary["validation_passed"])
            self.assertTrue(summary["raw_extraction_complete"])
            self.assertTrue(summary["originals_unchanged"])
            with gzip.open(output / "raw-cells.jsonl.gz", "rt") as f:
                self.assertEqual(len(list(f)), 6)
            self.assertTrue(summary["failures"])
            with gzip.open(output / "standard-records.jsonl.gz", "rt") as f:
                records = [json.loads(line) for line in f]
            # API failure must not silently fall back to a rule-based/guessed value.
            self.assertIsNone(financial({"records": records})["value"]["normalized_value"])
            self.assertEqual(financial({"records": records})["mapping"]["status"], "unmapped")


class LayoutSizeReductionTests(unittest.TestCase):
    """The layout call is the one that was observed hanging: it must not
    carry real numeric amounts (the row/cell stage still gets real numbers,
    unaffected — see ContextGroundingTests), and must be scoped to one
    sheet at a time rather than resending the whole workbook every time."""

    def test_layout_context_masks_numbers_but_keeps_structure(self):
        w, _ = fixture(value="999999")
        ctx = sheet_layout_context(w, "Sheet")
        serialized = json.dumps(ctx, ensure_ascii=False)
        self.assertNotIn("999999", serialized)  # real amount withheld at layout stage
        self.assertEqual(ctx["target_sheet"]["sheet"], "Sheet")
        self.assertIn("cross_sheet_reference", ctx)
        # Structure the layout stage needs is still present (format strings
        # are deduplicated into a lookup table rather than dropped).
        self.assertTrue(any(c.get("format_id") for c in ctx["target_sheet"]["cells"]))
        self.assertTrue(ctx["target_sheet"]["number_formats"])

    def test_repeated_number_formats_are_deduplicated_into_a_lookup_table(self):
        w, _ = fixture(value="1")
        # Give every cell in the fixture's single row the same format, to
        # exercise dedup (fixture already sets one format on all cells).
        ctx = sheet_layout_context(w, "Sheet")
        number_cells = [c for c in ctx["target_sheet"]["cells"] if c.get("format_id")]
        self.assertTrue(number_cells)
        format_ids = {c["format_id"] for c in number_cells}
        # Same source format string -> same format_id, looked up once.
        self.assertEqual(len(format_ids), len(set(ctx["target_sheet"]["number_formats"]) & format_ids))
        for format_id in format_ids:
            self.assertIn(format_id, ctx["target_sheet"]["number_formats"])

    def test_cross_sheet_reference_is_label_only_and_excludes_numbers(self):
        w, _ = fixture(value="123456789")
        original = w["sheets"][0]
        other_cells = [dict(c, sheet="Other") for c in original["cells"]]
        w["sheets"].append({"name": "Other", "cells": other_cells, "counts": dict(original["counts"])})
        ctx = sheet_layout_context(w, "Sheet")
        cross = next(r for r in ctx["cross_sheet_reference"] if r["sheet"] == "Other")
        serialized = json.dumps(cross, ensure_ascii=False)
        self.assertNotIn("123456789", serialized)
        self.assertTrue(cross["labels"])  # the string header labels are still visible

    def test_row_cell_stage_is_unaffected_and_still_gets_real_numbers(self):
        # Guard against a regression where the layout fix accidentally
        # masks numbers in rows_context too (that stage must keep them).
        w, s = fixture(value="123456")
        plan = s["layout"]["sheets"][0]
        ctx = rows_context(w, w["sheets"][0], plan, [4])
        self.assertIn("123456", json.dumps(ctx, ensure_ascii=False))


class SheetLayoutValidationTests(unittest.TestCase):
    def test_accepts_a_valid_single_sheet_plan(self):
        w, s = fixture()
        result = {"sheet_plan": s["layout"]["sheets"][0], "uncertainties": []}
        self.assertEqual(sheet_plan_errors(result, w, "Sheet"), [])

    def test_rejects_a_plan_for_the_wrong_sheet(self):
        w, s = fixture()
        result = {"sheet_plan": dict(s["layout"]["sheets"][0], sheet="WrongSheet"), "uncertainties": []}
        self.assertTrue(sheet_plan_errors(result, w, "Sheet"))

    def test_rejects_fabricated_evidence(self):
        w, s = fixture()
        plan = copy.deepcopy(s["layout"]["sheets"][0])
        plan["periods"][0]["evidence"][0]["cell"] = "Z999"
        result = {"sheet_plan": plan, "uncertainties": []}
        self.assertTrue(sheet_plan_errors(result, w, "Sheet"))


class PerSheetLayoutCallTests(unittest.TestCase):
    def test_infer_workbook_issues_one_layout_call_per_sheet_not_one_combined_call(self):
        w, _ = fixture(value="42")
        original = w["sheets"][0]
        w["sheets"].append({"name": "Sheet2", "cells": [dict(c, sheet="Sheet2") for c in original["cells"]],
                            "counts": dict(original["counts"])})
        tasks = []

        class FakeClient:
            model = "fake"
            effort = "high"

            def ask(self, task, instructions, context, schema):
                tasks.append(task)
                if task.startswith("workbook_layout"):
                    sheet_name = context["target_sheet"]["sheet"]
                    return {"sheet_plan": {"sheet": sheet_name, "kind": "unknown", "header_end_row": 0,
                                           "organization_code_column": None, "organization_name_column": None,
                                           "fallback_organization_code": None, "fallback_organization_name": None,
                                           "unit": None, "unit_evidence": [], "periods": [], "dimensions": [], "notes": []},
                            "uncertainties": []}
                rows_out = []
                for row in context["rows"]:
                    cells_out = [{"column": cell["cell"][0], "status": "not_numeric", "value_kind": "text",
                                 "source_value_read": str(cell["value"]) if cell.get("value") is not None else None,
                                 "normalized_value": None, "unit": None, "unit_evidence": [], "reason": "fake"}
                                for cell in row["cells"]]
                    rows_out.append({"row": row["row"], "role": "reference", "item_name": None, "item_code": None,
                                     "header_path": [], "metric_code": None, "cells": cells_out, "reason": "fake"})
                return {"rows": rows_out, "uncertainties": []}

        semantic = infer_workbook(w, FakeClient(), review=False, batch_size=64, concurrency=2)
        layout_calls = [t for t in tasks if t == "workbook_layout_sheet"]
        self.assertEqual(len(layout_calls), 2)  # one per sheet, not one combined call
        self.assertEqual({p["sheet"] for p in semantic["layout"]["sheets"]}, {"Sheet", "Sheet2"})


class RunPoolTests(unittest.TestCase):
    """_run_pool backs both the per-sheet layout calls and the row/cell
    batches; it must not add its own blocking wait on top of a failure or
    a Ctrl+C, since a worker thread's in-flight network call can't be
    forcibly cancelled either way."""

    def test_results_are_returned_for_every_item(self):
        self.assertEqual(sorted(_run_pool(lambda x: x * 2, [1, 2, 3], concurrency=2)), [2, 4, 6])

    def test_keyboard_interrupt_propagates_instead_of_hanging_or_being_swallowed(self):
        def job(x):
            if x == 0:
                raise KeyboardInterrupt
            time.sleep(0.05)
            return x
        with self.assertRaises(KeyboardInterrupt):
            _run_pool(job, [0, 1, 2], concurrency=1)


class ProgressVisibilityTests(unittest.TestCase):
    """Request start, stage/task, attempt count, completion and elapsed
    time must be visible — the original complaint was a silent hang."""

    def test_ask_reports_start_and_completion_with_task_name(self):
        messages = []
        with tempfile.TemporaryDirectory() as tmp:
            def transport(body):
                return {"status": "completed", "output": [{"content": [{"type": "output_text", "text": '{"sheets":[],"uncertainties":[]}'}]}]}
            c = ResponsesClient(tmp, model="m", transport=transport, progress=messages.append)
            c.ask("workbook_layout_sheet", "x", {}, LAYOUT_SCHEMA)
        joined = " ".join(messages)
        self.assertIn("workbook_layout_sheet", joined)
        self.assertTrue(any("처리 시작" in m for m in messages))
        self.assertTrue(any("완료" in m for m in messages))
        for secret in ("Authorization", "Bearer"):
            self.assertNotIn(secret, joined)

    def test_cache_hit_is_reported_without_a_second_api_call(self):
        with tempfile.TemporaryDirectory() as tmp:
            calls = []
            def transport(body):
                calls.append(body)
                return {"status": "completed", "output": [{"content": [{"type": "output_text", "text": '{"sheets":[],"uncertainties":[]}'}]}]}
            messages = []
            c = ResponsesClient(tmp, model="m", transport=transport, progress=messages.append)
            c.ask("workbook_layout_sheet", "x", {}, LAYOUT_SCHEMA)
            messages.clear()
            c.ask("workbook_layout_sheet", "x", {}, LAYOUT_SCHEMA)
        self.assertEqual(len(calls), 1)
        self.assertTrue(any("캐시" in m for m in messages))

    def test_real_request_path_reports_attempt_and_retry_without_secrets(self):
        headers = Message()
        err_body = json.dumps({"error": {"code": "rate_limit_exceeded", "type": "rate_limit_error", "message": "slow down"}}).encode()
        http_error = urllib.error.HTTPError("https://api.openai.com/v1/responses", 429, "error", headers, io.BytesIO(err_body))
        success_bytes = json.dumps({"status": "completed", "output": [{"content": [{"type": "output_text", "text": '{"sheets":[],"uncertainties":[]}'}]}]}).encode()
        attempts = {"n": 0}

        def fake_urlopen(request, timeout=None):
            attempts["n"] += 1
            if attempts["n"] == 1:
                raise http_error
            return io.BytesIO(success_bytes)

        messages = []
        with tempfile.TemporaryDirectory() as tmp:
            with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-should-never-appear-anywhere"}):
                c = ResponsesClient(tmp, model="gpt-6-astra", progress=messages.append)
                with patch("excel_parser_ai.client.urllib.request.urlopen", side_effect=fake_urlopen), patch("time.sleep", return_value=None):
                    c.ask("workbook_layout_sheet", "x", {}, LAYOUT_SCHEMA)
        self.assertEqual(attempts["n"], 2)
        joined = " ".join(messages)
        self.assertIn("재시도", joined)
        self.assertIn("429", joined)
        self.assertIn("시도", joined)
        for secret in ("sk-should-never-appear-anywhere", "Authorization", "Bearer"):
            self.assertNotIn(secret, joined)

    def test_connection_timeout_retries_are_bounded_not_four(self):
        attempts = {"n": 0}

        def fake_urlopen(request, timeout=None):
            attempts["n"] += 1
            raise TimeoutError("simulated timeout")

        with tempfile.TemporaryDirectory() as tmp:
            with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}):
                c = ResponsesClient(tmp, model="gpt-6-astra")
                with patch("excel_parser_ai.client.urllib.request.urlopen", side_effect=fake_urlopen), patch("time.sleep", return_value=None):
                    with self.assertRaises(ModelError):
                        c._request({"model": "gpt-6-astra"})
        self.assertEqual(attempts["n"], 2)  # bounded, not the old 4 — avoids repeating an already-slow request


class KeyboardInterruptHandlingTests(unittest.TestCase):
    def test_main_reports_clean_interrupt_message_without_traceback(self):
        import excel_parser_ai.cli as cli_module
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "input.xlsx"
            source.write_bytes(b"x")
            buffer = io.StringIO()
            with patch("excel_parser_ai.cli.REPO", root), \
                 patch("excel_parser_ai.cli.read_env", return_value=None), \
                 patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}), \
                 patch("excel_parser_ai.cli.run", side_effect=KeyboardInterrupt), \
                 patch("excel_parser_ai.cli.ResponsesClient"), \
                 contextlib.redirect_stderr(buffer):
                code = cli_module.main(["run", str(source)])
        self.assertEqual(code, 130)
        output = buffer.getvalue()
        self.assertIn("중단됨", output)
        self.assertIn("부분 결과 위치", output)
        self.assertNotIn("Traceback", output)


def wide_sheet(num_rows=20, num_columns=12):
    """A synthetic financial-statement-shaped sheet: many rows, each with
    many value columns, to exercise size-based (not just row-count-based)
    batching. Mirrors the shape that produced a 900KB+ real request."""
    columns = [chr(ord("C") + i) for i in range(num_columns)]
    cells = []
    for row in range(1, num_rows + 1):
        cells.append({"record_id": f"source/1/B{row}", "file_id": "hash", "sheet_id": "1", "sheet": "Wide",
                      "address": f"B{row}", "row": row, "column": "B", "column_index": 2,
                      "value": f"항목{row}", "data_type": "string", "formula": None,
                      "has_content": True, "style_id": 0, "number_format": None})
        for col in columns:
            address = f"{col}{row}"
            cells.append({"record_id": f"source/1/{address}", "file_id": "hash", "sheet_id": "1", "sheet": "Wide",
                          "address": address, "row": row, "column": col, "column_index": ord(col) - 64,
                          "value": str(row * 1000), "data_type": "number", "formula": None,
                          "has_content": True, "style_id": 0, "number_format": "#,##0;[Red](#,##0)"})
    sheet = {"name": "Wide", "cells": cells, "counts": {"content_cells": len(cells), "formulas": 0}}
    plan = {"sheet": "Wide", "kind": "financial_statement", "header_end_row": 0,
           "organization_code_column": None, "organization_name_column": None,
           "fallback_organization_code": None, "fallback_organization_name": None,
           "unit": None, "unit_evidence": [], "periods": [], "dimensions": [], "notes": []}
    return sheet, plan


class BatchSizingTests(unittest.TestCase):
    def test_batches_are_split_by_estimated_size_not_just_row_count(self):
        sheet, plan = wide_sheet(num_rows=20, num_columns=12)
        numbers = list(range(1, 21))
        # max_rows alone would allow all 20 rows in one batch; a byte cap
        # tight enough for a handful of 12-column rows must split it.
        batches = plan_row_batches(sheet, plan, numbers, max_rows=20, max_output_cells=10_000, max_context_bytes=3000)
        self.assertGreater(len(batches), 1)
        self.assertEqual(sorted(n for b in batches for n in b), numbers)  # no row lost or duplicated

    def test_batches_are_split_by_estimated_output_cell_count(self):
        sheet, plan = wide_sheet(num_rows=20, num_columns=12)
        numbers = list(range(1, 21))
        batches = plan_row_batches(sheet, plan, numbers, max_rows=20, max_output_cells=30, max_context_bytes=10**9)
        self.assertGreater(len(batches), 1)
        for batch in batches:
            self.assertLessEqual(len(batch) * 12, 30 + 12)  # at most ~one row over the cap (never an empty batch)

    def test_a_single_oversized_row_is_never_dropped(self):
        sheet, plan = wide_sheet(num_rows=1, num_columns=12)
        batches = plan_row_batches(sheet, plan, [1], max_rows=20, max_output_cells=1, max_context_bytes=1)
        self.assertEqual(batches, [[1]])  # still processed, just alone in its own batch

    def test_default_batch_size_is_much_smaller_than_the_old_64(self):
        import excel_parser_ai.context as context_module
        self.assertLess(context_module.DEFAULT_MAX_BATCH_ROWS, 64)

    def test_preceding_hierarchy_does_not_grow_unboundedly_across_batches(self):
        # Previously this resent EVERY earlier row as "preceding_hierarchy"
        # context for each later batch of the same sheet — a measured
        # contributor to a 900KB+ real request.
        sheet, plan = wide_sheet(num_rows=200, num_columns=1)
        ctx_early = rows_context({"sheets": [sheet]}, sheet, plan, [5, 6])
        ctx_late = rows_context({"sheets": [sheet]}, sheet, plan, [195, 196])
        self.assertLessEqual(len(ctx_late["preceding_hierarchy"]), 40)
        self.assertLessEqual(len(ctx_early["preceding_hierarchy"]), len(ctx_late["preceding_hierarchy"]) + 1)


class PartialBatchFailureTests(unittest.TestCase):
    def test_run_pool_partial_keeps_successes_and_stops_pending_work_after_a_failure(self):
        # A near-instant failing job racing the worker thread picking up the
        # very next queued item is an inherent, harmless timing window (real
        # network calls take far longer than that race ever could) — so this
        # asserts the documented guarantee (pending work is meaningfully
        # curtailed, successes are kept), not an exact per-item outcome.
        started = []
        def job(x):
            started.append(x)
            if x == 2:
                raise RuntimeError("simulated batch failure")
            time.sleep(0.2)
            return x
        results, failures = _run_pool_partial(job, list(range(20)), concurrency=2)
        self.assertIn(2, [item for item, _ in failures])
        self.assertTrue(results)  # successes before/around the failure were kept, not discarded
        self.assertLess(len(started), 20)  # most pending work was never started, not just left to finish

    def test_run_pool_partial_returns_no_failures_on_full_success(self):
        results, failures = _run_pool_partial(lambda x: x * 2, [1, 2, 3], concurrency=2)
        self.assertEqual(sorted(results), [2, 4, 6])
        self.assertEqual(failures, [])

    def test_infer_workbook_keeps_successful_row_batches_when_a_later_one_fails(self):
        w, _ = fixture(value="1")
        original = w["sheets"][0]
        # Two independent sheets so each becomes its own row batch/job.
        w["sheets"].append({"name": "Sheet2", "cells": [dict(c, sheet="Sheet2") for c in original["cells"]],
                            "counts": dict(original["counts"])})

        def make_plan(sheet_name):
            return {"sheet": sheet_name, "kind": "unknown", "header_end_row": 0,
                   "organization_code_column": None, "organization_name_column": None,
                   "fallback_organization_code": None, "fallback_organization_name": None,
                   "unit": None, "unit_evidence": [], "periods": [], "dimensions": [], "notes": []}

        class FlakyClient:
            model, effort = "fake", "high"
            def ask(self, task, instructions, context, schema):
                if task.startswith("workbook_layout"):
                    return {"sheet_plan": make_plan(context["target_sheet"]["sheet"]), "uncertainties": []}
                sheet_name = context["sheet"]
                if sheet_name == "Sheet2":
                    raise ModelError("simulated row batch failure for Sheet2")
                rows_out = []
                for row in context["rows"]:
                    cells_out = [{"column": c["cell"][0], "status": "not_numeric", "value_kind": "text",
                                 "source_value_read": str(c["value"]) if c.get("value") is not None else None,
                                 "normalized_value": None, "unit": None, "unit_evidence": [], "reason": "fake"}
                                for c in row["cells"]]
                    rows_out.append({"row": row["row"], "role": "reference", "item_name": None, "item_code": None,
                                     "header_path": [], "metric_code": None, "cells": cells_out, "reason": "fake"})
                return {"rows": rows_out, "uncertainties": []}

        semantic = infer_workbook(w, FlakyClient(), review=False, batch_size=64, concurrency=1)
        self.assertEqual(semantic["row_batches"]["failed"], 1)
        self.assertGreaterEqual(semantic["row_batches"]["completed"], 1)
        self.assertTrue(semantic["rows"]["Sheet"])  # Sheet's successful batch was kept
        self.assertEqual(semantic["rows"]["Sheet2"], {})  # Sheet2's failed batch left unclassified, not invented

    def test_pipeline_preserves_successful_records_when_one_file_partially_fails(self):
        w, s = fixture(value="-100", normalized="-100000")
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "input.xlsx"
            source.write_bytes(b"read-only test input")

            class PartialFailClient:
                model, effort = "fake", "high"
                def __init__(self):
                    self.events = []
                def ask(self, task, instructions, context, schema):
                    if task.startswith("workbook_layout"):
                        return {"sheet_plan": s["layout"]["sheets"][0], "uncertainties": []}
                    raise ModelError("simulated row batch failure")

            client = PartialFailClient()
            with patch("excel_parser_ai.pipeline.REPO", root), \
                 patch("excel_parser_ai.pipeline.extract_workbook", return_value=w), \
                 patch("excel_parser_ai.pipeline.validate_extraction", return_value={"passed": True}):
                summary = run([source], root / "parser_ai/output/run", client=client)
            self.assertFalse(summary["llm_inference_completed"])  # partial, must not claim completion
            self.assertTrue(any(f["error_type"] == "PartialRowBatchFailure" for f in summary["failures"]))
            with gzip.open(root / "parser_ai/output/run/standard-records.jsonl.gz", "rt") as f:
                records = [json.loads(line) for line in f]
            self.assertTrue(records)  # raw/standard candidates still preserved, not thrown away


class ConnectionErrorHandlingTests(unittest.TestCase):
    """RemoteDisconnected (seen in a real failed run) is a ConnectionResetError
    subclass raised by http.client.getresponse(), which urlopen() does NOT
    wrap into URLError — it previously went completely uncaught here."""

    def test_remote_disconnected_isinstance_of_connection_error(self):
        # RemoteDisconnected has multiple inheritance: it is simultaneously a
        # ConnectionResetError (-> ConnectionError -> OSError) AND a
        # http.client.BadStatusLine (-> HTTPException). Both are true; this
        # pins down the specific fact the `except (..., ConnectionError, ...)`
        # clause in client.py actually relies on, rather than asserting it
        # only indirectly through the behavior of a mocked request.
        import http.client
        inst = http.client.RemoteDisconnected("Remote end closed connection without response")
        self.assertIsInstance(inst, ConnectionError)
        self.assertIsInstance(inst, ConnectionResetError)
        self.assertIsInstance(inst, http.client.BadStatusLine)  # both parents hold at once
        self.assertNotIsInstance(inst, urllib.error.URLError)  # confirms urlopen() does NOT wrap this one

    def test_remote_disconnected_is_caught_with_a_distinct_message_and_bounded_retries(self):
        import http.client
        attempts = {"n": 0}
        def fake_urlopen(request, timeout=None):
            attempts["n"] += 1
            raise http.client.RemoteDisconnected("Remote end closed connection without response")
        messages = []
        with tempfile.TemporaryDirectory() as tmp:
            with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}):
                c = ResponsesClient(tmp, model="gpt-6-astra", progress=messages.append)
                with patch("excel_parser_ai.client.urllib.request.urlopen", side_effect=fake_urlopen), patch("time.sleep", return_value=None):
                    with self.assertRaises(ModelError) as ctx:
                        c._request({"model": "gpt-6-astra"})
        self.assertEqual(attempts["n"], 2)  # bounded, not retried forever
        self.assertIn("RemoteDisconnected", str(ctx.exception))
        self.assertIn("끊김", " ".join(messages) + str(ctx.exception))

    def test_incomplete_read_during_body_parsing_is_caught(self):
        import http.client
        attempts = {"n": 0}
        def fake_urlopen(request, timeout=None):
            attempts["n"] += 1
            raise http.client.IncompleteRead(b"partial")
        with tempfile.TemporaryDirectory() as tmp:
            with patch.dict("os.environ", {"OPENAI_API_KEY": "sk-test"}):
                c = ResponsesClient(tmp, model="gpt-6-astra")
                with patch("excel_parser_ai.client.urllib.request.urlopen", side_effect=fake_urlopen), patch("time.sleep", return_value=None):
                    with self.assertRaises(ModelError):
                        c._request({"model": "gpt-6-astra"})
        self.assertEqual(attempts["n"], 2)

    def test_failed_call_is_recorded_as_an_event_with_no_usage_data(self):
        with tempfile.TemporaryDirectory() as tmp:
            def transport(body):
                raise ModelError("simulated failure")
            c = ResponsesClient(tmp, model="m", transport=transport)
            with self.assertRaises(ModelError):
                c.ask("workbook_layout_sheet", "x", {}, LAYOUT_SCHEMA)
        self.assertEqual(len(c.events), 1)
        self.assertTrue(c.events[0]["failed"])
        self.assertEqual(c.events[0]["usage"], {})
        self.assertEqual(c.events[0]["error_type"], "ModelError")


def _generic_grounded_transport(body):
    """A real OpenAI Responses transport stand-in: always returns a
    schema-valid, source-grounded answer for whichever task/context it is
    given, so a REAL ResponsesClient's caching behaves authentically (this
    is not a client.ask() override — it is the HTTP-call replacement, so
    the normal cache-check/digest/call_count logic in ask() still runs for
    real)."""
    task = body["text"]["format"]["name"]
    context = json.loads(body["input"][1]["content"])
    if "source" in context:  # repair/review wrapper
        context = context["source"]
    if task.startswith("workbook_layout"):
        sheet_name = context["target_sheet"]["sheet"]
        data = {"sheet_plan": {"sheet": sheet_name, "kind": "unknown", "header_end_row": 0,
                               "organization_code_column": None, "organization_name_column": None,
                               "fallback_organization_code": None, "fallback_organization_name": None,
                               "unit": None, "unit_evidence": [], "periods": [], "dimensions": [], "notes": []},
                "uncertainties": []}
    else:
        rows_out = []
        for row in context["rows"]:
            cells_out = [{"column": c["cell"][0], "status": "not_numeric", "value_kind": "text",
                         "source_value_read": str(c["value"]) if c.get("value") is not None else None,
                         "normalized_value": None, "unit": None, "unit_evidence": [], "reason": "fake"}
                        for c in row["cells"]]
            rows_out.append({"row": row["row"], "role": "reference", "item_name": None, "item_code": None,
                             "header_path": [], "metric_code": None, "cells": cells_out, "reason": "fake"})
        data = {"rows": rows_out, "uncertainties": []}
    return {"status": "completed", "output": [{"content": [{"type": "output_text", "text": json.dumps(data)}]}]}


class MaxNewRowBatchesTests(unittest.TestCase):
    """Covers the exact scenario asked for: reuse cached layout/row batches,
    process only a small, exact number of genuinely NEW row batches, then
    stop cleanly (not as a failure) — verified against a REAL ResponsesClient
    so the cache-hit/call_count accounting is authentic, not simulated."""

    def test_reuses_cache_fully_processes_exactly_n_new_batches_then_stops_cleanly(self):
        sheet_a, _ = wide_sheet(num_rows=4, num_columns=1)
        sheet_a["name"] = "A"
        for c in sheet_a["cells"]:
            c["sheet"] = "A"
        workbook_a = {"file": {"file_id": "h", "source_id": "s", "path": "/x.xlsx", "name": "x.xlsx"},
                     "sheets": [sheet_a], "diagnostics": []}

        with tempfile.TemporaryDirectory() as tmp:
            cache_dir = Path(tmp) / "cache"

            # Run 1: full, unrestricted — populates the cache for sheet A's
            # layout and both of its row batches (batch_size=2 -> 2 batches).
            client1 = ResponsesClient(cache_dir, model="m", transport=_generic_grounded_transport)
            infer_workbook(workbook_a, client1, review=False, batch_size=2, concurrency=2)
            calls_in_run1 = client1.call_count
            self.assertGreater(calls_in_run1, 0)

            # Run 2: sheet A unchanged (must hit cache) plus two BRAND NEW
            # sheets B and C (never seen before), each exactly one batch.
            # With max_new_row_batches=1: A's 2 batches must be free cache
            # hits, B's 1 batch is the single allowed new one, and C's batch
            # must be left unattempted — stopped, not failed.
            def make_sheet(name, rows):
                sheet, _ = wide_sheet(num_rows=rows, num_columns=1)
                sheet["name"] = name
                for c in sheet["cells"]:
                    c["sheet"] = name
                return sheet

            sheet_b, sheet_c = make_sheet("B", 2), make_sheet("C", 2)
            workbook_2 = {"file": workbook_a["file"], "sheets": [sheet_a, sheet_b, sheet_c], "diagnostics": []}
            client2 = ResponsesClient(cache_dir, model="m", transport=_generic_grounded_transport)
            semantic = infer_workbook(workbook_2, client2, review=False, batch_size=2, concurrency=2,
                                      max_new_row_batches=1)

        row_batches = semantic["row_batches"]
        self.assertTrue(row_batches["stopped_by_new_batch_limit"])
        self.assertEqual(row_batches["failed"], 0)  # a deliberate stop, not a failure
        self.assertTrue(semantic["rows"]["A"])   # fully present, served from cache
        self.assertTrue(semantic["rows"]["B"] or semantic["rows"]["C"])  # exactly one new sheet got its batch
        self.assertFalse(semantic["rows"]["B"] and semantic["rows"]["C"])  # ...but not both

    def test_pipeline_marks_a_deliberate_stop_as_incomplete_but_not_a_failure(self):
        w, s = fixture(value="1")
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "input.xlsx"
            source.write_bytes(b"x")
            client = ResponsesClient(root / "cache", model="m", transport=_generic_grounded_transport)
            with patch("excel_parser_ai.pipeline.REPO", root), \
                 patch("excel_parser_ai.pipeline.extract_workbook", return_value=w), \
                 patch("excel_parser_ai.pipeline.validate_extraction", return_value={"passed": True}):
                summary = run([source], root / "parser_ai/output/run", client=client,
                              review=False, max_new_row_batches=0)
        self.assertFalse(summary["llm_inference_completed"])
        self.assertFalse(any(f.get("error_type") == "PartialRowBatchFailure" for f in summary["failures"]))
        self.assertEqual(summary["failures"], [])  # a deliberate test stop is not logged as a failure


if __name__ == "__main__":
    unittest.main()
