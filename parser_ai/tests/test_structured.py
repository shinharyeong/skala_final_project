import copy
import tempfile
import unittest
from pathlib import Path

from excel_parser_ai.client import ModelError
from excel_parser_ai.structured import apply_plans, infer_structure, structure_fingerprint, validate_plan


def cell(address, value, data_type="string", *, formula=None, style=1):
    column = "".join(x for x in address if x.isalpha())
    row = int("".join(x for x in address if x.isdigit()))
    number = 0
    for x in column:
        number = number * 26 + ord(x) - 64
    return {"record_id": f"source/1/{address}", "sheet": "Sheet", "sheet_id": "1", "address": address,
            "column": column, "column_index": number, "row": row, "value": value, "data_type": data_type,
            "cached_value": value if formula and formula.get("cached_status") == "stored_unverified" else None,
            "formula": formula, "has_content": True, "style_id": style,
            "number_format": "#,##0;(#,##0)", "row_hidden": False, "column_hidden": False, "merge": None}


def fixture(*, organization="사업부1", code="F10", year=2025):
    cells = [cell("B1", f"{year}년 1월"), cell("C1", f"{year}년 2월"),
             cell("B2", "실적"), cell("C2", "추정"), cell("D1", code), cell("E1", organization),
             cell("F1", "단위: 천원"), cell("A3", "매출"), cell("A4", "비용"),
             cell("B3", "1", "number"), cell("C3", "0", "number"), cell("B4", "-2", "number"),
             cell("C4", None, "number", formula={"text": "=1+2", "cached_status": "missing"})]
    return {"file": {"file_id": "file", "source_id": "source", "path": "/mock.xlsx", "name": "mock.xlsx"},
            "sheets": [{"name": "Sheet", "dimension": "A1:F4", "merges": [], "cells": cells, "state": "visible"}],
            "diagnostics": []}


def plan():
    def period(column, month, scenario):
        return {"start_col": column, "end_col": column, "year": 2025,
                "year_ref": {"sheet": "Sheet", "cell": f"{column}1"}, "aggregation": "month",
                "month_header_row": 1, "start_month": month, "scenario": scenario,
                "scenario_ref": {"sheet": "Sheet", "cell": f"{column}2"},
                "evidence": [{"sheet": "Sheet", "cell": f"{column}1"}, {"sheet": "Sheet", "cell": f"{column}2"}]}
    return {"sheet": "Sheet", "kind": "financial_statement", "tables": [{
        "start_row": 3, "end_row": 4, "start_col": "B", "end_col": "C",
        "header_start_row": 1, "header_end_row": 2,
        "organization_code_col": None, "organization_name_col": None,
        "organization_code_ref": {"sheet": "Sheet", "cell": "D1"},
        "organization_name_ref": {"sheet": "Sheet", "cell": "E1"},
        "periods": [period("B", 1, "actual"), period("C", 2, "estimate")],
        "rows": [{"start_row": 3, "end_row": 4, "role": "detail", "item_col": "A",
                  "item_code_col": None, "hierarchy_cols": ["A"], "ancestor_rows": [],
                  "metric_code": None, "reason": "labels"}],
        "unit": "천원", "normalization_mode": "amount_to_won",
        "unit_evidence": [{"sheet": "Sheet", "cell": "F1"}],
        "unit_overrides": [], "exclusions": [], "exceptions": []}],
        "excluded_areas": [], "uncertainties": []}


class StructureApplicationTests(unittest.TestCase):
    def test_full_coverage_and_decimal_zero_negative_formula_cache(self):
        workbook, declaration = fixture(), plan()
        self.assertEqual(validate_plan(declaration, workbook, workbook["sheets"][0]), [])
        records = {r["source"]["cell"]: r for r in apply_plans(workbook, [declaration])["records"]}
        self.assertEqual(len(records), 4)
        self.assertEqual(str(records["B3"]["value"]["normalized_value"]), "1000")
        self.assertEqual(str(records["C3"]["value"]["normalized_value"]), "0")
        self.assertEqual(str(records["B4"]["value"]["normalized_value"]), "-2000")
        self.assertEqual(records["C4"]["value"]["status"], "unresolved")
        self.assertIsNone(records["C4"]["value"]["normalized_value"])
        self.assertIn("PYTHON_RULE_APPLICATION", records["B3"]["mapping"]["rule_ids"])

    def test_missing_duplicate_and_nonexistent_evidence_rejected(self):
        workbook, declaration = fixture(), plan()
        declaration["tables"][0]["periods"].pop()
        self.assertTrue(any("period coverage" in x for x in validate_plan(declaration, workbook, workbook["sheets"][0])))
        declaration = plan()
        declaration["tables"][0]["periods"].append(copy.deepcopy(declaration["tables"][0]["periods"][0]))
        self.assertTrue(any("period coverage" in x for x in validate_plan(declaration, workbook, workbook["sheets"][0])))
        declaration = plan()
        declaration["tables"][0]["unit_evidence"][0]["cell"] = "Z99"
        self.assertTrue(any("missing evidence" in x for x in validate_plan(declaration, workbook, workbook["sheets"][0])))

    def test_reuse_rebinds_organization_year_and_values(self):
        first, second, declaration = fixture(), fixture(organization="사업부2", code="F12", year=2027), plan()
        second["sheets"][0]["cells"][-4]["value"] = "9"  # B3, current file only
        next(c for c in second["sheets"][0]["cells"] if c["address"] == "C2")["value"] = "계획"
        self.assertEqual(structure_fingerprint(first, declaration), structure_fingerprint(second, declaration))
        records = {r["source"]["cell"]: r for r in apply_plans(second, [declaration], reused_sheets=["Sheet"],
            metadata={"submission_version": "V2.01", "submission_month": "2027-02"})["records"]}
        self.assertEqual(records["B3"]["organization"]["code"], "F12")
        self.assertEqual(records["B3"]["organization"]["name"], "사업부2")
        self.assertEqual(records["B3"]["period"]["year"], 2027)
        self.assertEqual(str(records["B3"]["value"]["normalized_value"]), "9000")
        self.assertEqual(records["C3"]["period"]["scenario"], "plan")
        self.assertEqual(records["B3"]["submission"]["version_raw"], "V2.01")

    def test_column_move_changes_fingerprint_and_is_uncovered(self):
        first, second, declaration = fixture(), fixture(), plan()
        moved = next(c for c in second["sheets"][0]["cells"] if c["address"] == "B3")
        moved.update(address="G3", column="G", column_index=7, record_id="source/1/G3")
        self.assertNotEqual(structure_fingerprint(first, declaration), structure_fingerprint(second, declaration))
        self.assertTrue(any("uncovered" in x for x in validate_plan(declaration, second, second["sheets"][0])))
        styled = fixture()
        next(c for c in styled["sheets"][0]["cells"] if c["address"] == "B3")["style_id"] = 99
        self.assertNotEqual(structure_fingerprint(first, declaration), structure_fingerprint(styled, declaration))

    def test_repeated_header_and_mid_table_unit_change_remain_unresolved(self):
        for extra in ([cell("D4", "1월"), cell("E4", "2월"), cell("F4", "3월")],
                      [cell("D4", "단위: 원")]):
            workbook, declaration = fixture(), plan()
            workbook["sheets"][0]["cells"].extend(extra)
            records = {r["source"]["cell"]: r for r in apply_plans(workbook, [declaration])["records"]}
            self.assertEqual(records["B4"]["mapping"]["status"], "unmapped")
            self.assertEqual(records["B3"]["mapping"]["status"], "mapped")

    def test_declared_mid_table_unit_override_is_applied(self):
        workbook, declaration = fixture(), plan()
        workbook["sheets"][0]["cells"].append(cell("D4", "단위: 원"))
        declaration["tables"][0]["unit_overrides"] = [{"start_row": 4, "end_row": 4,
            "start_col": "B", "end_col": "C", "unit": "원", "normalization_mode": "amount_to_won",
            "evidence": [{"sheet": "Sheet", "cell": "D4"}]}]
        self.assertEqual(validate_plan(declaration, workbook, workbook["sheets"][0]), [])
        records = {r["source"]["cell"]: r for r in apply_plans(workbook, [declaration])["records"]}
        self.assertEqual(str(records["B4"]["value"]["normalized_value"]), "-2")
        self.assertEqual(records["B4"]["mapping"]["status"], "mapped")

    def test_zero_is_distinct_from_blank_and_dash(self):
        declaration = plan()
        zero = fixture()
        blank = fixture()
        dash = fixture()
        next(c for c in blank["sheets"][0]["cells"] if c["address"] == "C3")["value"] = None
        dash_cell = next(c for c in dash["sheets"][0]["cells"] if c["address"] == "C3")
        dash_cell.update(value="-", data_type="string")
        values = []
        for workbook in (zero, blank, dash):
            record = next(r for r in apply_plans(workbook, [declaration])["records"] if r["source"]["cell"] == "C3")
            values.append((record["value"]["status"], record["value"]["normalized_value"]))
        self.assertEqual(values[0][0], "confirmed")
        self.assertEqual(str(values[0][1]), "0")
        self.assertEqual(values[1][0], "missing")
        self.assertEqual(values[2][0], "unresolved")
        self.assertIsNone(values[2][1])

    def test_ambiguous_sign_text_is_kept_unresolved(self):
        workbook, declaration = fixture(), plan()
        ambiguous = next(c for c in workbook["sheets"][0]["cells"] if c["address"] == "C3")
        ambiguous.update(value="△12", data_type="string")
        record = next(r for r in apply_plans(workbook, [declaration])["records"] if r["source"]["cell"] == "C3")
        self.assertEqual(record["raw_value"], "△12")
        self.assertEqual(record["value"]["status"], "unresolved")
        self.assertIsNone(record["value"]["normalized_value"])

    def test_declared_ancestor_row_builds_current_file_hierarchy(self):
        workbook, declaration = fixture(), plan()
        workbook["sheets"][0]["cells"].append(cell("A2", "영업손익"))
        declaration["tables"][0]["rows"][0]["ancestor_rows"] = [2]
        record = next(r for r in apply_plans(workbook, [declaration])["records"] if r["source"]["cell"] == "B3")
        self.assertEqual(record["item"]["header_path"], ["영업손익", "매출"])

    def test_reference_row_never_becomes_amount(self):
        workbook, declaration = fixture(), plan()
        declaration["tables"][0]["rows"][0]["role"] = "reference"
        record = next(r for r in apply_plans(workbook, [declaration])["records"] if r["source"]["cell"] == "B3")
        self.assertEqual(record["value"]["status"], "not_numeric")
        self.assertIsNone(record["value"]["normalized_value"])

    def test_ratio_mode_keeps_stored_fraction(self):
        workbook, declaration = fixture(), plan()
        next(c for c in workbook["sheets"][0]["cells"] if c["address"] == "F1")["value"] = "단위: %"
        next(c for c in workbook["sheets"][0]["cells"] if c["address"] == "B3")["value"] = "0.125"
        table = declaration["tables"][0]
        table["unit"] = "%"
        table["normalization_mode"] = "ratio_fraction"
        table["rows"][0]["role"] = "ratio"
        record = next(r for r in apply_plans(workbook, [declaration])["records"] if r["source"]["cell"] == "B3")
        self.assertEqual(str(record["value"]["normalized_value"]), "0.125")

    def test_scenario_conflict_and_exception_failure_are_unresolved(self):
        workbook, declaration = fixture(), plan()
        next(c for c in workbook["sheets"][0]["cells"] if c["address"] == "C2")["value"] = "실적/추정"
        records = {r["source"]["cell"]: r for r in apply_plans(workbook, [declaration])["records"]}
        self.assertEqual(records["C3"]["period"]["scenario"], "unknown")
        self.assertEqual(records["C3"]["mapping"]["status"], "partial")
        declaration["tables"][0]["exceptions"] = [{"start_row": 3, "end_row": 3, "start_col": "B", "end_col": "B",
                                                    "reason": "ambiguous", "evidence": [{"sheet": "Sheet", "cell": "B3"}]}]
        records = {r["source"]["cell"]: r for r in apply_plans(workbook, [declaration])["records"]}
        self.assertEqual(records["B3"]["mapping"]["status"], "unmapped")
        records = apply_plans(workbook, [])["records"]
        self.assertTrue(all(r["mapping"]["status"] == "unmapped" for r in records))

    def test_template_cache_reuses_only_matching_structure(self):
        class FakeClient:
            def __init__(self, cache):
                self.cache_dir, self.events, self.calls = cache, [], 0
            def ask(self, task, prompt, context, schema, **kwargs):
                self.calls += 1
                self.events.append({"task": task, "cached": False})
                return {"plan": plan()}
        with tempfile.TemporaryDirectory() as tmp:
            client = FakeClient(Path(tmp))
            infer_structure(fixture(), client)
            infer_structure(fixture(organization="사업부2", code="F12", year=2027), client)
            self.assertEqual(client.calls, 1)
            changed = fixture()
            changed["sheets"][0]["cells"].append(cell("F4", "새 머리글"))
            infer_structure(changed, client)
            self.assertEqual(client.calls, 2)

    def test_grouped_exception_call_is_bounded_and_traceable(self):
        declaration = plan()
        declaration["tables"][0]["exceptions"] = [{"start_row": 3, "end_row": 3, "start_col": "B", "end_col": "B",
            "reason": "ambiguous", "evidence": [{"sheet": "Sheet", "cell": "B3"}]}]
        class FakeClient:
            def __init__(self, path):
                self.cache_dir, self.events, self.calls = path, [], []
            def ask(self, task, prompt, context, schema, **kwargs):
                self.calls.append(task)
                self.events.append({"task": task, "cached": False})
                if task == "sheet_structure_v1":
                    return {"plan": declaration}
                return {"resolutions": [{"sheet": "Sheet", "start_row": 3, "end_row": 3,
                    "start_col": "B", "end_col": "B", "unit": "천원", "normalization_mode": "amount_to_won",
                    "unit_evidence": [{"sheet": "Sheet", "cell": "F1"}], "scenario": "actual",
                    "scenario_evidence": [{"sheet": "Sheet", "cell": "B2"}], "role": "detail",
                    "reason": "source headers"}], "unresolved": []}
        with tempfile.TemporaryDirectory() as tmp:
            client = FakeClient(Path(tmp))
            mapped, semantic = infer_structure(fixture(), client, max_exception_calls=1)
            self.assertEqual(client.calls, ["sheet_structure_v1", "structure_exceptions_v1"])
            self.assertEqual(semantic["exception_calls"], 1)
            b3 = next(r for r in mapped["records"] if r["source"]["cell"] == "B3")
            self.assertIn("LLM_EXCEPTION_REGION", b3["mapping"]["rule_ids"])

    def test_structure_api_failure_preserves_candidates_without_retry_loop(self):
        class FailingClient:
            def __init__(self, path):
                self.cache_dir, self.events, self.calls = path, [], 0
            def ask(self, task, prompt, context, schema, **kwargs):
                self.calls += 1
                raise RuntimeError("mock failure")
        with tempfile.TemporaryDirectory() as tmp:
            client = FailingClient(Path(tmp))
            mapped, semantic = infer_structure(fixture(), client)
            self.assertEqual(client.calls, 1)
            self.assertEqual(len(semantic["failures"]), 1)
            self.assertTrue(all(r["mapping"]["status"] == "unmapped" for r in mapped["records"]))

    def test_sheet_without_numeric_candidate_needs_no_api_call(self):
        workbook = fixture()
        workbook["sheets"].append({"name": "Notes", "dimension": "A1:A1", "merges": [],
                                   "cells": [{**cell("A1", "설명"), "sheet": "Notes", "record_id": "source/2/A1"}],
                                   "state": "visible"})
        class FakeClient:
            def __init__(self, path):
                self.cache_dir, self.events, self.calls = path, [], 0
            def ask(self, task, prompt, context, schema, **kwargs):
                self.calls += 1
                self.events.append({"task": task, "cached": False})
                return {"plan": plan()}
        with tempfile.TemporaryDirectory() as tmp:
            client = FakeClient(Path(tmp))
            _, semantic = infer_structure(workbook, client)
            self.assertEqual(client.calls, 1)
            self.assertEqual(len(semantic["structure_plans"]), 2)


def sheet_cells(name, cells):
    return [dict(c, sheet=name) for c in cells]


def big_sheet_fixture(num_rows):
    """A tall detail table: header rows 1-2, then `num_rows` data rows each
    with one numeric cell in column B. Used to exercise row-range-block
    splitting (both proactive, for large sheets, and reactive, after a
    simulated timeout on a smaller one)."""
    cells = [cell("B1", "2025년 1월"), cell("B2", "실적")]
    for row in range(3, 3 + num_rows):
        cells.append(cell(f"A{row}", f"항목{row}"))
        cells.append(cell(f"B{row}", str(row), "number"))
    return {"file": {"file_id": "file", "source_id": "source", "path": "/mock.xlsx", "name": "mock.xlsx"},
            "sheets": [{"name": "Sheet", "dimension": f"A1:B{2 + num_rows}", "merges": [], "cells": cells, "state": "visible"}],
            "diagnostics": []}


def block_plan(lo, hi):
    """A minimal, valid single-table plan declaring exactly rows [lo, hi] of
    column B in big_sheet_fixture — what a real model would return for one
    row-range block request."""
    ref = {"sheet": "Sheet", "cell": "B1"}
    return {"sheet": "Sheet", "kind": "financial_statement",
            "tables": [{"start_row": lo, "end_row": hi, "start_col": "B", "end_col": "B",
                       "header_start_row": 1, "header_end_row": 2,
                       "organization_code_col": None, "organization_name_col": None,
                       "organization_code_ref": ref, "organization_name_ref": ref,
                       "periods": [{"start_col": "B", "end_col": "B", "year": 2025, "year_ref": ref,
                                   "aggregation": "month", "month_header_row": 1, "start_month": 1,
                                   "scenario": "actual", "scenario_ref": {"sheet": "Sheet", "cell": "B2"},
                                   "evidence": [ref, {"sheet": "Sheet", "cell": "B2"}]}],
                       "rows": [{"start_row": lo, "end_row": hi, "role": "detail", "item_col": "A",
                                "item_code_col": None, "hierarchy_cols": [], "ancestor_rows": [],
                                "metric_code": None, "reason": "labels"}],
                       "unit": "원", "normalization_mode": "amount_to_won", "unit_evidence": [ref],
                       "unit_overrides": [], "exclusions": [], "exceptions": []}],
            "excluded_areas": [], "uncertainties": []}


class LargeSheetBlockSplittingTests(unittest.TestCase):
    """Covers requirement 2: a sheet with more data rows than
    LARGE_SHEET_ROW_THRESHOLD is split into row-range blocks instead of one
    whole-sheet call."""

    def test_large_sheet_is_split_into_multiple_block_requests(self):
        workbook = big_sheet_fixture(num_rows=160)  # > LARGE_SHEET_ROW_THRESHOLD (150)

        class BlockAwareClient:
            def __init__(self, cache_dir):
                self.cache_dir, self.events, self.calls = cache_dir, [], []
            def ask(self, task, prompt, context, schema, **kwargs):
                self.calls.append(task)
                self.events.append({"task": task, "cached": False})
                lo, hi = context["declared_row_range"]
                return {"plan": block_plan(lo, hi)}

        with tempfile.TemporaryDirectory() as tmp:
            client = BlockAwareClient(Path(tmp))
            mapped, semantic = infer_structure(workbook, client)

        self.assertNotIn("sheet_structure_v1", client.calls)  # never sent as one whole-sheet call
        block_calls = [t for t in client.calls if t.startswith("sheet_structure_v1_block")]
        self.assertGreaterEqual(len(block_calls), 2)  # 160 rows / 100-row blocks
        self.assertEqual(sum(1 for r in mapped["records"] if r["mapping"]["status"] == "mapped"), 160)
        self.assertEqual(mapped["diagnostics"], [])  # merged blocks give full, gapless coverage

    def test_small_sheet_is_not_split(self):
        workbook = big_sheet_fixture(num_rows=5)  # well under the threshold

        class WholeSheetClient:
            def __init__(self, cache_dir):
                self.cache_dir, self.events, self.calls = cache_dir, [], []
            def ask(self, task, prompt, context, schema, **kwargs):
                self.calls.append(task)
                self.events.append({"task": task, "cached": False})
                return {"plan": block_plan(3, 7)}

        with tempfile.TemporaryDirectory() as tmp:
            client = WholeSheetClient(Path(tmp))
            infer_structure(workbook, client)
        self.assertEqual(client.calls, ["sheet_structure_v1"])


class FormulaHeavySheetSkipTests(unittest.TestCase):
    """Covers requirement 1: a sheet whose formula ratio is at/above
    FORMULA_SKIP_RATIO never gets an LLM call, and its SUM formulas are
    cross-checked against the raw values they actually reference instead."""

    def _workbook(self, *, matching=True):
        detail = sheet_cells("Detail", [cell("C10", "100", "number"), cell("C11", "200", "number")])
        match_formula = {"text": "=SUM(Detail!C10:C11)", "cached_status": "stored_unverified"}
        bad_formula = {"text": "=SUM(Detail!C10:C11)", "cached_status": "stored_unverified"}
        summary = sheet_cells("Summary", [
            cell("D5", "300" if matching else "300", "number", formula=match_formula),
            cell("D6", "999", "number", formula=bad_formula),
            cell("E1", "레이블"),
        ])
        return {"file": {"file_id": "f", "source_id": "s", "path": "/m.xlsx", "name": "m.xlsx"},
               "sheets": [{"name": "Detail", "dimension": "C10:C11", "merges": [], "cells": detail, "state": "visible"},
                          {"name": "Summary", "dimension": "D5:E1", "merges": [], "cells": summary, "state": "visible"}],
               "diagnostics": []}

    def test_formula_heavy_sheet_gets_no_llm_call_and_is_recorded_as_skipped(self):
        workbook = self._workbook()
        ref = {"sheet": "Detail", "cell": "C10"}

        class DetailOnlyClient:
            def __init__(self, cache_dir):
                self.cache_dir, self.events, self.calls = cache_dir, [], []
            def ask(self, task, prompt, context, schema, **kwargs):
                self.calls.append(task)
                self.events.append({"task": task, "cached": False})
                return {"plan": {"sheet": context["sheet"], "kind": "revenue_detail",
                                 "tables": [{"start_row": 10, "end_row": 11, "start_col": "C", "end_col": "C",
                                            "header_start_row": 1, "header_end_row": 1,
                                            "organization_code_col": None, "organization_name_col": None,
                                            "organization_code_ref": ref, "organization_name_ref": ref,
                                            "periods": [{"start_col": "C", "end_col": "C", "year": 2025, "year_ref": ref,
                                                        "aggregation": "annual", "month_header_row": None, "start_month": None,
                                                        "scenario": "actual", "scenario_ref": ref, "evidence": [ref]}],
                                            "rows": [{"start_row": 10, "end_row": 11, "role": "detail", "item_col": None,
                                                     "item_code_col": None, "hierarchy_cols": [], "ancestor_rows": [],
                                                     "metric_code": None, "reason": "x"}],
                                            "unit": "원", "normalization_mode": "amount_to_won", "unit_evidence": [ref],
                                            "unit_overrides": [], "exclusions": [], "exceptions": []}],
                                 "excluded_areas": [], "uncertainties": []}}

        with tempfile.TemporaryDirectory() as tmp:
            client = DetailOnlyClient(Path(tmp))
            mapped, semantic = infer_structure(workbook, client)

        # Only Detail (formula_ratio 0.0) was ever asked; Summary (ratio
        # 2/3 >= 0.3) got no LLM call at all.
        self.assertEqual(client.calls, ["sheet_structure_v1"])
        self.assertEqual(len(semantic["skipped_sheets"]), 1)
        self.assertEqual(semantic["skipped_sheets"][0]["sheet"], "Summary")
        self.assertGreaterEqual(semantic["skipped_sheets"][0]["formula_ratio"], 0.3)

    def test_matching_and_mismatching_sum_formulas_are_crosschecked_without_changing_records(self):
        workbook = self._workbook()

        class DetailOnlyClient:
            def __init__(self, cache_dir):
                self.cache_dir, self.events = cache_dir, []
            def ask(self, task, prompt, context, schema, **kwargs):
                self.events.append({"task": task, "cached": False})
                ref = {"sheet": "Detail", "cell": "C10"}
                return {"plan": {"sheet": context["sheet"], "kind": "revenue_detail",
                                 "tables": [{"start_row": 10, "end_row": 11, "start_col": "C", "end_col": "C",
                                            "header_start_row": 1, "header_end_row": 1,
                                            "organization_code_col": None, "organization_name_col": None,
                                            "organization_code_ref": ref, "organization_name_ref": ref,
                                            "periods": [{"start_col": "C", "end_col": "C", "year": 2025, "year_ref": ref,
                                                        "aggregation": "annual", "month_header_row": None, "start_month": None,
                                                        "scenario": "actual", "scenario_ref": ref, "evidence": [ref]}],
                                            "rows": [{"start_row": 10, "end_row": 11, "role": "detail", "item_col": None,
                                                     "item_code_col": None, "hierarchy_cols": [], "ancestor_rows": [],
                                                     "metric_code": None, "reason": "x"}],
                                            "unit": "원", "normalization_mode": "amount_to_won", "unit_evidence": [ref],
                                            "unit_overrides": [], "exclusions": [], "exceptions": []}],
                                 "excluded_areas": [], "uncertainties": []}}

        with tempfile.TemporaryDirectory() as tmp:
            client = DetailOnlyClient(Path(tmp))
            mapped, semantic = infer_structure(workbook, client)

        diagnostics_by_cell = {d["record_id"]: d for d in mapped["diagnostics"] if d.get("code", "").startswith("FORMULA_RECOMPUTE")}
        d5 = next(r for r in mapped["records"] if r["source"]["cell"] == "D5")
        d6 = next(r for r in mapped["records"] if r["source"]["cell"] == "D6")
        self.assertEqual(diagnostics_by_cell[d5["record_id"]]["code"], "FORMULA_RECOMPUTE_MATCH")   # 100+200 == 300
        self.assertEqual(diagnostics_by_cell[d6["record_id"]]["code"], "FORMULA_RECOMPUTE_MISMATCH")  # 100+200 != 999
        # The crosscheck is diagnostic-only: skipped-sheet records never
        # become mapped/confirmed, no matter what the recompute found.
        for record in (d5, d6):
            self.assertEqual(record["mapping"]["status"], "unmapped")
            self.assertIsNone(record["value"]["normalized_value"])


class TimeoutFallbackTests(unittest.TestCase):
    """Covers requirement 3: a connection-class timeout on a whole-sheet
    request is NOT retried with the same payload — it falls back to
    smaller row-range blocks, and the size reduction is logged."""

    def test_timeout_falls_back_to_blocks_instead_of_repeating_the_same_request(self):
        workbook = big_sheet_fixture(num_rows=120)  # under the threshold: whole-sheet is attempted first

        class TimeoutThenBlocksClient:
            def __init__(self, cache_dir):
                self.cache_dir, self.events, self.calls, self.notifications = cache_dir, [], [], []
            def _notify(self, message):
                self.notifications.append(message)
            def ask(self, task, prompt, context, schema, **kwargs):
                self.calls.append(task)
                if task == "sheet_structure_v1":
                    raise ModelError("simulated timeout", connection_error=True)
                self.events.append({"task": task, "cached": False})
                lo, hi = context["declared_row_range"]
                return {"plan": block_plan(lo, hi)}

        with tempfile.TemporaryDirectory() as tmp:
            client = TimeoutThenBlocksClient(Path(tmp))
            mapped, semantic = infer_structure(workbook, client)

        # The whole-sheet request was attempted exactly once — not repeated
        # as-is after the timeout.
        self.assertEqual(client.calls.count("sheet_structure_v1"), 1)
        block_calls = [t for t in client.calls if t.startswith("sheet_structure_v1_block")]
        self.assertGreaterEqual(len(block_calls), 2)  # 120 rows / 100-row blocks -> smaller requests
        self.assertEqual(len(semantic["timeout_fallbacks"]), 1)
        fallback = semantic["timeout_fallbacks"][0]
        self.assertEqual(fallback["sheet"], "Sheet")
        self.assertGreater(fallback["original_request_bytes"], 0)
        self.assertEqual(fallback["block_count"], len(block_calls))
        self.assertTrue(any("블록" in message and "KB" in message for message in client.notifications))
        self.assertEqual(sum(1 for r in mapped["records"] if r["mapping"]["status"] == "mapped"), 120)

    def test_non_connection_errors_are_not_treated_as_a_shrink_opportunity(self):
        # An HTTP-error-class ModelError (connection_error=False, the
        # default) must propagate normally — only a genuine connection
        # timeout triggers the block fallback.
        workbook = big_sheet_fixture(num_rows=5)

        class HttpErrorClient:
            def __init__(self, cache_dir):
                self.cache_dir, self.events = cache_dir, []
            def ask(self, task, prompt, context, schema, **kwargs):
                raise ModelError("simulated HTTP 403")

        with tempfile.TemporaryDirectory() as tmp:
            client = HttpErrorClient(Path(tmp))
            mapped, semantic = infer_structure(workbook, client)
        self.assertEqual(len(semantic["failures"]), 1)
        self.assertEqual(semantic["timeout_fallbacks"], [])


def plan_for(sheet_name):
    """plan(), retargeted (table + every {sheet, cell} pointer) to a
    different sheet name — for building multi-sheet fixtures."""
    def retarget(value):
        if isinstance(value, dict):
            if set(value) == {"sheet", "cell"}:
                value["sheet"] = sheet_name
            else:
                for child in value.values():
                    retarget(child)
        elif isinstance(value, list):
            for child in value:
                retarget(child)
    declaration = copy.deepcopy(plan())
    declaration["sheet"] = sheet_name
    retarget(declaration)
    return declaration


def two_sheet_fixture():
    """Two independent sheets, each shaped like fixture()'s single sheet,
    so each needs its own structure-interpretation call."""
    workbook = fixture()
    sheet_one = dict(workbook["sheets"][0], name="SheetOne",
                      cells=sheet_cells("SheetOne", workbook["sheets"][0]["cells"]))
    other = fixture()["sheets"][0]
    sheet_two = dict(other, name="SheetTwo", cells=sheet_cells("SheetTwo", other["cells"]))
    return {"file": workbook["file"], "sheets": [sheet_one, sheet_two], "diagnostics": []}


class PerSheetIndependentFailureTests(unittest.TestCase):
    """infer_structure() must not let one sheet's final failure stop every
    other sheet from being attempted — only a now-exhausted call budget
    should do that (see structured.py's main per-sheet loop)."""

    def test_one_sheets_failure_does_not_block_the_next_sheets_attempt(self):
        workbook = two_sheet_fixture()

        class PartiallyFailingClient:
            def __init__(self, cache_dir):
                self.cache_dir, self.events, self.attempted = cache_dir, [], []

            def ask(self, task, prompt, context, schema, **kwargs):
                self.attempted.append(context["sheet"])
                self.events.append({"task": task, "cached": False})
                if context["sheet"] == "SheetOne":
                    raise RuntimeError("mock sheet-specific failure")
                return {"plan": plan_for("SheetTwo")}

        with tempfile.TemporaryDirectory() as tmp:
            client = PartiallyFailingClient(Path(tmp))
            mapped, semantic = infer_structure(workbook, client)

        # Both sheets were attempted — SheetTwo was not skipped just
        # because SheetOne failed first.
        self.assertEqual(client.attempted, ["SheetOne", "SheetTwo"])
        self.assertEqual(len(semantic["failures"]), 1)
        self.assertEqual(semantic["failures"][0]["sheet"], "SheetOne")
        self.assertEqual(len(semantic["structure_plans"]), 1)
        self.assertEqual(semantic["structure_plans"][0]["sheet"], "SheetTwo")
        sheet_one_records = [r for r in mapped["records"] if r["source"]["sheet"] == "SheetOne"]
        sheet_two_records = [r for r in mapped["records"] if r["source"]["sheet"] == "SheetTwo"]
        self.assertTrue(all(r["mapping"]["status"] == "unmapped" for r in sheet_one_records))
        self.assertTrue(any(r["mapping"]["status"] == "mapped" for r in sheet_two_records))

    def test_budget_exhaustion_stops_remaining_sheets_without_attempting_them(self):
        workbook = two_sheet_fixture()

        class BudgetedClient:
            """Mimics ResponsesClient's call_count/max_calls bookkeeping
            closely enough to exercise the budget-aware break: the cap is
            reached by the first (failing) call, so the second sheet must
            never even be attempted."""
            max_calls = 1

            def __init__(self, cache_dir):
                self.cache_dir, self.events, self.attempted, self.call_count = cache_dir, [], [], 0

            def ask(self, task, prompt, context, schema, **kwargs):
                self.attempted.append(context["sheet"])
                self.call_count += 1
                self.events.append({"task": task, "cached": False})
                raise RuntimeError("mock failure that exhausts the budget")

        with tempfile.TemporaryDirectory() as tmp:
            client = BudgetedClient(Path(tmp))
            mapped, semantic = infer_structure(workbook, client)

        self.assertEqual(client.attempted, ["SheetOne"])  # SheetTwo never attempted
        self.assertEqual(len(semantic["failures"]), 1)
        self.assertEqual(semantic["structure_plans"], [])


class RectangleValidationMessageTests(unittest.TestCase):
    """The 'start_row must equal header_end_row + 1' mistake (observed in a
    real run against 매출_세부_IFRS15後, where the model repeated the exact
    same off-by-two rectangle on repair) must produce a specific, numeric
    error — not the old generic 'invalid table/header rectangle' bucket
    message — so a repair retry has an actionable signal to correct it."""

    def test_start_row_equal_to_header_start_row_gets_a_specific_actionable_message(self):
        # Reproduces the real failure shape seen against 매출_세부_IFRS15後:
        # a multi-row header (header_start_row < header_end_row) where the
        # model set start_row equal to header_start_row instead of
        # header_end_row + 1 — both the first LLM response and its repair
        # made this exact mistake.
        workbook, declaration = fixture(), plan()
        table = declaration["tables"][0]
        table["start_row"] = table["header_start_row"]  # 1, instead of 3
        errors = validate_plan(declaration, workbook, workbook["sheets"][0])
        self.assertFalse(any(e == "invalid table/header rectangle" for e in errors))
        specific = [e for e in errors if "header_end_row" in e and "start_row" in e]
        self.assertTrue(specific, f"expected a specific header_end_row/start_row message, got: {errors}")
        self.assertIn("(1)", specific[0])
        self.assertIn("merged", specific[0].lower())

    def test_correcting_start_row_to_header_end_row_plus_one_resolves_it(self):
        workbook, declaration = fixture(), plan()
        table = declaration["tables"][0]
        table["start_row"] = table["header_start_row"]  # the broken declaration
        self.assertTrue(validate_plan(declaration, workbook, workbook["sheets"][0]))
        table["start_row"] = table["header_end_row"] + 1  # the fix: back to 3
        self.assertEqual(validate_plan(declaration, workbook, workbook["sheets"][0]), [])

    def test_reversed_row_and_column_bounds_get_distinct_specific_messages(self):
        workbook, declaration = fixture(), plan()
        declaration["tables"][0]["end_row"] = 2  # now < start_row (3)
        errors = validate_plan(declaration, workbook, workbook["sheets"][0])
        self.assertTrue(any("end_row" in e and "start_row" in e and "(2)" in e for e in errors))

        workbook, declaration = fixture(), plan()
        declaration["tables"][0]["start_col"], declaration["tables"][0]["end_col"] = "C", "B"
        errors = validate_plan(declaration, workbook, workbook["sheets"][0])
        self.assertTrue(any("end_col" in e and "start_col" in e for e in errors))


if __name__ == "__main__":
    unittest.main()
