import copy
import tempfile
import unittest
from pathlib import Path

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
            def ask(self, task, prompt, context, schema):
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
            def ask(self, task, prompt, context, schema):
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
            def ask(self, task, prompt, context, schema):
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
            def ask(self, task, prompt, context, schema):
                self.calls += 1
                self.events.append({"task": task, "cached": False})
                return {"plan": plan()}
        with tempfile.TemporaryDirectory() as tmp:
            client = FakeClient(Path(tmp))
            _, semantic = infer_structure(workbook, client)
            self.assertEqual(client.calls, 1)
            self.assertEqual(len(semantic["structure_plans"]), 2)


if __name__ == "__main__":
    unittest.main()
