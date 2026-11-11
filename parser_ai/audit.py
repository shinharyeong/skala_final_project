"""Read-only, source-linked audit of one parser_ai output directory.

This is a diagnostic comparison, not a rule-based correction or an accuracy score.
"""
import argparse
import csv
import gzip
import json
import re
import sys
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
from excel_parser_ai.shared import normalize_number


def rows(path):
    with gzip.open(path, "rt", encoding="utf-8") as stream:
        for line in stream:
            yield json.loads(line)


def equal_number(a, b):
    try:
        return Decimal(str(a)) == Decimal(str(b))
    except (InvalidOperation, TypeError):
        return a == b


def audit(directory):
    directory = Path(directory).resolve()
    summary = json.loads((directory / "summary.json").read_text())
    raw = {r["record_id"]: r for r in rows(directory / "raw-cells.jsonl.gz")}
    records = list(rows(directory / "standard-records.jsonl.gz"))
    diagnostics = list(rows(directory / "diagnostics.jsonl.gz"))
    by_record = defaultdict(list)
    for diagnostic in diagnostics:
        if diagnostic.get("record_id"):
            by_record[diagnostic["record_id"]].append(diagnostic)
    counts = Counter()
    seen = Counter()
    output = directory / "source-audit.csv"
    fields = ["check", "issues", "source_sheet", "source_cell", "raw_value", "formula", "ai_organization_code",
              "ai_organization_name", "ai_item_code", "ai_item_name", "ai_item_path", "ai_year", "ai_month",
              "ai_aggregation", "ai_scenario", "header_evidence", "ai_role", "ai_value_kind", "ai_normalized_value",
              "ai_unit", "ai_multiplier", "unit_evidence", "mapping_status", "value_status", "diagnostics", "record_id"]
    with output.open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader()
        for record in records:
            source = record["source"]
            rid = source["raw_record_id"]
            original = raw.get(rid)
            seen[rid] += 1
            issues = []
            if original is None:
                issues.append("SOURCE_MISSING")
            else:
                if not equal_number(original.get("value"), record.get("raw_value")):
                    issues.append("RAW_VALUE_MISMATCH")
                if original["sheet"] != source["sheet"] or original["address"] != source["cell"]:
                    issues.append("SOURCE_COORDINATE_MISMATCH")
            if seen[rid] > 1:
                issues.append("DUPLICATE_SOURCE")
            value = record["value"]
            period = record["period"]
            evidence = [str(x) for x in period.get("raw_headers", []) if x is not None]
            evidence_text = " ".join(evidence)
            if record["mapping"]["status"] == "mapped":
                year_tokens = set(int(x) for x in re.findall(r"(?<!\d)(20\d{2})(?:년)?(?!\d)", evidence_text))
                if year_tokens and period.get("year") not in year_tokens:
                    issues.append("YEAR_EVIDENCE_CONFLICT")
                month_tokens = set(int(x) for x in re.findall(r"(?<!\d)(1[0-2]|0?[1-9])월", evidence_text))
                if month_tokens and period.get("aggregation") == "month" and period.get("month") not in month_tokens:
                    issues.append("MONTH_EVIDENCE_CONFLICT")
                for token, scenario in [("실적", "actual"), ("계획", "plan"), ("추정", "estimate")]:
                    if token in evidence_text and period.get("scenario") not in {scenario, "mixed"}:
                        issues.append("SCENARIO_EVIDENCE_CONFLICT")
                        break
            if original and value.get("status") == "confirmed" and value.get("normalized_value") is not None:
                check = normalize_number(original.get("value"), data_type=original["data_type"],
                                         value_kind=value.get("value_kind"), unit=value.get("source_unit"),
                                         number_format=original.get("number_format"))
                if check.get("status") == "confirmed" and check.get("normalized_value") is not None:
                    if not equal_number(check["normalized_value"], value["normalized_value"]):
                        issues.append("NORMALIZATION_DISAGREEMENT")
            if value.get("source_unit"):
                unit_refs = record.get("unit_provenance") or []
                if not unit_refs:
                    issues.append("UNIT_EVIDENCE_MISSING")
                elif not any(ref.get("raw_record_id") in raw for ref in unit_refs):
                    issues.append("UNIT_EVIDENCE_SOURCE_MISSING")
                elif not any(str(value["source_unit"]) in str(raw[ref["raw_record_id"]].get("value"))
                             for ref in unit_refs if ref.get("raw_record_id") in raw):
                    issues.append("UNIT_EVIDENCE_TEXT_REVIEW")
            if record["mapping"]["status"] == "mapped":
                for target, refs, issue in [
                    (record["organization"].get("code"), record["organization"].get("provenance", []), "ORG_CODE_EVIDENCE_REVIEW"),
                    (record["organization"].get("name"), record["organization"].get("provenance", []), "ORG_NAME_EVIDENCE_REVIEW"),
                    (record["item"].get("name"), record["item"].get("provenance", []), "ITEM_EVIDENCE_REVIEW"),
                ]:
                    if target and refs and not any(str(target) in str(raw[ref["raw_record_id"]].get("value"))
                            for ref in refs if ref.get("raw_record_id") in raw):
                        issues.append(issue)
            if record["mapping"]["status"] != "mapped" or value.get("status") != "confirmed":
                issues.append("AI_UNRESOLVED")
            if not period.get("provenance") and record["mapping"]["status"] == "mapped":
                issues.append("PERIOD_EVIDENCE_MISSING")
            check_status = "issue" if any(i not in {"AI_UNRESOLVED"} for i in issues) else "review" if issues else "checks_passed"
            counts[check_status] += 1
            counts.update(issues)
            writer.writerow({
                "check": check_status, "issues": " | ".join(issues), "source_sheet": source["sheet"],
                "source_cell": source["cell"], "raw_value": original.get("value") if original else "",
                "formula": (original.get("formula") or {}).get("text") if original else "",
                "ai_organization_code": record["organization"].get("code"),
                "ai_organization_name": record["organization"].get("name"),
                "ai_item_code": record["item"].get("code"), "ai_item_name": record["item"].get("name"),
                "ai_item_path": " / ".join(str(x) for x in record["item"].get("header_path", []) if x),
                "ai_year": period.get("year"), "ai_month": period.get("month"),
                "ai_aggregation": period.get("aggregation"), "ai_scenario": period.get("scenario"),
                "header_evidence": " / ".join(evidence), "ai_role": record["role"].get("kind"),
                "ai_value_kind": value.get("value_kind"), "ai_normalized_value": value.get("normalized_value"),
                "ai_unit": value.get("source_unit"), "ai_multiplier": value.get("multiplier"),
                "unit_evidence": " / ".join(f"{p.get('sheet')}!{p.get('cell')}" for p in record.get("unit_provenance", [])),
                "mapping_status": record["mapping"]["status"], "value_status": value.get("status"),
                "diagnostics": " | ".join(
                    [d.get("code", "") + ": " + d.get("message", "") for d in by_record[record["record_id"]]]
                    + record["mapping"].get("reasons", []) + value.get("diagnostics", [])),
                "record_id": record["record_id"],
            })
        for rid, original in raw.items():
            if original.get("has_content", True) and (original["data_type"] == "number" or original.get("formula")) and seen[rid] == 0:
                counts["MISSING_CANDIDATE"] += 1
                counts["issue"] += 1
                writer.writerow({"check": "issue", "issues": "MISSING_CANDIDATE", "source_sheet": original["sheet"],
                                 "source_cell": original["address"], "raw_value": original.get("value"),
                                 "formula": (original.get("formula") or {}).get("text"), "record_id": rid})
    report = {"run": directory.name, "llm_inference_completed": summary["llm_inference_completed"],
              "records": len(records), "candidate_raw_cells": sum(1 for x in raw.values() if x.get("has_content", True) and
                    (x["data_type"] == "number" or x.get("formula"))), "counts": dict(counts),
              "workbook_diagnostics": [d for d in diagnostics if not d.get("record_id")][:100],
              "csv": str(output)}
    (directory / "source-audit-summary.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({k: v for k, v in report.items() if k != "workbook_diagnostics"}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_dir", type=Path)
    audit(parser.parse_args().output_dir)
