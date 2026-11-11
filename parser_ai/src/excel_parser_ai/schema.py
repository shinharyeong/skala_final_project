"""Strict Responses API schemas for the historical legacy-cell engine.

The model reads real source numbers, formulas and cached values (Python no
longer masks amounts) and returns its own normalized value directly for every
financial candidate cell. Python never recomputes or overwrites that value;
it only validates grounding (do the model's claims match the real source?)
and coverage (did it answer for every required row/cell?).
"""


def obj(properties):
    return {"type": "object", "properties": properties, "required": list(properties), "additionalProperties": False}


def arr(items):
    return {"type": "array", "items": items}


STR = {"type": "string"}
NULL_STR = {"type": ["string", "null"]}
NULL_INT = {"type": ["integer", "null"]}
REF = obj({"sheet": STR, "cell": STR})
NULL_REF = {"anyOf": [REF, {"type": "null"}]}
UNIT = {"type": ["string", "null"], "enum": [None, "원", "천원", "만원", "백만원", "억원", "%", "명", "개", "건", "대"]}
PERIOD = obj({
    "column": STR, "year": NULL_INT, "month": NULL_INT, "quarter": NULL_INT,
    "aggregation": {"type": "string", "enum": ["month", "annual", "unknown"]},
    "scenario": {"type": "string", "enum": ["actual", "plan", "estimate", "mixed", "unknown"]},
    "evidence": arr(REF),
})
SHEET = obj({
    "sheet": STR,
    "kind": {"type": "string", "enum": ["revenue_detail", "expense_detail", "financial_statement", "organization_summary", "reference", "unknown"]},
    "header_end_row": {"type": "integer"},
    "organization_code_column": NULL_STR, "organization_name_column": NULL_STR,
    "fallback_organization_code": NULL_REF, "fallback_organization_name": NULL_REF,
    "unit": UNIT, "unit_evidence": arr(REF), "periods": arr(PERIOD),
    "dimensions": arr(obj({"name": STR, "column": STR})),
    "notes": arr(STR),
})
LAYOUT_SCHEMA = obj({"sheets": arr(SHEET), "uncertainties": arr(STR)})
# One sheet per layout API call (see context.sheet_layout_context): smaller,
# faster requests than one combined call for the whole workbook, with
# cross-sheet evidence still available via a separate label-only field.
SHEET_LAYOUT_SCHEMA = obj({"sheet_plan": SHEET, "uncertainties": arr(STR)})

VALUE_KIND = {"type": "string", "enum": ["amount", "ratio", "quantity", "number", "identifier", "code", "date", "text"]}
VALUE_STATUS = {"type": "string", "enum": ["confirmed", "unresolved", "not_numeric", "missing"]}
# One descriptor per financial-candidate cell in a row. The model performs the
# full interpretation here: sign, unit, percent handling and the final
# normalized decimal string. There is no separate Python normalization step
# for these cells; Python only checks that source_value_read matches what it
# actually extracted, and that normalized_value is a well-formed decimal
# string when status is "confirmed".
CELL = obj({
    "column": STR,
    "status": VALUE_STATUS,
    "value_kind": VALUE_KIND,
    "source_value_read": NULL_STR,
    "normalized_value": NULL_STR,
    "unit": UNIT, "unit_evidence": arr(REF),
    "reason": STR,
})
ROW = obj({
    "row": {"type": "integer"},
    "role": {"type": "string", "enum": ["detail", "subtotal", "total", "ratio", "reference", "unknown"]},
    "item_name": NULL_REF, "item_code": NULL_REF, "header_path": arr(REF),
    "metric_code": {"type": ["string", "null"], "enum": [None, "OP", "CONTINGENCY", "OP_EX_CONTINGENCY"]},
    "cells": arr(CELL),
    "reason": STR,
})
ROWS_SCHEMA = obj({"rows": arr(ROW), "uncertainties": arr(STR)})
