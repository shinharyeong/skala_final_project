"""Only the reader, numeric normalizer and source validators are shared.

Rule-based semantic mapping is not imported. The model does not receive the
baseline output, a supplied master, or an answer sheet from another workbook.
"""
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[3]
BASELINE = REPO / "parser"
sys.path.insert(0, str(BASELINE / "src"))
from excel_parser.extraction import extract_workbook, file_fingerprint
from excel_parser.normalization import normalize_number
from excel_parser.validation import validate_extraction, validate_provenance


def cell_index(workbook):
    return {(s["name"], c["address"]): c for s in workbook["sheets"] for c in s["cells"]}


def ref(cell):
    return {"sheet": cell["sheet"], "cell": cell["address"], "raw_record_id": cell["record_id"]}


def column_number(label):
    if not isinstance(label, str) or not label.isalpha() or not label.isupper():
        raise ValueError(f"Invalid column: {label!r}")
    n = 0
    for character in label:
        n = n * 26 + ord(character) - 64
    return n


def source_value(cell):
    return cell.get("value") if cell else None
