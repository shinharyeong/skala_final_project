"""Run from any current directory: python3 parser_ai/run.py --help."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
from excel_parser_ai.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
