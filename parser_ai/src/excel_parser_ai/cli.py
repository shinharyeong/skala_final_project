"""CLI deliberately offers parsing only. Future comparison is a separate task."""
import argparse
import json
import os
import sys
import uuid
from datetime import datetime
from pathlib import Path

from .client import ResponsesClient, read_env
from .pipeline import run
from .shared import REPO


def main(argv=None):
    read_env(REPO / "parser_ai" / ".env")
    parser = argparse.ArgumentParser(description="OpenAI 의미 해석 엑셀 파서 — 기존 출력 계약 호환. 평가·마스터 생성 없음")
    sub = parser.add_subparsers(dest="command", required=True)
    for command in ("prepare", "run"):
        p = sub.add_parser(command, help="API 없이 원본 추출·입력 준비" if command == "prepare" else "OpenAI로 의미 해석하여 표준 데이터 저장")
        p.add_argument("inputs", nargs="+", type=Path)
        p.add_argument("--recursive", action="store_true")
        p.add_argument("--output", type=Path)
        p.add_argument("--metadata", type=Path)
        p.add_argument("--max-files", type=int)
        p.add_argument("--engine", choices=["structure", "legacy-cell"], default="structure",
                       help="기본: LLM 구조 해석 + Python Decimal 적용. legacy-cell은 이전 실험 재현용")
        p.add_argument("--max-exception-calls", type=int, default=0,
                       help="구조 예외 일괄 LLM 호출 상한(기본 0: 자동 호출 없이 미해결 보존)")
        p.add_argument("--model", default=os.getenv("PARSER_AI_MODEL", "gpt-6-astra"))
        p.add_argument("--effort", choices=["low", "medium", "high", "xhigh", "max"], default=os.getenv("PARSER_AI_REASONING_EFFORT", "high"))
        p.add_argument("--review", choices=["on", "off"], default=os.getenv("PARSER_AI_REVIEW", "on"))
        p.add_argument("--batch-size", type=int, default=12,
                       help="행 배치당 최대 행 수(기본 12). 실제 요청 크기는 --max-batch-cells/--max-batch-bytes로도 제한됨")
        p.add_argument("--max-batch-cells", type=int, default=150, help="배치당 모델이 해석할 예상 셀 수 상한")
        p.add_argument("--max-batch-bytes", type=int, default=40_000, help="배치 요청 맥락의 예상 바이트 상한")
        p.add_argument("--concurrency", type=int, default=int(os.getenv("PARSER_AI_CONCURRENCY", "2")))
        p.add_argument("--max-calls", type=int, default=300)
        p.add_argument("--max-new-row-batches", type=int, default=None,
                       help=("테스트용: 캐시된 레이아웃/행 배치는 그대로 재사용하고, 실제로 새로 호출되는 "
                             "행 배치를 이 수만큼만 처리한 뒤 그 파일 처리를 의도적으로 중단한다(실패 아님, "
                             "llm_inference_completed=false). 동시 실행 경합을 피하려고 이 옵션을 쓰면 해당 "
                             "파일의 행 배치는 순차 처리된다. 기본은 비활성(제한 없음)"))
    args = parser.parse_args(argv)
    if (args.batch_size < 1 or args.max_batch_cells < 1 or args.max_batch_bytes < 1 or args.concurrency < 1
            or args.max_calls < 1 or (args.max_files is not None and args.max_files < 1)
            or (args.max_new_row_batches is not None and args.max_new_row_batches < 1)
            or args.max_exception_calls < 0):
        parser.error("batch-size, max-batch-cells, max-batch-bytes, concurrency, max-calls, max-files, max-new-row-batches는 양수여야 합니다")
    if args.engine == "structure" and args.max_new_row_batches is not None:
        parser.error("--max-new-row-batches는 이전 셀별 엔진 전용입니다")
    prepare = args.command == "prepare"
    if not prepare and not os.getenv("OPENAI_API_KEY"):
        print("OPENAI_API_KEY가 없습니다. parser_ai/.env에 키를 설정하세요. API 없이 준비하려면 prepare를 사용하세요.", file=sys.stderr)
        return 2
    metadata = json.loads(args.metadata.read_text(encoding="utf-8")) if args.metadata else {}
    name = datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:8]
    output = args.output or REPO / "parser_ai" / "output" / name
    report = lambda m: print(m, file=sys.stderr, flush=True)
    client = None if prepare else ResponsesClient(REPO / "parser_ai" / "cache", model=args.model, effort=args.effort, max_calls=args.max_calls, progress=report)
    try:
        summary = run(args.inputs, output, client=client, prepare=prepare, metadata=metadata,
                      recursive=args.recursive, review=args.review == "on", batch_size=args.batch_size,
                      max_batch_output_cells=args.max_batch_cells, max_batch_context_bytes=args.max_batch_bytes,
                      concurrency=args.concurrency, max_files=args.max_files,
                      max_new_row_batches=args.max_new_row_batches, progress=report,
                      structure_mode=args.engine == "structure", max_exception_calls=args.max_exception_calls)
    except (OSError, ValueError) as error:
        print(f"처리 중단: {error}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        # No long traceback: say where it stopped and where partial results
        # (raw extraction already written before any API call per file) live.
        print("", file=sys.stderr)
        print("중단됨 (Ctrl+C). 위에 표시된 마지막 단계에서 멈췄습니다.", file=sys.stderr)
        print(f"부분 결과 위치: {output}", file=sys.stderr)
        print("raw-cells.jsonl.gz/workbooks.jsonl.gz 등 그 시점까지 처리된 파일은 저장되어 있을 수 있으나, "
              "summary.json/inference.json은 아직 생성되지 않았을 수 있습니다. 추가 재시도는 하지 않았습니다.", file=sys.stderr)
        return 130
    print(json.dumps({k: v for k, v in summary.items() if k != "files"}, ensure_ascii=False, indent=2))
    return 0 if summary["validation_passed"] and (prepare or summary["llm_inference_completed"]) else 2


if __name__ == "__main__":
    raise SystemExit(main())
