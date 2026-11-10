"""경로 조회와 실제 파싱·검증 명령을 콘솔에 연결한다.

저장은 명시한 옵션에서만 수행하며 기존 결과를 덮어쓰지 않는다.
"""

import argparse
import json
import re
import uuid
from pathlib import Path

from .paths import ProjectPaths


def main() -> int:
    """CLI 조회 또는 파싱·검증을 실행하고 종료 상태를 반환한다.

    Args:
        없음. sys.argv에서 --root와 paths/files/parse 명령을 읽는다.
        parse의 inputs는 파일·폴더 목록이며 --write-output은 결과 저장 요청이다.

    Returns:
        정상 완료 또는 의미 미확정이 있는 원본 추출 완료는 0.
        parse의 파일 처리·독립 검증 실패는 1. 프로세스 종료는 호출자가 처리한다.

    Raises:
        SystemExit: 도움말은 0, 인자 오류는 2, 포착한 OSError/ValueError는 1로
            종료한다. 도움말·인자 오류는 설정 파일을 읽기 전에 처리한다.
        RuntimeError: 경로 해석 등에서 발생하면 이 함수가 포착하지 않고 전파한다.

    Side Effects:
        조회는 메타데이터만, parse는 원본 Excel 내용을 읽는다.
        --write-output이 있을 때만 output 하위에 새 실행 결과를 생성한다.
        Data 내부와 기존 출력에는 쓰지 않는다. 콘솔 출력이 있다.
        Python 바이트코드 캐시는 호출 시 -B 또는 환경변수로 별도 차단해야 한다.
    """
    parser = argparse.ArgumentParser(description="엑셀 데이터 파서 디렉토리 관리")
    parser.add_argument("--root", help="프로젝트 루트 경로 (생략하면 자동 탐색)")
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("paths", help="설정된 디렉토리 경로 확인")
    files = commands.add_parser("files", help="입력 엑셀 파일 목록 확인")
    files.add_argument("--recursive", action="store_true", help="입력 하위 폴더도 탐색")
    parse = commands.add_parser("parse", help="전체 원본 추출·의미 매핑·독립 검증")
    parse.add_argument("inputs", nargs="+", type=Path, help="원본 Excel 파일 또는 폴더")
    parse.add_argument("--recursive", action="store_true", help="입력 하위 폴더도 탐색")
    parse.add_argument(
        "--write-output",
        action="store_true",
        help="사용자가 허용한 범위에서 결과 파일 저장",
    )
    parse.add_argument(
        "--output-name", help="output 아래 새 실행 폴더명 (기본: 고유 ID)"
    )
    parse.add_argument(
        "--metadata", type=Path, help="defaults/files 구조의 JSON 메타데이터"
    )
    parse.add_argument("--submission-version", help="제출본 버전 원문. 예: V1.01")
    parse.add_argument(
        "--submission-month", help="숫자 대상 월과 별개인 제출 월. 예: 2026-11"
    )
    # 도움말과 잘못된 인자는 설정 파일에 접근하기 전에 처리한다.
    args = parser.parse_args()

    try:
        paths = ProjectPaths.load(args.root)
        if args.command == "parse":
            from .pipeline import run_parser

            metadata = (
                json.loads(args.metadata.read_text(encoding="utf-8"))
                if args.metadata
                else {}
            )
            if not isinstance(metadata, dict):
                raise ValueError("메타데이터는 JSON 객체여야 합니다.")
            defaults = metadata.setdefault("defaults", {})
            if not isinstance(defaults, dict):
                raise ValueError("metadata.defaults는 객체여야 합니다.")
            if args.submission_version:
                defaults["submission_version"] = args.submission_version
            if args.submission_month:
                defaults["submission_month"] = args.submission_month
            output_dir = None
            if args.output_name and not args.write_output:
                raise ValueError("--output-name은 --write-output과 함께 사용하세요.")
            if args.write_output:
                name = args.output_name or "run-" + uuid.uuid4().hex[:12]
                if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]*", name):
                    raise ValueError(
                        "출력 폴더명은 영문·숫자로 시작하며 영문·숫자·점·밑줄·하이픈만 사용합니다."
                    )
                output_dir = paths.output_dir / name
            summary = run_parser(
                args.inputs,
                root=paths.root,
                recursive=args.recursive,
                output_dir=output_dir,
                metadata=metadata,
            )
            print(
                json.dumps(
                    {key: value for key, value in summary.items() if key != "files"},
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 0 if summary["validation_passed"] else 1
        if args.command == "files":
            workbooks = paths.list_input_files(recursive=args.recursive)
            for workbook in workbooks:
                print(workbook.relative_to(paths.root))
            print(f"총 {len(workbooks)}개")
        else:
            print(f"프로젝트: {paths.root}")
            print(f"입력:     {paths.input_dir}")
            print(f"출력 예정 경로 (저장 시 허락 필요): {paths.output_dir}")
            print(f"로그 예정 경로 (저장 시 허락 필요): {paths.logs_dir}")
    except (OSError, ValueError) as error:
        # 사용자에게 보여줄 오류와 종료 코드를 맞추며 별도 로그 파일은 만들지 않는다.
        parser.exit(1, f"오류: {error}\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
