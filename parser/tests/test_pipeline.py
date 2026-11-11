"""합성 메모리 자료로 출력 보호·실패 처리·십진 직렬화를 검증한다."""

import hashlib
import json
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from helpers import PROJECT_ROOT, extract_memory, make_package, open_memory_archive

from excel_parser.pipeline import collect_inputs, json_value, run_parser

_ORIGINAL_PATH_OPEN = Path.open


def _read_only_open(path: Path, mode: str = "r", *args, **kwargs):
    """설정 파일의 읽기만 허용하고 테스트 중 파일 쓰기를 차단한다.

    Args:
        path: 열려는 경로.
        mode: 파일 열기 모드.
        args: Path.open에 전달할 추가 위치 인자.
        kwargs: Path.open에 전달할 추가 이름 인자.
    Returns:
        원래 Path.open이 반환한 읽기 스트림.
    Raises:
        AssertionError: 쓰기 또는 갱신 모드로 파일을 열려는 경우.
        OSError: 읽기 파일에 접근할 수 없는 경우.
    Side Effects:
        기존 설정 파일만 읽는다. 파일을 만들거나 변경하지 않는다.
    """
    if any(flag in mode for flag in "wax+"):
        raise AssertionError("테스트 중 디스크 쓰기를 허용하지 않습니다.")
    return _ORIGINAL_PATH_OPEN(path, mode, *args, **kwargs)


class PipelineTests(unittest.TestCase):
    """원본 파일 없이 입출력 경계의 사용자 보호 계약을 검사한다."""

    def _run_memory(
        self,
        *,
        inputs: list[Path] | None = None,
        output_dir: Path | None = None,
    ) -> dict:
        """파일 입출력 경계만 대체하고 실제 파이프라인을 실행한다.

        Args:
            inputs: 실제로 생성하지 않는 가상 입력 경로 목록.
            output_dir: 거부 동작을 검사할 출력 경로. 기본값은 저장하지 않음.
        Returns:
            실제 run_parser가 반환한 실행 요약.
        Raises:
            AssertionError: 실행 중 디스크 쓰기를 시도한 경우.
            Exception: 파이프라인이 호출자에게 전파하는 입력 오류.
        Side Effects:
            설정 파일만 읽는다. OOXML과 해시는 메모리에 두며 파일을 생성하지 않는다.
        """
        package = make_package()
        sources = inputs or [PROJECT_ROOT / "__memory__/synthetic.xlsx"]
        with (
            patch("excel_parser.pipeline.collect_inputs", return_value=sources),
            patch(
                "excel_parser.pipeline.file_fingerprint",
                return_value=hashlib.sha256(package).hexdigest(),
            ),
            patch(
                "excel_parser.pipeline.extract_workbook",
                side_effect=lambda path, **kwargs: extract_memory(
                    package, name=path.name
                ),
            ),
            patch(
                "excel_parser.validation.ZipFile",
                side_effect=lambda _: open_memory_archive(package),
            ),
            patch.object(Path, "open", autospec=True, side_effect=_read_only_open),
            patch.object(Path, "mkdir", side_effect=AssertionError("mkdir 금지")),
            patch.object(
                Path, "write_text", side_effect=AssertionError("텍스트 쓰기 금지")
            ),
            patch.object(
                Path, "write_bytes", side_effect=AssertionError("바이트 쓰기 금지")
            ),
            patch(
                "excel_parser.pipeline.gzip.open",
                side_effect=AssertionError("압축 결과 쓰기 금지"),
            ),
        ):
            return run_parser(sources, root=PROJECT_ROOT, output_dir=output_dir)

    def test_decimal_serialization_is_text_and_exact(self) -> None:
        """긴 십진수를 JSON 문자열로 정확히 보존한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 숫자 타입이나 자릿수가 손실된 경우.
        Side Effects:
            메모리 문자열만 생성한다.
        """
        value = Decimal("123456789012345678901234567890.00100")
        result = json.loads(json.dumps({"value": value}, default=json_value))
        self.assertEqual(result["value"], "123456789012345678901234567890.00100")

    def test_explicit_inputs_are_deduplicated(self) -> None:
        """같은 해석 경로를 여러 번 명시해도 한 번만 수집한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 동일 경로가 중복 처리되는 경우.
        Side Effects:
            파일 존재 여부만 대체하며 실제 파일을 생성하거나 읽지 않는다.
        """
        source = PROJECT_ROOT / "__memory__/synthetic.xlsx"
        with patch.object(Path, "is_file", return_value=True):
            result = collect_inputs([source, source])
        self.assertEqual(result, [source.resolve()])

    def test_data_output_is_rejected_before_creation(self) -> None:
        """Data 내부 결과 생성 요청을 디스크 쓰기 전에 거부한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 원본 경로 쓰기를 시도하거나 경로를 허용한 경우.
        Side Effects:
            설정만 읽고 합성 메모리 자료를 사용한다.
        """
        with self.assertRaises(ValueError):
            self._run_memory(output_dir=PROJECT_ROOT / "Data/forbidden")

    def test_read_only_run_never_opens_output(self) -> None:
        """기본 실행은 실제 추출·매핑·검증을 하되 결과 파일을 열지 않는다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 저장을 시도하거나 독립 검증이 실패한 경우.
        Side Effects:
            설정을 읽고 메모리 OOXML을 처리한다. 파일 쓰기는 차단한다.
        """
        result = self._run_memory()
        self.assertTrue(result["raw_extraction_complete"])
        self.assertTrue(result["validation_passed"])
        self.assertTrue(result["originals_unchanged"])
        self.assertEqual(result["content_cell_count"], 5)
        self.assertEqual(result["formula_count"], 1)
        self.assertEqual(result["standard_record_count"], 2)
        self.assertIsNone(result["output_dir"])

    def test_mapping_failure_does_not_discard_raw_extraction(self) -> None:
        """의미 매핑에 실패해도 완료된 원본 추출을 유지한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 원본 집계를 잃거나 매핑 실패를 성공으로 보고한 경우.
        Side Effects:
            합성 OOXML을 메모리에서 처리하며 결과 파일은 만들지 않는다.
        """
        with patch(
            "excel_parser.pipeline.map_workbook",
            side_effect=ValueError("mapping fixture"),
        ):
            result = self._run_memory()
        self.assertTrue(result["raw_extraction_complete"])
        self.assertEqual(result["content_cell_count"], 5)
        self.assertEqual(result["standard_record_count"], 0)
        self.assertFalse(result["validation_passed"])
        self.assertEqual(result["failures"][0]["stage"], "mapping")

    def test_unsupported_input_is_reported_without_skipping_others(self) -> None:
        """미지원 확장자를 진단하고 지원 입력의 추출은 계속한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 미지원 입력이 누락되거나 전체 완료로 잘못 보고된 경우.
        Side Effects:
            가상 입력 경로와 메모리 OOXML만 사용한다.
        """
        result = self._run_memory(
            inputs=[
                PROJECT_ROOT / "__memory__/supported.xlsx",
                PROJECT_ROOT / "__memory__/unsupported.xlsb",
            ]
        )
        self.assertEqual(result["input_file_count"], 2)
        self.assertEqual(result["processed_file_count"], 1)
        self.assertFalse(result["raw_extraction_complete"])
        self.assertFalse(result["validation_passed"])
        self.assertEqual(result["failures"][0]["stage"], "extraction")


if __name__ == "__main__":
    unittest.main()
