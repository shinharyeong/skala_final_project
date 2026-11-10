"""원본 저장값과 독립적으로 정한 경계 기대값으로 숫자 정규화를 검증한다.

임시 파일·샘플 워크북·캐시를 생성하지 않고 메모리에서만 실행한다.
"""

import unittest
from decimal import Decimal

from excel_parser.normalization import inspect_number_format, normalize_number


class NormalizationTests(unittest.TestCase):
    """숫자·부호·단위·보호 자료형의 정규화 계약을 확인한다."""

    def assert_value(self, original: object, expected: str, **options: object) -> dict:
        """정규화 결과를 사전에 정한 정확한 십진 기대값과 비교한다.

        Args:
            original: 검증할 원본 값.
            expected: 독립적으로 정한 기대값의 십진 문자열.
            options: normalize_number에 넘길 명시적 필드 조건.

        Returns:
            추가 검증에 사용할 정규화 결과 사전.

        Raises:
            AssertionError: 결과가 확정 상태 또는 기대값과 다를 때 발생한다.

        Side Effects:
            없음. 파일·원본 데이터를 변경하지 않는다.
        """
        result = normalize_number(original, **options)
        self.assertEqual(result["status"], "confirmed", result)
        self.assertIsInstance(result["normalized_value"], Decimal)
        self.assertEqual(result["normalized_value"], Decimal(expected))
        self.assertEqual(result["original_value"], original)
        return result

    def test_stored_sign_and_percent_ignore_display(self) -> None:
        """저장 부호와 0·비율이 시각 서식에 의해 바뀌지 않는지 확인한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 저장값이 부호·백분율 표시로 재변환된 경우.
        Side Effects:
            없음.
        """
        cases = [
            ("-100", "0;0", "-100"),
            ("100", "[Red](0);0", "100"),
            ("100", "[Red]0", "100"),
            ("0", '#,##0;(#,##0);"-"', "0"),
            ("-100", '#,##0;"△"#,##0', "-100"),
            ("400000000", '[Red]"↓"0', "400000000"),
            ("1000000", "#,##0,,", "1000000"),
            ("0.125", "0.00%", "0.125"),
        ]
        for original, format_code, expected in cases:
            with self.subTest(original=original, format=format_code):
                self.assert_value(
                    original, expected, data_type="number", number_format=format_code
                )
        self.assert_value(
            "0.125",
            "0.125",
            data_type="number",
            value_kind="ratio",
            unit="%",
            number_format="0.0%",
        )

    def test_text_sign_variants(self) -> None:
        """앞뒤 부호·전각·괄호·지수부의 독립 기대값을 확인한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 등록된 숫자 문자열의 결과가 다를 때 발생한다.
        Side Effects:
            없음.
        """
        negative_cases = [
            "-1,000",
            "- 1,000",
            "1,000-",
            "1,000 -",
            "−1,000",
            "－１，０００",
            "(1,000)",
            "(₩1,000)",
            "₩(1,000)",
            "(1,000)원",
            "원(1,000)",
            "(1,000원)",
        ]
        for original in negative_cases:
            with self.subTest(original=original):
                self.assert_value(
                    original,
                    "-1000",
                    data_type="string",
                    value_kind="amount",
                    locale="ko_KR",
                )
        self.assert_value(
            "+ 1,000", "1000", data_type="string", value_kind="amount", locale="ko_KR"
        )
        self.assert_value("1E-3", "0.001", data_type="string", value_kind="amount")
        self.assert_value("-1E-3", "-0.001", data_type="string", value_kind="amount")
        self.assert_value("(12.5%)", "-0.125", data_type="string", value_kind="ratio")

    def test_ambiguous_signs_and_symbols_are_not_guessed(self) -> None:
        """충돌 부호와 확인되지 않은 삼각형·화살표를 확정하지 않는다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 미확정 표기를 임의 숫자로 바꾼 경우.
        Side Effects:
            없음.
        """
        cases = [
            "(-100)",
            "--100",
            "+-100",
            "△-100",
            "▲100",
            "▽100",
            "▼100",
            "↑100",
            "↓100",
        ]
        for original in cases:
            with self.subTest(original=original):
                result = normalize_number(
                    original, data_type="string", value_kind="amount", locale="ko_KR"
                )
                self.assertEqual(result["status"], "unresolved")
                self.assertIsNone(result["normalized_value"])
        rules = {"△": {"sign": -1, "rule_id": "TEST-EXPLICIT-LEGEND"}}
        result = self.assert_value(
            "△100", "-100", data_type="string", value_kind="amount", symbol_rules=rules
        )
        self.assertIn("TEST-EXPLICIT-LEGEND", result["rule_ids"])
        for original in ("△-100", "△△100", "1△000"):
            with self.subTest(original=original):
                result = normalize_number(
                    original,
                    data_type="string",
                    value_kind="amount",
                    symbol_rules=rules,
                )
                self.assertEqual(result["status"], "unresolved")

    def test_locale_and_grouping(self) -> None:
        """로케일·비분리 공백과 모호한 단일 구분자 표기를 구분한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 그룹·소수 구분을 잘못 해석한 경우.
        Side Effects:
            없음.
        """
        cases = [
            ("1,234.56", "en_US", "1234.56"),
            ("1.234,56", "de_DE", "1234.56"),
            ("1 234,56", "fr_FR", "1234.56"),
            ("1\u00a0234,56", "fr_FR", "1234.56"),
            ("1\u202f234,56", "fr_FR", "1234.56"),
            ("1,234.56", None, "1234.56"),
            ("1.234,56", None, "1234.56"),
            ("1 234", None, "1234"),
            ("1,234,567", None, "1234567"),
            (".5", None, "0.5"),
        ]
        for original, locale, expected in cases:
            with self.subTest(original=original, locale=locale):
                self.assert_value(
                    original,
                    expected,
                    data_type="string",
                    value_kind="amount",
                    locale=locale,
                )
        for original in ("1,234", "1.234", "12,34,567", "1 23", "1,2.3", "100."):
            with self.subTest(original=original):
                self.assertEqual(
                    normalize_number(original, data_type="string", value_kind="amount")[
                        "status"
                    ],
                    "unresolved",
                )

    def test_units_are_applied_once(self) -> None:
        """단위 배율과 백분율을 한 번만 적용하고 충돌을 보류한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 배율을 중복 적용하거나 차원 충돌을 무시한 경우.
        Side Effects:
            없음.
        """
        cases = [
            ("1.25", "number", "amount", "천원", "1250"),
            ("1.25", "number", "amount", "억원", "125000000"),
            ("1.25천원", "string", "amount", "천원", "1250"),
            ("천원 1.25", "string", "amount", None, "1250"),
            ("1.25 천 원", "string", "amount", "천원", "1250"),
            ("KRW1.25천원", "string", "amount", "천원", "1250"),
            ("(₩1.25천원)", "string", "amount", "천원", "-1250"),
            ("12.5%", "string", "ratio", "%", "0.125"),
            ("12.5", "string", "ratio", "%", "0.125"),
            ("0.125", "number", "ratio", "%", "0.125"),
            ("1.5", "number", "quantity", "만명", "15000"),
        ]
        for original, data_type, kind, unit, expected in cases:
            with self.subTest(original=original, unit=unit):
                self.assert_value(
                    original, expected, data_type=data_type, value_kind=kind, unit=unit
                )
        for original, unit, kind in (
            ("1천원", "백만원", "amount"),
            ("1%", "원", "amount"),
            ("1", "만명", "amount"),
            ("1원000", None, "amount"),
        ):
            with self.subTest(original=original, unit=unit):
                self.assertEqual(
                    normalize_number(
                        original, data_type="string", value_kind=kind, unit=unit
                    )["status"],
                    "unresolved",
                )

    def test_non_numeric_and_missing_states(self) -> None:
        """코드·날짜·불리언·빈 값·대시·실제 0을 별개 상태로 보존한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 비금액 입력을 숫자로 바꾸거나 빈 값을 0으로 바꾼 경우.
        Side Effects:
            없음.
        """
        protected = [
            ("00123", "string", "code"),
            ("00123", "number", "identifier"),
            ("2027-01-01", "string", "amount"),
            ("A-100", "string", "code"),
            (True, "boolean", "amount"),
            ("#DIV/0!", "error", "amount"),
            ("2027-01-01", "date", "amount"),
            ("100", "string", None),
        ]
        for original, data_type, kind in protected:
            with self.subTest(original=original):
                result = normalize_number(
                    original, data_type=data_type, value_kind=kind
                )
                self.assertEqual(result["status"], "not_numeric")
                self.assertIsNone(result["normalized_value"])
        states = [
            (None, "missing", "MISSING-NULL"),
            ("", "missing", "MISSING-EMPTY-TEXT"),
            ("미입력", "missing", "MISSING-NOT-ENTERED"),
            ("-", "unresolved", "TEXT-DASH-PLACEHOLDER"),
        ]
        for original, state, rule in states:
            result = normalize_number(original, data_type="string", value_kind="amount")
            self.assertEqual(result["status"], state)
            self.assertIn(rule, result["rule_ids"])
        self.assert_value("0", "0", data_type="number", value_kind="amount")
        self.assert_value("0", "0", data_type="string", value_kind="amount")
        result = normalize_number(
            "45292", data_type="number", number_format="yyyy-mm-dd"
        )
        self.assertEqual(result["status"], "not_numeric")

    def test_currency_and_precision(self) -> None:
        """통화 충돌을 표시하고 28자리를 넘는 십진 원문을 보존한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 통화를 추측하거나 Decimal 기본 정밀도로 잘린 경우.
        Side Effects:
            없음.
        """
        self.assert_value("USD 100", "100", data_type="string", value_kind="amount")
        self.assert_value("$100 USD", "100", data_type="string", value_kind="amount")
        for original in ("$100", "¥100", "₩100 USD"):
            self.assertEqual(
                normalize_number(original, data_type="string", value_kind="amount")[
                    "status"
                ],
                "unresolved",
            )
        original = "123456789012345678901234567890.123456789"
        self.assert_value(
            original,
            "123456789012345678901234567890123.456789",
            data_type="number",
            value_kind="amount",
            unit="천원",
        )
        self.assert_value(
            "-" + original,
            "-123456789012345678901234567890.123456789",
            data_type="string",
            value_kind="amount",
        )
        self.assertEqual(
            normalize_number("NaN", data_type="number")["status"], "unresolved"
        )
        self.assertEqual(
            normalize_number("Infinity", data_type="number")["status"], "unresolved"
        )
        self.assertEqual(
            normalize_number("1E-2000000", data_type="number")["status"], "unresolved"
        )
        self.assertEqual(
            normalize_number("₩12.5%", data_type="string", value_kind="ratio")[
                "status"
            ],
            "unresolved",
        )

    def test_format_metadata_is_not_a_calculation_engine(self) -> None:
        """구간·색·조건·표시 배율은 메타데이터로만 반환하는지 확인한다.

        Args:
            없음.
        Returns:
            없음.
        Raises:
            AssertionError: 서식 정보를 잃거나 표시 배율을 값에 적용한 경우.
        Side Effects:
            없음.
        """
        result = inspect_number_format('[Red][>=100]0.00%;[Blue](0);"-";@')
        self.assertEqual(len(result["sections"]), 4)
        self.assertEqual(result["sections"][0]["condition_tokens"], [">=100"])
        self.assertEqual(result["sections"][0]["color_tokens"], ["Red"])
        self.assertTrue(result["percent_display"])
        self.assertFalse(inspect_number_format('0"%"')["percent_display"])
        self.assertEqual(
            inspect_number_format("#,##0,,")["sections"][0]["display_scale_divisor"],
            "1000000",
        )
        self.assertFalse(inspect_number_format("General")["date_or_time_display"])
        self.assertTrue(inspect_number_format("[h]:mm:ss")["date_or_time_display"])


if __name__ == "__main__":
    unittest.main()
