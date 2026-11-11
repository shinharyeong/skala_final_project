"""저장 숫자와 문자열 숫자를 구분하여 정밀한 십진 값으로 정규화한다.

표시 서식은 해석 근거로 보존하지만 저장 숫자의 부호와 값을 다시 계산하지 않는다.
모든 함수는 메모리에서만 동작하며 원본·결과 파일을 읽거나 쓰지 않는다.
"""

import re
import unicodedata
from decimal import (
    Decimal,
    DecimalException,
    Inexact,
    InvalidOperation,
    Underflow,
    localcontext,
)

NUMERIC_KINDS = frozenset({"amount", "ratio", "quantity", "number"})
PROTECTED_KINDS = frozenset(
    {
        "code",
        "identifier",
        "account_code",
        "organization_code",
        "date",
        "datetime",
        "boolean",
        "error",
        "text",
        "title",
        "version",
    }
)
SYMBOLS = frozenset("△▲▽▼↑↓↗↘↔↕→←⇧⇩")
UNIT_DEFINITIONS = {
    "원": ("KRW", "1", "amount"),
    "천원": ("KRW", "1000", "amount"),
    "만원": ("KRW", "10000", "amount"),
    "백만원": ("KRW", "1000000", "amount"),
    "억원": ("KRW", "100000000", "amount"),
    "KRW": ("KRW", "1", "amount"),
    "USD": ("USD", "1", "amount"),
    "EUR": ("EUR", "1", "amount"),
    "GBP": ("GBP", "1", "amount"),
    "JPY": ("JPY", "1", "amount"),
    "CNY": ("CNY", "1", "amount"),
    "CAD": ("CAD", "1", "amount"),
    "AUD": ("AUD", "1", "amount"),
    "NZD": ("NZD", "1", "amount"),
    "달러": ("USD", "1", "amount"),
    "천달러": ("USD", "1000", "amount"),
    "백만달러": ("USD", "1000000", "amount"),
    "%": (None, "0.01", "ratio"),
    "퍼센트": (None, "0.01", "ratio"),
    "명": (None, "1", "quantity"),
    "천명": (None, "1000", "quantity"),
    "만명": (None, "10000", "quantity"),
    "개": (None, "1", "quantity"),
    "건": (None, "1", "quantity"),
    "대": (None, "1", "quantity"),
}
CURRENCY_SYMBOLS = {
    "₩": "KRW",
    "€": "EUR",
    "£": "GBP",
    "$": None,
    "¥": None,
    "KRW": "KRW",
    "USD": "USD",
    "EUR": "EUR",
    "GBP": "GBP",
    "JPY": "JPY",
    "CNY": "CNY",
    "CAD": "CAD",
    "AUD": "AUD",
    "NZD": "NZD",
}
AFFIX_PATTERN = re.compile(
    r"백만\s*달러|천\s*달러|백만\s*원|천\s*원|만\s*원|억\s*원|"
    r"천\s*명|만\s*명|퍼센트|KRW|USD|EUR|GBP|JPY|CNY|CAD|AUD|NZD|"
    r"달러|원|명|개|건|대|[%₩€£$¥]"
)
LOCALES = {
    "ko_KR": (".", ","),
    "ko": (".", ","),
    "en_US": (".", ","),
    "en": (".", ","),
    "de_DE": (",", "."),
    "de": (",", "."),
    "fr_FR": (",", " "),
    "fr": (",", " "),
}


def inspect_number_format(number_format: str | None) -> dict:
    """숫자 표시 형식의 구간·조건·색·표시 배율을 값 변경 없이 요약한다.

    Args:
        number_format: 원본 Excel 표시 형식. 없으면 None.

    Returns:
        원문, 구간별 메타데이터, 백분율·날짜 표시 여부를 담은 사전.

    Raises:
        없음. 알 수 없는 형식은 원문을 보존하며 완전한 렌더링을 시도하지 않는다.

    Side Effects:
        없음. 파일을 읽거나 쓰지 않는다.
    """
    raw = number_format or ""
    sections = []
    current = ""
    quoted = False
    escaped = False
    for character in raw:
        if character == ";" and not quoted and not escaped:
            sections.append(current)
            current = ""
            continue
        current += character
        if escaped:
            escaped = False
        elif character == "\\":
            escaped = True
        elif character == '"':
            quoted = not quoted
    sections.append(current)
    details = []
    for index, section in enumerate(sections):
        mask = re.sub(r'"[^\"]*"|\\.|_.|\*.', "", section)
        brackets = re.findall(r"\[([^]]+)\]", mask)
        conditions = [value for value in brackets if re.match(r"[<>=]", value)]
        colors = [
            value
            for value in brackets
            if re.fullmatch(
                r"(?i:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\d+)", value
            )
        ]
        clean = re.sub(r"\[[^]]+\]", "", mask)
        scale_match = re.search(r"[0#?](,+)(?:%|\s)*$", clean)
        scale_commas = len(scale_match.group(1)) if scale_match else 0
        date_tokens = bool(re.search(r"(?i)[ydh]|s{1,2}|m{1,5}", clean))
        elapsed = any(re.fullmatch(r"(?i)[hms]+", value) for value in brackets)
        details.append(
            {
                "index": index,
                "raw": section,
                "condition_tokens": conditions,
                "color_tokens": colors,
                "percent_display": "%" in clean,
                "display_scale_divisor": str(1000**scale_commas),
                "date_or_time_display": date_tokens or elapsed,
            }
        )
    return {
        "raw": number_format,
        "sections": details,
        "percent_display": any(item["percent_display"] for item in details),
        "date_or_time_display": any(item["date_or_time_display"] for item in details),
        "interpretation": "metadata_only_not_rendered_or_applied",
    }


def _base_result(
    value: object, value_kind: str | None, unit: str | None, number_format: str | None
) -> dict:
    """정규화 상태를 담을 독립된 결과 객체를 만든다.

    Args:
        value: 원본 저장값 또는 문자열.
        value_kind: 호출자가 확인한 필드 종류.
        unit: 호출자가 확인한 원본 단위.
        number_format: 원본 표시 형식.

    Returns:
        정규화값이 없는 초기 결과 사전.

    Raises:
        없음.

    Side Effects:
        없음. 입력 객체를 변경하지 않는다.
    """
    return {
        "original_value": value,
        "normalized_value": None,
        "status": "unresolved",
        "value_kind": value_kind,
        "currency": None,
        "source_unit": unit,
        "multiplier": Decimal("1"),
        "rule_ids": [],
        "diagnostics": [],
        "format_analysis": inspect_number_format(number_format),
    }


def _stop(result: dict, status: str, code: str, message: str) -> dict:
    """값을 확정하지 못한 이유와 상태를 결과에 기록한다.

    Args:
        result: 현재 호출에서 생성한 결과 사전.
        status: missing, not_numeric, unresolved 중 하나.
        code: 기계 판별에 사용하는 규칙·진단 코드.
        message: 사람이 읽을 수 있는 미확정 사유.

    Returns:
        상태와 진단이 추가된 같은 결과 객체.

    Raises:
        없음.

    Side Effects:
        해당 결과 사전만 변경하며 원본 입력이나 파일은 변경하지 않는다.
    """
    result["status"] = status
    result["rule_ids"].append(code)
    result["diagnostics"].append({"code": code, "message": message})
    return result


def _unit_info(unit: str | None) -> tuple[str | None, Decimal, str | None] | None:
    """등록된 단위의 통화·배율·값 종류를 찾는다.

    Args:
        unit: 단위 문자열. None은 배율이 지정되지 않았음을 뜻한다.

    Returns:
        통화, Decimal 배율, 값 종류의 튜플. 미등록 단위이면 None.

    Raises:
        없음.

    Side Effects:
        없음.
    """
    if unit is None:
        return None, Decimal("1"), None
    key = re.sub(r"\s+", "", unicodedata.normalize("NFKC", unit))
    definition = UNIT_DEFINITIONS.get(key)
    if definition is None:
        return None
    currency, multiplier, kind = definition
    return currency, Decimal(multiplier), kind


def _remove_affixes(text: str) -> tuple[str, list[str], list[str], str | None]:
    """숫자 바깥의 등록된 통화·단위 표기만 분리한다.

    Args:
        text: 유니코드와 공백을 정규화한 전체 문자열.

    Returns:
        숫자 부분, 단위 목록, 통화 기호 목록, 오류 코드의 튜플.

    Raises:
        없음.

    Side Effects:
        없음.
    """
    digits = [match.start() for match in re.finditer(r"\d", text)]
    units = []
    currencies = []
    for match in AFFIX_PATTERN.finditer(text):
        if digits and digits[0] < match.start() < digits[-1]:
            return text, [], [], "TEXT-EMBEDDED-UNIT"
        token = re.sub(r"\s+", "", match.group())
        if token in CURRENCY_SYMBOLS:
            currencies.append(token)
        else:
            units.append(token)
    return AFFIX_PATTERN.sub("", text).strip(), units, currencies, None


def _parse_decimal_token(token: str, locale: str | None) -> Decimal | None:
    """부호 없는 숫자 토큰을 로케일과 전체 구분자 패턴으로 해석한다.

    Args:
        token: 부호와 단위가 제거된 숫자 문자열. 지수부의 부호는 허용한다.
        locale: 등록된 로케일 또는 None. None에서는 모호한 단일 구분자를 거부한다.

    Returns:
        유효한 유한 Decimal. 구분자나 전체 형태가 모호하면 None.

    Raises:
        없음. Decimal 변환 실패는 None으로 반환한다.

    Side Effects:
        없음.
    """
    match = re.fullmatch(r"([\d., ]+)([eE][+-]?\d+)?", token)
    if not match:
        return None
    mantissa, exponent = match.group(1), match.group(2) or ""
    if locale is not None:
        separators = LOCALES.get(locale)
        if separators is None:
            return None
        decimal_separator, grouping_separator = separators
    elif "," in mantissa and "." in mantissa:
        decimal_separator = "," if mantissa.rfind(",") > mantissa.rfind(".") else "."
        grouping_separator = "." if decimal_separator == "," else ","
    elif " " in mantissa:
        decimal_separator = "," if "," in mantissa else "."
        grouping_separator = " "
    else:
        separator = "," if "," in mantissa else "."
        if mantissa.count(separator) > 1:
            decimal_separator = "." if separator == "," else ","
            grouping_separator = separator
        else:
            pieces = mantissa.split(separator)
            if len(pieces) == 2 and 1 <= len(pieces[0]) <= 3 and len(pieces[1]) == 3:
                return None
            decimal_separator = separator
            grouping_separator = "," if separator == "." else "."
    if mantissa.count(decimal_separator) > 1:
        return None
    pieces = mantissa.split(decimal_separator)
    integer = pieces[0]
    fraction = pieces[1] if len(pieces) == 2 else None
    if fraction is not None and not fraction.isdigit():
        return None
    if grouping_separator in integer:
        groups = integer.split(grouping_separator)
        if not (groups[0].isdigit() and 1 <= len(groups[0]) <= 3):
            return None
        if not all(group.isdigit() and len(group) == 3 for group in groups[1:]):
            return None
        integer = "".join(groups)
    if not integer and fraction is not None:
        integer = "0"
    if not integer.isdigit():
        return None
    canonical = integer + ("." + fraction if fraction is not None else "") + exponent
    try:
        value = Decimal(canonical)
    except InvalidOperation:
        return None
    return value if value.is_finite() else None


def normalize_number(
    value: object,
    *,
    data_type: str,
    value_kind: str | None = None,
    unit: str | None = None,
    number_format: str | None = None,
    locale: str | None = None,
    symbol_rules: dict | None = None,
) -> dict:
    """원본 자료형을 우선하여 숫자·문자열 부호·단위를 한 번만 정규화한다.

    Args:
        value: 원본 저장값. 숫자 셀의 XML 숫자 문자열도 허용한다.
        data_type: number/string/boolean/error/date/blank 등 원본 자료형.
        value_kind: amount/ratio/quantity/number 또는 보호할 code/date/text 등.
        unit: 원본 제목·헤더 등에서 확인한 단위. 표시 배율과 별개이다.
        number_format: 원본 숫자 표시 형식. 저장 숫자의 부호를 바꾸지 않는다.
        locale: 문자열 숫자의 확인된 로케일. 미지정 모호 표기는 거부한다.
        symbol_rules: 확인된 기호별 부호 규칙. 값은 1/-1 또는 sign을 가진 사전.

    Returns:
        원본·Decimal 정규화값·상태·단위·배율·규칙·진단을 담은 사전.
        confirmed만 정규화값을 가진다. JSON 직렬화 시 Decimal은 문자열로 쓴다.

    Raises:
        없음. 지원하지 않는 입력·표기는 미확정 진단으로 반환한다.

    Side Effects:
        없음. 파일·DB·콘솔을 사용하거나 원본 객체를 변경하지 않는다.
    """
    result = _base_result(value, value_kind, unit, number_format)
    if (
        value_kind in PROTECTED_KINDS
        or data_type in {"boolean", "bool", "b", "error", "e", "date", "d"}
        or isinstance(value, bool)
    ):
        return _stop(
            result,
            "not_numeric",
            "TYPE-PROTECTED",
            "코드·날짜·불리언·오류·설명 필드는 숫자로 바꾸지 않습니다.",
        )
    if value is None:
        return _stop(
            result,
            "missing",
            "MISSING-NULL",
            "원본 값이 없습니다. 0으로 대체하지 않습니다.",
        )
    if isinstance(value, str) and not value.strip():
        return _stop(
            result, "missing", "MISSING-EMPTY-TEXT", "빈 문자열을 실제 0과 구분합니다."
        )
    if isinstance(value, str) and value.strip() == "미입력":
        return _stop(
            result,
            "missing",
            "MISSING-NOT-ENTERED",
            "명시적인 미입력 표기를 보존합니다.",
        )
    if result["format_analysis"]["date_or_time_display"]:
        return _stop(
            result,
            "not_numeric",
            "TYPE-DATE-FORMAT",
            "날짜·시간 표시 형식은 금액으로 정규화하지 않습니다.",
        )
    explicit_unit = _unit_info(unit)
    if explicit_unit is None:
        return _stop(
            result,
            "unresolved",
            "UNIT-UNKNOWN",
            "등록되지 않은 단위의 배율은 추정하지 않습니다.",
        )
    currency, multiplier, unit_kind = explicit_unit
    stored_number = data_type in {"number", "numeric", "n"}
    if stored_number:
        try:
            number = Decimal(str(value))
        except (InvalidOperation, ValueError):
            return _stop(
                result,
                "unresolved",
                "RAW-INVALID-NUMBER",
                "저장 숫자를 유한 십진수로 해석할 수 없습니다.",
            )
        if not number.is_finite():
            return _stop(
                result,
                "unresolved",
                "RAW-NON-FINITE",
                "비유한 숫자는 표준 금액으로 확정하지 않습니다.",
            )
        result["rule_ids"].append("RAW-NUMBER-001")
        if number_format:
            result["rule_ids"].append("RAW-FORMAT-001")
        if unit_kind == "ratio":
            multiplier = Decimal("1")
            result["rule_ids"].append("RAW-PERCENT-001")
        if isinstance(value, float):
            result["rule_ids"].append("RAW-FLOAT-REPRESENTATION")
    else:
        if not isinstance(value, str):
            return _stop(
                result,
                "not_numeric",
                "TYPE-UNSUPPORTED",
                "지원하지 않는 원본 자료형입니다.",
            )
        if value_kind not in NUMERIC_KINDS:
            return _stop(
                result,
                "not_numeric",
                "TEXT-CONTEXT-REQUIRED",
                "문자열 숫자는 확인된 숫자 필드에서만 해석합니다.",
            )
        text = unicodedata.normalize("NFKC", value).replace("−", "-").strip()
        if text != value.strip():
            result["rule_ids"].append("TEXT-UNICODE-001")
        if text in {"-", "—", "–"}:
            return _stop(
                result,
                "unresolved",
                "TEXT-DASH-PLACEHOLDER",
                "대시 표기는 미입력이나 0으로 확정하지 않습니다.",
            )
        if re.fullmatch(r"\d{2,4}[-/]\d{1,2}(?:[-/]\d{1,2})?", text):
            return _stop(
                result,
                "not_numeric",
                "TEXT-DATE-LIKE",
                "날짜처럼 보이는 문자열은 금액으로 바꾸지 않습니다.",
            )
        text, units, currency_symbols, error = _remove_affixes(text)
        if error:
            return _stop(
                result,
                "unresolved",
                error,
                "숫자 중간의 단위 표기는 지원하지 않습니다.",
            )
        embedded = [_unit_info(token) for token in units]
        if embedded:
            first_unit = embedded[0]
            if any(info != first_unit for info in embedded[1:]):
                return _stop(
                    result,
                    "unresolved",
                    "UNIT-CONFLICT",
                    "문자열에 서로 다른 단위가 함께 있습니다.",
                )
            if unit is not None and explicit_unit != first_unit:
                return _stop(
                    result,
                    "unresolved",
                    "UNIT-CONFLICT",
                    "헤더 단위와 문자열 단위가 일치하지 않습니다.",
                )
            currency, multiplier, unit_kind = first_unit
            result["source_unit"] = unit or units[0]
        declared_currencies = {
            CURRENCY_SYMBOLS[symbol]
            for symbol in currency_symbols
            if CURRENCY_SYMBOLS[symbol] is not None
        }
        if currency is not None:
            declared_currencies.add(currency)
        if len(declared_currencies) > 1:
            return _stop(
                result,
                "unresolved",
                "CURRENCY-CONFLICT",
                "서로 다른 통화 표기가 함께 있습니다.",
            )
        if declared_currencies:
            currency = next(iter(declared_currencies))
        for symbol in currency_symbols:
            if CURRENCY_SYMBOLS[symbol] is None:
                candidates = (
                    {"USD", "CAD", "AUD", "NZD"} if symbol == "$" else {"JPY", "CNY"}
                )
                if currency not in candidates:
                    return _stop(
                        result,
                        "unresolved",
                        "CURRENCY-AMBIGUOUS",
                        "통화 기호만으로 통화를 확정하지 않습니다.",
                    )
        parenthesized = text.startswith("(") and text.endswith(")")
        sign_markers = []
        if parenthesized:
            if value_kind not in {"amount", "ratio"}:
                return _stop(
                    result,
                    "unresolved",
                    "TEXT-PARENTHESES-CONTEXT",
                    "괄호 음수는 확인된 금액·비율 필드에만 적용합니다.",
                )
            text = text[1:-1].strip()
            sign_markers.append(-1)
            result["rule_ids"].append("TEXT-ACCOUNTING-001")
        markers = [character for character in text if character in SYMBOLS]
        for symbol in markers:
            position = text.index(symbol)
            if re.search(r"\d", text[:position]) and re.search(
                r"\d", text[position + 1 :]
            ):
                return _stop(
                    result,
                    "unresolved",
                    "TEXT-SYMBOL-POSITION",
                    "숫자 중간의 기호를 제거해 값을 이어 붙이지 않습니다.",
                )
            definition = (symbol_rules or {}).get(symbol)
            sign = (
                definition.get("sign") if isinstance(definition, dict) else definition
            )
            if sign not in (-1, 1) or isinstance(sign, bool):
                return _stop(
                    result,
                    "unresolved",
                    "TEXT-SYMBOL-UNCONFIRMED",
                    "삼각형·화살표의 부호 의미가 확인되지 않았습니다.",
                )
            sign_markers.append(sign)
            text = text.replace(symbol, "", 1).strip()
            result["rule_ids"].append("TEXT-SYMBOL-001")
            if isinstance(definition, dict) and definition.get("rule_id"):
                result["rule_ids"].append(str(definition["rule_id"]))
        if text.startswith(("+", "-")):
            sign_markers.append(-1 if text[0] == "-" else 1)
            text = text[1:].strip()
            result["rule_ids"].append("TEXT-SIGN-LEADING")
        if text.endswith(("+", "-")):
            sign_markers.append(-1 if text[-1] == "-" else 1)
            text = text[:-1].strip()
            result["rule_ids"].append("TEXT-SIGN-TRAILING")
        if len(sign_markers) > 1:
            return _stop(
                result,
                "unresolved",
                "TEXT-SIGN-CONFLICT",
                "중복·충돌 부호를 곱셈으로 처리하지 않습니다.",
            )
        number = _parse_decimal_token(text, locale)
        if number is None:
            return _stop(
                result,
                "unresolved",
                "TEXT-NUMBER-AMBIGUOUS",
                "전체 숫자 형태나 소수·천 단위 구분을 확정할 수 없습니다.",
            )
        if sign_markers:
            if sign_markers[0] == -1:
                number = number.copy_negate()
        result["rule_ids"].append("TEXT-NUMBER-001")
        if locale is not None:
            result["rule_ids"].append("TEXT-LOCALE-001")
        if unit_kind == "ratio":
            result["rule_ids"].append("TEXT-PERCENT-001")
    if (
        unit_kind
        and value_kind in NUMERIC_KINDS - {"number"}
        and value_kind != unit_kind
    ):
        return _stop(
            result,
            "unresolved",
            "UNIT-KIND-CONFLICT",
            "필드 종류와 단위의 차원이 다릅니다.",
        )
    if currency is not None and value_kind in {"ratio", "quantity"}:
        return _stop(
            result,
            "unresolved",
            "CURRENCY-KIND-CONFLICT",
            "비율·수량 필드에 금액 통화가 함께 있습니다.",
        )
    result["currency"] = currency
    result["multiplier"] = multiplier
    result["value_kind"] = (
        value_kind
        if value_kind not in {None, "number"}
        else unit_kind or value_kind or "number"
    )
    if unit is not None or result["source_unit"] is not None:
        result["rule_ids"].append("UNIT-001")
    try:
        with localcontext() as context:
            context.prec = max(
                50,
                len(number.as_tuple().digits) + len(multiplier.as_tuple().digits) + 10,
            )
            context.traps[Inexact] = True
            context.traps[Underflow] = True
            result["normalized_value"] = number * multiplier
    except DecimalException:
        return _stop(
            result,
            "unresolved",
            "DECIMAL-RANGE",
            "십진수 연산의 지원 범위를 벗어난 값입니다.",
        )
    result["status"] = "confirmed"
    return result
