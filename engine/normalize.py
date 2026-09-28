"""열 이름과 셀 값을 정규화한다 (순수 함수만 둔다).

- normalize_header: 열 이름 비교용 (공백·대소문자·특수문자 무시)
- is_blank: 빈칸 판정
- parse_value: 시나리오 형식(문자/정수/실수/날짜)에 맞춰 값을 읽는다.
  읽을 수 없으면 ok=False 와 한국어 사유, 가능하면 수정 제안값을 돌려준다.
"""

from __future__ import annotations

import math
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from typing import Any

# 형식 이름 (CLAUDE.md "시나리오 파일 형식"의 어휘)
TEXT = "문자"
INT = "정수"
REAL = "실수"
DATE = "날짜"
FORMATS = (TEXT, INT, REAL, DATE)

# 날짜로 받아들이는 연도 범위
MIN_YEAR, MAX_YEAR = 1900, 2100
# 날짜 칸에 숫자가 들어왔을 때 "엑셀 날짜 일련번호"로 보고 수정 제안할 연도 범위
SERIAL_SUGGEST_MIN_YEAR, SERIAL_SUGGEST_MAX_YEAR = 1980, 2100
EXCEL_EPOCH = date(1899, 12, 30)


@dataclass(frozen=True)
class Parsed:
    ok: bool
    value: Any = None          # 정규화된 값 (ok일 때)
    reason: str | None = None  # 형식 오류 사유 (ok가 아닐 때)
    suggestion: Any = None     # 수정 제안값 (있을 때만)


# ---------------------------------------------------------------- 이름
_NON_WORD = re.compile(r"[\W_]+", re.UNICODE)


def normalize_header(name: Any) -> str:
    """열 이름 비교용 키. 'Insp. Date' -> 'inspdate', 'P/N' -> 'pn', '입고 수량' -> '입고수량'."""
    if name is None:
        return ""
    s = unicodedata.normalize("NFKC", str(name)).casefold()
    return _NON_WORD.sub("", s)


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


# ---------------------------------------------------------------- 값 표시
def display(value: Any) -> str:
    """오류 설명에 넣을 짧은 표기."""
    if value is None:
        return "(빈칸)"
    if isinstance(value, datetime):
        if value.time() == datetime.min.time():
            return value.date().isoformat()
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def format_number(n: Any) -> str:
    if isinstance(n, float) and n.is_integer():
        n = int(n)
    return str(n)


# ---------------------------------------------------------------- 숫자
_NUMBER_RE = re.compile(r"^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)?(?:\.\d+)?$")
_NUMBER_CHARS = set("0123456789+-,. ")


def _number_from_text(s: str) -> Decimal | None:
    s = s.strip()
    if not s or not _NUMBER_RE.match(s) or not any(ch.isdigit() for ch in s):
        return None
    try:
        return Decimal(s.replace(",", ""))
    except InvalidOperation:
        return None


def _foreign_chars(s: str) -> str:
    """숫자 표기에 쓰이지 않는 문자만 순서대로 모은다. '약 120' -> '약', '20EA' -> 'EA'."""
    seen: list[str] = []
    for ch in s.strip():
        if ch not in _NUMBER_CHARS:
            seen.append(ch)
    return "".join(seen)


def _text_number_reason(s: str, fmt: str) -> str:
    foreign = _foreign_chars(s)
    if foreign:
        return f"{fmt} 칸에 문자('{foreign}')가 섞임"
    return f"{fmt}로 읽을 수 없는 값"


def parse_int(raw: Any) -> Parsed:
    if isinstance(raw, bool):
        return Parsed(False, reason="정수 칸에 참/거짓 값이 있음")
    if isinstance(raw, int):
        return Parsed(True, raw)
    if isinstance(raw, float):
        if not math.isfinite(raw):
            return Parsed(False, reason="정수로 읽을 수 없는 값")
        if raw.is_integer():
            return Parsed(True, int(raw))
        return Parsed(False, reason=f"정수 칸에 소수({format_number(raw)})가 있음")
    if isinstance(raw, (datetime, date)):
        return Parsed(False, reason="정수 칸에 날짜가 있음")
    s = str(raw)
    d = _number_from_text(s)
    if d is None:
        return Parsed(False, reason=_text_number_reason(s, INT))
    if d != d.to_integral_value():
        return Parsed(False, reason=f"정수 칸에 소수({s.strip()})가 있음")
    return Parsed(True, int(d))


def parse_real(raw: Any) -> Parsed:
    if isinstance(raw, bool):
        return Parsed(False, reason="실수 칸에 참/거짓 값이 있음")
    if isinstance(raw, (int, float)):
        if isinstance(raw, float) and not math.isfinite(raw):
            return Parsed(False, reason="실수로 읽을 수 없는 값")
        return Parsed(True, raw)
    if isinstance(raw, (datetime, date)):
        return Parsed(False, reason="실수 칸에 날짜가 있음")
    s = str(raw)
    d = _number_from_text(s)
    if d is None:
        return Parsed(False, reason=_text_number_reason(s, REAL))
    value = int(d) if d == d.to_integral_value() and "." not in s else float(d)
    return Parsed(True, value)


# ---------------------------------------------------------------- 날짜
_TIME = r"(?:[ T]+\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?"
_YMD_SEP = re.compile(
    r"^(?P<y>\d{4})\s*(?P<sep>[-./])\s*(?P<m>\d{1,2})\s*(?P=sep)\s*(?P<d>\d{1,2})\s*\.?" + _TIME + r"$")
_YMD_KO = re.compile(r"^(?P<y>\d{4})\s*년\s*(?P<m>\d{1,2})\s*월\s*(?P<d>\d{1,2})\s*일$")
_YMD_COMPACT = re.compile(r"^(?P<y>\d{4})(?P<m>\d{2})(?P<d>\d{2})$")
_YM_SEP = re.compile(r"^(?P<y>\d{4})\s*[-./]\s*(?P<m>\d{1,2})\s*\.?$")
_YM_KO = re.compile(r"^(?P<y>\d{4})\s*년\s*(?P<m>\d{1,2})\s*월$")

_DAY_PATTERNS = (_YMD_SEP, _YMD_KO, _YMD_COMPACT)
_MONTH_PATTERNS = (_YM_SEP, _YM_KO)


def _make_date(y: int, m: int, d: int) -> Parsed:
    if not (MIN_YEAR <= y <= MAX_YEAR):
        return Parsed(False, reason=f"연도가 허용 범위({MIN_YEAR}~{MAX_YEAR})를 벗어남({y}년)")
    if not (1 <= m <= 12):
        return Parsed(False, reason=f"존재하지 않는 날짜({m}월)")
    try:
        return Parsed(True, date(y, m, d))
    except ValueError:
        return Parsed(False, reason=f"존재하지 않는 날짜({m}월 {d}일)")


def _date_from_number(n: float) -> date | None:
    """숫자가 8자리 날짜(20260803)나 엑셀 날짜 일련번호로 해석되면 그 날짜."""
    if not math.isfinite(n) or n <= 0:
        return None
    if n.is_integer() and 10_000_000 <= n <= 99_999_999:
        s = str(int(n))
        p = _make_date(int(s[:4]), int(s[4:6]), int(s[6:]))
        return p.value if p.ok else None
    try:
        d = EXCEL_EPOCH + timedelta(days=int(n))
    except OverflowError:
        return None
    if SERIAL_SUGGEST_MIN_YEAR <= d.year <= SERIAL_SUGGEST_MAX_YEAR:
        return d
    return None


def parse_date(raw: Any) -> Parsed:
    if isinstance(raw, datetime):
        return Parsed(True, raw.date())
    if isinstance(raw, date):
        return Parsed(True, raw)
    if isinstance(raw, bool):
        return Parsed(False, reason="날짜 칸에 참/거짓 값이 있음")
    if isinstance(raw, (int, float)):
        guess = _date_from_number(float(raw))
        if guess is not None:
            return Parsed(False, reason=f"날짜 칸에 숫자({format_number(raw)})가 있음",
                          suggestion=guess)
        return Parsed(False, reason=f"날짜 칸에 숫자({format_number(raw)})가 있음")
    s = str(raw).strip()
    for pat in _DAY_PATTERNS:
        m = pat.match(s)
        if m:
            return _make_date(int(m["y"]), int(m["m"]), int(m["d"]))
    for pat in _MONTH_PATTERNS:
        m = pat.match(s)
        if m:
            return _make_date(int(m["y"]), int(m["m"]), 1)
    return Parsed(False, reason="날짜로 읽을 수 없는 표기")


# ---------------------------------------------------------------- 문자
def to_text(raw: Any) -> str:
    """문자 칸 값. 공백은 건드리지 않는다 (허용값은 원래 값 그대로 엄격히 비교)."""
    if isinstance(raw, str):
        return raw
    return display(raw)


def parse_value(raw: Any, fmt: str) -> Parsed:
    """빈칸이 아닌 값을 형식에 맞춰 읽는다. 빈칸 처리는 호출하는 쪽에서 먼저 한다."""
    if fmt == TEXT:
        return Parsed(True, to_text(raw))
    if fmt == INT:
        return parse_int(raw)
    if fmt == REAL:
        return parse_real(raw)
    if fmt == DATE:
        return parse_date(raw)
    raise ValueError(f"알 수 없는 형식: {fmt}")
