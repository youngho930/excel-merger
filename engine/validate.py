"""검증 5종: 필수값 빈칸 / 형식 오류 / 범위 밖 값 / 허용값 아닌 값 / 중복 행.

한 셀은 빈칸 -> 형식 -> 범위·허용값 순서로 검사하고, 먼저 걸린 오류 하나만 낸다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterable

from .normalize import TEXT, display, format_number, is_blank, parse_value
from .scenario import ColumnSpec

# 오류 종류 (CLAUDE.md "검증 종류"의 이름)
REQUIRED = "필수값 빈칸"
FORMAT = "형식 오류"
RANGE = "범위 밖 값"
ALLOWED = "허용값 아닌 값"
DUPLICATE = "중복 행"
KINDS = (REQUIRED, FORMAT, RANGE, ALLOWED, DUPLICATE)

WHOLE_ROW = "(행 전체)"
WHOLE_FILE = "(전체)"


@dataclass
class Issue:
    file: str
    row: int | None               # 엑셀 행 번호. 파일 전체 오류면 None
    column: str | None            # 원본 열 이름 (중복은 "(행 전체)", 열이 없으면 None)
    standard: str                 # 기준명 (중복은 "품번 + 로트번호")
    kind: str
    value: Any                    # 원래 값
    message: str                  # 설명 (한국어)
    suggestion: Any = None        # 수정 제안값
    dup_group: str | None = None  # "중복-1" …


@dataclass(frozen=True)
class CellResult:
    value: Any                    # 정규화된 값 (빈칸·형식 오류면 None, 범위·허용값 오류는 읽은 값)
    kind: str | None = None
    message: str | None = None
    suggestion: Any = None

    @property
    def ok(self) -> bool:
        return self.kind is None


_SPACES = re.compile(r"\s+")


def _allowed_text(allowed: Iterable[Any]) -> str:
    return "/".join(display(a) for a in allowed)


def _allowed_issue(raw: Any, value: Any, spec: ColumnSpec) -> CellResult:
    allowed = spec.allowed or ()
    shown = display(raw)
    if spec.fmt == TEXT and isinstance(value, str):
        stripped = value.strip()
        if stripped in allowed:
            return CellResult(value, ALLOWED, f'앞뒤 공백 있음 → "{stripped}"으로 수정 제안', stripped)
        squashed = _SPACES.sub("", value)
        for a in allowed:
            if _SPACES.sub("", str(a)) == squashed:
                return CellResult(value, ALLOWED, f'공백 차이 → "{a}"일 가능성', a)
    return CellResult(value, ALLOWED, f"허용값({_allowed_text(allowed)})이 아닌 '{shown}'")


def validate_cell(raw: Any, spec: ColumnSpec) -> CellResult:
    """셀 하나를 검사한다. 오류가 없으면 정규화된 값을 돌려준다."""
    if is_blank(raw):
        if spec.required:
            return CellResult(None, REQUIRED, f"필수 열 '{spec.name}'이(가) 비어 있음")
        return CellResult(None)

    parsed = parse_value(raw, spec.fmt)
    if not parsed.ok:
        msg = parsed.reason or "형식에 맞지 않는 값"
        sug = parsed.suggestion
        if sug is not None:
            msg += f" → {display(sug)}으로 수정 제안"
        return CellResult(None, FORMAT, msg, sug)
    value = parsed.value

    if spec.range is not None:
        lo, hi = spec.range
        if not (lo <= value <= hi):
            return CellResult(value, RANGE,
                              f"허용 범위({format_number(lo)}~{format_number(hi)})를 벗어남: {display(raw)}")

    if spec.allowed is not None and value not in spec.allowed:
        return _allowed_issue(raw, value, spec)

    return CellResult(value)


def dup_key_part(value: Any, spec: ColumnSpec) -> Any:
    """중복 판정용 값. 문자는 앞뒤 공백만 무시한다."""
    if spec.fmt == TEXT and isinstance(value, str):
        return value.strip()
    return value


def dup_key_text(names: list[str], values: list[Any]) -> str:
    return ", ".join(f"{n}={display(v)}" for n, v in zip(names, values))
