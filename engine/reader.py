"""엑셀 파일을 읽고 머리글 행을 자동으로 찾는다.

원본 파일은 읽기 전용으로만 연다 (절대 저장하지 않는다).
행 번호는 사용자가 엑셀 화면에서 보는 번호(1부터)를 그대로 쓴다.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .normalize import is_blank, normalize_header
from .scenario import Scenario

HEADER_SCAN_ROWS = 20           # 머리글을 찾을 때 위에서부터 훑는 행 수
MAX_FILE_BYTES = 20 * 1024 * 1024
MAX_ROWS = 100_000              # 파일 하나에서 읽는 최대 행 수
MAX_COLS = 200
EXCEL_SUFFIXES = (".xlsx", ".xlsm")


class ReadError(Exception):
    """엑셀을 읽을 수 없을 때. 메시지는 한국어."""


@dataclass
class SourceRow:
    excel_row: int              # 엑셀 화면의 행 번호
    cells: tuple[Any, ...]      # 열 위치(0부터) 순서의 원래 값


@dataclass
class SourceTable:
    file_name: str
    path: Path
    sheet_name: str
    header_row: int                     # 머리글 행 번호 (엑셀 화면 기준)
    headers: list[str]                  # 표시용 열 이름 (빈 머리글은 "(빈 머리글 C열)")
    raw_headers: list[Any]              # 원래 머리글 값 (매칭에 사용)
    rows: list[SourceRow] = field(default_factory=list)
    header_confidence: str = "높음"     # "높음" | "낮음" | "지정"

    def column_label(self, index: int) -> str:
        """'C열 'Part No'' 같은 표시."""
        return f"{get_column_letter(index + 1)}열 '{self.headers[index]}'"


@dataclass(frozen=True)
class HeaderGuess:
    row: int            # 1부터
    score: int          # 기준명·동의어와 일치한 서로 다른 기준열 수
    confident: bool


def header_vocabulary(scenario: Scenario) -> dict[str, str]:
    """정규화한 이름 -> 기준명."""
    vocab: dict[str, str] = {}
    for col in scenario.columns:
        for n in col.names:
            vocab[normalize_header(n)] = col.name
    return vocab


def detect_header_row(rows: list[tuple[Any, ...]], scenario: Scenario) -> HeaderGuess:
    """위쪽 행들 중 기준명·동의어와 가장 많이 일치하는 행을 머리글로 본다."""
    vocab = header_vocabulary(scenario)
    threshold = min(2, len(scenario.columns))
    best_row, best_score = 0, -1
    for i, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        matched = {vocab[k] for k in (normalize_header(v) for v in row if not is_blank(v)) if k in vocab}
        if len(matched) > best_score:
            best_row, best_score = i, len(matched)
    if best_score >= threshold:
        return HeaderGuess(best_row + 1, best_score, True)
    # 찾지 못함: 채워진 칸이 가장 많은 행 (동점이면 위쪽)
    fallback, most = 0, -1
    for i, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        n = sum(1 for v in row if not is_blank(v))
        if n > most:
            fallback, most = i, n
    return HeaderGuess(fallback + 1, max(best_score, 0), False)


def _open_workbook(path: Path):
    if path.suffix.lower() not in EXCEL_SUFFIXES:
        raise ReadError(f"'{path.name}'은 엑셀(.xlsx) 파일이 아닙니다. .xls 파일은 엑셀에서 .xlsx로 다시 저장해 주세요.")
    if not path.is_file():
        raise ReadError(f"'{path.name}' 파일을 찾을 수 없습니다.")
    if path.stat().st_size > MAX_FILE_BYTES:
        raise ReadError(f"'{path.name}'이 너무 큽니다 (최대 {MAX_FILE_BYTES // (1024 * 1024)}MB).")
    try:
        # read_only: 원본을 수정할 수 없고 메모리도 적게 쓴다. data_only: 수식 대신 저장된 계산값.
        return load_workbook(path, read_only=True, data_only=True)
    except (zipfile.BadZipFile, KeyError, OSError, ValueError, TypeError):
        raise ReadError(f"'{path.name}'을 엑셀 파일로 열 수 없습니다. 파일이 손상되었거나 암호가 걸려 있을 수 있습니다.") from None
    except Exception:  # openpyxl 내부 예외 종류가 다양하다
        raise ReadError(f"'{path.name}'을 엑셀 파일로 열 수 없습니다.") from None


def _sheet_rows(ws) -> list[tuple[Any, ...]]:
    """시트의 모든 행 값을 1행부터 빈 행 포함, 순서대로 읽는다."""
    ws.reset_dimensions()  # 파일에 적힌 크기 정보를 믿지 않고 실제 셀을 읽는다
    rows: list[tuple[Any, ...]] = []
    for row in ws.iter_rows():
        cells = [c for c in row if getattr(c, "row", None) is not None]
        if not cells:
            continue
        r = cells[0].row
        if r > MAX_ROWS + HEADER_SCAN_ROWS:
            raise ReadError(f"'{ws.title}' 시트의 행이 너무 많습니다 (최대 {MAX_ROWS:,}행).")
        while len(rows) < r - 1:
            rows.append(())
        values: dict[int, Any] = {}
        for c in cells:
            if c.column > MAX_COLS:
                if c.value is not None:
                    raise ReadError(f"'{ws.title}' 시트의 열이 너무 많습니다 (최대 {MAX_COLS}열).")
                continue
            values[c.column] = c.value
        width = max(values) if values else 0
        rows.append(tuple(values.get(i) for i in range(1, width + 1)))
    return rows


def read_table(path: str | Path, scenario: Scenario, header_row: int | None = None,
               sheet: str | None = None) -> SourceTable:
    """엑셀 파일 하나를 읽는다.

    header_row를 주면 그 행을 머리글로 쓰고, 없으면 자동으로 찾는다.
    sheet를 주지 않으면 머리글 점수가 가장 높은 시트를 쓴다(동점이면 앞 시트).
    """
    path = Path(path)
    wb = _open_workbook(path)
    try:
        candidates = []
        names = [sheet] if sheet else wb.sheetnames
        for name in names:
            if name not in wb.sheetnames:
                raise ReadError(f"'{path.name}'에 '{name}' 시트가 없습니다.")
            rows = _sheet_rows(wb[name])
            guess = detect_header_row(rows, scenario)
            candidates.append((guess, name, rows))
    finally:
        wb.close()

    best = candidates[0]
    for c in candidates[1:]:
        if (c[0].confident, c[0].score) > (best[0].confident, best[0].score):
            best = c
    guess, sheet_name, rows = best

    if header_row is not None:
        if header_row < 1 or header_row > max(len(rows), 1):
            raise ReadError(f"'{path.name}'에 {header_row}행이 없습니다.")
        hr, confidence = header_row, "지정"
    else:
        hr, confidence = guess.row, ("높음" if guess.confident else "낮음")

    data_rows = rows[hr:] if rows else []
    header_vals = list(rows[hr - 1]) if rows else []
    width = max([len(header_vals)] + [len(r) for r in data_rows]) if (header_vals or data_rows) else 0
    raw_headers = header_vals + [None] * (width - len(header_vals))
    headers = [
        str(v).strip() if not is_blank(v) else f"(빈 머리글 {get_column_letter(i + 1)}열)"
        for i, v in enumerate(raw_headers)
    ]

    table_rows = []
    for offset, r in enumerate(data_rows):
        if all(is_blank(v) for v in r):
            continue  # 완전히 빈 행은 건너뛴다
        cells = tuple(r) + (None,) * (width - len(r))
        table_rows.append(SourceRow(excel_row=hr + 1 + offset, cells=cells))

    return SourceTable(file_name=path.name, path=path, sheet_name=sheet_name, header_row=hr,
                       headers=headers, raw_headers=raw_headers, rows=table_rows,
                       header_confidence=confidence)
