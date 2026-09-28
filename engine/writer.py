"""결과 엑셀을 새 파일로 저장한다 (원본은 절대 건드리지 않는다).

시트 1 "취합결과": 출처 파일, 원래 행 + 기준열. 오류 셀은 오류 종류별 색.
시트 2 "오류목록": 파일, 행, 열, 기준열, 오류 종류, 값, 설명, 수정 제안값, 중복 그룹 + 색 범례.
"""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .merge import MergeResult
from .validate import ALLOWED, DUPLICATE, FORMAT, KINDS, RANGE, REQUIRED, WHOLE_FILE

RESULT_SHEET = "취합결과"
ERROR_SHEET = "오류목록"
SOURCE_COLS = ["출처 파일", "원래 행"]
ERROR_COLS = ["파일", "행", "열", "기준열", "오류 종류", "값", "설명", "수정 제안값", "중복 그룹"]

COLORS = {
    REQUIRED: "FFC7CE",   # 연한 빨강
    FORMAT: "FFEB9C",     # 연한 노랑
    RANGE: "F8CBAD",      # 연한 주황
    ALLOWED: "D9D2E9",    # 연한 보라
    DUPLICATE: "BDD7EE",  # 연한 파랑
}
FILLS = {k: PatternFill("solid", start_color=v, end_color=v) for k, v in COLORS.items()}
BOLD = Font(bold=True)

_SAFE_NAME = re.compile(r"[^0-9A-Za-z가-힣_\-]+")


def _put(ws, row: int, col: int, value: Any):
    """셀에 값을 쓴다. 문자열은 수식으로 해석되지 않게 항상 문자 타입으로 둔다."""
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value)
    cell = ws.cell(row=row, column=col, value=value)
    if isinstance(value, str):
        cell.data_type = "s"   # '=' '+' '-' '@'로 시작해도 수식이 되지 않는다
    elif isinstance(value, datetime):
        cell.number_format = "yyyy-mm-dd hh:mm:ss" if value.time() != datetime.min.time() else "yyyy-mm-dd"
    elif isinstance(value, date):
        cell.number_format = "yyyy-mm-dd"
    return cell


def safe_output_path(out_dir: str | Path, file_name: str, sources: list[Path]) -> Path:
    """output 폴더 안의 경로만 허용하고, 원본 파일 경로와 같으면 거부한다."""
    out_dir = Path(out_dir).resolve()
    stem = _SAFE_NAME.sub("_", Path(file_name).stem).strip("_") or "취합결과"
    path = (out_dir / f"{stem}.xlsx").resolve()
    if path.parent != out_dir:
        raise ValueError("결과 파일은 출력 폴더 안에만 저장할 수 있습니다.")
    src = {Path(s).resolve() for s in sources}
    n = 2
    base = path
    while path.exists() or path in src:
        path = base.with_name(f"{base.stem}_{n}.xlsx")
        n += 1
    return path


def _autosize(ws, widths: list[int]):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(8, min(w + 2, 60))


def _text_len(v: Any) -> int:
    if v is None:
        return 0
    s = v.isoformat() if isinstance(v, (date, datetime)) else str(v)
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in s)


def write_result(result: MergeResult, out_dir: str | Path = "output",
                 file_name: str | None = None) -> Path:
    scenario = result.scenario
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    if file_name is None:
        file_name = f"{scenario.key}_취합결과_{datetime.now():%Y%m%d_%H%M%S}"
    sources = [p.table.path for p in result.plans]
    path = safe_output_path(out_dir, file_name, sources)

    wb = Workbook()
    ws = wb.active
    ws.title = RESULT_SHEET
    headers = SOURCE_COLS + scenario.column_names
    widths = [_text_len(h) for h in headers]
    for c, h in enumerate(headers, start=1):
        _put(ws, 1, c, h).font = BOLD
    for r, row in enumerate(result.rows, start=2):
        _put(ws, r, 1, row.source_file)
        _put(ws, r, 2, row.excel_row)
        for c, name in enumerate(scenario.column_names, start=3):
            kind = row.errors.get(name)
            # 오류 셀은 원래 값 그대로 (자동으로 고치지 않는다), 정상 셀은 정규화된 값
            value = row.raw[name] if kind else row.values[name]
            cell = _put(ws, r, c, value)
            if kind:
                cell.fill = FILLS[kind]
            elif row.dup_group and name in scenario.dup_keys:
                cell.fill = FILLS[DUPLICATE]
            widths[c - 1] = max(widths[c - 1], _text_len(value))
        if row.dup_group:
            ws.cell(row=r, column=1).fill = FILLS[DUPLICATE]
            ws.cell(row=r, column=2).fill = FILLS[DUPLICATE]
        widths[0] = max(widths[0], _text_len(row.source_file))
    ws.freeze_panes = "C2"
    if result.rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(result.rows) + 1}"
    _autosize(ws, widths)

    es = wb.create_sheet(ERROR_SHEET)
    widths = [_text_len(h) for h in ERROR_COLS]
    for c, h in enumerate(ERROR_COLS, start=1):
        _put(es, 1, c, h).font = BOLD
    for r, issue in enumerate(result.issues, start=2):
        vals = [issue.file,
                issue.row if issue.row is not None else WHOLE_FILE,
                issue.column if issue.column is not None else "(없음)",
                issue.standard, issue.kind,
                issue.value if issue.value is not None else "(빈칸)",
                issue.message, issue.suggestion, issue.dup_group]
        for c, v in enumerate(vals, start=1):
            _put(es, r, c, v)
            widths[c - 1] = max(widths[c - 1], _text_len(v))
        es.cell(row=r, column=5).fill = FILLS[issue.kind]
    es.freeze_panes = "A2"
    if result.issues:
        es.auto_filter.ref = f"A1:{get_column_letter(len(ERROR_COLS))}{len(result.issues) + 1}"
    _autosize(es, widths)

    # 범례: 오류목록 오른쪽
    lc = len(ERROR_COLS) + 2
    _put(es, 1, lc, "색 범례").font = BOLD
    for i, kind in enumerate(KINDS, start=2):
        _put(es, i, lc, kind).fill = FILLS[kind]
    _put(es, len(KINDS) + 3, lc, "취합결과 시트의 오류 셀은 원래 값 그대로 두고 색만 칠했습니다.")
    es.column_dimensions[get_column_letter(lc)].width = 16

    wb.save(path)
    return path
