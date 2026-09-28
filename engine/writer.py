"""결과 엑셀을 새 파일로 저장한다 (원본은 절대 건드리지 않는다).

시트 1 "취합결과": 출처 파일, 원래 행 + 기준열. 오류 셀은 오류 종류별 색.
시트 2 "오류목록": 파일, 행, 열, 기준열, 오류 종류, 값, 설명, 수정 제안값, 중복 그룹
                  (+ 제안값을 적용했을 때만 "처리" 열).
시트 3 "범례": 색의 뜻. 오류목록과 섞이지 않게 따로 둔다 (pandas로 오류목록을 읽을 때 이름 없는 열이 생기지 않도록).

apply_suggestions=True 이면 수정 제안값이 있는 오류 셀을 제안값으로 바꾸고 초록색으로 칠한다.
기본값(False)은 오류 셀을 원래 값 그대로 두고 색만 칠한다.
"""

from __future__ import annotations

import re
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .merge import MergeResult
from .validate import ALLOWED, DUPLICATE, FORMAT, KINDS, RANGE, REQUIRED, WHOLE_FILE

RESULT_SHEET = "취합결과"
ERROR_SHEET = "오류목록"
LEGEND_SHEET = "범례"
SOURCE_COLS = ["출처 파일", "원래 행"]
ERROR_COLS = ["파일", "행", "열", "기준열", "오류 종류", "값", "설명", "수정 제안값", "중복 그룹"]
ACTION_COL = "처리"            # 제안값을 적용했을 때만 오류목록 끝에 붙는 열
AUTO_FIXED = "자동 수정됨"

COLORS = {
    REQUIRED: "FFC7CE",   # 연한 빨강
    FORMAT: "FFEB9C",     # 연한 노랑
    RANGE: "F8CBAD",      # 연한 주황
    ALLOWED: "D9D2E9",    # 연한 보라
    DUPLICATE: "BDD7EE",  # 연한 파랑
}
FIXED_COLOR = "C6EFCE"    # 연한 초록: 제안값으로 자동 수정한 셀
FILLS = {k: PatternFill("solid", start_color=v, end_color=v) for k, v in COLORS.items()}
FIXED_FILL = PatternFill("solid", start_color=FIXED_COLOR, end_color=FIXED_COLOR)
BOLD = Font(bold=True)

LEGEND_TEXT = {
    REQUIRED: "꼭 채워야 하는 칸이 비어 있습니다.",
    FORMAT: "숫자 칸에 글자가 있거나, 날짜로 읽을 수 없는 값입니다.",
    RANGE: "허용 범위를 벗어난 숫자입니다.",
    ALLOWED: "정해진 값 목록에 없는 값입니다.",
    DUPLICATE: "중복기준 열의 값이 다른 행과 같습니다 (중복기준 열에 색칠).",
}

# 오류목록 시트에 적는 최대 행 수. 넘으면 안내 행을 남긴다 (보안 검토 4번: 메모리 폭주 방지)
MAX_ERROR_ROWS = 20_000

_SAFE_NAME = re.compile(r"[^0-9A-Za-z가-힣_\-]+")


def _cell(ws, value: Any, fill: PatternFill | None = None, bold: bool = False) -> WriteOnlyCell:
    """write-only 시트용 셀. 문자열은 수식으로 해석되지 않게 항상 문자 타입으로 둔다."""
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value)
    cell = WriteOnlyCell(ws, value=value)
    if isinstance(value, str):
        cell.data_type = "s"   # '=' '+' '-' '@'로 시작해도 수식이 되지 않는다
    elif isinstance(value, datetime):
        cell.number_format = "yyyy-mm-dd hh:mm:ss" if value.time() != datetime.min.time() else "yyyy-mm-dd"
    elif isinstance(value, date):
        cell.number_format = "yyyy-mm-dd"
    if fill is not None:
        cell.fill = fill
    if bold:
        cell.font = BOLD
    return cell


def safe_output_path(out_dir: str | Path, file_name: str, sources: list[Path]) -> Path:
    """output 폴더 안의 경로만 허용하고, 원본 파일이나 기존 파일과 겹치지 않는 새 파일을 만든다.

    파일은 배타적으로(없을 때만) 만들어 자리를 잡아 두므로, 동시에 실행돼도 서로 덮어쓰지 않는다.
    """
    out_dir = Path(out_dir).resolve()
    stem = _SAFE_NAME.sub("_", Path(file_name).stem).strip("_") or "취합결과"
    base = (out_dir / f"{stem}.xlsx").resolve()
    if base.parent != out_dir:
        raise ValueError("결과 파일은 출력 폴더 안에만 저장할 수 있습니다.")
    src = {Path(s).resolve() for s in sources}
    path, n = base, 2
    while True:
        if path not in src:
            try:
                with open(path, "xb"):
                    pass
                return path
            except FileExistsError:
                pass
        path = base.with_name(f"{base.stem}_{n}.xlsx")
        n += 1


def _set_widths(ws, widths: list[int]):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(8, min(w + 2, 60))


def _text_len(v: Any) -> int:
    if v is None:
        return 0
    s = v.isoformat() if isinstance(v, (date, datetime)) else str(v)
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in s)


def _error_values(issue) -> list[Any]:
    return [issue.file,
            issue.row if issue.row is not None else WHOLE_FILE,
            issue.column if issue.column is not None else "(없음)",
            issue.standard, issue.kind,
            issue.value if issue.value is not None else "(빈칸)",
            issue.message, issue.suggestion, issue.dup_group]


def write_result(result: MergeResult, out_dir: str | Path = "output",
                 file_name: str | None = None, apply_suggestions: bool = False) -> Path:
    """결과 엑셀을 새 파일로 저장한다.

    write-only 모드로 한 행씩 흘려 쓰므로 행이 많아도 메모리를 적게 쓴다.
    열 너비는 행을 쓰기 전에 정해야 해서 먼저 한 번 훑어 계산한다.
    """
    scenario = result.scenario
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    if file_name is None:
        tag = "_제안값적용" if apply_suggestions else ""
        file_name = f"{scenario.key}_취합결과{tag}_{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:6]}"
    sources = [p.table.path for p in result.plans]
    path = safe_output_path(out_dir, file_name, sources)

    fixes = result.suggestion_map() if apply_suggestions else {}

    def value_of(row, name: str) -> tuple[Any, bool]:
        """(결과 파일에 쓸 값, 자동 수정했는지)."""
        key = (row.source_file, row.excel_row, name)
        if key in fixes:
            return fixes[key], True
        # 오류 셀은 원래 값 그대로 (자동으로 고치지 않는다), 정상 셀은 정규화된 값
        return (row.raw[name] if name in row.errors else row.values[name]), False

    wb = Workbook(write_only=True)

    # ---- 시트 1: 취합결과
    ws = wb.create_sheet(RESULT_SHEET)
    headers = SOURCE_COLS + scenario.column_names
    widths = [_text_len(h) for h in headers]
    for row in result.rows:
        widths[0] = max(widths[0], _text_len(row.source_file))
        for c, name in enumerate(scenario.column_names, start=2):
            widths[c] = max(widths[c], _text_len(value_of(row, name)[0]))
    _set_widths(ws, widths)
    ws.freeze_panes = "C2"
    if result.rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(result.rows) + 1}"
    ws.append([_cell(ws, h, bold=True) for h in headers])
    for row in result.rows:
        dup_fill = FILLS[DUPLICATE] if row.dup_group else None
        cells = [_cell(ws, row.source_file, dup_fill), _cell(ws, row.excel_row, dup_fill)]
        for name in scenario.column_names:
            value, fixed = value_of(row, name)
            kind = row.errors.get(name)
            if fixed:
                fill = FIXED_FILL
            elif kind:
                fill = FILLS[kind]
            elif row.dup_group and name in scenario.dup_keys:
                fill = FILLS[DUPLICATE]
            else:
                fill = None
            cells.append(_cell(ws, value, fill))
        ws.append(cells)

    # ---- 시트 2: 오류목록
    es = wb.create_sheet(ERROR_SHEET)
    cols = ERROR_COLS + ([ACTION_COL] if apply_suggestions else [])
    shown = result.issues[:MAX_ERROR_ROWS]
    omitted = len(result.issues) - len(shown)
    widths = [_text_len(h) for h in cols]
    for issue in shown:
        for c, v in enumerate(_error_values(issue)):
            widths[c] = max(widths[c], _text_len(v))
    if apply_suggestions:
        widths[-1] = max(widths[-1], _text_len(AUTO_FIXED))
    _set_widths(es, widths)
    es.freeze_panes = "A2"
    if shown:
        es.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(shown) + 1}"

    kind_col = ERROR_COLS.index("오류 종류")
    es.append([_cell(es, h, bold=True) for h in cols])
    for issue in shown:
        cells = [_cell(es, v, FILLS[issue.kind] if c == kind_col else None)
                 for c, v in enumerate(_error_values(issue))]
        if apply_suggestions:
            if (issue.file, issue.row, issue.standard) in fixes and issue.kind != DUPLICATE:
                cells.append(_cell(es, AUTO_FIXED, FIXED_FILL))
            else:
                cells.append(None)
        es.append(cells)
    if omitted:
        es.append([_cell(es, f"이하 {omitted:,}건은 생략했습니다 (최대 {MAX_ERROR_ROWS:,}건 표시). "
                             "원본 파일을 먼저 정리한 뒤 다시 취합해 주세요.")])

    # ---- 시트 3: 범례
    ls = wb.create_sheet(LEGEND_SHEET)
    _set_widths(ls, [16, 70])
    ls.append([_cell(ls, "색", bold=True), _cell(ls, "뜻", bold=True)])
    for k in KINDS:
        ls.append([_cell(ls, k, FILLS[k]), _cell(ls, LEGEND_TEXT[k])])
    if apply_suggestions:
        ls.append([_cell(ls, AUTO_FIXED, FIXED_FILL),
                   _cell(ls, "수정 제안값으로 바꾼 셀입니다. 원래 값은 오류목록의 '값' 열에 있습니다.")])
    ls.append([])
    if apply_suggestions:
        note = ("수정 제안값이 있는 오류 셀은 제안값으로 바꾸고 초록색으로 칠했습니다. "
                "제안값이 없는 오류 셀은 원래 값 그대로 두고 색만 칠했습니다.")
    else:
        note = "취합결과 시트의 오류 셀은 원래 값 그대로 두고 색만 칠했습니다."
    ls.append([_cell(ls, "참고", bold=True), _cell(ls, note)])

    wb.save(path)
    return path
