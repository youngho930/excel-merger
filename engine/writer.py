"""결과 엑셀을 새 파일로 저장한다 (원본은 절대 건드리지 않는다).

시트 순서
1. "취합결과": 출처 파일, 원래 행 + 기준열. 오류 셀은 오류 종류별 색. 사용자가 뺀 행은 없다.
2. "오류목록": 파일, 행, 열, 기준열, 오류 종류, 값, 설명, 수정 제안값, 중복 그룹, 처리.
   "처리" 열은 항상 있다 (양식이 매번 같도록). 처리하지 않은 오류는 빈칸이고,
   처리한 오류는 "자동 수정됨" / "행 제외됨" / "중복 해소(이 행을 남김)".
3. "제외된 행": 사용자가 중복 그룹에서 뺀 행의 기록 (뺀 행이 있을 때만).
4. "요약": 파일 수, 행 수, 제외한 행 수, 오류 종류별 전체 / 처리됨 / 남은 오류.
5. "범례": 색의 뜻. 오류목록과 섞이지 않게 따로 둔다 (pandas로 읽을 때 이름 없는 열이 생기지 않도록).

apply_suggestions=True 이면 수정 제안값이 있는 오류 셀을 제안값으로 바꾸고 초록색으로 칠한다.
excluded 에 (파일 이름, 원래 행)을 주면 그 행을 취합결과에서 빼고 "제외된 행" 시트로 옮긴다.
둘 다 주지 않으면 오류 셀은 원래 값 그대로 두고 색만 칠한다.
"""

from __future__ import annotations

import re
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .merge import AUTO_FIXED, DUP_RESOLVED, EXCLUDED, MergeResult
from .validate import ALLOWED, DUPLICATE, FORMAT, KINDS, RANGE, REQUIRED, WHOLE_FILE

RESULT_SHEET = "취합결과"
ERROR_SHEET = "오류목록"
EXCLUDED_SHEET = "제외된 행"
SUMMARY_SHEET = "요약"
LEGEND_SHEET = "범례"
SOURCE_COLS = ["출처 파일", "원래 행"]
ACTION_COL = "처리"
ERROR_COLS = ["파일", "행", "열", "기준열", "오류 종류", "값", "설명", "수정 제안값", "중복 그룹", ACTION_COL]
EXCLUDED_EXTRA_COLS = ["중복 그룹", "오류"]
SUMMARY_KIND_COLS = ["오류 종류", "전체", "처리됨", "남은 오류"]

COLORS = {
    REQUIRED: "FFC7CE",   # 연한 빨강
    FORMAT: "FFEB9C",     # 연한 노랑
    RANGE: "F8CBAD",      # 연한 주황
    ALLOWED: "D9D2E9",    # 연한 보라
    DUPLICATE: "BDD7EE",  # 연한 파랑
}
FIXED_COLOR = "C6EFCE"    # 연한 초록: 제안값으로 자동 수정한 셀
DONE_COLOR = "E7E6E6"     # 연한 회색: 행 제외됨, 중복 해소
FILLS = {k: PatternFill("solid", start_color=v, end_color=v) for k, v in COLORS.items()}
FIXED_FILL = PatternFill("solid", start_color=FIXED_COLOR, end_color=FIXED_COLOR)
DONE_FILL = PatternFill("solid", start_color=DONE_COLOR, end_color=DONE_COLOR)
ACTION_FILLS = {AUTO_FIXED: FIXED_FILL, EXCLUDED: DONE_FILL, DUP_RESOLVED: DONE_FILL}
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


def _error_values(issue, action: str) -> list[Any]:
    return [issue.file,
            issue.row if issue.row is not None else WHOLE_FILE,
            issue.column if issue.column is not None else "(없음)",
            issue.standard, issue.kind,
            issue.value if issue.value is not None else "(빈칸)",
            issue.message, issue.suggestion, issue.dup_group, action or None]


def write_result(result: MergeResult, out_dir: str | Path = "output",
                 file_name: str | None = None, apply_suggestions: bool = False,
                 excluded: Iterable[tuple[str, int]] = ()) -> Path:
    """결과 엑셀을 새 파일로 저장한다.

    write-only 모드로 한 행씩 흘려 쓰므로 행이 많아도 메모리를 적게 쓴다.
    열 너비는 행을 쓰기 전에 정해야 해서 먼저 한 번 훑어 계산한다.
    """
    scenario = result.scenario
    res = result.resolve(excluded, apply_suggestions)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    if file_name is None:
        tag = ("_제안값적용" if apply_suggestions else "") + ("_행제외" if res.excluded else "")
        file_name = f"{scenario.key}_취합결과{tag}_{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:6]}"
    sources = [p.table.path for p in result.plans]
    path = safe_output_path(out_dir, file_name, sources)

    fixes = res.fixes
    kept = [r for r in result.rows if (r.source_file, r.excel_row) not in res.excluded]
    dropped = [r for r in result.rows if (r.source_file, r.excel_row) in res.excluded]

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
    for row in kept:
        widths[0] = max(widths[0], _text_len(row.source_file))
        for c, name in enumerate(scenario.column_names, start=2):
            widths[c] = max(widths[c], _text_len(value_of(row, name)[0]))
    _set_widths(ws, widths)
    ws.freeze_panes = "C2"
    if kept:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(kept) + 1}"
    ws.append([_cell(ws, h, bold=True) for h in headers])
    for row in kept:
        is_dup = bool(row.dup_group) and (row.source_file, row.excel_row) not in res.resolved_rows
        dup_fill = FILLS[DUPLICATE] if is_dup else None
        cells = [_cell(ws, row.source_file, dup_fill), _cell(ws, row.excel_row, dup_fill)]
        for name in scenario.column_names:
            value, fixed = value_of(row, name)
            kind = row.errors.get(name)
            if fixed:
                fill = FIXED_FILL
            elif kind:
                fill = FILLS[kind]
            elif is_dup and name in scenario.dup_keys:
                fill = FILLS[DUPLICATE]
            else:
                fill = None
            cells.append(_cell(ws, value, fill))
        ws.append(cells)

    # ---- 시트 2: 오류목록
    es = wb.create_sheet(ERROR_SHEET)
    shown = list(zip(result.issues, res.actions))[:MAX_ERROR_ROWS]
    omitted = len(result.issues) - len(shown)
    widths = [_text_len(h) for h in ERROR_COLS]
    for issue, action in shown:
        for c, v in enumerate(_error_values(issue, action)):
            widths[c] = max(widths[c], _text_len(v))
    _set_widths(es, widths)
    es.freeze_panes = "A2"
    if shown:
        es.auto_filter.ref = f"A1:{get_column_letter(len(ERROR_COLS))}{len(shown) + 1}"

    kind_col = ERROR_COLS.index("오류 종류")
    action_col = ERROR_COLS.index(ACTION_COL)
    es.append([_cell(es, h, bold=True) for h in ERROR_COLS])
    for issue, action in shown:
        cells = []
        for c, v in enumerate(_error_values(issue, action)):
            if c == kind_col:
                cells.append(_cell(es, v, FILLS[issue.kind]))
            elif c == action_col:
                cells.append(_cell(es, v, ACTION_FILLS[v]) if v else None)
            else:
                cells.append(_cell(es, v))
        es.append(cells)
    if omitted:
        es.append([_cell(es, f"이하 {omitted:,}건은 생략했습니다 (최대 {MAX_ERROR_ROWS:,}건 표시). "
                             "원본 파일을 먼저 정리한 뒤 다시 취합해 주세요.")])

    # ---- 시트 3: 제외된 행 (있을 때만) — 원래 값 그대로
    if dropped:
        xs = wb.create_sheet(EXCLUDED_SHEET)
        xheaders = SOURCE_COLS + EXCLUDED_EXTRA_COLS + scenario.column_names
        by_row: dict[tuple[str, int], list[str]] = {}
        for issue in result.issues:
            if issue.row is not None:
                label = issue.kind if issue.kind == DUPLICATE else f"{issue.kind}: {issue.standard}"
                by_row.setdefault((issue.file, issue.row), []).append(label)
        lines = []
        for row in dropped:
            summary = ", ".join(by_row.get((row.source_file, row.excel_row), []))
            lines.append((row, summary))
        widths = [_text_len(h) for h in xheaders]
        for row, summary in lines:
            vals = [row.source_file, row.excel_row, row.dup_group, summary] + \
                   [row.raw[n] for n in scenario.column_names]
            for c, v in enumerate(vals):
                widths[c] = max(widths[c], _text_len(v))
        _set_widths(xs, widths)
        xs.freeze_panes = "C2"
        xs.append([_cell(xs, h, bold=True) for h in xheaders])
        for row, summary in lines:
            cells = [_cell(xs, row.source_file), _cell(xs, row.excel_row),
                     _cell(xs, row.dup_group), _cell(xs, summary)]
            for name in scenario.column_names:
                kind = row.errors.get(name)
                cells.append(_cell(xs, row.raw[name], FILLS[kind] if kind else None))
            xs.append(cells)

    # ---- 시트 4: 요약
    ss = wb.create_sheet(SUMMARY_SHEET)
    _set_widths(ss, [22, 10, 10, 12])
    total, done, remaining = res.totals()
    for label, value in [("파일 수", len(result.plans)),
                         ("읽은 데이터 행 수", len(result.rows)),
                         ("제외한 행 수", len(dropped)),
                         ("취합결과 행 수", len(kept)),
                         ("수정 제안값 일괄 적용", "예" if apply_suggestions else "아니오")]:
        ss.append([_cell(ss, label, bold=True), _cell(ss, value)])
    ss.append([])
    ss.append([_cell(ss, h, bold=True) for h in SUMMARY_KIND_COLS])
    for k, (t, d, r) in res.counts().items():
        ss.append([_cell(ss, k, FILLS[k]), _cell(ss, t), _cell(ss, d), _cell(ss, r, bold=r > 0)])
    ss.append([_cell(ss, "합계", bold=True), _cell(ss, total, bold=True), _cell(ss, done, bold=True),
               _cell(ss, remaining, bold=True)])
    ss.append([])
    ss.append([_cell(ss, "처리됨: 오류목록의 '처리' 열에 내용이 있는 오류. 남은 오류 = 전체 - 처리됨. "
                         "확인이 필요한 것은 '남은 오류'입니다.")])

    # ---- 시트 5: 범례
    ls = wb.create_sheet(LEGEND_SHEET)
    _set_widths(ls, [24, 70])
    ls.append([_cell(ls, "색", bold=True), _cell(ls, "뜻", bold=True)])
    for k in KINDS:
        ls.append([_cell(ls, k, FILLS[k]), _cell(ls, LEGEND_TEXT[k])])
    ls.append([_cell(ls, AUTO_FIXED, FIXED_FILL),
               _cell(ls, "수정 제안값으로 바꾼 셀입니다. 원래 값은 오류목록의 '값' 열에 있습니다.")])
    ls.append([_cell(ls, EXCLUDED, DONE_FILL),
               _cell(ls, "중복 그룹에서 뺀 행의 오류입니다. 그 행은 '제외된 행' 시트에 원래 값으로 남아 있습니다.")])
    ls.append([_cell(ls, DUP_RESOLVED, DONE_FILL),
               _cell(ls, "같은 그룹의 다른 행을 빼서 중복이 풀린 행입니다. 취합결과에서 파란 색을 뺐습니다.")])
    ls.append([])
    if apply_suggestions:
        note = ("수정 제안값이 있는 오류 셀은 제안값으로 바꾸고 초록색으로 칠했습니다. "
                "제안값이 없는 오류 셀은 원래 값 그대로 두고 색만 칠했습니다.")
    else:
        note = "취합결과 시트의 오류 셀은 원래 값 그대로 두고 색만 칠했습니다."
    ls.append([_cell(ls, "참고", bold=True), _cell(ls, note)])

    wb.save(path)
    return path
