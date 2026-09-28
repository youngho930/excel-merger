"""취합 파이프라인.

    prepare(scenario, paths)  -> list[FilePlan]   # 읽기 + 머리글 탐지 + 열 매칭
    (화면/CLI가 plan.match 를 확인·수정하고, pre_run_warnings()로 경고를 보여준다)
    execute(scenario, plans)  -> MergeResult      # 정규화 + 검증 + 중복 + 취합
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from .matching import MatchResult, match_columns
from .reader import ReadError, SourceTable, read_table
from .scenario import Scenario
from .validate import (DUPLICATE, KINDS, REQUIRED, WHOLE_ROW, Issue, dup_key_part,
                       dup_key_text, validate_cell)

MAX_DUP_REFS = 5   # 중복 설명에 적는 상대 행 최대 개수
MAX_FILES = 50     # 한 번에 취합할 수 있는 최대 파일 수
FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")


def escape_formula(value: Any) -> Any:
    """CSV·표 내보내기용: 수식으로 해석될 수 있는 문자열 앞에 '를 붙인다 (보안 검토 8번)."""
    if isinstance(value, str) and value.startswith(FORMULA_PREFIXES):
        return "'" + value
    return value


@dataclass
class FilePlan:
    table: SourceTable
    match: MatchResult

    @property
    def file_name(self) -> str:
        return self.table.file_name

    def warnings(self) -> list[str]:
        out = []
        if self.table.header_confidence == "낮음":
            out.append(f"[{self.file_name}] 머리글 행을 확실히 찾지 못해 {self.table.header_row}행을 머리글로 "
                       "가정했습니다. 매칭 결과를 확인해 주세요.")
        out.extend(self.match.warnings())
        return out


@dataclass
class MergedRow:
    source_file: str
    excel_row: int
    values: dict[str, Any]             # 기준명 -> 정규화된 값 (빈칸·형식 오류는 None)
    raw: dict[str, Any]                # 기준명 -> 원래 값
    errors: dict[str, str] = field(default_factory=dict)   # 기준명 -> 오류 종류
    dup_group: str | None = None


# "처리" 열에 적는 말 (오류목록)
AUTO_FIXED = "자동 수정됨"
EXCLUDED = "행 제외됨"
DUP_RESOLVED = "중복 해소(이 행을 남김)"


@dataclass
class Resolution:
    """MergeResult.resolve()의 결과: 사용자가 처리한 내용."""
    excluded: frozenset                  # 뺀 행 (파일, 원래 행)
    apply_suggestions: bool
    fixes: dict                          # (파일, 행, 기준명) -> 결과 파일에 넣을 제안값 (뺀 행 제외)
    actions: list[str]                   # result.issues와 같은 순서. 처리 안 한 오류는 ""
    resolved_rows: frozenset             # 중복이 풀려 파란 색을 빼는 행 (파일, 원래 행)
    fully_excluded_groups: list[str]     # 행이 모두 빠진 중복 그룹
    kinds: list[str]                     # result.issues의 오류 종류 (counts 계산용)

    def counts(self) -> dict[str, tuple[int, int, int]]:
        """오류 종류 -> (전체, 처리됨, 남은 오류)."""
        total, done = Counter(self.kinds), Counter(k for k, a in zip(self.kinds, self.actions) if a)
        return {k: (total[k], done[k], total[k] - done[k]) for k in KINDS}

    def totals(self) -> tuple[int, int, int]:
        done = sum(1 for a in self.actions if a)
        return len(self.actions), done, len(self.actions) - done


@dataclass
class MergeResult:
    scenario: Scenario
    plans: list[FilePlan]
    rows: list[MergedRow]
    issues: list[Issue]

    def counts(self) -> dict[str, int]:
        c = Counter(i.kind for i in self.issues)
        return {k: c.get(k, 0) for k in KINDS}

    def suggestion_map(self) -> dict[tuple[str, int, str], Any]:
        """셀 단위 오류 중 수정 제안값이 있는 것: (파일, 행, 기준명) -> 제안값.

        "수정 제안값 일괄 적용"에서 결과 파일에 반영할 대상이다. 파일 전체 오류와 중복 행은 셀이 아니므로 뺀다.
        """
        return {(i.file, i.row, i.standard): i.suggestion for i in self.issues
                if i.suggestion is not None and i.row is not None and i.kind != DUPLICATE}

    def dup_groups(self) -> dict[str, list[MergedRow]]:
        """중복 그룹 이름 -> 그 그룹의 행들 (취합 순서)."""
        groups: dict[str, list[MergedRow]] = {}
        for r in self.rows:
            if r.dup_group:
                groups.setdefault(r.dup_group, []).append(r)
        return groups

    def resolve(self, excluded: Iterable[tuple[str, int]] = (),
                apply_suggestions: bool = False) -> "Resolution":
        """사용자의 처리(행 제외, 제안값 적용)를 반영한 결과를 계산한다. 원래 결과는 바꾸지 않는다.

        excluded: 취합결과에서 뺄 행 (파일 이름, 원래 행 번호)
        """
        excluded = frozenset((f, int(r)) for f, r in excluded)
        groups = self.dup_groups()
        kept_count = {g: sum(1 for r in rows if (r.source_file, r.excel_row) not in excluded)
                      for g, rows in groups.items()}
        resolved_rows = frozenset((r.source_file, r.excel_row) for g, rows in groups.items()
                                  if kept_count[g] == 1 for r in rows
                                  if (r.source_file, r.excel_row) not in excluded)
        fully_excluded = [g for g, n in kept_count.items() if n == 0]
        fixes = {k: v for k, v in self.suggestion_map().items()
                 if (k[0], k[1]) not in excluded} if apply_suggestions else {}

        actions = []
        for i in self.issues:
            pos = (i.file, i.row)
            if i.row is not None and pos in excluded:
                actions.append(EXCLUDED)
            elif i.kind != DUPLICATE and (i.file, i.row, i.standard) in fixes:
                actions.append(AUTO_FIXED)
            elif i.kind == DUPLICATE and pos in resolved_rows:
                actions.append(DUP_RESOLVED)
            else:
                actions.append("")
        return Resolution(excluded=excluded, apply_suggestions=apply_suggestions, fixes=fixes,
                          actions=actions, resolved_rows=resolved_rows, fully_excluded_groups=fully_excluded,
                          kinds=[i.kind for i in self.issues])

    def to_dataframe(self, escape_formulas: bool = False):
        """화면 표시용 표. CSV로 내보낼 때는 escape_formulas=True로 수식 주입을 막는다."""
        import pandas as pd
        cols = ["출처 파일", "원래 행"] + self.scenario.column_names
        esc = escape_formula if escape_formulas else (lambda v: v)
        data = []
        for r in self.rows:
            data.append([esc(r.source_file), r.excel_row] +
                        [esc(r.raw[c] if c in r.errors else r.values[c]) for c in self.scenario.column_names])
        return pd.DataFrame(data, columns=cols)


def prepare(scenario: Scenario, paths: list[str | Path],
            header_rows: dict[str, int] | None = None) -> list[FilePlan]:
    """파일들을 읽고 열을 매칭한다. 파일 순서는 주어진 순서를 그대로 쓴다."""
    header_rows = header_rows or {}
    if len(paths) > MAX_FILES:
        raise ReadError(f"한 번에 취합할 수 있는 파일은 최대 {MAX_FILES}개입니다.")
    plans = []
    for p in paths:
        p = Path(p)
        table = read_table(p, scenario, header_row=header_rows.get(p.name))
        plans.append(FilePlan(table=table, match=match_columns(table, scenario)))
    return plans


def pre_run_warnings(plans: list[FilePlan]) -> list[str]:
    return [w for p in plans for w in p.warnings()]


def execute(scenario: Scenario, plans: list[FilePlan]) -> MergeResult:
    rows: list[MergedRow] = []
    issues: list[Issue] = []
    dup_candidates: list[tuple[MergedRow, tuple]] = []
    spec_by_name = {c.name: c for c in scenario.columns}

    for plan in plans:
        match = plan.match
        missing = match.missing_required
        dup_blocked = [s for s in missing if s in scenario.dup_keys]
        if missing:
            msg = "파일에 필수 열이 없음: " + ", ".join(f"'{s}'" for s in missing)
            if dup_blocked:
                msg += f" — {', '.join(dup_blocked)} 열이 없어 이 파일은 중복 검사를 하지 못함"
            issues.append(Issue(file=plan.file_name, row=None, column=None,
                                standard=", ".join(missing), kind=REQUIRED, value=None, message=msg))

        file_issues: list[Issue] = []
        for src in plan.table.rows:
            values: dict[str, Any] = {}
            raw: dict[str, Any] = {}
            errors: dict[str, str] = {}
            for m in match.matches:
                spec = spec_by_name[m.standard]
                if m.source_index is None:
                    values[m.standard] = None
                    raw[m.standard] = None
                    continue
                cell = src.cells[m.source_index] if m.source_index < len(src.cells) else None
                raw[m.standard] = cell
                res = validate_cell(cell, spec)
                values[m.standard] = res.value
                if not res.ok:
                    errors[m.standard] = res.kind
                    file_issues.append(Issue(
                        file=plan.file_name, row=src.excel_row, column=m.source_name,
                        standard=m.standard, kind=res.kind, value=cell, message=res.message,
                        suggestion=res.suggestion))
            row = MergedRow(plan.file_name, src.excel_row, values, raw, errors)
            rows.append(row)

            # 중복 판정: 중복기준 열이 없거나, 빈칸이거나, 형식 오류인 행은 뺀다
            if scenario.dup_keys and not dup_blocked:
                # (범위·허용값 오류인 값은 읽을 수는 있으므로 판정에 쓴다)
                if all(values.get(k) is not None for k in scenario.dup_keys):
                    key = tuple(dup_key_part(values[k], spec_by_name[k]) for k in scenario.dup_keys)
                    dup_candidates.append((row, key))
        issues.extend(file_issues)

    issues.extend(_find_duplicates(scenario, dup_candidates, rows))
    order = {p.file_name: i for i, p in enumerate(plans)}
    col_order = {c: i for i, c in enumerate(scenario.column_names)}
    issues.sort(key=lambda i: (order.get(i.file, 0), i.row or 0, i.kind == DUPLICATE,
                               col_order.get(i.standard, -1)))
    return MergeResult(scenario=scenario, plans=plans, rows=rows, issues=issues)


def _find_duplicates(scenario: Scenario, candidates: list[tuple[MergedRow, tuple]],
                     rows: list[MergedRow]) -> list[Issue]:
    groups: dict[tuple, list[MergedRow]] = {}
    for row, key in candidates:
        groups.setdefault(key, []).append(row)
    label = " + ".join(scenario.dup_keys)
    short = "+".join(scenario.dup_keys)
    issues = []
    n = 0
    for key, members in groups.items():   # dict는 처음 나온 순서를 지킨다
        if len(members) < 2:
            continue
        n += 1
        group = f"중복-{n}"
        text = dup_key_text(scenario.dup_keys, list(key))
        # 설명에는 상대 행을 앞에서부터 최대 MAX_DUP_REFS개만 적는다.
        # 그룹 전체를 행마다 다시 훑지 않도록 앞쪽 몇 개만 잘라 둔다 (보안 검토 3번: O(k²) 방지)
        head = members[:MAX_DUP_REFS + 1]
        for row in members:
            row.dup_group = group
            shown = [o for o in head if o is not row][:MAX_DUP_REFS]
            refs = ", ".join(f"{o.source_file} {o.excel_row}행" for o in shown)
            rest = len(members) - 1 - len(shown)
            if rest > 0:
                refs += f" 외 {rest}행"
            issues.append(Issue(file=row.source_file, row=row.excel_row, column=WHOLE_ROW,
                                standard=label, kind=DUPLICATE, value=text,
                                message=f"{refs}과 중복기준({short}) 값이 같음", dup_group=group))
    return issues


__all__ = ["AUTO_FIXED", "DUP_RESOLVED", "EXCLUDED", "FilePlan", "MergedRow", "MergeResult", "Resolution",
           "prepare", "pre_run_warnings", "execute"]
