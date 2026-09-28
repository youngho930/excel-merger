"""원본 열과 기준열을 짝짓는다.

순서: 정확히 일치(기준명) -> 동의어. 비교 전 공백·대소문자·특수문자를 무시한다.
결과(MatchResult)는 바로 적용하지 않고, 화면에서 사용자가 assign()으로 고칠 수 있다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from openpyxl.utils import get_column_letter

from .normalize import normalize_header
from .reader import SourceTable
from .scenario import Scenario

EXACT = "정확히 일치"
ALIAS = "동의어"
USER = "사용자 지정"
AI = "AI 추천"          # engine/ai_match.py 가 채운다 (사용자가 확인표에서 확인·수정)
NONE = "매칭 안 됨"


@dataclass
class ColumnMatch:
    standard: str                  # 기준명
    required: bool
    source_index: int | None       # 원본 열 위치(0부터). 머리글 이름이 겹칠 수 있어 위치로 구분
    source_name: str | None        # 원본 열 이름 (표시용)
    method: str                    # EXACT | ALIAS | USER | NONE
    matched_alias: str | None = None


@dataclass
class Conflict:
    """같은 기준열에 원본 열 후보가 둘 이상일 때."""
    standard: str
    chosen: int
    ignored: list[int]


@dataclass
class MatchResult:
    file_name: str
    headers: list[str]
    matches: list[ColumnMatch]                 # 기준열 순서, 기준열마다 1개
    conflicts: list[Conflict] = field(default_factory=list)

    # ------------------------------------------------ 조회
    def get(self, standard: str) -> ColumnMatch:
        for m in self.matches:
            if m.standard == standard:
                return m
        raise KeyError(standard)

    def column_label(self, index: int) -> str:
        return f"{get_column_letter(index + 1)}열 '{self.headers[index]}'"

    @property
    def unmatched_sources(self) -> list[int]:
        """어느 기준열에도 쓰이지 않은 원본 열 위치."""
        used = {m.source_index for m in self.matches if m.source_index is not None}
        return [i for i in range(len(self.headers)) if i not in used]

    @property
    def missing_required(self) -> list[str]:
        return [m.standard for m in self.matches if m.required and m.source_index is None]

    @property
    def missing_optional(self) -> list[str]:
        return [m.standard for m in self.matches if not m.required and m.source_index is None]

    # ------------------------------------------------ 경고 (실행 전)
    def warnings(self) -> list[str]:
        out = []
        unmatched = ", ".join(self.column_label(i) for i in self.unmatched_sources) or "없음"
        if self.missing_required:
            names = ", ".join(f"'{s}'" for s in self.missing_required)
            out.append(f"[{self.file_name}] 필수 열 {names}을(를) 찾지 못했습니다. "
                       f"매칭 안 된 원본 열: {unmatched}")
        for c in self.conflicts:
            out.append(
                f"[{self.file_name}] {c.standard} 후보가 {1 + len(c.ignored)}개: "
                f"{self.column_label(c.chosen)} 사용, "
                + ", ".join(f"{self.column_label(i)} 무시" for i in c.ignored))
        return out

    # ------------------------------------------------ 사용자 수정
    def assign(self, standard: str, source_index: int | None) -> None:
        """기준열에 원본 열을 직접 지정한다 (None이면 매칭 해제)."""
        target = self.get(standard)
        if source_index is not None:
            if not (0 <= source_index < len(self.headers)):
                raise ValueError(f"원본 열 위치 {source_index}가 파일에 없습니다.")
            for m in self.matches:
                if m is not target and m.source_index == source_index:
                    raise ValueError(
                        f"{self.column_label(source_index)}은 이미 '{m.standard}'에 쓰이고 있습니다. "
                        "먼저 그 매칭을 해제해 주세요.")
        target.source_index = source_index
        target.source_name = self.headers[source_index] if source_index is not None else None
        target.method = USER if source_index is not None else NONE
        target.matched_alias = None
        self.conflicts = [c for c in self.conflicts if c.standard != standard]

    def suggest(self, standard: str, source_index: int) -> bool:
        """비어 있는 기준열에 원본 열을 "AI 추천"으로 채운다.

        기준열이 이미 매칭돼 있거나, 원본 열이 다른 기준열에 쓰이고 있거나, 위치가 파일에 없으면
        아무것도 하지 않고 False. 사용자는 나중에 assign()으로 바꾸거나 해제할 수 있다.
        """
        try:
            target = self.get(standard)
        except KeyError:
            return False
        if target.source_index is not None:
            return False
        if not (isinstance(source_index, int) and 0 <= source_index < len(self.headers)):
            return False
        if any(m.source_index == source_index for m in self.matches):
            return False
        target.source_index = source_index
        target.source_name = self.headers[source_index]
        target.method = AI
        target.matched_alias = None
        return True

    def to_records(self) -> list[dict[str, Any]]:
        """화면 표(st.data_editor)·JSON용."""
        return [{"기준열": m.standard, "필수": m.required, "원본열 위치": m.source_index,
                 "원본열": m.source_name, "매칭 방법": m.method} for m in self.matches]

    def apply_records(self, records: list[dict[str, Any]]) -> None:
        """to_records() 형태의 수정본을 적용한다. 바뀐 행만 assign한다."""
        for r in records:
            m = self.get(r["기준열"])
            idx = r.get("원본열 위치")
            idx = None if idx is None or idx != idx else int(idx)  # NaN -> None
            if idx != m.source_index:
                self.assign(m.standard, None)
        for r in records:
            m = self.get(r["기준열"])
            idx = r.get("원본열 위치")
            idx = None if idx is None or idx != idx else int(idx)
            if idx != m.source_index:
                self.assign(m.standard, idx)


def match_columns(table: SourceTable, scenario: Scenario) -> MatchResult:
    """정확히 일치 -> 동의어 순으로 매칭. 후보가 여럿이면 정확히 일치 우선, 그다음 왼쪽 열."""
    keys = [normalize_header(h) for h in table.raw_headers]
    matches: list[ColumnMatch] = []
    conflicts: list[Conflict] = []
    for col in scenario.columns:
        exact_key = normalize_header(col.name)
        alias_keys = {normalize_header(a): a for a in col.aliases}
        candidates: list[tuple[int, int, str | None]] = []   # (우선순위, 위치, 맞은 동의어)
        for i, k in enumerate(keys):
            if not k:
                continue
            if k == exact_key:
                candidates.append((0, i, None))
            elif k in alias_keys:
                candidates.append((1, i, alias_keys[k]))
        if not candidates:
            matches.append(ColumnMatch(col.name, col.required, None, None, NONE))
            continue
        candidates.sort()
        prio, idx, alias = candidates[0]
        matches.append(ColumnMatch(col.name, col.required, idx, table.headers[idx],
                                   EXACT if prio == 0 else ALIAS, alias))
        if len(candidates) > 1:
            conflicts.append(Conflict(col.name, idx, sorted(i for _, i, _ in candidates[1:])))
    return MatchResult(file_name=table.file_name, headers=list(table.headers),
                       matches=matches, conflicts=conflicts)
