"""자유 양식 모드: 미리 정한 시나리오 없이, 올린 파일들의 머리글로 결과 양식을 만든다.

흐름
    tables = scan_tables(paths)                  # 머리글 자동 탐지 (시나리오 어휘 없이, 여러 파일 후보 비교)
    groups = regroup(tables, merges)             # 정규화 이름이 같은 열은 자동으로 한 묶음 + 사용자·AI 묶기
    problems = check_choices(groups, choices)    # 결과 열 이름 검사 (비면 통과)
    sc = build_scenario(groups, choices, date_gids, name)   # 메모리 안 Scenario (YAML 을 읽은 것과 같다)
    plans = make_plans(tables, sc)               # 기존 execute / write_result 를 그대로 쓴다
    text = export_yaml(sc, key)                  # "이 양식 저장": scenarios/ 에 넣으면 정식 시나리오

엔진 원칙 그대로: 시나리오 이름·열 이름을 코드에 적지 않는다. 모든 이름은 업로드된 파일과 사용자 입력에서 온다.
형식은 "문자"가 기본이고, 빈칸을 뺀 값이 모두 날짜로 읽히는 열만 "날짜"로 둔다 (그래서 형식 오류는 생기지 않는다).
"""

from __future__ import annotations

import hashlib
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, Iterable, Sequence

import yaml

from .matching import match_columns
from .merge import MAX_FILES, FilePlan
from .normalize import DATE, TEXT, _YMD_COMPACT, MAX_DATE_TEXT, display, is_blank, normalize_header, parse_date
from .reader import ReadError, SourceTable, detect_header_row_generic, read_table, row_labels, scan_tops
from .scenario import (K_ALIASES, K_COLUMNS, K_DESC, K_DUP, K_FORMAT, K_NAME, K_REQUIRED, K_STD, MAX_ALIASES,
                       MAX_COLUMNS, MAX_YAML_BYTES, SCENARIO_KEY_RE, Scenario, ScenarioError, parse_scenario)
from .writer import SOURCE_COLS

FREE_KEY = "free_form"            # 메모리 안 Scenario 의 key (결과 파일 이름에만 쓰인다)
DEFAULT_NAME = "자유 양식"
MAX_RESULT_NAME = 50               # 결과 열 이름 최대 글자 수
MAX_FORM_NAME = 60                 # 양식 이름 최대 글자 수
MAX_BAD_EXAMPLES = 3               # 날짜 안내에 보여주는 예시 값 수

AUTO = "자동"
USER = "사용자"
AI = "AI 추천"


class FreeFormError(Exception):
    """자유 양식 설정 오류. 메시지는 한국어 (열 이름이 들어갈 수 있으므로 화면에서는 이스케이프해서 보여준다)."""


# ------------------------------------------------------------------ 읽기
def scan_tables(paths: Sequence[str | Path], max_rows: int | None = None,
                header_rows: dict[str, int] | None = None) -> list[SourceTable]:
    """파일들을 읽는다. 머리글은 시나리오 없이 찾고, 후보가 동점이면 다른 파일의 머리글 후보와 겹치는 행을 고른다.

    1차: 파일마다 위쪽 행만 읽어 가장 그럴듯한 머리글 후보의 이름을 모은다.
    2차: 다른 파일들의 후보 이름(peer)을 넘겨 머리글을 정하고 데이터를 읽는다.
    header_rows: 파일 이름 -> 사용자가 지정한 머리글 행.
    """
    header_rows = header_rows or {}
    if len(paths) > MAX_FILES:
        raise ReadError(f"한 번에 취합할 수 있는 파일은 최대 {MAX_FILES}개입니다.")
    paths = [Path(p) for p in paths]
    first: list[set[str]] = []
    for p in paths:
        best: tuple[Any, set[str]] | None = None
        for _, top in scan_tops(p):
            g = detect_header_row_generic(top)
            row = header_rows.get(p.name, g.row)   # 사용자가 지정한 행이 있으면 그 행의 이름을 후보로
            labels = row_labels(top[row - 1]) if 0 < row <= len(top) else set()
            if best is None or (g.confident, g.score) > (best[0].confident, best[0].score):
                best = (g, labels)
        first.append(best[1] if best else set())
    tables = []
    for i, p in enumerate(paths):
        peers = frozenset().union(*(s for j, s in enumerate(first) if j != i)) if len(paths) > 1 else frozenset()
        tables.append(read_table(p, None, header_row=header_rows.get(p.name), max_rows=max_rows, peer_names=peers))
    return tables


# ------------------------------------------------------------------ 묶음
@dataclass(frozen=True)
class SourceColumn:
    file_index: int
    col_index: int
    name: str        # 표시용 원래 이름 (앞뒤 공백 제거)
    norm: str        # 정규화한 이름


@dataclass
class ColumnGroup:
    members: list[SourceColumn]
    origin: str = AUTO

    @property
    def norms(self) -> tuple[str, ...]:
        return tuple(sorted({m.norm for m in self.members}))

    @property
    def gid(self) -> str:
        """묶음 id: 정규화 이름 목록에서 만든다 (머리글을 다시 읽어도 같은 묶음이면 같은 id)."""
        return "g" + hashlib.sha1("\x1f".join(self.norms).encode("utf-8")).hexdigest()[:10]

    @property
    def names(self) -> list[str]:
        out: list[str] = []
        for m in self.members:
            if m.name not in out:
                out.append(m.name)
        return out

    @property
    def files(self) -> frozenset[int]:
        return frozenset(m.file_index for m in self.members)

    @property
    def label(self) -> str:
        """가장 많이 쓰인 이름 (같으면 먼저 나온 이름)."""
        counts = Counter(m.name for m in self.members)
        return max(self.names, key=lambda n: (counts[n], -self.names.index(n)))


def _columns(tables: Sequence[SourceTable]) -> list[SourceColumn]:
    out = []
    for fi, t in enumerate(tables):
        for ci, raw in enumerate(t.raw_headers):
            if is_blank(raw):
                continue
            norm = normalize_header(raw)
            if norm:
                out.append(SourceColumn(fi, ci, str(raw).strip(), norm))
    return out


def skipped_columns(tables: Sequence[SourceTable]) -> list[tuple[int, int]]:
    """머리글이 비었거나 글자·숫자가 없어 묶을 수 없는 열 (파일 번호, 열 위치)."""
    used = {(c.file_index, c.col_index) for c in _columns(tables)}
    return [(fi, ci) for fi, t in enumerate(tables) for ci in range(len(t.raw_headers)) if (fi, ci) not in used]


def group_headers(tables: Sequence[SourceTable]) -> list[ColumnGroup]:
    """정규화한 이름이 같은 열을 한 묶음으로. 처음 나온 순서(파일 순서, 열 순서)."""
    by_norm: dict[str, ColumnGroup] = {}
    for c in _columns(tables):
        by_norm.setdefault(c.norm, ColumnGroup([])).members.append(c)
    return list(by_norm.values())


def _first_pos(g: ColumnGroup) -> tuple[int, int]:
    return min((m.file_index, m.col_index) for m in g.members)


def shared_file(a: ColumnGroup, b: ColumnGroup) -> int | None:
    """두 묶음이 함께 들어 있는 파일 번호 (없으면 None)."""
    common = a.files & b.files
    return min(common) if common else None


def merge_groups(groups: list[ColumnGroup], gids: Iterable[str], origin: str = USER,
                 file_names: Sequence[str] | None = None) -> list[ColumnGroup]:
    """묶음 여러 개를 하나로 합친다. 같은 파일에 함께 있는 열끼리는 합칠 수 없다 (FreeFormError)."""
    gids = list(dict.fromkeys(gids))
    chosen = [g for g in groups if g.gid in gids]
    if len(chosen) < 2:
        raise FreeFormError("묶을 열을 두 개 이상 골라 주세요.")
    for i, a in enumerate(chosen):
        for b in chosen[i + 1:]:
            f = shared_file(a, b)
            if f is not None:
                where = file_names[f] if file_names and f < len(file_names) else f"{f + 1}번째 파일"
                raise FreeFormError(f"'{a.label}'과(와) '{b.label}'은 같은 파일({where})에 함께 있는 서로 다른 열이라 "
                                    "묶을 수 없습니다.")
    merged = ColumnGroup(sorted((m for g in chosen for m in g.members), key=lambda m: (m.file_index, m.col_index)),
                         origin)
    out, placed = [], False
    for g in groups:
        if g.gid in gids:
            if not placed:
                out.append(merged)
                placed = True
            continue
        out.append(g)
    return out


def regroup(tables: Sequence[SourceTable], merges: Sequence[tuple[Iterable[str], str]]) -> list[ColumnGroup]:
    """자동 묶음에 사용자·AI 묶기(merges: (정규화 이름들, 방법))를 차례로 다시 적용한다.

    머리글 행을 바꿔 파일을 다시 읽어도 묶기 결정이 남도록, 묶기는 열 위치가 아니라 정규화 이름으로 기억한다.
    지금 파일에 없는 이름이나 같은 파일에 함께 있게 된 묶기는 건너뛴다.
    """
    groups = group_headers(tables)
    for norms, origin in merges:
        norms = set(norms)
        gids = [g.gid for g in groups if norms & set(g.norms)]
        if len(gids) < 2:
            continue
        try:
            groups = merge_groups(groups, gids, origin)
        except FreeFormError:
            continue
    return groups


def remove_merges(merges: Sequence[tuple[Iterable[str], str]], group: ColumnGroup) -> list[tuple[frozenset, str]]:
    """묶음 풀기: 이 묶음의 이름이 들어 있는 묶기 기록을 모두 지운다."""
    norms = set(group.norms)
    return [(frozenset(n), o) for n, o in merges if not (set(n) & norms)]


# ------------------------------------------------------------------ 날짜 판단
@dataclass
class DateCheck:
    is_date: bool                  # 빈칸을 뺀 값이 모두 날짜로 읽힘
    total: int                     # 빈칸을 뺀 값 수
    bad: int = 0                   # 날짜로 읽을 수 없는 값 수
    examples: list[str] = field(default_factory=list)   # 날짜로 읽을 수 없는 값 예시 (사용자 입력: 화면에서 이스케이프)

    @property
    def mixed(self) -> bool:
        """대부분 날짜인데 읽을 수 없는 값이 섞여 문자로 비교하는 열 (안내 대상)."""
        good = self.total - self.bad
        return self.bad > 0 and good > 0 and good >= self.bad


def _date_ok(v: Any) -> bool:
    """자유 양식에서 날짜로 받아들이는 값: 시간이 없는 날짜 셀, 또는 구분자(- . /)나 년·월·일이 있는 날짜 글자.

    8자리 숫자 글자(20260901)는 주문번호·코드일 수 있어 날짜로 보지 않는다. 숫자 셀도 보지 않는다.
    """
    if isinstance(v, datetime):
        return v.time() == time.min
    if isinstance(v, date):
        return True
    if not isinstance(v, str):
        return False
    s = v.strip()
    if len(s) > MAX_DATE_TEXT or _YMD_COMPACT.match(s):
        return False
    return parse_date(s).ok


def check_dates(tables: Sequence[SourceTable], group: ColumnGroup) -> DateCheck:
    total = bad = 0
    examples: list[str] = []
    for m in group.members:
        for row in tables[m.file_index].rows:
            v = row.cells[m.col_index] if m.col_index < len(row.cells) else None
            if is_blank(v):
                continue
            total += 1
            if not _date_ok(v):
                bad += 1
                shown = display(v)
                if len(examples) < MAX_BAD_EXAMPLES and shown not in examples:
                    examples.append(shown[:40])
    return DateCheck(is_date=total > 0 and bad == 0, total=total, bad=bad, examples=examples)


# ------------------------------------------------------------------ 사용자 설정 -> Scenario
@dataclass(frozen=True)
class ColumnChoice:
    gid: str
    name: str             # 결과 열 이름
    required: bool = False
    dup: bool = False     # 중복 판단 기준 (체크하면 필수도 켜진다)


def clean_name(name: Any) -> str:
    """결과 열 이름·양식 이름 정리: 줄바꿈·연속 공백을 공백 하나로, 앞뒤 공백 제거."""
    return " ".join(str(name or "").split())


def check_choices(groups: Sequence[ColumnGroup], choices: Sequence[ColumnChoice]) -> list[str]:
    """결과 열 설정의 문제 목록 (한국어). 비어 있으면 실행·저장할 수 있다."""
    problems: list[str] = []
    if not choices:
        return ["결과에 넣을 열을 하나 이상 골라 주세요."]
    if len(choices) > MAX_COLUMNS:
        problems.append(f"결과에 넣을 수 있는 열은 최대 {MAX_COLUMNS}개입니다 (지금 {len(choices)}개).")
    by_gid = {g.gid: g for g in groups}
    reserved = {normalize_header(n): n for n in SOURCE_COLS}
    owner: dict[str, str] = {}          # 모든 묶음의 원본 이름(정규화) -> 묶음 id
    for g in groups:
        for n in g.norms:
            owner[n] = g.gid
    seen: dict[str, str] = {}
    for pos, c in enumerate(choices, start=1):
        g = by_gid.get(c.gid)
        if g is None:
            problems.append(f"{pos}번째 열의 묶음을 찾을 수 없습니다. 파일을 다시 올려 주세요.")
            continue
        name = clean_name(c.name)
        where = f"{pos}번째 열('{g.label}')"
        if not name:
            problems.append(f"{where}의 결과 열 이름이 비어 있습니다.")
            continue
        if len(name) > MAX_RESULT_NAME:
            problems.append(f"{where}의 결과 열 이름이 너무 깁니다 (최대 {MAX_RESULT_NAME}자).")
            continue
        key = normalize_header(name)
        if not key:
            problems.append(f"{where}의 결과 열 이름 '{name}'에 글자나 숫자가 없습니다.")
            continue
        if key in reserved:
            problems.append(f"결과 열 이름 '{name}'은 결과 파일이 쓰는 '{reserved[key]}' 열과 겹쳐 쓸 수 없습니다.")
            continue
        if key in seen:
            problems.append(f"결과 열 이름 '{name}'이 두 번 쓰였습니다 ('{seen[key]}'과 공백·대소문자·기호만 다름). "
                            "다른 이름을 적어 주세요.")
            continue
        seen[key] = name
        other = owner.get(key)
        if other is not None and other != c.gid:
            problems.append(f"결과 열 이름 '{name}'은 다른 묶음('{by_gid[other].label}')의 원본 열 이름과 같아, "
                            "그 열의 데이터를 잘못 가져올 수 있습니다. 두 묶음을 합치거나 다른 이름을 적어 주세요.")
            continue
        if len(g.names) > MAX_ALIASES:
            problems.append(f"{where}에 묶인 원본 이름이 너무 많습니다 (최대 {MAX_ALIASES}개).")
    return problems


def scenario_dict(groups: Sequence[ColumnGroup], choices: Sequence[ColumnChoice], date_gids: Iterable[str],
                  name: str = DEFAULT_NAME, description: str = "") -> dict[str, Any]:
    """시나리오 YAML 과 같은 모양의 dict. 동의어는 묶음의 원본 이름(결과 열 이름과 같은 것은 뺌)."""
    by_gid = {g.gid: g for g in groups}
    date_gids = set(date_gids)
    columns = []
    for c in choices:
        g = by_gid[c.gid]
        std = clean_name(c.name)
        col: dict[str, Any] = {K_STD: std}
        aliases = [n for n in g.names if n != std]
        if aliases:
            col[K_ALIASES] = aliases
        col[K_REQUIRED] = bool(c.required or c.dup)
        col[K_FORMAT] = DATE if c.gid in date_gids else TEXT
        columns.append(col)
    return {K_NAME: clean_name(name) or DEFAULT_NAME, K_DESC: clean_name(description),
            K_COLUMNS: columns, K_DUP: [clean_name(c.name) for c in choices if c.dup]}


def build_scenario(groups: Sequence[ColumnGroup], choices: Sequence[ColumnChoice], date_gids: Iterable[str] = (),
                   name: str = DEFAULT_NAME, description: str = "", key: str = FREE_KEY) -> Scenario:
    """사용자가 고른 열로 메모리 안 Scenario 를 만든다. YAML 과 같은 검사(parse_scenario)를 통과해야 한다."""
    problems = check_choices(groups, choices)
    if problems:
        raise FreeFormError(problems[0])
    return parse_scenario(scenario_dict(groups, choices, date_gids, name, description), key=key, where="자유 양식 설정")


def make_plans(tables: Sequence[SourceTable], scenario: Scenario) -> list[FilePlan]:
    """이미 읽은 표에 열 매칭만 한다 (파일을 다시 읽지 않는다)."""
    return [FilePlan(table=t, match=match_columns(t, scenario)) for t in tables]


# ------------------------------------------------------------------ YAML 내보내기
def scenario_to_dict(sc: Scenario) -> dict[str, Any]:
    columns = []
    for c in sc.columns:
        col: dict[str, Any] = {K_STD: c.name}
        if c.aliases:
            col[K_ALIASES] = list(c.aliases)
        col[K_REQUIRED] = c.required
        col[K_FORMAT] = c.fmt
        columns.append(col)
    return {K_NAME: sc.name, K_DESC: sc.description, K_COLUMNS: columns, K_DUP: list(sc.dup_keys)}


def check_file_key(key: str) -> str | None:
    """'파일 이름'(확장자 제외) 검사. 문제가 있으면 한국어 이유."""
    if not key:
        return "파일 이름을 적어 주세요."
    if len(key) > 60:
        return "파일 이름이 너무 깁니다 (최대 60자)."
    if not SCENARIO_KEY_RE.fullmatch(key):
        return "파일 이름에는 영문·숫자·_·- 만 쓸 수 있습니다."
    return None


def export_yaml(sc: Scenario, key: str = FREE_KEY) -> str:
    """시나리오를 YAML 글로. yaml.safe_dump 로만 만들고(글자를 이어 붙이지 않음), 다시 읽어 같은지 확인한다."""
    err = check_file_key(key)
    if err:
        raise FreeFormError(err)
    data = scenario_to_dict(sc)
    text = ("# 엑셀 자동 취합·검증기의 자유 양식에서 저장한 시나리오입니다. scenarios/ 폴더에 넣으면 정식 시나리오로 쓸 수 있습니다.\n"
            + yaml.safe_dump(data, allow_unicode=True, sort_keys=False, default_flow_style=None, width=4096))
    if len(text.encode("utf-8")) > MAX_YAML_BYTES:
        raise FreeFormError(f"양식 파일이 너무 커서 저장할 수 없습니다 (최대 {MAX_YAML_BYTES // 1024}KB). "
                            "열 수를 줄여 주세요.")
    try:
        again = parse_scenario(yaml.safe_load(text), key=key, where="저장할 양식")
    except ScenarioError as e:
        raise FreeFormError(f"저장할 양식을 다시 읽는 중 문제가 생겼습니다. {e}") from None
    if scenario_to_dict(again) != data:
        raise FreeFormError("저장할 양식을 다시 읽었을 때 내용이 달라 저장하지 않았습니다.")
    return text


__all__ = ["AI", "AUTO", "USER", "ColumnChoice", "ColumnGroup", "DateCheck", "FREE_KEY", "FreeFormError",
           "SourceColumn", "build_scenario", "check_choices", "check_dates", "check_file_key", "export_yaml",
           "group_headers", "make_plans", "merge_groups", "regroup", "remove_merges", "scan_tables",
           "scenario_dict", "scenario_to_dict", "skipped_columns"]
