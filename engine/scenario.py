"""시나리오 YAML을 읽고 형식을 검사한다.

엔진은 시나리오 이름·열 이름을 전혀 모른다. 모든 것은 이 파일이 YAML에서 읽어 온다.
형식이 잘못된 YAML은 ScenarioError(한국어 메시지)로 거부한다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

from .normalize import FORMATS, INT, REAL, TEXT, normalize_header, parse_value

# YAML 키 (CLAUDE.md "시나리오 파일 형식")
K_NAME, K_DESC, K_COLUMNS, K_DUP = "name", "description", "columns", "중복기준"
K_STD, K_ALIASES, K_REQUIRED, K_FORMAT, K_RANGE, K_ALLOWED = "기준명", "동의어", "필수", "형식", "범위", "허용값"

TOP_KEYS = {K_NAME, K_DESC, K_COLUMNS, K_DUP}
COLUMN_KEYS = {K_STD, K_ALIASES, K_REQUIRED, K_FORMAT, K_RANGE, K_ALLOWED}

MAX_YAML_BYTES = 256 * 1024
SCENARIO_KEY_RE = re.compile(r"^[A-Za-z0-9_\-]+$")


class ScenarioError(Exception):
    """시나리오 파일 형식 오류. 메시지는 비개발자가 읽을 수 있는 한국어."""


@dataclass(frozen=True)
class ColumnSpec:
    name: str                              # 기준명
    aliases: tuple[str, ...] = ()          # 동의어
    required: bool = False                 # 필수
    fmt: str = TEXT                        # 형식: 문자/정수/실수/날짜
    range: tuple[float, float] | None = None
    allowed: tuple[Any, ...] | None = None  # 형식에 맞춰 정규화된 허용값

    @property
    def names(self) -> tuple[str, ...]:
        return (self.name, *self.aliases)


@dataclass
class Scenario:
    key: str                               # 파일 이름(확장자 제외)
    name: str
    description: str
    columns: list[ColumnSpec]
    dup_keys: list[str] = field(default_factory=list)
    path: Path | None = None

    def column(self, name: str) -> ColumnSpec:
        for c in self.columns:
            if c.name == name:
                return c
        raise KeyError(name)

    @property
    def column_names(self) -> list[str]:
        return [c.name for c in self.columns]


def _is_number(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _parse_column(i: int, raw: Any, where: str) -> ColumnSpec:
    pos = f"{where}의 columns {i}번째 항목"
    if not isinstance(raw, dict):
        raise ScenarioError(f"{pos}이 '기준명: 열 이름' 형태가 아닙니다.")
    unknown = [str(k) for k in raw if k not in COLUMN_KEYS]
    if unknown:
        raise ScenarioError(
            f"{pos}에 알 수 없는 키가 있습니다: {', '.join(unknown)}. "
            f"쓸 수 있는 키: {', '.join(sorted(COLUMN_KEYS))}")

    name = raw.get(K_STD)
    if name is None or str(name).strip() == "":
        raise ScenarioError(f"{pos}에 '{K_STD}'이 없습니다.")
    name = str(name).strip()
    pos = f"{where}의 '{name}' 열"

    fmt = raw.get(K_FORMAT)
    if fmt is None:
        raise ScenarioError(f"{pos}에 '{K_FORMAT}'이 없습니다. {'/'.join(FORMATS)} 중 하나를 적어 주세요.")
    if fmt not in FORMATS:
        raise ScenarioError(f"{pos}의 형식 '{fmt}'은 쓸 수 없습니다. {'/'.join(FORMATS)} 중 하나를 적어 주세요.")

    aliases_raw = raw.get(K_ALIASES) or []
    if not isinstance(aliases_raw, list):
        raise ScenarioError(f"{pos}의 '{K_ALIASES}'은 [이름1, 이름2] 목록으로 적어 주세요.")
    aliases = tuple(str(a).strip() for a in aliases_raw if a is not None and str(a).strip())

    required = raw.get(K_REQUIRED, False)
    if not isinstance(required, bool):
        raise ScenarioError(f"{pos}의 '{K_REQUIRED}'은 true 또는 false로 적어 주세요.")

    rng = raw.get(K_RANGE)
    if rng is not None:
        if fmt not in (INT, REAL):
            raise ScenarioError(f"{pos}은 형식이 '{fmt}'이라 '{K_RANGE}'을 쓸 수 없습니다 (정수·실수 열만 가능).")
        if not (isinstance(rng, list) and len(rng) == 2 and all(_is_number(x) for x in rng)):
            raise ScenarioError(f"{pos}의 '{K_RANGE}'은 [최소, 최대] 숫자 두 개로 적어 주세요.")
        if rng[0] > rng[1]:
            raise ScenarioError(f"{pos}의 '{K_RANGE}'에서 최소({rng[0]})가 최대({rng[1]})보다 큽니다.")
        rng = (rng[0], rng[1])

    allowed_raw = raw.get(K_ALLOWED)
    allowed = None
    if allowed_raw is not None:
        if not isinstance(allowed_raw, list) or not allowed_raw:
            raise ScenarioError(f"{pos}의 '{K_ALLOWED}'은 [값1, 값2] 목록으로 적어 주세요.")
        norm = []
        for a in allowed_raw:
            p = parse_value(a, fmt) if a is not None else None
            if p is None or not p.ok:
                raise ScenarioError(f"{pos}의 허용값 '{a}'이 형식 '{fmt}'에 맞지 않습니다.")
            norm.append(p.value)
        allowed = tuple(norm)

    return ColumnSpec(name=name, aliases=aliases, required=required, fmt=fmt, range=rng, allowed=allowed)


def parse_scenario(data: Any, key: str, where: str = "시나리오 파일") -> Scenario:
    """이미 읽은 YAML 데이터(dict)를 검사해 Scenario로 만든다."""
    if not isinstance(data, dict):
        raise ScenarioError(f"{where}의 내용이 비어 있거나 'name: ...' 형태가 아닙니다.")
    unknown = [str(k) for k in data if k not in TOP_KEYS]
    if unknown:
        raise ScenarioError(
            f"{where}에 알 수 없는 키가 있습니다: {', '.join(unknown)}. "
            f"쓸 수 있는 키: {', '.join(sorted(TOP_KEYS))}")
    for k in (K_NAME, K_COLUMNS, K_DUP):
        if k not in data:
            raise ScenarioError(f"{where}에 필수 항목 '{k}'이 없습니다.")

    name = data[K_NAME]
    if name is None or str(name).strip() == "":
        raise ScenarioError(f"{where}의 '{K_NAME}'이 비어 있습니다.")
    cols_raw = data[K_COLUMNS]
    if not isinstance(cols_raw, list) or not cols_raw:
        raise ScenarioError(f"{where}의 '{K_COLUMNS}'에 열이 하나도 없습니다.")

    columns = [_parse_column(i, c, where) for i, c in enumerate(cols_raw, start=1)]

    # 기준명 중복, 정규화한 이름(기준명·동의어)이 두 기준열에 겹치는지 검사
    owner: dict[str, str] = {}
    seen_names: set[str] = set()
    for c in columns:
        if c.name in seen_names:
            raise ScenarioError(f"{where}에 기준명 '{c.name}'이 두 번 나옵니다.")
        seen_names.add(c.name)
        for n in c.names:
            key_n = normalize_header(n)
            if not key_n:
                raise ScenarioError(f"{where}의 '{c.name}' 열 이름 '{n}'에 글자나 숫자가 없습니다.")
            other = owner.get(key_n)
            if other is not None and other != c.name:
                raise ScenarioError(
                    f"{where}에서 이름 '{n}'이 '{other}' 열과 '{c.name}' 열에 함께 쓰였습니다 "
                    "(공백·대소문자·특수문자를 무시하고 비교). 한 열에만 남겨 주세요.")
            owner[key_n] = c.name

    dup = data[K_DUP]
    if dup is None:
        dup = []
    if not isinstance(dup, list):
        raise ScenarioError(f"{where}의 '{K_DUP}'은 [열1, 열2] 목록으로 적어 주세요 (없으면 []).")
    dup = [str(d).strip() for d in dup]
    for d in dup:
        if d not in seen_names:
            raise ScenarioError(f"{where}의 '{K_DUP}'에 있는 '{d}'이 columns의 기준명에 없습니다.")
    if len(set(dup)) != len(dup):
        raise ScenarioError(f"{where}의 '{K_DUP}'에 같은 열이 두 번 들어 있습니다.")

    desc = data.get(K_DESC) or ""
    return Scenario(key=key, name=str(name).strip(), description=str(desc).strip(),
                    columns=columns, dup_keys=dup)


def load_scenario(path: str | Path) -> Scenario:
    """YAML 파일 하나를 읽는다."""
    path = Path(path)
    where = f"시나리오 파일 '{path.name}'"
    if not path.is_file():
        raise ScenarioError(f"{where}을 찾을 수 없습니다.")
    if path.stat().st_size > MAX_YAML_BYTES:
        raise ScenarioError(f"{where}이 너무 큽니다 (최대 {MAX_YAML_BYTES // 1024}KB).")
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        raise ScenarioError(f"{where}을 UTF-8로 읽을 수 없습니다. UTF-8로 저장해 주세요.") from None
    try:
        data = yaml.safe_load(text)
    except yaml.YAMLError as e:
        mark = getattr(e, "problem_mark", None)
        at = f" ({mark.line + 1}번째 줄 근처)" if mark is not None else ""
        raise ScenarioError(f"{where}의 YAML 문법이 잘못되었습니다{at}. 들여쓰기와 콜론(:)을 확인해 주세요.") from None
    sc = parse_scenario(data, key=path.stem, where=where)
    sc.path = path
    return sc


def find_scenario(key: str, scenarios_dir: str | Path) -> Scenario:
    """scenarios/ 폴더에서 이름(확장자 제외)으로 시나리오를 찾는다."""
    if not SCENARIO_KEY_RE.match(key):
        raise ScenarioError(f"시나리오 이름 '{key}'에는 영문·숫자·_·- 만 쓸 수 있습니다.")
    return load_scenario(Path(scenarios_dir) / f"{key}.yaml")


def list_scenarios(scenarios_dir: str | Path) -> list[Scenario]:
    """폴더 안의 모든 시나리오 (형식이 틀린 파일이 있으면 그 오류를 그대로 낸다)."""
    return [load_scenario(p) for p in sorted(Path(scenarios_dir).glob("*.yaml"))]
