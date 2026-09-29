"""엔진 결과를 정답지(samples/expected_errors.json)와 대조하는 도우미.

대조 기준
- 일반 오류: (파일, 행, 기준열, 오류 종류). 원본 열 이름이 정답지와 같은지도 따로 확인한다.
- 중복 행: (파일, 행, "중복 행"). 정답지는 뒤에 나온 행만 적었으므로,
  정답지 중복 항목의 상대 행(duplicate_of)에 엔진이 낸 "중복 행"은 정답으로 본다.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ANSWER_JSON = ROOT / "samples" / "expected_errors.json"
DUP = "중복 행"


def load_answer() -> dict:
    return json.loads(ANSWER_JSON.read_text(encoding="utf-8"))


def _key(file, row, standard, kind):
    return (file, row, DUP) if kind == DUP else (file, row, standard, kind)


@dataclass
class Comparison:
    missed: list = field(default_factory=list)          # 정답지에 있는데 엔진이 놓친 것
    false_positive: list = field(default_factory=list)  # 정답지에 없는데 엔진이 잡은 것
    wrong_column: list = field(default_factory=list)    # 잡았지만 원본 열 이름이 다른 것
    matched: int = 0
    accepted_partners: int = 0

    @property
    def ok(self) -> bool:
        return not (self.missed or self.false_positive or self.wrong_column)

    def report(self) -> str:
        lines = [f"정답 일치 {self.matched}건, 중복 상대 행 {self.accepted_partners}건"]
        if self.missed:
            lines.append("놓친 오류:")
            lines += [f"  - {e['file']} {e['row']}행 {e['column']}({e['standard']}) {e['kind']} "
                      f"값={e['shown']}" for e in self.missed]
        if self.false_positive:
            lines.append("잘못 잡은 오류:")
            lines += [f"  - {i.file} {i.row}행 {i.column}({i.standard}) {i.kind} 값={i.value!r} "
                      f"설명={i.message}" for i in self.false_positive]
        if self.wrong_column:
            lines.append("열 이름이 다른 오류:")
            lines += [f"  - {e['file']} {e['row']}행 정답 '{e['column']}' / 엔진 '{i.column}'"
                      for e, i in self.wrong_column]
        return "\n".join(lines)


def compare(issues, expected: list[dict]) -> Comparison:
    exp = {_key(e["file"], e["row"], e["standard"], e["kind"]): e for e in expected}
    partners = {(e["duplicate_of"]["file"], e["duplicate_of"]["row"], DUP)
                for e in expected if e.get("duplicate_of")}
    got = {}
    for i in issues:
        got.setdefault(_key(i.file, i.row, i.standard, i.kind), i)
    cmp = Comparison()
    for k, e in exp.items():
        if k not in got:
            cmp.missed.append(e)
            continue
        cmp.matched += 1
        if e["kind"] != DUP and got[k].column != e["column"]:
            cmp.wrong_column.append((e, got[k]))
    for k, i in got.items():
        if k in exp:
            continue
        if k in partners:
            cmp.accepted_partners += 1
            continue
        cmp.false_positive.append(i)
    return cmp


def run_merge_main(argv: list[str]) -> tuple[int, str]:
    """scripts/run_merge.py 의 main()을 실행하고 (종료 코드, 표준 출력)을 돌려준다."""
    import contextlib
    import importlib.util
    import io

    spec = importlib.util.spec_from_file_location("run_merge", ROOT / "scripts" / "run_merge.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        code = module.main(argv)
    return code, buf.getvalue()


# ------------------------------------------------------------------ 자유 양식 정답지 (기존 37건과 별개)
FREE_ANSWER_JSON = ROOT / "samples" / "free_form_expected_errors.json"


def load_free_answer() -> dict:
    return json.loads(FREE_ANSWER_JSON.read_text(encoding="utf-8"))


def free_form_scenario(config: dict, tables, key: str = "free_form"):
    """정답지의 config(결과 열·묶을 원본 이름·필수·중복기준)대로 자유 양식 Scenario 를 만든다.

    화면에서 사용자가 하는 일(묶기 -> 열 고르기 -> 날짜 판단)을 그대로 따른다.
    """
    from engine.free_form import USER, ColumnChoice, build_scenario, check_dates, regroup
    from engine.normalize import normalize_header

    merges = [({normalize_header(s) for s in c["sources"]}, USER) for c in config["columns"] if len(c["sources"]) > 1]
    groups = regroup(tables, merges)
    by_norm = {n: g for g in groups for n in g.norms}
    choices = [ColumnChoice(by_norm[normalize_header(c["sources"][0])].gid, c["name"], c["required"], c["dup"])
               for c in config["columns"]]
    dates = {g.gid for g in groups if check_dates(tables, g).is_date}
    return build_scenario(groups, choices, dates, name=config["name"], key=key), groups
