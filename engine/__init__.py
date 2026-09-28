"""엑셀 자동 취합·검증 엔진.

시나리오 이름·열 이름은 전혀 모른다. 모든 설정은 scenarios/*.yaml 에서 읽는다.

    from engine import load_scenario, prepare, pre_run_warnings, execute, write_result
    sc = load_scenario("scenarios/stock_count.yaml")
    plans = prepare(sc, paths)            # 읽기 + 머리글 탐지 + 열 매칭
    warnings = pre_run_warnings(plans)    # 실행 전 경고 (화면에서 보여주고 매칭 수정)
    result = execute(sc, plans)           # 정규화 + 검증 5종 + 취합
    path = write_result(result, "output")
"""

from .matching import MatchResult, match_columns
from .merge import FilePlan, MergeResult, execute, pre_run_warnings, prepare
from .reader import ReadError, read_table
from .scenario import Scenario, ScenarioError, find_scenario, list_scenarios, load_scenario
from .validate import KINDS, Issue
from .writer import write_result

__all__ = [
    "FilePlan", "Issue", "KINDS", "MatchResult", "MergeResult", "ReadError", "Scenario",
    "ScenarioError", "execute", "find_scenario", "list_scenarios", "load_scenario",
    "match_columns", "pre_run_warnings", "prepare", "read_table", "write_result",
]
