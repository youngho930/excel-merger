import pytest

from answer_key import ROOT
from engine.scenario import ScenarioError, find_scenario, list_scenarios, load_scenario

BASE = """name: 테스트
description: 설명
columns:
  - 기준명: 코드
    동의어: [Code]
    필수: true
    형식: 문자
  - 기준명: 수량
    형식: 정수
    범위: [0, 10]
중복기준: [코드]
"""


def write(tmp_path, text, name="t.yaml"):
    p = tmp_path / name
    p.write_text(text, encoding="utf-8")
    return p


def test_real_scenarios_load():
    scs = list_scenarios(ROOT / "scenarios")
    assert {s.key for s in scs} == {"incoming_inspection", "stock_count", "monthly_report"}


def test_base_loads(tmp_path):
    sc = load_scenario(write(tmp_path, BASE))
    assert sc.column_names == ["코드", "수량"] and sc.dup_keys == ["코드"]
    assert sc.column("수량").range == (0, 10)


@pytest.mark.parametrize("old,new,msg", [
    ("name: 테스트\n", "", "필수 항목 'name'"),
    ("중복기준: [코드]\n", "", "필수 항목 '중복기준'"),
    ("    형식: 문자\n", "", "'형식'이 없습니다"),
    ("형식: 문자", "형식: 글자", "형식 '글자'은 쓸 수 없습니다"),
    ("  - 기준명: 수량\n", "  - 이름: 수량\n", "알 수 없는 키"),
    ("필수: true", "필수: 예", "true 또는 false"),
    ("범위: [0, 10]", "범위: [10, 0]", "최소(10)가 최대(0)보다 큽니다"),
    ("범위: [0, 10]", "범위: [0]", "[최소, 최대]"),
    ("중복기준: [코드]", "중복기준: [품목코드]", "'품목코드'이 columns의 기준명에 없습니다"),
    ("동의어: [Code]", "동의어: [수 량]", "'코드' 열과 '수량' 열에 함께 쓰였습니다"),
    ("description: 설명\n", "description: 설명\nextra: 1\n", "알 수 없는 키가 있습니다: extra"),
])
def test_invalid_yaml_rejected_in_korean(tmp_path, old, new, msg):
    assert old in BASE
    with pytest.raises(ScenarioError) as e:
        load_scenario(write(tmp_path, BASE.replace(old, new)))
    assert msg in str(e.value)


def test_range_on_text_rejected(tmp_path):
    text = BASE.replace("    형식: 문자\n", "    형식: 문자\n    범위: [0, 1]\n")
    with pytest.raises(ScenarioError, match="정수·실수 열만"):
        load_scenario(write(tmp_path, text))


def test_yaml_syntax_error(tmp_path):
    with pytest.raises(ScenarioError, match="YAML 문법"):
        load_scenario(write(tmp_path, "name: [깨짐\ncolumns: x"))


def test_empty_file(tmp_path):
    with pytest.raises(ScenarioError, match="비어 있거나"):
        load_scenario(write(tmp_path, ""))


def test_find_scenario_rejects_path_tricks():
    with pytest.raises(ScenarioError, match="영문"):
        find_scenario("../CLAUDE", ROOT / "scenarios")
