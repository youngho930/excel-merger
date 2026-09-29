"""클라우드 자유 양식 체험에서 나온 문제의 회귀 테스트 (2026-09-29).

1. 결과 열이 말없이 빠짐: 머리글 행을 바꿨다가 되돌리면 사용자가 빼지 않은 열이 결과에서 사라졌다.
2. 적용하지 않은 AI 묶기 추천이 남은 채 실행할 수 있었다 -> 경고 + '추천 적용하고 실행'.
3. AI가 추천하지 않은 쌍(접수일/등록일)을 직접 묶기 -> '사용자 지정' 표시, 풀 수 있음.
4. '자유 양식 체험'으로 들어오면 정답지대로 필수·중복 기준을 미리 켠다 (직접 올린 파일은 켜지 않음).
"""

import json

from streamlit.testing.v1 import AppTest

from answer_key import ROOT, compare, load_free_answer
from engine import ai_match
from engine.free_form import regroup

import app

APP = str(ROOT / "app.py")
ANSWER = load_free_answer()
AI_PAIRS = [("고객명", "성명"), ("연락처", "전화번호"), ("제품명", "모델"), ("증상", "고장내용")]   # 실제 Gemini 가 준 4건
HARNESS = '''
import sys
sys.path.insert(0, {root!r})
import streamlit as st
import app
app.init_state()
st.session_state.scenario = app.FREE_OPTION
if not st.session_state.get("tried"):
    st.session_state.tried = True
    app.load_files(None, [(n, open(p, "rb").read()) for n, p in {items!r}], "upload")
app.show_messages()
if st.session_state.ff_tables:
    app.render_free_steps(set())
'''


def free_demo() -> AppTest:
    at = AppTest.from_file(APP, default_timeout=120)
    at.run()
    at.button(key="free_demo_btn").click().run()
    assert not at.exception, at.exception
    return at


def groups_of(at):
    return regroup(at.session_state["ff_tables"], at.session_state["ff_merges"])


def by_label(at):
    return {g.label: g for g in groups_of(at)}


def key(at, kind, label):
    return f"ff_{kind}_{at.session_state['load_id']}_{by_label(at)[label].gid}"


def order_labels(at):
    labels = {g.gid: g.label for g in groups_of(at)}
    return [labels[g] for g in at.session_state["ff_order"]]


def fake_ai(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "test-key")

    def fake(prompt, api_key, model, timeout):
        data = json.loads(prompt.rsplit("<data>", 1)[1].split("</data>")[0])
        ids = {n: c["id"] for c in data["columns"] for n in c["names"]}
        return json.dumps({"groups": [[ids[a], ids[b]] for a, b in AI_PAIRS]})
    monkeypatch.setattr(ai_match, "call_gemini", fake)


# ------------------------------------------------------------------ 1. 결과 열이 말없이 빠지지 않음
def test_columns_survive_header_row_change_and_revert():
    at = free_demo()
    before = set(order_labels(at))
    assert "고객명" in before and "담당기사" in before
    fi = next(i for i, t in enumerate(at.session_state["ff_tables"]) if "강남" in t.file_name)
    hr = f"ff_hr_{at.session_state['load_id']}_{fi}"
    at.number_input(key=hr).set_value(2).run()                  # 실수로 바뀜 (데이터 행이 머리글로)
    at.number_input(key=hr).set_value(1).run()                  # 되돌림
    assert not at.exception, at.exception
    assert set(order_labels(at)) == before                      # 예전: '담당기사'가 말없이 빠졌다
    assert {g.gid for g in groups_of(at)} == set(at.session_state["ff_order"])


def test_all_groups_in_result_after_ai_apply_and_run(monkeypatch):
    fake_ai(monkeypatch)
    at = free_demo()
    at.button(key="ff_ai_btn").click().run()
    at.button(key="ff_apply_props").click().run()
    assert {g.gid for g in groups_of(at)} == set(at.session_state["ff_order"])
    assert "고객명" in order_labels(at)
    at.button(key="run_btn").click().run()
    assert "고객명" in at.session_state["result"].scenario.column_names


def test_only_columns_the_user_removed_stay_out():
    at = free_demo()
    pick = f"ff_pick_{at.session_state['load_id']}"
    keep = [g for g in at.session_state["ff_order"] if g != by_label(at)["담당기사"].gid]
    at.multiselect(key=pick).set_value(keep).run()             # 사용자가 직접 뺌
    assert "담당기사" not in order_labels(at)
    assert any("결과에 넣지 않은 열: 담당기사" in c.value.replace("\\", "") for c in at.caption)
    fi = next(i for i, t in enumerate(at.session_state["ff_tables"]) if "강남" in t.file_name)
    hr = f"ff_hr_{at.session_state['load_id']}_{fi}"
    at.number_input(key=hr).set_value(2).run()
    at.number_input(key=hr).set_value(1).run()
    assert "담당기사" not in order_labels(at)                    # 직접 뺀 열은 머리글을 다시 읽어도 빠진 채로
    at.multiselect(key=pick).set_value(at.session_state["ff_order"] + [by_label(at)["담당기사"].gid]).run()
    assert "담당기사" in order_labels(at)                        # 다시 넣으면 들어간다


# ------------------------------------------------------------------ 2. 적용하지 않은 추천 경고
def test_pending_proposals_warn_and_apply_then_run(monkeypatch):
    fake_ai(monkeypatch)
    at = free_demo()
    at.button(key="ff_ai_btn").click().run()
    warn = " ".join(w.value for w in at.warning)
    assert "적용하지 않은 묶기 추천이 4건 있습니다" in warn
    assert at.button(key="ff_apply_run").label == "추천 적용하고 실행"
    assert not at.button(key="run_btn").disabled               # 그냥 실행도 가능
    at.button(key="ff_apply_run").click().run()
    assert not at.exception, at.exception
    groups = by_label(at)
    assert groups["고객명"].origin == "AI 추천" and "성명" in groups["고객명"].names
    assert at.session_state["result"] is not None               # 묶은 뒤 바로 실행됐다
    assert "고객명" in at.session_state["result"].scenario.column_names
    assert "성명" not in at.session_state["result"].scenario.column_names
    assert not [w for w in at.warning if "적용하지 않은 묶기 추천" in w.value]
    assert "ff_apply_run" not in [b.key for b in at.button]


def test_pending_proposals_can_be_ignored(monkeypatch):
    fake_ai(monkeypatch)
    at = free_demo()
    at.button(key="ff_ai_btn").click().run()
    at.button(key="run_btn").click().run()                      # 추천을 적용하지 않고 실행
    assert not at.exception, at.exception
    cols = at.session_state["result"].scenario.column_names
    assert "고객명" in cols and "성명" in cols                  # 묶지 않았으므로 따로따로


# ------------------------------------------------------------------ 3. 직접 묶기
def test_direct_merge_is_user_defined_and_can_be_split():
    at = free_demo()
    g = by_label(at)
    at.multiselect(key=f"ff_merge_pick_{at.session_state['load_id']}").set_value(
        [g["접수일"].gid, g["등록일"].gid])
    at.button(key="ff_merge_btn").click().run()
    assert not at.exception, at.exception
    merged = by_label(at)["접수일"]
    assert merged.origin == "사용자 지정" and "등록일" in merged.names and "등록일" not in by_label(at)
    table = next(d.value for d in at.dataframe if "묶은 방법" in d.value.columns)   # 묶음 표
    assert table.set_index("묶음").loc["접수일", "묶은 방법"] == "사용자 지정"
    at.selectbox(key=f"ff_split_pick_{at.session_state['load_id']}").set_value(merged.gid).run()
    at.button(key="ff_split_btn").click().run()
    assert "등록일" in by_label(at)                              # 풀었다


# ------------------------------------------------------------------ 4. 체험 미리 설정
def test_demo_preset_matches_answer_key():
    at = free_demo()
    preset = {c["name"]: c for c in ANSWER["config"]["columns"]}
    for name in ("접수일", "고객명", "연락처", "제품명"):
        assert at.session_state[key(at, "req", name)] == preset[name]["required"], name
        assert at.session_state[key(at, "dup", name)] == preset[name]["dup"], name
    assert at.session_state[key(at, "dup", "등록일")] is False    # 짝 열은 묶으면 설정을 이어받는다
    # 5쌍만 묶고 설정은 건드리지 않은 채 실행하면 정답지 오류가 모두 나온다
    for a, b in [("고객명", "성명"), ("연락처", "전화번호"), ("제품명", "모델"), ("증상", "고장내용"), ("접수일", "등록일")]:
        g = by_label(at)
        at.multiselect(key=f"ff_merge_pick_{at.session_state['load_id']}").set_value([g[a].gid, g[b].gid])
        at.button(key="ff_merge_btn").click().run()
    at.button(key="run_btn").click().run()
    assert not at.exception, at.exception
    cmp = compare(at.session_state["result"].issues, ANSWER["errors"])
    assert cmp.ok, cmp.report()


def test_uploaded_files_get_no_preset():
    items = [(p.name, str(p)) for p in app.free_sample_files()]
    at = AppTest.from_string(HARNESS.format(root=str(ROOT), items=items), default_timeout=120)
    at.run()
    assert not at.exception, at.exception
    assert at.session_state["load_source"] == "upload"
    lid = at.session_state["load_id"]
    flags = [at.session_state[f"ff_{kind}_{lid}_{g.gid}"] for g in groups_of(at) for kind in ("req", "dup")]
    assert len(flags) == 2 * len(groups_of(at)) and not any(flags)   # 직접 올린 파일은 아무것도 켜지 않는다
