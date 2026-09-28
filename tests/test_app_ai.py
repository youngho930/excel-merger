"""화면의 AI 매칭 추천 — AppTest + 가짜 Gemini (실제 API는 부르지 않는다).

AppTest 는 같은 프로세스에서 app.py 를 돌리므로 engine.ai_match.call_gemini 를 바꿔 끼우면
화면의 "AI에게 매칭 추천 받기"가 가짜 응답을 받는다.
"""

import json

import pytest
import yaml
from streamlit.testing.v1 import AppTest

from answer_key import ROOT
from engine import ai_match

APP = str(ROOT / "app.py")
MANIFEST = yaml.safe_load((ROOT / "samples" / "ai_demo" / "demo.yaml").read_text(encoding="utf-8"))


class FakeGemini:
    def __init__(self, reply=None, error=None):
        self.reply, self.error, self.prompts = reply, error, []

    def __call__(self, prompt, api_key, model, timeout):
        self.prompts.append(prompt)
        if self.error:
            raise self.error
        return self.reply(prompt) if callable(self.reply) else self.reply


def correct_reply(prompt):
    """프롬프트의 <data> 에서 열 번호를 찾아 정답 짝으로 답한다."""
    data = json.loads(prompt.rsplit("<data>", 1)[1].split("</data>")[0])
    cols = {c["name"]: c["col"] for c in data["files"][0]["columns"]}
    return json.dumps({"matches": [{"file": "F1", "standard": std, "col": cols[src]}
                                   for src, std in MANIFEST["ai_targets"].items()]}, ensure_ascii=False)


@pytest.fixture
def no_key(monkeypatch):
    monkeypatch.delenv(ai_match.KEY_NAME, raising=False)


@pytest.fixture
def with_key(monkeypatch):
    monkeypatch.setenv(ai_match.KEY_NAME, "test-key")


def open_demo() -> AppTest:
    at = AppTest.from_file(APP, default_timeout=120)
    at.run()
    at.button(key="ai_demo_btn_0").click().run()
    assert not at.exception, at.exception
    assert at.selectbox(key="scenario").value == MANIFEST["scenario"]
    assert any("AI 매칭 체험 파일 1개를 읽었습니다" in s.value for s in at.success)
    return at


def methods(at):
    return {cm.standard: cm.method for cm in at.session_state["plans"][0].match.matches}


def texts(elements):
    return " ".join(e.value for e in elements)


def test_without_key_shows_notice_and_still_works(no_key):
    at = open_demo()
    assert "AI 없이 동의어 매칭만 사용 중" in texts(at.info)
    assert at.button(key="ai_btn").disabled
    assert "필수 열" in texts(at.warning)            # 동의어만으로는 필수 열을 못 찾는다
    at.button(key="run_btn").click().run()
    assert not at.exception and not at.error
    assert next(m.value for m in at.metric if m.label == "전체 오류") == "1건"   # 필수 열 누락 1건


def test_mock_ai_recommendation_then_run(with_key, monkeypatch):
    fake = FakeGemini(correct_reply)
    monkeypatch.setattr(ai_match, "call_gemini", fake)
    at = open_demo()
    assert not at.info or "AI 없이" not in texts(at.info)
    assert "남은 AI 호출: 5/5회" in texts(at.caption)
    assert any("열 이름만" in c.value for c in at.caption)     # 무엇이 전송되는지 표시

    at.button(key="ai_btn").click().run()
    assert not at.exception, at.exception
    assert len(fake.prompts) == 1 and "NP-3001" not in fake.prompts[0]   # 기본은 값 없이 이름만
    assert "AI 추천 5건을 확인표에" in texts(at.success)
    assert "남은 AI 호출: 4/5회" in texts(at.caption)
    got = methods(at)
    assert all(got[std] == "AI 추천" for std in MANIFEST["ai_targets"].values())
    assert "AI 추천" in texts(at.markdown)                     # 확인표에 "AI 추천" 표시
    assert "필수 열" not in texts(at.warning)

    # 드롭다운 값도 AI 추천대로 채워져 있고, 사용자가 해제할 수 있다
    load_id = at.session_state["load_id"]
    m = at.session_state["plans"][0].match
    si = next(i for i, cm in enumerate(m.matches) if cm.standard == "판정")
    box = at.selectbox(key=f"map_{load_id}_0_{si}")
    assert box.value == m.headers.index("합부")
    box.set_value(None).run()
    assert methods(at)["판정"] == "매칭 안 됨"
    box = at.selectbox(key=f"map_{load_id}_0_{si}")
    box.set_value(m.headers.index("합부")).run()
    assert methods(at)["판정"] == "사용자 지정"

    at.button(key="run_btn").click().run()
    assert not at.exception and not at.error
    assert next(x.value for x in at.metric if x.label == "전체 오류") == "2건"   # 체험 파일에 심은 오류 2건


def test_examples_are_sent_only_when_checked(with_key, monkeypatch):
    fake = FakeGemini(correct_reply)
    monkeypatch.setattr(ai_match, "call_gemini", fake)
    at = open_demo()
    assert "NP-3001" not in texts(at.code)                    # 미리 보기: 이름만
    at.checkbox(key="ai_examples").check().run()
    assert "NP-3001" in texts(at.code)                        # 미리 보기에 예시값
    assert any("예시값 3개" in c.value for c in at.caption)
    at.button(key="ai_btn").click().run()
    assert "NP-3001" in fake.prompts[0]


def test_unknown_standard_from_ai_is_dropped(with_key, monkeypatch):
    def reply(prompt):
        good = json.loads(correct_reply(prompt))["matches"][:1]
        return json.dumps({"matches": good + [{"file": "F1", "standard": "가짜 기준열", "col": 1}]},
                          ensure_ascii=False)
    monkeypatch.setattr(ai_match, "call_gemini", FakeGemini(reply))
    at = open_demo()
    at.button(key="ai_btn").click().run()
    assert "AI 추천 1건" in texts(at.success) and "기준열 목록에 없는 이름 1건" in texts(at.success)
    assert "가짜 기준열" not in {cm.standard for cm in at.session_state["plans"][0].match.matches}


@pytest.mark.parametrize("fake, expect", [
    (FakeGemini(error=RuntimeError("boom")), "AI를 호출하지 못했습니다"),
    (FakeGemini("{깨진 json"), "JSON"),
])
def test_failure_falls_back_to_synonyms(with_key, monkeypatch, fake, expect):
    monkeypatch.setattr(ai_match, "call_gemini", fake)
    at = open_demo()
    before = methods(at)
    at.button(key="ai_btn").click().run()
    assert not at.exception and not at.error
    msg = texts(at.warning)
    assert expect in msg and "AI 없이 동의어 매칭만 사용 중" in msg and "boom" not in msg
    assert methods(at) == before


def test_session_call_limit(with_key, monkeypatch):
    fake = FakeGemini(error=RuntimeError("down"))
    monkeypatch.setattr(ai_match, "call_gemini", fake)
    at = open_demo()
    for _ in range(ai_match.MAX_CALLS_PER_SESSION):
        assert not at.button(key="ai_btn").disabled
        at.button(key="ai_btn").click().run()
    assert len(fake.prompts) == ai_match.MAX_CALLS_PER_SESSION
    assert at.button(key="ai_btn").disabled
    assert "남은 AI 호출: 0/5회" in texts(at.caption)
    assert "AI 호출 횟수(5회)를 모두 썼습니다" in texts(at.warning)
