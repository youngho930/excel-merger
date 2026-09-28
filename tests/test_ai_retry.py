"""AI 호출 오류 처리: 일시적 오류 재시도, 원인별 문구, 세션 횟수 차감 (가짜 응답, 실제 API는 부르지 않음).

실제 원인(2026-09-29): 앱에서 "AI를 호출하지 못했습니다"만 보였는데, 실제로는 Google 쪽 503 과부하였다.
"""

import json

import pytest
import yaml
from streamlit.testing.v1 import AppTest

from answer_key import ROOT
from engine import ai_match, load_scenario, prepare

APP = str(ROOT / "app.py")
DEMO_DIR = ROOT / "samples" / "ai_demo"
MANIFEST = yaml.safe_load((DEMO_DIR / "demo.yaml").read_text(encoding="utf-8"))
NOTICE = "AI 없이 동의어 매칭만 사용 중"


class ApiError(Exception):
    """google-genai APIError 흉내: HTTP 상태 코드를 code 속성에 둔다."""

    def __init__(self, code, message="secret-detail-should-not-show"):
        super().__init__(f"{code} {message}")
        self.code = code


def correct_reply(prompt):
    data = json.loads(prompt.rsplit("<data>", 1)[1].split("</data>")[0])
    cols = {c["name"]: c["col"] for c in data["files"][0]["columns"]}
    return json.dumps({"matches": [{"file": "F1", "standard": std, "col": cols[src]}
                                   for src, std in MANIFEST["ai_targets"].items()]}, ensure_ascii=False)


class Scripted:
    """호출될 때마다 정해 둔 순서대로 예외를 내거나 정답을 돌려준다."""

    def __init__(self, *steps):
        self.steps, self.calls = list(steps), 0

    def __call__(self, prompt, api_key, model, timeout):
        step = self.steps[min(self.calls, len(self.steps) - 1)]
        self.calls += 1
        if isinstance(step, BaseException):
            raise step
        return correct_reply(prompt)


@pytest.fixture
def sleeps(monkeypatch):
    waited = []
    monkeypatch.setattr(ai_match, "_sleep", waited.append)
    ai_match.GLOBAL_BUDGET.reset()
    yield waited
    ai_match.GLOBAL_BUDGET.reset()


def demo_plans():
    sc = load_scenario(ROOT / "scenarios" / f"{MANIFEST['scenario']}.yaml")
    return sc, prepare(sc, sorted(DEMO_DIR.glob("*.xlsx")))


# ------------------------------------------------------------------ 엔진
def test_503_then_success_retries_with_1s_then_3s(sleeps):
    sc, plans = demo_plans()
    fake = Scripted(ApiError(503), ApiError(503), "ok")
    out = ai_match.recommend(sc, plans, api_key="k", caller=fake, budget=ai_match.CallBudget())
    assert fake.calls == 3 and sleeps == [1, 3]
    assert out.applied == len(MANIFEST["ai_targets"])


def test_429_is_retried_too(sleeps):
    sc, plans = demo_plans()
    fake = Scripted(ApiError(429), "ok")
    out = ai_match.recommend(sc, plans, api_key="k", caller=fake, budget=ai_match.CallBudget())
    assert fake.calls == 2 and sleeps == [1] and out.applied > 0


def test_all_retries_fail_raises_busy_with_clear_message(sleeps):
    sc, plans = demo_plans()
    fake = Scripted(ApiError(503))
    with pytest.raises(ai_match.AiBusy) as e:
        ai_match.recommend(sc, plans, api_key="k", caller=fake, budget=ai_match.CallBudget())
    assert fake.calls == 3 and sleeps == [1, 3]          # 처음 1번 + 재시도 2번
    assert str(e.value) == "Google AI 서버가 붐빕니다. 잠시 후 다시 눌러 주세요."
    assert all(m.method != "AI 추천" for m in plans[0].match.matches)


def test_retries_count_against_app_wide_budget(sleeps):
    sc, plans = demo_plans()
    budget = ai_match.CallBudget()
    with pytest.raises(ai_match.AiBusy):
        ai_match.recommend(sc, plans, api_key="k", caller=Scripted(ApiError(503)), budget=budget)
    assert len(budget._calls) == 3                       # 실제 호출 3번은 앱 전체 한도에 센다


@pytest.mark.parametrize("error, kind, text", [
    (ApiError(503), ai_match.AiBusy, "Google AI 서버가 붐빕니다"),
    (ApiError(429), ai_match.AiBusy, "사용 한도에 잠시 걸렸습니다"),
    (ApiError(401), ai_match.AiError, "AI 키가 올바르지 않거나"),
    (ApiError(400, "API key not valid"), ai_match.AiError, "AI 키가 올바르지 않거나"),
    (ApiError(404), ai_match.AiError, "AI 모델 이름을 찾을 수 없습니다"),
    (ConnectionError("socket"), ai_match.AiError, "네트워크 연결을 확인해 주세요"),
    (type("ConnectError", (Exception,), {})("socket"), ai_match.AiError, "네트워크 연결을 확인해 주세요"),
    (ModuleNotFoundError("No module named 'google'"), ai_match.AiSetupError, "구성 요소가 설치되지 않았습니다"),
    (UnicodeEncodeError("ascii", "키“", 1, 2, "bad"), ai_match.AiSetupError, "AI 키 형식이 올바르지 않습니다"),
    (TypeError("socket"), ai_match.AiError, "예상하지 못한 오류"),
])
def test_errors_are_explained_by_cause(error, kind, text):
    err = ai_match.explain_failure(error)
    assert type(err) is kind and text in str(err)
    assert "secret-detail" not in str(err) and "socket" not in str(err)


def test_install_problem_is_not_network_and_not_counted(monkeypatch, sleeps):
    # 예전: ImportError 도 "네트워크 연결을 확인해 주세요"로 나오고 세션 횟수가 깎였다
    monkeypatch.setenv(ai_match.KEY_NAME, "test-key")

    def missing(*a):
        raise ModuleNotFoundError("No module named 'google'")
    monkeypatch.setattr(ai_match, "call_gemini", missing)
    at = open_demo()
    at.button(key="ai_btn").click().run()
    msg = texts(at.warning)
    assert "AI 기능 구성 요소가 설치되지 않았습니다" in msg and NOTICE in msg and "네트워크" not in msg
    assert "남은 AI 호출: 5/5회" in texts(at.caption)
    assert sleeps == []                                     # 설치 문제는 재시도하지 않는다


@pytest.mark.parametrize("bad_key", ["“AIzaSyExample”", "AIza SyExample", "AIza키Example", "AIzaExample\n"])
def test_bad_key_format_is_caught_before_calling(bad_key, sleeps):
    sc, plans = demo_plans()
    fake = Scripted("ok")
    with pytest.raises(ai_match.AiSetupError, match="AI 키 형식이 올바르지 않습니다"):
        ai_match.recommend(sc, plans, api_key=bad_key, caller=fake, budget=ai_match.CallBudget())
    assert fake.calls == 0


class Captured:
    def __init__(self):
        import logging
        self.records = []
        self.handler = logging.Handler()
        self.handler.emit = lambda r: self.records.append(r.getMessage())

    def __enter__(self):
        ai_match.log.addHandler(self.handler)
        return self.records

    def __exit__(self, *exc):
        ai_match.log.removeHandler(self.handler)


def test_failures_are_logged_with_key_masked(sleeps):
    sc, plans = demo_plans()
    key = "AIzaSyTESTKEY0123456789abcdefghijkl"
    leaky = ApiError(503, f"quota for key {key} url https://x/?key={key}")
    with Captured() as logs, pytest.raises(ai_match.AiBusy) as e:
        ai_match.recommend(sc, plans, api_key=key, caller=Scripted(leaky), budget=ai_match.CallBudget())
    assert len(logs) == 3                                    # 시도마다 한 줄
    assert all("ApiError" in line and "503" in line for line in logs)
    assert all(key not in line and "***" in line for line in logs)
    assert key not in str(e.value) and "503" not in str(e.value)   # 화면 문구에는 원본 메시지 없음


def test_success_after_retry_is_logged(sleeps):
    sc, plans = demo_plans()
    with Captured() as logs:
        ai_match.recommend(sc, plans, api_key="k", caller=Scripted(ApiError(503), "ok"),
                           budget=ai_match.CallBudget())
    assert any("실패 (시도 1/3)" in line for line in logs) and any("성공 (시도 2/3" in line for line in logs)


def test_redact_masks_key_shapes():
    assert ai_match.redact("x AIzaSyABCDEFGHIJKLMNOP y") == "x *** y"
    assert ai_match.redact("https://h/v1?key=abc123&x=1") == "https://h/v1?key=***&x=1"
    assert ai_match.redact("plain-secret here", "plain-secret") == "*** here"


def test_non_transient_error_is_not_retried(sleeps):
    sc, plans = demo_plans()
    fake = Scripted(ApiError(404))
    with pytest.raises(ai_match.AiError, match="모델 이름"):
        ai_match.recommend(sc, plans, api_key="k", caller=fake, budget=ai_match.CallBudget())
    assert fake.calls == 1 and sleeps == []


# ------------------------------------------------------------------ 화면
def open_demo() -> AppTest:
    at = AppTest.from_file(APP, default_timeout=120)
    at.run()
    at.button(key="ai_demo_btn_0").click().run()
    assert not at.exception, at.exception
    return at


def texts(elements):
    """화면 문구. 앱은 마크다운 특수문자를 이스케이프하므로 비교 전에 백슬래시를 뺀다."""
    return " ".join(e.value for e in elements).replace("\\", "")


def block_with_keys(node, keys):
    """직접 자식 버튼의 key 가 keys 를 모두 포함하는 블록을 찾는다."""
    children = list(getattr(node, "children", {}).values())
    if keys <= {getattr(c, "key", None) for c in children}:
        return node
    for c in children:
        found = block_with_keys(c, keys)
        if found is not None:
            return found
    return None


def test_app_503_then_success_fills_ai_recommendations(monkeypatch, sleeps):
    monkeypatch.setenv(ai_match.KEY_NAME, "test-key")
    fake = Scripted(ApiError(503), "ok")
    monkeypatch.setattr(ai_match, "call_gemini", fake)
    at = open_demo()
    at.button(key="ai_btn").click().run()
    assert not at.exception, at.exception
    assert fake.calls == 2 and sleeps == [1]
    assert "AI 추천 5건을 확인표에" in texts(at.success)
    assert "남은 AI 호출: 4/5회" in texts(at.caption)      # 성공한 시도는 1회로 센다


def test_after_ai_fills_everything_the_notice_matches_the_situation(monkeypatch, sleeps):
    # 예전: AI 추천을 받은 뒤에도 "AI 추천이 필요 없습니다"가 떠서 헷갈렸다
    monkeypatch.setenv(ai_match.KEY_NAME, "test-key")
    monkeypatch.setattr(ai_match, "call_gemini", Scripted("ok"))
    at = open_demo()
    at.button(key="ai_btn").click().run()
    captions = texts(at.caption)
    assert "모든 열이 매칭되었습니다" in captions and "AI 추천이 필요 없습니다" not in captions


def test_samples_matched_by_synonyms_say_ai_not_needed():
    at = AppTest.from_file(APP, default_timeout=120)
    at.run()
    at.button(key="sample_btn").click().run()
    captions = texts(at.caption)
    assert "AI 추천이 필요 없습니다" in captions and "모든 열이 매칭되었습니다" not in captions


def test_app_busy_message_and_session_count_not_used(monkeypatch, sleeps):
    monkeypatch.setenv(ai_match.KEY_NAME, "test-key")
    fake = Scripted(ApiError(503))
    monkeypatch.setattr(ai_match, "call_gemini", fake)
    at = open_demo()
    at.button(key="ai_btn").click().run()
    assert not at.exception and not at.error
    msg = texts(at.warning)
    assert "Google AI 서버가 붐빕니다. 잠시 후 다시 눌러 주세요." in msg and NOTICE in msg
    assert "secret-detail" not in msg
    assert "남은 AI 호출: 5/5회" in texts(at.caption)      # 일시적 오류는 세션 횟수에서 빼지 않는다
    assert all(cm.method != "AI 추천" for cm in at.session_state["plans"][0].match.matches)


def test_app_no_key_message(monkeypatch):
    monkeypatch.delenv(ai_match.KEY_NAME, raising=False)
    at = open_demo()
    assert NOTICE in texts(at.info) and "GEMINI_API_KEY" in texts(at.info)
    assert at.button(key="ai_btn").disabled
    # 버튼을 거치지 않고 엔진에 키 없이 부르면 원인이 분명한 문구
    sc, plans = demo_plans()
    with pytest.raises(ai_match.AiNotSent, match="AI 키\\(GEMINI_API_KEY\\)가 설정되지 않았습니다"):
        ai_match.recommend(sc, plans, api_key=None)


def test_demo_button_sits_next_to_sample_button():
    at = AppTest.from_file(APP, default_timeout=120)
    at.run()
    keys = [b.key for b in at.button]
    assert keys.index("ai_demo_btn_0") == keys.index("sample_btn") + 1
    # 두 버튼이 같은 블록(가로 컨테이너)의 바로 아래에 나란히 있다 (예전: 서로 다른 같은 너비 칸)
    block = block_with_keys(at.main, {"sample_btn", "ai_demo_btn_0"})
    assert block is not None
