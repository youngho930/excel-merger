import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


@pytest.fixture(autouse=True)
def _no_real_secrets_or_ai(monkeypatch):
    """테스트는 로컬 .streamlit/secrets.toml(실제 키)을 읽지 않고, 실제 Gemini도 부르지 않는다.

    - AppTest 는 at.secrets 를 비워 두면 실제 파일의 st.secrets 를 그대로 쓴다. 빈 secrets 로 바꿔 끼운다.
      키가 필요한 테스트는 환경변수(GEMINI_API_KEY=test-key)로 넣는다.
    - 가짜 응답을 끼우지 않은 테스트가 실제 API를 부르면 바로 실패하게 한다 (사용자 키·할당량 보호).
    """
    import streamlit as st
    from streamlit.runtime.secrets import Secrets

    from engine import ai_match

    empty = Secrets()
    empty._secrets = {}
    monkeypatch.setattr(st, "secrets", empty)

    def refuse(*args, **kwargs):
        raise AssertionError("테스트에서 실제 Gemini API를 부르면 안 됩니다. call_gemini 를 가짜로 바꿔 끼우세요.")
    monkeypatch.setattr(ai_match, "call_gemini", refuse)


@pytest.fixture(autouse=True)
def _reset_ai_budget():
    """앱 전체 AI 호출 한도(분당·하루)는 프로세스 전체가 공유하므로 테스트마다 비운다."""
    from engine import ai_match
    ai_match.GLOBAL_BUDGET.reset()
    yield
    ai_match.GLOBAL_BUDGET.reset()
