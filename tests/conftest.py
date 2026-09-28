import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


@pytest.fixture(autouse=True)
def _reset_ai_budget():
    """앱 전체 AI 호출 한도(분당·하루)는 프로세스 전체가 공유하므로 테스트마다 비운다."""
    from engine import ai_match
    ai_match.GLOBAL_BUDGET.reset()
    yield
    ai_match.GLOBAL_BUDGET.reset()
