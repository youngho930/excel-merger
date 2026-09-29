"""화면 숫자 카드의 '자동 테스트 N개'가 실제 테스트 개수와 같은지 확인한다.

값은 project_stats.toml 한 곳에서 관리한다. 테스트를 추가·삭제했는데 이 값을 안 바꾸면 여기서 실패한다.
(이 파일의 테스트도 개수에 들어간다.)
"""

import os
import re
import subprocess
import sys

from answer_key import ROOT

import app


def collected_test_count() -> int:
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    out = subprocess.run(
        [sys.executable, "-m", "pytest", "--collect-only", "-q", "-p", "no:cacheprovider", str(ROOT / "tests")],
        cwd=ROOT, capture_output=True, text=True, encoding="utf-8", timeout=300, env=env)
    m = re.search(r"^(\d+) tests? collected", out.stdout, re.M)
    assert m, out.stdout[-2000:] + out.stderr[-2000:]
    return int(m.group(1))


def test_project_stats_test_count_matches_collected_tests():
    shown = app.load_test_count()
    assert shown is not None, "project_stats.toml 의 [tests] count 를 읽지 못했습니다."
    actual = collected_test_count()
    assert shown == actual, (f"project_stats.toml 의 테스트 개수({shown})가 실제 수집된 개수({actual})와 다릅니다. "
                             "project_stats.toml 을 고쳐 주세요.")


def test_measurement_matches_the_answer_key():
    # 측정에 쓴 수입검사 샘플의 오류 수는 정답지와 같아야 하고, 찾은 수는 그보다 클 수 없다
    import json
    m = app.load_measurement()
    assert m is not None, "project_stats.toml 의 [measurement] 를 읽지 못했습니다."
    answer = json.loads((ROOT / "samples" / "expected_errors.json").read_text(encoding="utf-8"))
    assert m["errors_total"] == len(answer["scenarios"]["incoming_inspection"]["errors"])
    assert m["manual_found"] <= m["errors_total"] and m["tool_found"] <= m["errors_total"]


def test_measurement_doc_shows_the_same_values():
    # docs/measurement.md 의 결과 표는 project_stats.toml 과 같은 값을 초 단위로 적는다
    m = app.load_measurement()
    doc = (ROOT / "docs" / "measurement.md").read_text(encoding="utf-8")

    def mmss(seconds):
        return f"{seconds // 60}분 {seconds % 60}초"
    for text in (mmss(m["manual_seconds"]), mmss(m["tool_seconds"]),
                 f"{m['manual_found']} / {m['errors_total']}", f"{m['tool_found']} / {m['errors_total']}"):
        assert text in doc, text


def test_readme_does_not_contradict_the_test_count():
    # README 에 테스트 개수를 적는다면 project_stats.toml 과 같은 값이어야 한다
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    numbers = {int(n.replace(",", "")) for n in re.findall(r"([\d,]+)\s*개(?:의)?\s*(?:자동\s*)?테스트", readme)}
    assert numbers <= {app.load_test_count()}, numbers
