"""D 단계 보안 검토(AI-Generated Code Security Auditor) 지적사항의 회귀 테스트.

D-1 [중간] 업로드 파일의 열 이름·파일 이름이 마크다운으로 렌더링됨
D-2 [낮음] 운영체제가 저장할 수 없는 업로드 파일 이름에서 처리되지 않은 예외
D-3 [낮음] 중복 선택 조합마다 결과 파일이 쌓임
D-4 [낮음] 오래된 세션 정리가 폴더 mtime만 봐서 사용 중인 세션을 지울 수 있음
D-5 [낮음] 세션당 AI 한도 우회(새 세션), 한 번 호출의 크기 상한 없음
D-6 [정보] Windows에서 대소문자만 다른 두 파일이 같은 파일로 저장됨
D-7 [낮음] 하위 프로세스에 API 키 환경변수가 전달됨
D-8 [정보] 동시에 도는 하위 프로세스 수 상한 없음
D-10 [정보] 업로드 개수·크기를 메모리로 읽은 뒤에 검사
"""

import os
import re
import threading
import time
from pathlib import Path

import pytest
from openpyxl import Workbook
from streamlit.testing.v1 import AppTest

from answer_key import ROOT
from engine import ai_match, jobs, load_scenario, prepare

import app

SC_YAML = ROOT / "scenarios" / "incoming_inspection.yaml"
EVIL_HEADER = "![p](https://attacker.invalid/b.png) [세션 만료 - 다시 로그인](https://attacker.invalid/x)"
EVIL_NOTICE = ":red-background[**관리자 공지: 비밀번호를 입력하세요**]"


def xlsx(path, rows):
    wb = Workbook()
    for r in rows:
        wb.active.append(list(r))
    wb.save(path)
    return path


HARNESS = '''
import sys
sys.path.insert(0, {root!r})
import streamlit as st
import app
from engine import load_scenario
app.init_state()
sc = load_scenario({yaml!r})
if not st.session_state.get("tried"):
    st.session_state.tried = True
    app.load_files(sc, [(n, open(p, "rb").read()) for n, p in {items!r}], "upload")
app.show_messages()
if st.session_state.plans:
    app.render_matching(sc)
'''


def harness(items) -> AppTest:
    """업로드와 같은 경로(app.load_files -> app.render_matching)로 파일을 올린 화면."""
    script = HARNESS.format(root=str(ROOT), yaml=str(SC_YAML), items=[(n, str(p)) for n, p in items])
    at = AppTest.from_string(script, default_timeout=120)
    at.run()
    assert not at.exception, at.exception
    return at


def rendered_texts(at) -> list[str]:
    out = []
    for kind in ("warning", "caption", "error", "markdown", "success", "info"):
        out += [e.value for e in getattr(at, kind)]
    out += [e.label for e in at.expander]
    out += [e.label for e in at.selectbox] + [e.label for e in at.checkbox]
    return out


# ------------------------------------------------------------------ D-1
def test_md_escapes_markdown():
    s = app.md(EVIL_HEADER + " " + EVIL_NOTICE + "\n# 제목")
    unescaped = re.sub(r"\\.", "", s)
    for bad in ("](", "![", ":red-background[", "**", "#", "\n"):
        assert bad not in unescaped
    assert app.md("수입검사_대성정공.xlsx") == "수입검사\\_대성정공\\.xlsx"


def test_column_and_file_names_are_not_rendered_as_markdown(tmp_path):
    f = xlsx(tmp_path / "a.xlsx", [("품목코드", "로트번호", EVIL_HEADER, EVIL_NOTICE), ("A-1", "L1", "x", "y")])
    at = harness([("__굵은 파일명__ [링크](evil).xlsx", f)])
    texts = rendered_texts(at)
    joined = " ".join(texts)
    assert "attacker" in joined and "굵은 파일명" in joined   # 실제로 화면에 나왔는지
    for t in texts:
        unescaped = re.sub(r"\\.", "", t)
        assert "](" not in unescaped and "![" not in unescaped, t
        assert ":red-background[" not in unescaped and "__굵은" not in unescaped, t


# ------------------------------------------------------------------ D-2
@pytest.mark.parametrize("name, expected", [
    ("**굵은** _x_:?.xlsx", "__굵은__ _x___.xlsx"),
    ('a<b>c|d"e.xlsx', "a_b_c_d_e.xlsx"),
    ("CON.xlsx", "_CON.xlsx"),
    ("끝에 점. .xlsx", "끝에 점.xlsx"),
])
def test_disk_name(name, expected):
    assert app.disk_name(name) == expected


def test_long_name_is_shortened_below_linux_limit():
    n = app.disk_name("가" * 86 + ".xlsx")
    assert len(n.encode("utf-8")) <= app.MAX_NAME_BYTES and n.endswith(".xlsx")


def test_unsavable_names_do_not_crash(tmp_path):
    f = xlsx(tmp_path / "a.xlsx", [("품목코드", "로트번호"), ("A-1", "L1")])
    g = xlsx(tmp_path / "b.xlsx", [("품목코드", "로트번호"), ("A-2", "L2")])
    at = harness([("**굵은 파일명** _x_:?.xlsx", f), ("나" * 86 + ".xlsx", g)])
    assert not at.error
    names = [p.file_name for p in at.session_state["plans"]]
    assert names[0] == "__굵은 파일명__ _x___.xlsx"
    assert all(len(n.encode("utf-8")) <= app.MAX_NAME_BYTES for n in names)


def test_write_failure_gives_korean_message(tmp_path, monkeypatch):
    f = xlsx(tmp_path / "a.xlsx", [("품목코드", "로트번호"), ("A-1", "L1")])

    def boom(self, data):
        raise OSError(22, "Invalid argument")
    monkeypatch.setattr(Path, "write_bytes", boom)
    at = harness([("a.xlsx", f)])
    assert not at.exception
    assert any("저장하지 못했습니다" in e.value for e in at.error)


# ------------------------------------------------------------------ D-6
def test_names_differing_only_in_case_are_rejected(tmp_path):
    f = xlsx(tmp_path / "a.xlsx", [("품목코드", "로트번호"), ("A-1", "L1")])
    g = xlsx(tmp_path / "b.xlsx", [("품목코드", "로트번호"), ("A-2", "L2")])
    at = harness([("A.xlsx", f), ("a.xlsx", g)])
    assert any("같은 이름의 파일" in e.value for e in at.error)
    assert not at.session_state["plans"]


# ------------------------------------------------------------------ D-10
def test_upload_batch_checked_before_reading():
    assert app.upload_batch_error([1] * app.MAX_FILES) is None
    assert "파일 수 제한" in app.upload_batch_error([1] * (app.MAX_FILES + 1))
    assert "전체 크기 제한" in app.upload_batch_error([app.MAX_TOTAL_UPLOAD_BYTES, 1])


# ------------------------------------------------------------------ D-3
def test_result_files_do_not_pile_up():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=120)
    at.run()
    at.selectbox(key="scenario").set_value("stock_count").run()
    at.button(key="sample_btn").click().run()
    at.button(key="run_btn").click().run()
    rid = at.session_state["result_id"]
    keys = [c.key for c in at.checkbox if c.key and c.key.startswith(f"keep_{rid}_")]
    for k in keys:                                    # 체크를 하나씩 풀기
        at.checkbox(key=k).uncheck().run()
    at.checkbox(key="apply_suggestions").check().run()
    for k in keys:                                    # 다시 체크
        at.checkbox(key=k).check().run()
    assert not at.exception and not at.error
    out = Path(at.session_state["workspace"]) / "out"
    files = list(out.glob("*.xlsx"))
    assert len(files) <= app.MAX_KEPT_OUTPUTS, files   # 수정 전: 9개
    assert Path(at.session_state["current_output"]).is_file()


# ------------------------------------------------------------------ D-4
def test_cleanup_keeps_session_with_recent_inner_file(tmp_path, monkeypatch):
    monkeypatch.setattr(app, "SESSION_ROOT", tmp_path)
    now = time.time()
    old = now - app.STALE_SECONDS - 3600
    busy, idle = tmp_path / "s_busy", tmp_path / "s_idle"
    for d in (busy, idle):
        (d / "out").mkdir(parents=True)
        (d / "out" / "r.xlsx").write_bytes(b"x")
    for p in (idle / "out" / "r.xlsx", idle / "out", idle, busy / "out", busy):
        os.utime(p, (old, old))
    # busy: 세션 폴더와 out/ 은 오래됐지만 안에 방금 쓴 결과 파일이 있다 (수정 전에는 지워졌음)
    cleaned = []
    app.cleanup_stale(now)
    assert busy.exists() and not idle.exists(), cleaned


# ------------------------------------------------------------------ D-5
def test_global_budget_per_minute_and_day(monkeypatch):
    b = ai_match.CallBudget()
    monkeypatch.setattr(ai_match, "GLOBAL_PER_MINUTE", 2)
    monkeypatch.setattr(ai_match, "GLOBAL_PER_DAY", 3)
    t = 1_000_000.0
    assert b.try_acquire(t) is None and b.try_acquire(t + 1) is None
    assert "1분 뒤" in b.try_acquire(t + 2)
    assert b.try_acquire(t + 70) is None
    assert "오늘" in b.try_acquire(t + 200)
    assert b.try_acquire(t + 86_500) is None     # 하루가 지나면 다시


def test_global_budget_is_thread_safe(monkeypatch):
    b = ai_match.CallBudget()
    monkeypatch.setattr(ai_match, "GLOBAL_PER_MINUTE", 10)
    got = []
    threads = [threading.Thread(target=lambda: got.append(b.try_acquire())) for _ in range(50)]
    for th in threads:
        th.start()
    for th in threads:
        th.join()
    assert got.count(None) == 10


def _demo_plans():
    import yaml
    manifest = yaml.safe_load((ROOT / "samples" / "ai_demo" / "demo.yaml").read_text(encoding="utf-8"))
    sc = load_scenario(ROOT / "scenarios" / f"{manifest['scenario']}.yaml")
    return sc, prepare(sc, [ROOT / "samples" / "ai_demo" / manifest["files"][0]])


def test_budget_exhausted_blocks_call_before_sending():
    sc, plans = _demo_plans()
    calls = []
    budget = ai_match.CallBudget()
    for _ in range(ai_match.GLOBAL_PER_MINUTE):
        budget.try_acquire()
    with pytest.raises(ai_match.AiNotSent, match="1분 뒤"):
        ai_match.recommend(sc, plans, api_key="k", budget=budget, caller=lambda *a: calls.append(a) or "{}")
    assert calls == []


def test_request_size_is_capped(monkeypatch):
    sc, plans = _demo_plans()
    many = plans * 15                                   # 매칭 안 된 열이 있는 파일 15개
    req = ai_match.build_request(sc, many)
    assert len(req["files"]) == ai_match.MAX_AI_FILES and ai_match.request_is_trimmed(many)
    monkeypatch.setattr(ai_match, "MAX_AI_COLUMNS", 2)
    assert all(len(f["columns"]) <= 2 for f in ai_match.build_request(sc, plans)["files"])
    monkeypatch.setattr(ai_match, "MAX_PROMPT_CHARS", 100)
    calls = []
    with pytest.raises(ai_match.AiNotSent, match="너무 깁니다"):
        ai_match.recommend(sc, plans, api_key="k", caller=lambda *a: calls.append(a) or "{}")
    assert calls == []


def test_global_limit_shared_across_sessions(monkeypatch):
    # 수정 전: 새 세션마다 5회씩 -> 세션 2개로 10회. 이제 앱 전체 한도가 막는다
    monkeypatch.setenv(ai_match.KEY_NAME, "test-key")
    monkeypatch.setattr(ai_match, "GLOBAL_PER_MINUTE", 3)
    sent = []
    monkeypatch.setattr(ai_match, "call_gemini", lambda *a: sent.append(1) or "{깨진")
    for _ in range(2):                                  # 세션 2개
        at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=120)
        at.run()
        at.button(key="ai_demo_btn_0").click().run()
        for _ in range(ai_match.MAX_CALLS_PER_SESSION):
            if not at.button(key="ai_btn").disabled:
                at.button(key="ai_btn").click().run()
        assert not at.exception
    assert len(sent) == 3
    msgs = " ".join(w.value for w in at.warning)
    assert "1분 뒤" in msgs
    assert "남은 AI 호출: 5/5회" in " ".join(c.value for c in at.caption)   # 보내지 못한 시도는 세션 횟수에서 빼지 않음


# ------------------------------------------------------------------ D-7
def test_child_process_does_not_get_secrets(monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "should-not-leak")
    monkeypatch.setenv("SOME_ACCESS_TOKEN", "x")
    names = jobs.run_job("env_names", timeout=60)
    assert "GEMINI_API_KEY" not in names and "SOME_ACCESS_TOKEN" not in names
    assert any(n.upper() == "PATH" for n in names)


# ------------------------------------------------------------------ D-8
def test_concurrent_jobs_are_limited(monkeypatch):
    slots = threading.BoundedSemaphore(1)
    monkeypatch.setattr(jobs, "_SLOTS", slots)
    slots.acquire()                                     # 다른 작업이 자리를 차지하고 있다
    try:
        t = time.perf_counter()
        with pytest.raises(jobs.TimeLimitError, match="작업이 많아"):
            jobs.run_job("sleep", 0, timeout=1)
        assert time.perf_counter() - t < 3
    finally:
        slots.release()
    assert jobs.run_job("sleep", 0, timeout=30) == 0    # 자리가 나면 바로 돈다
