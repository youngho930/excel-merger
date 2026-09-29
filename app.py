"""엑셀 자동 취합·검증기 — 화면 (Streamlit).

흐름: 시나리오 선택 -> 파일 올리기(또는 샘플 체험) -> 열 매칭 확인·수정 -> 실행 -> 결과 확인·다운로드

- 시나리오 이름·열 이름은 이 파일에도 없다. 모두 scenarios/*.yaml 에서 읽는다.
- 파일 읽기와 실행은 engine.jobs.run_job 으로 하위 프로세스에서 돌리고, 단계마다 시간 제한을 둔다.
- 올린 파일은 세션마다 따로 만든 임시 폴더에 복사해서 쓴다. 원본(샘플 포함)은 절대 수정하지 않는다.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import time
import tomllib
import uuid
from pathlib import Path, PurePosixPath
from typing import Any

import pandas as pd
import streamlit as st
import yaml

from engine import KINDS, ReadError, ScenarioError, list_scenarios, pre_run_warnings
from engine import ai_match
from engine import limits as limits_mod
from engine.jobs import JobError, TimeLimitError, run_job, set_max_concurrent
from engine.matching import AI, USER
from engine.normalize import display
from engine.reader import EXCEL_SUFFIXES, read_table
from engine.writer import ACTION_COL, ERROR_COLS

ROOT = Path(__file__).resolve().parent
SCENARIO_DIR = ROOT / "scenarios"
SAMPLE_DIR = ROOT / "samples"



def setting_values(names) -> dict[str, Any]:
    """st.secrets 에서 names 만 골라 dict 로. 설정 파일이 없어도 오류 없이 빈 dict."""
    found: dict[str, Any] = {}
    try:
        for name in names:
            if name in st.secrets:
                found[name] = st.secrets[name]
    except Exception:
        return {}
    return found


# ---- 제한: secrets -> 환경변수 -> 프로필 기본값 (engine/limits.py). EXCEL_MERGER_PROFILE="cloud" 면 보수적인 값
LIMITS, LIMIT_NOTES = limits_mod.load_limits(setting_values(limits_mod.SETTING_NAMES))
set_max_concurrent(LIMITS.max_concurrent_jobs)
TIME_LIMIT = LIMITS.time_limit_seconds          # 초. 파일 읽기, 실행 단계마다
MAX_FILES = LIMITS.max_files                    # 한 번에 올리는 파일 수
MAX_FILE_BYTES = LIMITS.max_file_bytes          # 파일당 크기
MAX_ROWS = LIMITS.max_rows_per_file             # 파일당 데이터 행 수
MAX_TOTAL_ROWS = LIMITS.max_total_rows          # 모든 파일의 데이터 행 합계
MAX_TOTAL_UPLOAD_BYTES = LIMITS.max_total_upload_bytes   # 한 번에 올리는 파일 크기 합계
JOB_MEMORY_MB = LIMITS.job_memory_mb            # 처리 작업 하나의 메모리 상한 (Linux만, 0이면 없음)
MB = 1024 * 1024
SESSION_ROOT = Path(tempfile.gettempdir()) / "excel-merger-sessions"
STALE_SECONDS = 6 * 3600              # 이보다 오래 쓰지 않은 세션 폴더는 지운다
MAX_DUP_GROUPS_SHOWN = 50             # 화면에서 남길 행을 고를 수 있는 중복 그룹 수
DEMO_MANIFEST = "demo.yaml"           # samples/<폴더>/demo.yaml 이 있으면 AI 매칭 체험 폴더
MAX_KEPT_OUTPUTS = 2                  # 세션마다 남겨 두는 결과 파일 수 (나머지는 지운다)

NO_SOURCE = "(선택 안 함)"
REPO_URL = "https://github.com/youngho930/excel-merger"
PREVIEW_ROWS = 3                      # 샘플 미리보기에 보여줄 데이터 행 수

# ---- 용어 도움말 (물음표 아이콘). 한곳에 모아 두고 필요한 자리에서 짧게 보여준다
GLOSSARY = {
    "기준열": "결과 엑셀에 들어갈 표준 열입니다. 파일마다 열 이름이 달라도 이 열로 맞춰 모읍니다.",
    "허용값": "그 칸에 들어갈 수 있는 값의 목록입니다. 목록에 없는 값은 오류로 표시합니다. 예: 판정 = 합격/불합격",
    "중복기준": "같은 행인지 판단하는 열입니다. 이 열들의 값이 모두 같으면 중복 행으로 봅니다. "
               "날짜는 표기가 달라도 같은 날이면 같다고 봅니다.",
    "수정 제안값": "고칠 값이 분명할 때 보여주는 추천 값입니다. 예: '합격 '(뒤에 공백) → '합격'. "
                 "자동으로 바꾸지 않고, '수정 제안값 일괄 적용'을 켰을 때만 반영합니다.",
    "중복 그룹": "서로 중복인 행끼리 묶은 번호입니다(중복-1, 중복-2…). 같은 번호끼리 비교해서 남길 행을 고르면 됩니다.",
    "처리됨": "자동 수정됨, 행 제외됨, 중복 해소로 처리한 오류입니다.",
    "남은 오류": "직접 확인해야 할 오류입니다. 전체 오류에서 처리됨을 뺀 수입니다.",
    "전체 오류": "파일에서 찾은 모든 오류입니다. 처리해도 이 수는 줄지 않습니다.",
    "처리": "그 오류를 어떻게 처리했는지 적는 칸입니다. 자동 수정됨 / 행 제외됨 / 중복 해소. 빈칸이면 아직 남은 오류입니다.",
}


# ================================================================== 디자인 (HTML·CSS)
# 보안 원칙: unsafe HTML(st.html)로 내보내는 것은 아래의 "코드에 고정된 문자열"과
# 앱이 계산한 정수만 끼워 넣는 SAFE_HTML_FUNCS 의 결과뿐이다.
# 파일 이름·열 이름·시나리오 이름·설명·오류 메시지 같은 글자는 절대 이 HTML에 넣지 않는다
# (tests/test_design_safety.py 가 AST로 검사한다). 그런 글자는 st.markdown + md() 로만 보여준다.
PROJECT_STATS = ROOT / "project_stats.toml"   # 화면 숫자 카드의 테스트 개수 (한 곳에서 관리)

APP_CSS = """<style>
:root {
  --xm-accent: #21A366;      /* 강조색: 엑셀 초록 */
  --xm-accent-ink: #04150C;  /* 강조색 위 글자 (명도 대비 확보) */
  --xm-accent-2: #F2C14E;    /* 보조색: AI 관련 표시에만 */
  --xm-surface: #151D19;
  --xm-line: #2A3831;
  --xm-muted: #A3B0A8;
}
[data-testid="stMainBlockContainer"] { padding-top: 3.75rem; padding-bottom: 2rem; }
@media (max-width: 640px) {
  [data-testid="stMainBlockContainer"] { padding-top: 3.5rem; }
}

/* 제품 이름: 작은 머리 표시 */
.st-key-brand h1 {
  font-size: 0.95rem !important; font-weight: 600 !important; line-height: 1.4 !important;
  color: var(--xm-muted); padding: 0 !important; margin: 0 !important; letter-spacing: 0.01em;
}
.st-key-brand [data-testid="stHeaderActionElements"] { display: none; }

/* 히어로 */
.xm-wrap { container-type: inline-size; }
.xm-wrap ul, .xm-wrap ol { padding-left: 0 !important; margin-left: 0 !important; list-style: none; }
.xm-wrap li { margin-left: 0 !important; padding-left: 0; }
.xm-hero { display: grid; grid-template-columns: minmax(0, 1fr) minmax(250px, 300px);
  gap: 1rem 2rem; align-items: center; }
.xm-chips { display: flex; flex-wrap: wrap; gap: 0.375rem; list-style: none; margin: 0 0 0.75rem; padding: 0; }
.xm-chips li { font-size: 0.75rem; line-height: 1; padding: 0.35rem 0.6rem; margin: 0;
  border: 1px solid var(--xm-line); border-radius: 999px; color: var(--xm-muted); background: var(--xm-surface); }
.xm-chips li.xm-ai { border-color: rgba(242, 193, 78, 0.45); color: var(--xm-accent-2); }
.xm-title { font-size: clamp(1.45rem, 1rem + 1.8vw, 2.35rem); line-height: 1.25; font-weight: 800;
  letter-spacing: -0.02em; margin: 0; padding: 0; color: inherit; word-break: keep-all; }
.xm-title em { font-style: normal; color: var(--xm-accent); }
.xm-sub { margin: 0.7rem 0 0; color: var(--xm-muted); font-size: 0.975rem; line-height: 1.6;
  max-width: 40rem; word-break: keep-all; }
.xm-stats { list-style: none; margin: 0; padding: 0; display: grid; gap: 0.5rem; }
.xm-stats li { display: flex; align-items: baseline; gap: 0.75rem; margin: 0; padding: 0.65rem 1rem;
  border: 1px solid var(--xm-line); border-radius: 0.875rem; background: var(--xm-surface); }
.xm-stats b { font-size: 1.45rem; font-weight: 800; line-height: 1.2; color: var(--xm-accent);
  min-width: 3.4rem; font-variant-numeric: tabular-nums; }
.xm-stats span { font-size: 0.85rem; color: var(--xm-muted); }
@container (max-width: 640px) {
  .xm-hero { grid-template-columns: minmax(0, 1fr); gap: 0.9rem; }
  .xm-sub { font-size: 0.9rem; margin-top: 0.5rem; }
  .xm-stats { gap: 0.4rem; }
  .xm-stats li { padding: 0.45rem 0.85rem; }
  .xm-stats b { font-size: 1.2rem; }
}

/* 작동 흐름 띠 */
.xm-band { padding: 1rem 1.25rem 1.1rem; border: 1px solid var(--xm-line); border-radius: 1rem; }
.xm-band-title { font-size: 0.8rem; font-weight: 700; color: var(--xm-muted); margin: 0 0 0.8rem; }
.xm-flow { list-style: none; margin: 0; padding: 0; display: grid;
  grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 1.5rem; }
.xm-flow li { position: relative; display: flex; flex-direction: column; align-items: center;
  text-align: center; gap: 0.45rem; margin: 0; min-width: 0; }
.xm-flow li:not(:last-child)::after { content: "\\2192" / ""; position: absolute; right: -1.15rem;
  top: 0.6rem; color: var(--xm-muted); font-size: 1rem; line-height: 1; }
.xm-flow-icon { display: grid; place-items: center; flex: none; width: 2.5rem; height: 2.5rem;
  border-radius: 0.75rem; background: rgba(33, 163, 102, 0.15); color: var(--xm-accent);
  font-family: "Material Symbols Rounded"; font-size: 1.35rem; line-height: 1; font-weight: 400;
  font-feature-settings: "liga"; -webkit-font-feature-settings: "liga"; overflow: hidden; white-space: nowrap; }
.xm-flow-text { display: flex; flex-direction: column; gap: 0.15rem; min-width: 0; }
.xm-flow-text b { font-size: 0.9rem; font-weight: 700; line-height: 1.35; word-break: keep-all; }
.xm-flow-text small { font-size: 0.75rem; color: var(--xm-muted); line-height: 1.4; word-break: keep-all; }
.xm-flow-text i { font-style: normal; color: var(--xm-accent-2); }
@container (max-width: 640px) {
  .xm-flow { grid-template-columns: minmax(0, 1fr); gap: 0.85rem; }
  .xm-flow li { flex-direction: row; text-align: left; gap: 0.75rem; }
  .xm-flow li:not(:last-child)::after { content: "\\2193" / ""; right: auto; top: auto;
    left: 1.25rem; bottom: -0.85rem; transform: translateX(-50%); font-size: 0.8rem; }
}

/* 단계 카드와 원형 번호 배지 */
[class*="st-key-card_"] { background: var(--xm-surface); border-color: var(--xm-line) !important;
  border-radius: 1rem !important; }
[class*="st-key-card_"] h3 { font-size: 1.3rem; }
[class*="st-key-card_"] h3 [data-testid="stMarkdownBadge"],
[class*="st-key-card_"] h3 .stMarkdownBadge {
  display: inline-grid; place-items: center; width: 1.6rem; height: 1.6rem; padding: 0 !important;
  margin-right: 0.35rem; border-radius: 50% !important; vertical-align: 0.15em;
  background: var(--xm-accent) !important; color: var(--xm-accent-ink) !important;
  font-size: 0.85rem !important; font-weight: 800; line-height: 1;
}
.st-key-metric_remaining [data-testid="stMetric"] { border-color: var(--xm-accent) !important;
  background: rgba(33, 163, 102, 0.08); }

/* 강조색 버튼: 초록 위 흰 글자는 대비가 낮아 어두운 글자를 쓴다 */
[data-testid="stBaseButton-primary"], [data-testid="stBaseButton-primary"] p { color: var(--xm-accent-ink) !important; }

/* 푸터 */
.st-key-footer { margin-top: 1rem; padding-top: 0.75rem; border-top: 1px solid var(--xm-line); }
.st-key-footer [data-testid="stCaptionContainer"] { text-align: center; }
</style>"""

HERO_TEMPLATE = """<div class="xm-wrap"><section class="xm-hero" aria-label="소개">
<div class="xm-hero-text">
<ul class="xm-chips" aria-label="사용 기술"><li>Python</li><li>Streamlit</li><li class="xm-ai">Gemini AI</li><li>openpyxl</li></ul>
<h2 class="xm-title">양식이 제각각인 엑셀,<br>한 번에 <em>취합</em>하고 <em>오류까지</em></h2>
<p class="xm-sub">부서·협력사·지점마다 다른 엑셀 파일을 하나로 모으고, 빈칸·형식·중복 같은 오류를 자동으로 찾아 표시합니다. 원본 파일은 그대로 둡니다.</p>
</div>
<ul class="xm-stats" aria-label="숫자로 보는 특징">
<li><b>{kinds}종</b><span>자동 오류 검출</span></li>
<li><b>0줄</b><span>새 양식 추가 시 코드 수정</span></li>
{tests_card}</ul>
</section></div>"""
HERO_TESTS_CARD = "<li><b>{tests}개</b><span>자동 테스트</span></li>\n"

# 아이콘은 Streamlit에 들어 있는 Material Symbols 글꼴을 쓴다 (st.html 은 SVG를 지우고, 외부 요청도 늘지 않는다)
_ICON = '<span class="xm-flow-icon" aria-hidden="true">{}</span>'
FLOW_TEMPLATE = (
    '<div class="xm-wrap"><section class="xm-band" aria-label="작동 흐름">'
    '<p class="xm-band-title">작동 흐름</p><ol class="xm-flow">'
    '<li>' + _ICON.format("upload_file")
    + '<div class="xm-flow-text"><b>파일 업로드</b><small>양식이 다른 엑셀 여러 개</small></div></li>'
    '<li>' + _ICON.format("table_rows")
    + '<div class="xm-flow-text"><b>머리글 자동 탐지</b><small>제목 줄이 있어도 찾아냄</small></div></li>'
    '<li>' + _ICON.format("compare_arrows")
    + '<div class="xm-flow-text"><b>열 매칭(규칙→<i>AI</i>)</b><small>이름·동의어 먼저, 남은 열만 AI</small></div></li>'
    '<li>' + _ICON.format("fact_check")
    + '<div class="xm-flow-text"><b>오류 검증 {kinds}종</b><small>빈칸·형식·범위·허용값·중복</small></div></li>'
    '<li>' + _ICON.format("table_view")
    + '<div class="xm-flow-text"><b>결과 엑셀</b><small>오류 칸 색칠 + 오류목록 시트</small></div></li>'
    '</ol></section></div>')


def _whole_number(value: Any) -> int:
    """HTML에 끼워 넣을 수 있는 값은 앱이 계산한 0 이상의 정수뿐이다 (글자·bool은 거부)."""
    if type(value) is not int or value < 0:
        raise TypeError("디자인 HTML에는 0 이상의 정수만 넣을 수 있습니다.")
    return value


def hero_html(kinds: int, tests: int | None) -> str:
    """맨 위 히어로. kinds=오류 종류 수, tests=자동 테스트 수(없으면 그 카드를 뺀다)."""
    card = "" if tests is None else HERO_TESTS_CARD.format(tests=_whole_number(tests))
    return HERO_TEMPLATE.format(kinds=_whole_number(kinds), tests_card=card)


def flow_html(kinds: int) -> str:
    """작동 흐름 띠 (5단계)."""
    return FLOW_TEMPLATE.format(kinds=_whole_number(kinds))


SAFE_HTML_FUNCS = ("hero_html", "flow_html")   # st.html 인자로 쓸 수 있는 함수 (정수만 받는다)


def load_test_count() -> int | None:
    """project_stats.toml 의 테스트 개수. 파일이 없거나 값이 이상하면 None (카드를 숨긴다)."""
    try:
        n = tomllib.loads(PROJECT_STATS.read_text(encoding="utf-8"))["tests"]["count"]
    except (OSError, KeyError, TypeError, tomllib.TOMLDecodeError):
        return None
    return n if type(n) is int and n > 0 else None


def render_footer() -> None:
    with st.container(key="footer"):
        st.caption(f"만든 사람 신영호 · [GitHub 저장소]({REPO_URL})")


def step(num: int, title: str, summary: str, help: str | None = None) -> None:
    """단계 제목(원형 번호 배지) + 한 줄 설명. 자세한 내용은 details()로 접어 둔다."""
    st.subheader(f":green-badge[{_whole_number(num)}] {title}", help=help)
    st.caption(summary)


def scenario_help(by_key) -> str:
    """시나리오 선택 칸의 물음표 도움말: 시나리오마다 한 줄 설명."""
    return "\n\n".join(f"**{s.name}**: {s.description}" if s.description else f"**{s.name}**"
                       for s in by_key.values())


def details(text: str) -> None:
    with st.expander("자세히 보기"):
        st.markdown(text)


def render_sidebar() -> None:
    with st.sidebar:
        st.subheader("이 도구에 대해")
        st.markdown("부서·협력사·지점마다 양식이 다른 엑셀을 모으는 일은 손이 많이 가고 실수가 잦습니다.\n\n"
                    "이 도구는 파일을 한 번에 모으고, 빈칸·형식·중복 같은 오류를 자동으로 찾아 표시합니다.")
        st.markdown("**설계 원칙**\n"
                    "1. 시나리오는 설정 파일로 — 새 양식은 설정 파일만 추가하면 됩니다.\n"
                    "2. 규칙으로 먼저, 남은 열만 AI — 이름·동의어로 못 찾은 열만 AI에게 물어봅니다.\n"
                    "3. 원본은 그대로, 외부로는 열 이름만 — 올린 파일은 바꾸지 않고, AI에는 기본으로 열 이름만 보냅니다.")
        st.markdown(f"[GitHub 저장소]({REPO_URL})")
        st.caption("만든 사람: 신영호")


@st.cache_data(show_spinner=False, max_entries=32)
def sample_preview(path: str, mtime: float, scenario_key: str, _scenario) -> tuple[int, list[str], list[list[str]]]:
    """(머리글 행, 원래 열 이름, 앞쪽 데이터 몇 행). 저장소의 샘플 파일만 읽는다 (mtime·시나리오가 바뀌면 다시 읽음)."""
    table = read_table(path, _scenario)
    rows = [[cell_text(v) for v in r.cells] for r in table.rows[:PREVIEW_ROWS]]
    return table.header_row, list(table.headers), rows


def render_sample_preview(scenario, samples: list[Path]) -> None:
    with st.expander("샘플 파일 미리보기"):
        st.caption("파일마다 열 이름·열 순서·날짜 표기가 다릅니다. 이 도구가 그 차이를 맞춰서 모읍니다.")
        for p in samples:
            try:
                hr, headers, rows = sample_preview(str(p), p.stat().st_mtime, scenario.key, scenario)
            except (ReadError, OSError):
                st.caption(f"{md(p.name)}: 미리보기를 만들지 못했습니다.")
                continue
            note = f" · 위쪽 제목 줄 때문에 머리글이 {hr}행에 있습니다" if hr > 1 else ""
            st.markdown(f"**{md(p.name)}**{note}")
            st.dataframe(pd.DataFrame(rows, columns=headers), hide_index=True, width="stretch")


# ================================================================== 상태·작업 폴더
def reset_from(stage: str) -> None:
    """stage 이후 단계의 상태를 지운다. 'files' -> 파일·매칭·결과, 'result' -> 결과만."""
    ss = st.session_state
    if stage == "files":
        ss.plans = None
        ss.load_id = None
        ss.load_source = None
        ss.ai_message = None
    ss.result = None
    ss.result_id = None
    ss.outputs = {}
    ss.current_output = None
    ss.excluded = set()


def on_scenario_change() -> None:
    """시나리오가 바뀌면 매칭부터 다시 한다. 올려 둔 파일이 있으면 새 시나리오로 다시 읽는다."""
    reset_from("files")
    st.session_state.upload_sig = None


def on_apply_change() -> None:
    # 체크박스가 잠시 화면에서 사라져도(결과를 다시 만들 때) 선택을 기억한다
    st.session_state.apply_pref = st.session_state.apply_suggestions


def init_state() -> None:
    ss = st.session_state
    defaults = {"plans": None, "load_id": None, "load_source": None, "result": None, "outputs": {},
                "messages": [], "upload_sig": None, "uploader_n": 0, "map_error": None,
                "apply_pref": False, "excluded": set(), "result_id": None, "current_output": None,
                "pending_demo": None, "ai_calls": 0, "ai_message": None}
    for k, v in defaults.items():
        if k not in ss:
            ss[k] = v


def _last_used(d: Path) -> float:
    """세션 폴더를 마지막으로 쓴 시각: 폴더와 그 안 모든 파일·폴더의 mtime 중 가장 최근.

    하위 폴더(out/)에 파일을 써도 세션 폴더 자체의 mtime은 바뀌지 않으므로 안쪽까지 본다 (보안 검토 D-4).
    """
    newest = d.stat().st_mtime
    for p in d.rglob("*"):
        try:
            newest = max(newest, p.stat().st_mtime)
        except OSError:
            pass
    return newest


def cleanup_stale(now: float | None = None) -> None:
    now = time.time() if now is None else now
    try:
        entries = list(SESSION_ROOT.iterdir())
    except OSError:
        return
    for d in entries:
        try:
            if d.is_dir() and now - _last_used(d) > STALE_SECONDS:
                shutil.rmtree(d, ignore_errors=True)
        except OSError:
            pass


def workspace() -> Path:
    """이 세션만 쓰는 임시 폴더. 부를 때마다 사용 시각을 갱신한다."""
    ws = st.session_state.get("workspace")
    if ws is None or not Path(ws).is_dir():
        SESSION_ROOT.mkdir(parents=True, exist_ok=True)
        cleanup_stale()
        ws = tempfile.mkdtemp(prefix="s_", dir=SESSION_ROOT)
        st.session_state.workspace = ws
    try:
        os.utime(ws)
    except OSError:
        pass
    return Path(ws)


def fresh_dir(name: str) -> Path:
    """작업 폴더 안의 하위 폴더를 비우고 새로 만든다."""
    d = workspace() / name
    shutil.rmtree(d, ignore_errors=True)
    d.mkdir(parents=True)
    return d


def out_dir() -> Path:
    d = workspace() / "out"
    d.mkdir(parents=True, exist_ok=True)
    return d


def say(kind: str, text: str) -> None:
    """다음 화면 그리기에서 보여줄 안내 (error | warning | success)."""
    st.session_state.messages.append((kind, text))


def show_messages() -> None:
    # 안내에는 파일 이름·열 이름이 들어갈 수 있으므로 마크다운으로 해석되지 않게 한다
    for kind, text in st.session_state.messages:
        getattr(st, kind)(md(text))
    st.session_state.messages = []


# ================================================================== 사용자 문자열 표시
_MD_SPECIAL = set("\\`*_{}[]()#+-.!:<>|~$=")


def md(text: Any) -> str:
    """파일 이름·열 이름처럼 업로드된 파일에서 온 글자를 마크다운 요소(st.markdown, st.warning,
    st.caption, 위젯 라벨 등)에 넣기 전에 이스케이프한다.

    이스케이프하지 않으면 열 이름에 적힌 ![](외부 주소) 이미지가 열람자 브라우저에서 불러와지고,
    [링크](주소)나 :red-background[가짜 공지]가 앱 화면 안에 그려진다 (보안 검토 D-1).
    """
    s = " ".join(str(text).split("\n"))
    return "".join("\\" + ch if ch in _MD_SPECIAL else ch for ch in s)


# ================================================================== 파일 불러오기
_BAD_NAME_CHARS = set('<>:"/\\|?*') | {chr(i) for i in range(32)}
_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}
MAX_NAME_BYTES = 150       # 디스크에 저장할 파일 이름의 최대 UTF-8 바이트 (Linux 한도 255)


def safe_name(name: str) -> str:
    """업로드 파일 이름에서 폴더 부분을 떼어 낸다 (Windows 구분자 '\\'도 처리)."""
    return Path(PurePosixPath(name.replace("\\", "/")).name).name


def disk_name(name: str) -> str:
    """safe_name 결과를 Windows·Linux 모두에서 저장할 수 있는 이름으로 바꾼다 (보안 검토 D-2).

    금지 문자(< > : " / \\ | ? * 제어문자)는 '_'로, 끝의 점·공백은 제거, 예약 이름(CON 등)은 앞에 '_',
    이름이 너무 길면(UTF-8 150바이트) 줄인다. 확장자는 유지한다.
    """
    p = Path(name)
    suffix = p.suffix
    stem = "".join("_" if ch in _BAD_NAME_CHARS else ch for ch in p.stem).rstrip(" .") or "파일"
    if stem.split(".")[0].upper() in _RESERVED:
        stem = "_" + stem
    limit = MAX_NAME_BYTES - len(suffix.encode("utf-8"))
    while len(stem.encode("utf-8")) > limit:
        stem = stem[:-1]
    return stem + suffix


def upload_batch_error(sizes: list[int]) -> str | None:
    """업로드 목록을 메모리로 읽기 전에 개수·전체 크기를 검사한다 (보안 검토 D-10)."""
    if len(sizes) > MAX_FILES:
        return (f"파일 수 제한에 걸렸습니다: 한 번에 최대 {MAX_FILES}개까지 올릴 수 있습니다 "
                f"(올린 파일 {len(sizes)}개).")
    total = sum(sizes)
    if total > MAX_TOTAL_UPLOAD_BYTES:
        return (f"전체 크기 제한에 걸렸습니다: 올린 파일이 합계 {total / MB:.1f}MB입니다 "
                f"(한 번에 최대 {MAX_TOTAL_UPLOAD_BYTES // MB}MB).")
    return None


def load_files(scenario, items: list[tuple[str, Any]], source: str) -> None:
    """items: (파일 이름, bytes 또는 복사할 원본 경로). 검사 -> 세션 폴더에 복사 -> 읽기·매칭."""
    reset_from("files")
    if not items:
        return
    if len(items) > MAX_FILES:
        say("error", f"파일 수 제한에 걸렸습니다: 한 번에 최대 {MAX_FILES}개까지 올릴 수 있습니다 "
                     f"(올린 파일 {len(items)}개).")
        return
    names: list[str] = []
    for name, data in items:
        n = safe_name(name)
        if n and n not in (".", ".."):
            n = disk_name(n)
        if not n or n in (".", ".."):
            say("error", "파일 이름을 알 수 없는 파일이 있습니다. 이름을 바꿔 다시 올려 주세요.")
            return
        if Path(n).suffix.lower() not in EXCEL_SUFFIXES:
            say("error", f"'{n}'은 엑셀(.xlsx) 파일이 아닙니다. .xls 파일은 엑셀에서 .xlsx로 다시 저장해 주세요.")
            return
        size = len(data) if isinstance(data, (bytes, bytearray)) else Path(data).stat().st_size
        if size > MAX_FILE_BYTES:
            say("error", f"파일 크기 제한에 걸렸습니다: '{n}'은 {size / MB:.1f}MB입니다 "
                         f"(파일당 최대 {MAX_FILE_BYTES // MB}MB).")
            return
        # Windows는 대소문자만 다른 이름을 같은 파일로 저장하므로 대소문자를 무시하고 비교한다 (보안 검토 D-6)
        if n.casefold() in {x.casefold() for x in names}:
            say("error", f"같은 이름의 파일 '{n}'이 두 개 있습니다. 한 파일의 이름을 바꿔 다시 올려 주세요.")
            return
        names.append(n)

    in_dir = fresh_dir("in")
    shutil.rmtree(workspace() / "out", ignore_errors=True)
    paths = []
    for n, (_, data) in zip(names, items):
        dest = in_dir / n
        try:
            if isinstance(data, (bytes, bytearray)):
                dest.write_bytes(data)
            else:
                shutil.copyfile(data, dest)   # 샘플 원본은 건드리지 않고 복사본을 쓴다
        except OSError:
            say("error", f"'{n}' 파일을 저장하지 못했습니다. 파일 이름을 짧고 단순하게 바꿔 다시 올려 주세요.")
            return
        paths.append(dest)

    try:
        with st.spinner("파일을 읽고 열을 맞추는 중입니다…"):
            plans = run_job("prepare", scenario, paths, None, MAX_ROWS,
                            timeout=TIME_LIMIT, memory_mb=JOB_MEMORY_MB)
    except TimeLimitError as e:
        say("error", f"처리 시간 제한에 걸렸습니다. {e}")
        return
    except (ReadError, ScenarioError, JobError) as e:
        say("error", f"파일을 읽지 못했습니다. {e}")
        return

    total = sum(len(p.table.rows) for p in plans)
    if total > MAX_TOTAL_ROWS:
        say("error", f"행 수 제한에 걸렸습니다: 모든 파일의 데이터가 합계 {total:,}행입니다 "
                     f"(최대 {MAX_TOTAL_ROWS:,}행). 파일을 나눠서 취합해 주세요.")
        return
    ss = st.session_state
    ss.plans = plans
    ss.load_id = uuid.uuid4().hex[:8]
    ss.load_source = source
    ss.map_error = None


def sample_files(scenario) -> list[Path]:
    """samples/<시나리오 파일 이름>/ 안의 엑셀 파일 (없으면 빈 목록)."""
    folder = SAMPLE_DIR / scenario.key
    if not folder.is_dir() or (folder / DEMO_MANIFEST).exists():   # 체험 폴더는 시나리오 샘플이 아니다
        return []
    return sorted(p for p in folder.iterdir()
                  if p.is_file() and p.suffix.lower() in EXCEL_SUFFIXES and not p.name.startswith("~$"))


def ai_demos(scenario_keys) -> list[dict]:
    """samples/*/demo.yaml (AI 매칭 체험 설명서). 시나리오가 있고 파일이 모두 있는 것만."""
    out = []
    for manifest in sorted(SAMPLE_DIR.glob(f"*/{DEMO_MANIFEST}")):
        try:
            data = yaml.safe_load(manifest.read_text(encoding="utf-8"))
            files = [manifest.parent / safe_name(str(f)) for f in data["files"]]
            key = str(data["scenario"])
        except Exception:
            continue
        if key in scenario_keys and files and all(f.is_file() for f in files):
            out.append({"folder": manifest.parent.name, "scenario": key, "files": files,
                        "description": str(data.get("description") or "")})
    return out


def on_demo_click(demo: dict) -> None:
    ss = st.session_state
    ss.scenario = demo["scenario"]
    on_scenario_change()
    ss.uploader_n += 1
    ss.pending_demo = demo["folder"]


# ================================================================== 열 매칭
def map_key(fi: int, si: int) -> str:
    return f"map_{st.session_state.load_id}_{fi}_{si}"


def on_map_change(fi: int, standard: str, key: str) -> None:
    ss = st.session_state
    plan = ss.plans[fi]
    chosen = ss[key]
    try:
        plan.match.assign(standard, chosen)
        ss.map_error = None
        reset_from("result")
    except ValueError as e:
        ss.map_error = f"[{plan.file_name}] {e}"
        ss[key] = plan.match.get(standard).source_index   # 선택을 되돌린다


def ai_settings() -> tuple[str | None, str]:
    """(키, 모델). st.secrets -> 환경변수. 설정 파일이 없어도 오류 없이 넘어간다."""
    return ai_match.get_settings(setting_values((ai_match.KEY_NAME, ai_match.MODEL_KEY_NAME)))


def run_ai(scenario, api_key: str, model: str, with_examples: bool) -> None:
    ss = st.session_state
    ss.ai_calls += 1
    try:
        with st.spinner("AI에게 매칭 추천을 받는 중입니다…"):
            outcome = ai_match.recommend(scenario, ss.plans, api_key=api_key, model=model,
                                         with_examples=with_examples)
    except ai_match.AiError as e:
        if isinstance(e, (ai_match.AiNotSent, ai_match.AiBusy)):
            # 보내지 않았거나 Google 쪽 일시적 오류(503·429)로 끝났으면 세션 횟수에 넣지 않는다
            ss.ai_calls -= 1
        ss.ai_message = ("warning", f"AI 추천을 받지 못했습니다. {e} 지금은 {ai_match.NO_AI_NOTICE}입니다.")
        return
    # 채운 추천을 드롭다운에도 반영 (드롭다운은 이 아래에서 그려지므로 지금 바꿔도 된다)
    for s in outcome.accepted:
        m = ss.plans[s.file_index].match
        si = next(i for i, cm in enumerate(m.matches) if cm.standard == s.standard)
        ss[map_key(s.file_index, si)] = m.matches[si].source_index
    reset_from("result")
    dropped = f" 검증에서 버린 추천: {outcome.dropped_text()}." if outcome.dropped else ""
    if outcome.applied:
        ss.ai_message = ("success", f"AI 추천 {outcome.applied}건을 확인표에 'AI 추천'으로 채웠습니다. "
                                    f"아래 표에서 맞는지 확인하고, 틀리면 드롭다운에서 바꾸거나 해제해 주세요.{dropped}")
    else:
        ss.ai_message = ("info", f"AI가 확실한 짝을 찾지 못했습니다. 드롭다운에서 직접 지정해 주세요.{dropped}")


def render_ai(scenario) -> None:
    ss = st.session_state
    api_key, model = ai_settings()
    need = bool(ai_match.build_request(scenario, ss.plans)["files"])
    left = ai_match.MAX_CALLS_PER_SESSION - ss.ai_calls
    with st.container(border=True):
        st.markdown("**AI 매칭 추천 (선택)** — 동의어로도 찾지 못한 열만 AI에게 물어봅니다.")
        if not api_key:
            st.info(f"{ai_match.NO_AI_NOTICE}입니다. AI 추천을 쓰려면 관리자가 GEMINI_API_KEY를 설정해야 합니다.")
        if not need:
            matches = [m for p in ss.plans for m in p.match.matches]
            has_ai = any(m.method == AI for m in matches)
            all_filled = all(m.source_index is not None for m in matches)
            if has_ai and all_filled:
                st.caption("모든 열이 매칭되었습니다. 'AI 추천'으로 표시된 항목이 맞는지 아래 확인표에서 확인해 주세요.")
            elif has_ai:
                st.caption("AI 추천을 채웠고, 남은 원본 열이 없어 더 물어볼 열이 없습니다. "
                           "'AI 추천'으로 표시된 항목이 맞는지 아래 확인표에서 확인해 주세요.")
            else:
                st.caption("동의어로 매칭되지 않은 기준열·원본 열이 없어 AI 추천이 필요 없습니다.")
        else:
            with_examples = st.checkbox(
                "예시값 함께 보내기", key="ai_examples",
                help=f"켜면 매칭 안 된 원본 열마다 데이터 예시값 {ai_match.EXAMPLES_PER_COLUMN}개를 함께 보냅니다. "
                     "추천이 더 정확해질 수 있지만 파일의 실제 값이 외부(Google Gemini)로 전송됩니다.")
            request = ai_match.build_request(scenario, ss.plans, with_examples)
            what = ("매칭 안 된 기준열 이름과 원본 열 이름, 원본 열마다 예시값 "
                    f"{ai_match.EXAMPLES_PER_COLUMN}개" if with_examples
                    else "매칭 안 된 기준열 이름과 원본 열 이름만 (데이터 값은 보내지 않음)")
            st.caption(f"Google Gemini({md(model)})로 보내는 내용: {what}. 파일 이름은 F1, F2 같은 번호로 바꿔 보냅니다.")
            if ai_match.request_is_trimmed(ss.plans):
                st.caption(f"한 번에 파일 {ai_match.MAX_AI_FILES}개, 파일마다 원본 열 "
                           f"{ai_match.MAX_AI_COLUMNS}개까지만 보냅니다. 나머지는 드롭다운에서 지정해 주세요.")
            with st.expander("전송될 내용 미리 보기"):
                st.code(ai_match.preview_text(request), language="json")
            clicked = st.button("AI에게 매칭 추천 받기", key="ai_btn", disabled=not api_key or left <= 0)
            if clicked and api_key and left > 0:
                run_ai(scenario, api_key, model, with_examples)
                st.rerun()   # 남은 횟수·확인표를 새 상태로 다시 그린다
        st.caption(f"남은 AI 호출: {max(left, 0)}/{ai_match.MAX_CALLS_PER_SESSION}회 (세션당)")
        if left <= 0:
            st.warning(f"이 세션의 AI 호출 횟수({ai_match.MAX_CALLS_PER_SESSION}회)를 모두 썼습니다. "
                       "남은 열은 드롭다운에서 직접 지정해 주세요.")
        if ss.ai_message:
            kind, text = ss.ai_message
            getattr(st, kind)(md(text))


def render_matching(scenario) -> None:
    plans = st.session_state.plans
    step(3, "열 매칭 확인", "원본 열이 어느 기준열로 들어가는지 확인하고, 틀리면 드롭다운에서 고칩니다.",
         help=GLOSSARY["기준열"])
    details("**매칭 방법**\n"
            "- **정확히 일치**: 열 이름이 기준열과 같습니다. 띄어쓰기·대소문자·기호 차이는 무시합니다.\n"
            "- **동의어**: 시나리오 설정의 '다른 이름' 목록에 있는 이름입니다. 예: 품번 → 품목코드\n"
            "- **AI 추천**: 동의어로도 못 찾은 열을 AI가 추천했습니다. 맞는지 꼭 확인해 주세요.\n"
            "- **사용자 지정**: 드롭다운에서 직접 고른 열입니다.\n"
            "- **매칭 안 됨**: 짝을 찾지 못했습니다. 필수 열이면 실행 전에 경고가 나옵니다.\n\n"
            "AI에게는 기본으로 **열 이름만** 보냅니다. 파일의 실제 값은 '예시값 함께 보내기'를 켰을 때만 보냅니다.")
    if st.session_state.map_error:
        st.error(md(st.session_state.map_error))
    render_ai(scenario)
    # 경고·라벨에는 파일 이름·열 이름(업로드된 파일에서 온 글자)이 들어가므로 md()로 이스케이프한다
    warnings = pre_run_warnings(plans)
    if warnings:
        st.warning("실행 전에 확인해 주세요.\n\n" + "\n".join(f"- {md(w)}" for w in warnings))

    for fi, plan in enumerate(plans):
        t, m = plan.table, plan.match
        n_ai = sum(1 for cm in m.matches if cm.method == AI)
        flag = " · 확인 필요" if plan.warnings() or n_ai else ""
        if n_ai:
            flag += f" · AI 추천 {n_ai}개"
        label = f"{md(t.file_name)} — 머리글 {t.header_row}행, 데이터 {len(t.rows):,}행{flag}"
        with st.expander(label, expanded=bool(flag)):
            for w in plan.warnings():
                st.warning(md(w))
            head = st.columns([3, 4, 2])
            head[0].markdown("**원본 열**")
            head[1].markdown("**→ 기준열**")
            head[2].markdown("**매칭 방법**")
            options: list[int | None] = [None] + list(range(len(m.headers)))
            for si, cm in enumerate(m.matches):
                key = map_key(fi, si)
                if key not in st.session_state:
                    st.session_state[key] = cm.source_index
                row = st.columns([3, 4, 2], vertical_alignment="center")
                row[0].selectbox(
                    f"{md(cm.standard)}의 원본 열", options, key=key,
                    format_func=lambda i, m=m: NO_SOURCE if i is None else m.column_label(i),
                    on_change=on_map_change, args=(fi, cm.standard, key), label_visibility="collapsed")
                req = " (필수)" if cm.required else ""
                row[1].markdown(f"→ **{md(cm.standard)}**{req}")
                method = cm.method
                if cm.method == USER:
                    method = f":blue[{method}]"
                elif cm.method == AI:
                    method = f":violet[**{method}**] (확인 필요)"
                elif cm.source_index is None:
                    method = f":red[{method}]" if cm.required else f":gray[{method}]"
                row[2].markdown(method)
            unmatched = [m.column_label(i) for i in m.unmatched_sources]
            st.caption("매칭 안 된 원본 열: " + (md(", ".join(unmatched)) if unmatched else "없음"))


# ================================================================== 결과
def cell_text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, str) and v != v.strip():
        return f'"{v}"'   # 앞뒤 공백이 보이도록
    return display(v)


def issues_frame(result, res) -> pd.DataFrame:
    rows = []
    for i, action in zip(result.issues, res.actions):
        rows.append([i.file, "(전체)" if i.row is None else i.row, i.column or "(없음)", i.standard, i.kind,
                     "(빈칸)" if i.value is None else cell_text(i.value), i.message,
                     cell_text(i.suggestion), i.dup_group or "", action])
    df = pd.DataFrame(rows, columns=ERROR_COLS)
    df["행"] = df["행"].astype(str)
    return df


def current_resolution():
    ss = st.session_state
    return ss.result.resolve(ss.excluded, bool(ss.apply_pref))


def on_keep_change(key: str, file: str, row: int) -> None:
    ss = st.session_state
    if ss[key]:
        ss.excluded.discard((file, row))
    else:
        ss.excluded.add((file, row))


def render_duplicates(result, res) -> None:
    groups = result.dup_groups()
    if not groups:
        return
    with st.container(border=True, key="card_duplicates"):
        render_duplicate_groups(result, res, groups)


def render_duplicate_groups(result, res, groups) -> None:
    ss = st.session_state
    step(5, "중복 행 고르기", "중복 그룹마다 결과에 남길 행을 고릅니다. 기본은 전부 남김입니다.",
         help=GLOSSARY["중복 그룹"])
    details("- 체크를 끈 행은 취합결과에서 빠지고, **'제외된 행' 시트**에 원래 값 그대로 기록됩니다.\n"
            "- 그룹에 한 행만 남으면 그 행의 중복 오류는 **'중복 해소'**로 처리됩니다.\n"
            "- 한 그룹의 행을 모두 빼도 됩니다. 이때는 경고가 나옵니다.\n"
            "- 뺀 행의 다른 오류도 오류목록에 남고, '처리' 칸에 '행 제외됨'으로 표시됩니다.")
    names = result.scenario.column_names
    items = list(groups.items())
    if len(items) > MAX_DUP_GROUPS_SHOWN:
        st.info(f"중복 그룹이 {len(items):,}개라 앞의 {MAX_DUP_GROUPS_SHOWN}개만 보여줍니다. "
                "나머지 그룹은 모든 행을 남깁니다.")
    for gi, (group, rows) in enumerate(items[:MAX_DUP_GROUPS_SHOWN]):
        with st.container(border=True):
            st.markdown(f"**{md(group)}** · 중복기준 {md(' + '.join(result.scenario.dup_keys))} · {len(rows)}행")
            st.dataframe(pd.DataFrame(
                [{"출처 파일": r.source_file, "원래 행": str(r.excel_row), **{n: cell_text(r.raw[n]) for n in names}}
                 for r in rows]), hide_index=True, width="stretch")
            cols = st.columns(min(len(rows), 4))
            for j, r in enumerate(rows):
                key = f"keep_{ss.result_id}_{gi}_{j}"
                if key not in ss:
                    ss[key] = (r.source_file, r.excel_row) not in ss.excluded
                cols[j % len(cols)].checkbox(f"남김: {md(r.source_file)} {r.excel_row}행", key=key,
                                             on_change=on_keep_change, args=(key, r.source_file, r.excel_row))
            if group in res.fully_excluded_groups:
                st.warning(f"{md(group)}: 이 그룹의 행이 모두 빠집니다. 결과 파일에 이 그룹의 행이 하나도 남지 않습니다.")
            elif any((r.source_file, r.excel_row) in res.resolved_rows for r in rows) and \
                    any((r.source_file, r.excel_row) in res.excluded for r in rows):
                st.caption("한 행만 남아 중복이 풀렸습니다. 남은 행의 중복 오류는 '중복 해소(이 행을 남김)'로 처리됩니다.")


def run_merge(scenario) -> None:
    ss = st.session_state
    reset_from("result")
    applied = bool(ss.apply_pref)
    try:
        with st.spinner("오류를 검사하고 결과 파일을 만드는 중입니다…"):
            result, path = run_job("execute", scenario, ss.plans, out_dir(), applied,
                                   timeout=TIME_LIMIT, memory_mb=JOB_MEMORY_MB)
    except TimeLimitError as e:
        say("error", f"처리 시간 제한에 걸렸습니다. {e}")
        return
    except (ReadError, ScenarioError, JobError) as e:
        say("error", str(e))
        return
    for old in (workspace() / "out").glob("*.xlsx"):   # 이전 실행의 결과 파일 정리
        if old.resolve() != Path(path).resolve():
            old.unlink(missing_ok=True)
    ss.result = result
    ss.result_id = uuid.uuid4().hex[:8]
    ss.outputs = {output_key(applied, ()): str(path)}


def output_key(applied: bool, excluded) -> str:
    return f"{int(applied)}|" + "|".join(f"{f}:{r}" for f, r in sorted(excluded))


def output_for(res) -> Path | None:
    """지금 선택(제안값 적용, 뺀 행)에 맞는 결과 파일. 처음 고른 조합이면 새로 만든다."""
    ss = st.session_state
    key = output_key(res.apply_suggestions, res.excluded)
    if key not in ss.outputs:
        try:
            with st.spinner("결과 파일을 만드는 중입니다…"):
                ss.outputs[key] = str(run_job("write", ss.result, out_dir(), res.apply_suggestions,
                                              sorted(res.excluded), timeout=TIME_LIMIT,
                                              memory_mb=JOB_MEMORY_MB))
        except TimeLimitError as e:
            st.error(md(f"처리 시간 제한에 걸렸습니다. {e}"))
            return None
        except (ReadError, ScenarioError, JobError) as e:
            st.error(md(str(e)))
            return None
    else:
        ss.outputs[key] = ss.outputs.pop(key)   # 가장 최근에 쓴 것을 맨 뒤로
    prune_outputs()
    ss.current_output = ss.outputs[key]
    return Path(ss.outputs[key])


def prune_outputs() -> None:
    """최근 결과 파일 MAX_KEPT_OUTPUTS개만 남기고 지운다.

    제안값 적용·남길 행 조합마다 파일이 생기므로, 지우지 않으면 체크를 바꿀 때마다 디스크가 늘어난다 (보안 검토 D-3).
    """
    ss = st.session_state
    while len(ss.outputs) > MAX_KEPT_OUTPUTS:
        old_key = next(iter(ss.outputs))
        try:
            Path(ss.outputs.pop(old_key)).unlink(missing_ok=True)
        except OSError:
            pass


def render_result() -> None:
    result = st.session_state.result
    res = current_resolution()
    fixes = result.suggestion_map()
    total, done, remaining = res.totals()
    with st.container(border=True, key="card_result"):
        step(4, "실행 결과", "찾은 오류를 종류별로 셉니다. 직접 확인할 건수는 '남은 오류'입니다.")
        top = st.columns(4)
        with top[0].container(key="metric_remaining"):
            st.metric("남은 오류", f"{remaining:,}건", help=GLOSSARY["남은 오류"], border=True)
        top[1].metric("전체 오류", f"{total:,}건", help=GLOSSARY["전체 오류"], border=True)
        top[2].metric("처리됨", f"{done:,}건", help=GLOSSARY["처리됨"], border=True)
        top[3].metric("제외한 행", f"{len(res.excluded):,}개", help="중복 행 고르기에서 결과에서 뺀 행의 수입니다.",
                      border=True)
        info = st.columns(4)
        info[0].metric("파일 수", f"{len(result.plans)}개")
        info[1].metric("취합 행 수", f"{len(result.rows) - len(res.excluded):,}행")
        info[2].metric("수정 제안 있음", f"{len(fixes):,}건", help=GLOSSARY["수정 제안값"])
        details("**오류 종류**\n"
                "- **필수값 빈칸**: 꼭 있어야 하는 칸이 비어 있습니다.\n"
                "- **형식 오류**: 숫자 칸에 글자가 있거나(예: '12개'), 없는 날짜입니다(예: 13월).\n"
                "- **범위 밖 값**: 정해진 범위를 벗어났습니다(예: 음수 수량).\n"
                "- **허용값 아닌 값**: 정해진 값 목록에 없습니다(예: 'OK', '합격 ').\n"
                "- **중복 행**: 중복기준 열의 값이 같은 행이 둘 이상 있습니다. 다른 파일끼리도 찾습니다.")
        st.dataframe(pd.DataFrame([{"오류 종류": k, "전체": t, "처리됨": d, "남은 오류": r}
                                   for k, (t, d, r) in res.counts().items()]),
                     hide_index=True, key="kind_table")

    render_duplicates(result, res)

    with st.container(border=True, key="card_download"):
        step(6, "결과 엑셀 받기", "오류를 색칠한 결과 엑셀을 내려받습니다. 올린 파일은 바뀌지 않습니다.")
        details("**결과 엑셀의 시트**\n"
                "- **취합결과**: 모은 데이터. 출처 파일과 원래 행 번호가 붙고, 오류 칸은 종류별 색으로 칠합니다.\n"
                "- **오류목록**: 오류 하나당 한 줄. 파일, 행, 열, 오류 종류, 값, 설명, 수정 제안값, 중복 그룹, 처리.\n"
                "- **제외된 행**: 중복 행 고르기에서 뺀 행 (뺀 행이 있을 때만).\n"
                "- **요약**: 오류 종류별 전체 / 처리됨 / 남은 오류.\n"
                "- **범례**: 색의 뜻.")
        if "apply_suggestions" not in st.session_state:
            st.session_state.apply_suggestions = st.session_state.apply_pref
        applied = st.checkbox(
            "수정 제안값 일괄 적용", key="apply_suggestions", on_change=on_apply_change,
            help="켜면 결과 엑셀의 오류 셀 중 수정 제안값이 있는 셀을 제안값으로 바꾸고 초록색으로 표시합니다. "
                 "오류목록에는 '처리' 열에 '자동 수정됨'으로 남습니다. 원본 파일은 바뀌지 않습니다.")
        if applied:
            st.caption(f"제안값이 있는 {len(fixes):,}개 셀을 바꿔서 저장합니다. 나머지 오류 셀은 원래 값 그대로입니다.")
        else:
            st.caption("오류 셀은 원래 값 그대로 두고 색만 칠해서 저장합니다.")
        if res.excluded:
            st.caption(f"중복 그룹에서 뺀 {len(res.excluded):,}개 행은 '제외된 행' 시트로 옮겨 저장합니다.")
        path = output_for(res)
        if path is not None and path.is_file():
            st.download_button("결과 엑셀 내려받기", data=path.read_bytes(), file_name=path.name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="download", type="primary", on_click="ignore")

    with st.container(border=True, key="card_issues"):
        render_issue_list(result, res, remaining)


def render_issue_list(result, res, remaining: int) -> None:
    st.subheader("오류 목록")
    counts = result.counts()
    present = [k for k in KINDS if counts[k]]
    f1, f2 = st.columns([1, 3], vertical_alignment="bottom")
    only_remaining = f1.checkbox("남은 오류만 보기", value=True, key="only_remaining",
                                 help="끄면 처리한 오류(자동 수정됨·행 제외됨·중복 해소)도 함께 보여줍니다.")
    chosen = f2.multiselect("오류 종류로 거르기", options=list(KINDS), default=present, key="kind_filter",
                            placeholder="오류 종류를 고르세요")
    df = issues_frame(result, res)
    shown = df[df["오류 종류"].isin(chosen)]
    if only_remaining:
        shown = shown[shown[ACTION_COL] == ""]
    st.caption(f"{len(shown):,}건 표시 (전체 {len(df):,}건 중 남은 오류 {remaining:,}건)")
    st.dataframe(shown, hide_index=True, width="stretch", column_config={
        name: st.column_config.Column(help=GLOSSARY[name])
        for name in ("수정 제안값", "중복 그룹", ACTION_COL) if name in shown.columns and name in GLOSSARY})


# ================================================================== 화면
def main() -> None:
    st.set_page_config(page_title="엑셀 자동 취합·검증기", page_icon=":material/table_chart:", layout="wide")
    init_state()

    render_sidebar()
    st.html(APP_CSS)
    with st.container(key="brand"):
        st.title("엑셀 자동 취합·검증기")
    # 첫 화면(1280×720, 휴대폰 375×812)에서 "샘플 파일로 바로 체험" 버튼이 스크롤 없이 보이도록
    # 히어로는 짧게 두고, 작동 흐름 띠는 버튼 아래에 둔다
    st.html(hero_html(len(KINDS), load_test_count()))

    # ---- 1. 시나리오
    try:
        scenarios = list_scenarios(SCENARIO_DIR)
    except ScenarioError as e:
        st.error(f"시나리오 설정 파일에 문제가 있습니다. {e}")
        st.stop()
    if not scenarios:
        st.error("scenarios 폴더에 시나리오 설정 파일(.yaml)이 없습니다.")
        st.stop()
    by_key = {s.key: s for s in scenarios}
    with st.container(border=True, key="card_scenario"):
        scenario = render_scenario_card(by_key)
    st.html(flow_html(len(KINDS)))
    render_steps_after_scenario(scenario)
    render_footer()


def render_scenario_card(by_key: dict):
    """1단계: 시나리오 선택 + 샘플·AI 체험 버튼 + 접어 둔 설명. 고른 시나리오를 돌려준다."""
    ss = st.session_state
    step(1, "시나리오 선택", "어떤 종류의 엑셀을 모을지 고르고, 샘플로 바로 체험해 보세요.")
    key = st.selectbox("어떤 엑셀을 취합하나요?", list(by_key), key="scenario",
                       format_func=lambda k: by_key[k].name, on_change=on_scenario_change,
                       help=scenario_help(by_key))
    scenario = by_key[key]

    samples = sample_files(scenario)
    demos = ai_demos(set(by_key))
    # 체험 버튼은 왼쪽부터 나란히 붙인다 (같은 너비의 칸으로 나누면 두 번째 버튼이 화면 가운데로 떨어진다)
    sample_clicked = False
    with st.container(horizontal=True, gap="small"):
        if samples:
            sample_clicked = st.button("샘플 파일로 바로 체험", key="sample_btn", type="primary",
                                       help=f"이 시나리오의 샘플 엑셀 {len(samples)}개를 올린 것처럼 불러옵니다.")
        for i, demo in enumerate(demos):
            st.button("AI 매칭 체험", key=f"ai_demo_btn_{i}", on_click=on_demo_click, args=(demo,),
                      help=f"{demo['description']} ({by_key[demo['scenario']].name} 시나리오로 바뀝니다). "
                           "열 이름이 동의어 사전에 없어서 AI 추천이 필요한 파일입니다.")
    if sample_clicked:
        ss.uploader_n += 1          # 올려 둔 파일 목록은 비운다
        ss.upload_sig = None
        load_files(scenario, [(p.name, p) for p in samples], "sample")
    if ss.pending_demo:
        demo = next((d for d in demos if d["folder"] == ss.pending_demo), None)
        ss.pending_demo = None
        ss.upload_sig = None
        if demo is not None and demo["scenario"] == key:
            load_files(scenario, [(p.name, p) for p in demo["files"]], "demo")

    # 버튼 아래에 접어 둔다 (첫 화면에서 버튼이 위에 보이도록)
    if samples:
        render_sample_preview(scenario, samples)
    with st.expander("이 시나리오의 기준열 보기"):
        if scenario.description:
            st.caption(scenario.description)
        st.dataframe(pd.DataFrame([{
            "기준열": c.name, "필수": "예" if c.required else "", "형식": c.fmt,
            "범위": "" if c.range is None else f"{display(c.range[0])} ~ {display(c.range[1])}",
            "허용값": "" if c.allowed is None else ", ".join(display(a) for a in c.allowed),
            "다른 이름(동의어)": ", ".join(c.aliases)} for c in scenario.columns]),
            hide_index=True, width="stretch",
            column_config={"기준열": st.column_config.TextColumn(help=GLOSSARY["기준열"]),
                           "허용값": st.column_config.TextColumn(help=GLOSSARY["허용값"]),
                           "다른 이름(동의어)": st.column_config.TextColumn(
                               help="파일에 이 이름으로 적혀 있어도 같은 기준열로 알아봅니다.")})
        if scenario.dup_keys:
            st.markdown("중복기준: " + md(" + ".join(scenario.dup_keys)), help=GLOSSARY["중복기준"])
    details("- **시나리오**는 '어떤 엑셀을 모을지'에 대한 설정입니다. 결과 엑셀의 열(기준열), 필수 여부, "
            "형식, 허용값, 중복기준이 들어 있습니다.\n"
            "- **샘플 파일로 바로 체험**: 이 시나리오의 샘플 엑셀을 올린 것처럼 불러옵니다. 샘플에는 일부러 오류를 넣어 두었습니다.\n"
            "- **AI 매칭 체험**: 열 이름이 동의어 사전에 없는 파일입니다. AI 추천으로 열을 맞추는 과정을 볼 수 있습니다.")
    return scenario


def render_steps_after_scenario(scenario) -> None:
    """2단계(업로드)부터 결과까지. 파일이 없으면 안내만 보여주고 끝낸다."""
    ss = st.session_state
    # ---- 2. 업로드
    with st.container(border=True, key="card_upload"):
        step(2, "엑셀 파일 올리기", "모을 엑셀 파일(.xlsx)을 한 번에 여러 개 올립니다.")
        for note in LIMIT_NOTES:   # 관리자 설정(secrets·환경변수)이 잘못된 경우
            st.caption(f"설정 안내: {md(note)}")
        uploads = st.file_uploader("취합할 엑셀 파일(.xlsx)을 모두 골라 올려 주세요.", type=["xlsx", "xlsm"],
                                   accept_multiple_files=True, key=f"uploader_{ss.uploader_n}")
        sig = tuple((u.file_id, u.name, u.size) for u in uploads) if uploads else None
        if sig != ss.upload_sig:
            ss.upload_sig = sig
            if sig:
                err = upload_batch_error([u.size for u in uploads])   # 파일 내용을 꺼내기 전에 검사
                if err:
                    reset_from("files")
                    say("error", err)
                else:
                    load_files(scenario, [(u.name, u.getvalue()) for u in uploads], "upload")
            elif ss.load_source == "upload":
                reset_from("files")
        with st.expander("자세히 보기"):
            st.caption(f"제한: 한 번에 최대 {MAX_FILES}개 · 파일당 {MAX_FILE_BYTES // MB}MB · "
                       f"합계 {MAX_TOTAL_UPLOAD_BYTES // MB}MB · 파일당 {MAX_ROWS:,}행 · "
                       f"모든 파일 합계 {MAX_TOTAL_ROWS:,}행 · 단계마다 처리 시간 {TIME_LIMIT}초")
            st.markdown("- 열 이름·열 순서·날짜 표기가 파일마다 달라도 괜찮습니다.\n"
                        "- 위쪽에 제목 줄이 있어도 머리글 행을 자동으로 찾습니다.\n"
                        "- 올린 파일은 바꾸지 않습니다. 결과는 새 엑셀 파일로 만듭니다.\n"
                        "- .xls 파일은 엑셀에서 .xlsx로 다시 저장해서 올려 주세요.")
    show_messages()

    if not ss.plans:
        st.info("샘플 파일로 체험하거나, 엑셀 파일을 올리면 다음 단계가 나타납니다.")
        return
    src = {"sample": "샘플 파일", "demo": "AI 매칭 체험 파일"}.get(ss.load_source, "올린 파일")
    st.success(f"{src} {len(ss.plans)}개를 읽었습니다.")

    # ---- 3. 매칭 + 실행 버튼 (확인이 끝나면 같은 카드 안에서 바로 실행)
    with st.container(border=True, key="card_matching"):
        render_matching(scenario)
        run_clicked = st.button("취합·검증 실행", key="run_btn", type="primary")
    if run_clicked:
        run_merge(scenario)
        show_messages()
    if ss.result is not None:
        render_result()


if __name__ == "__main__":   # streamlit run / AppTest 는 이 파일을 __main__ 으로 실행한다
    main()
