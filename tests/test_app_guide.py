"""화면 사용 안내: 짧은 첫 안내(히어로), 작동 흐름 띠, 단계별 한 줄 설명 + 자세히 보기, 용어 도움말,
샘플 미리보기, 사이드바, 푸터."""

import re

from streamlit.testing.v1 import AppTest

from answer_key import ROOT

APP = str(ROOT / "app.py")
FLOW_STEPS = ["파일 업로드", "머리글 자동 탐지", "열 매칭(규칙→AI)", "오류 검증 5종", "결과 엑셀"]


def first_screen() -> AppTest:
    at = AppTest.from_file(APP, default_timeout=120)
    at.run()
    assert not at.exception, at.exception
    return at


def html_bodies(at) -> list[str]:
    return [e.proto.body for e in at.get("html")]


def visible_text(html: str) -> str:
    return re.sub(r"<[^>]+>", "", html)


def test_intro_is_short_and_shows_the_flow():
    at = first_screen()
    hero = next(b for b in html_bodies(at) if 'class="xm-hero"' in b)
    # 두 줄 제목 + 부제 한두 문장. 첫 안내가 길어지면 버튼이 첫 화면 밖으로 밀린다
    assert "양식이 제각각인 엑셀,<br>" in hero and "오류까지" in hero
    sub = re.search(r'<p class="xm-sub">(.*?)</p>', hero).group(1)
    assert len(sub) <= 120 and sub.count(".") <= 2
    for chip in ("Python", "Streamlit", "Gemini AI", "openpyxl"):
        assert f">{chip}</li>" in hero
    # 작동 흐름 5단계가 순서대로 (이전의 "30초 체험" 안내를 대신한다)
    flow = visible_text(next(b for b in html_bodies(at) if 'class="xm-flow"' in b))
    positions = [flow.index(s) for s in FLOW_STEPS]
    assert positions == sorted(positions)


def test_hero_numbers_are_real_values():
    from engine import KINDS
    import app
    at = first_screen()
    hero = next(b for b in html_bodies(at) if 'class="xm-hero"' in b)
    assert f"<b>{len(KINDS)}종</b>" in hero and "<b>0줄</b>" in hero
    assert f"<b>{app.load_test_count()}개</b><span>자동 테스트</span>" in hero


def test_buttons_come_right_after_the_hero():
    # 히어로 -> 시나리오 선택 -> 체험 버튼 -> 작동 흐름 띠 순서 (흐름 띠가 버튼을 밀어내지 않는다)
    at = first_screen()
    order = []
    for n in iter_nodes(at.main):
        if type(n).__name__ == "UnknownElement" and getattr(n, "type", None) == "html":
            body = n.proto.body
            order.append("hero" if 'class="xm-hero"' in body else "flow" if 'class="xm-flow"' in body else "css")
        else:
            order.append(getattr(n, "key", None))
    assert order.index("hero") < order.index("scenario") < order.index("sample_btn") \
        < order.index("ai_demo_btn_0") < order.index("flow")


def test_footer_shows_author_and_repo():
    at = first_screen()
    captions = " ".join(c.value for c in at.main.caption)
    assert "만든 사람 신영호" in captions and "(https://github.com/youngho930/excel-merger)" in captions


def test_sample_button_comes_before_the_folded_details():
    # 버튼이 위로 오고, 기준열·미리보기·자세히 보기는 그 아래 접혀 있다
    at = first_screen()
    labels = [e.label for e in at.expander]
    assert labels[:3] == ["샘플 파일 미리보기", "이 시나리오의 기준열 보기", "자세히 보기"]
    order = [getattr(n, "key", None) or getattr(n, "label", None) for n in iter_nodes(at.main)]
    assert order.index("sample_btn") < order.index("샘플 파일 미리보기")
    assert all(not e.proto.expanded for e in at.expander)          # 기본은 접힘


def iter_nodes(node):
    for child in getattr(node, "children", {}).values():
        yield child
        yield from iter_nodes(child)


def test_sample_preview_shows_each_file_with_its_own_headers():
    at = first_screen()
    preview = next(e for e in at.expander if e.label == "샘플 파일 미리보기")
    tables = [n for n in iter_nodes(preview) if type(n).__name__ == "Dataframe"]
    assert len(tables) == 3
    headers = [list(t.value.columns) for t in tables]
    assert len({tuple(h) for h in headers}) == 3                   # 파일마다 열 이름·순서가 다르다
    assert all(1 <= len(t.value) <= 3 for t in tables)             # 데이터는 몇 행만
    notes = " ".join(n.value for n in iter_nodes(preview) if type(n).__name__ == "Markdown")
    assert "머리글이 3행에 있습니다" in notes                       # 제목 줄이 있는 파일


def test_sidebar_about_section():
    at = first_screen()
    text = " ".join(m.value for m in at.sidebar.markdown) + " " + " ".join(c.value for c in at.sidebar.caption)
    assert "이 도구에 대해" in " ".join(s.value for s in at.sidebar.subheader)
    for phrase in ("시나리오는 설정 파일로", "규칙으로 먼저, 남은 열만 AI", "열 이름만",
                   "https://github.com/youngho930/excel-merger", "만든 사람: 신영호"):
        assert phrase in text, phrase


def test_terms_have_question_mark_help():
    at = first_screen()
    at.button(key="sample_btn").click().run()
    at.button(key="run_btn").click().run()
    assert not at.exception, at.exception
    helps = {m.label: m.help for m in at.metric}
    for label in ("남은 오류", "전체 오류", "처리됨", "수정 제안 있음"):
        assert helps[label], label
    subheaders = {s.value: s.proto.help for s in at.subheader}
    assert subheaders[":green-badge[5] 중복 행 고르기"]              # 중복 그룹 설명 (번호는 원형 배지)
    assert "자세히 보기" in [e.label for e in at.expander]
