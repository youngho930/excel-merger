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
    # 두 줄 제목 + 부제 세 문장 이내. 첫 안내가 길어지면 버튼이 첫 화면 밖으로 밀린다
    assert "양식이 제각각인 엑셀,<br>" in hero and "오류까지" in hero
    # 두 번째 줄은 "오류까지"만 강조색 ("취합"은 본문 글자색)
    assert "한 번에 취합하고 <em>오류까지</em>" in hero and re.findall(r"<em>(.*?)</em>", hero) == ["오류까지"]
    sub = re.search(r'<p class="xm-sub">(.*?)</p>', hero).group(1)
    assert len(sub) <= 120 and sub.count(".") <= 3
    assert sub.endswith("정해진 양식이 없어도 원하는 열과 순서를 골라 취합할 수 있습니다.")   # 자유 양식
    for chip in ("Python", "Streamlit", "Gemini AI", "openpyxl"):
        assert f">{chip}</li>" in hero
    # 작동 흐름 5단계가 순서대로 (이전의 "30초 체험" 안내를 대신한다)
    flow = visible_text(next(b for b in html_bodies(at) if 'class="xm-flow"' in b))
    positions = [flow.index(s) for s in FLOW_STEPS]
    assert positions == sorted(positions)


def test_hero_numbers_are_real_values():
    import app
    at = first_screen()
    hero = next(b for b in html_bodies(at) if 'class="xm-hero"' in b)
    m = app.load_measurement()
    # 직접 측정 결과 (project_stats.toml): 큰 숫자는 배율 702 ÷ 87 = 8.07 -> 8배(버림), 설명은 실제 분·초
    assert (m["manual_seconds"], m["tool_seconds"]) == (702, 87)
    assert m["ratio"] == m["manual_seconds"] // m["tool_seconds"] == 8
    assert ('<b>8배</b><span>11분&nbsp;42초 → 1분&nbsp;27초<i class="xm-hide-sm"> (직접 측정)</i></span>'
            in hero)
    assert "11분 → 1분" not in hero                               # 버림으로 차이가 커 보이던 옛 표기
    assert "<b>14 / 14</b><span>오류 검출 (수작업 8건)</span>" in hero
    assert f"<b>{app.load_test_count()}개</b><span>자동 테스트</span>" in hero
    assert "0줄" not in hero                                     # 설계 원칙 쪽(README)으로 옮김
    assert "샘플 3개 파일 기준 직접 측정 · 자세한 조건은 GitHub" in hero
    # 숫자 카드 3개가 하나의 패널(ul.xm-stats) 안에, 숫자(b) 다음에 설명(span)
    panel = re.search(r'<ul class="xm-stats"[^>]*>(.*?)</ul>', hero, re.S).group(1)
    # 설명(span) 안에는 모바일에서 숨기는 <i class="xm-hide-sm"> 만 올 수 있다
    assert len(re.findall(r'<li><b>[^<]+</b><span>[^<]+(?:<i class="xm-hide-sm">[^<]+</i>)?</span></li>',
                          panel)) == 3


def test_buttons_come_right_after_the_hero():
    # 히어로 -> 작동 흐름 띠 -> 시나리오 선택 -> 체험 버튼 순서 (버튼 앞에는 이것들만 온다)
    at = first_screen()
    order = []
    for n in iter_nodes(at.main):
        if type(n).__name__ == "UnknownElement" and getattr(n, "type", None) == "html":
            body = n.proto.body
            order.append("hero" if 'class="xm-hero"' in body else "flow" if 'class="xm-flow"' in body else "css")
        else:
            order.append(getattr(n, "key", None))
    assert order.index("hero") < order.index("flow") < order.index("scenario") < order.index("sample_btn") \
        < order.index("ai_demo_btn_0")
    # 버튼보다 앞에 접기 메뉴나 표 같은 다른 요소가 끼어들지 않는다 (첫 화면에 버튼이 들어오도록)
    before = list(iter_nodes(at.main))[:order.index("sample_btn")]
    assert not [n for n in before if type(n).__name__ in ("Expander", "Dataframe", "Metric")]


def test_footer_shows_author_and_repo():
    at = first_screen()
    captions = " ".join(c.value for c in at.main.caption)
    assert "만든 사람 신영호" in captions and "(https://github.com/youngho930/excel-merger)" in captions


def test_sample_button_comes_before_the_folded_details():
    # 버튼이 위로 오고, 미리보기와 "시나리오 자세히 보기"(기준열 표 + 중복기준 + 설명)는 그 아래 접혀 있다
    at = first_screen()
    labels = [e.label for e in at.expander]
    assert labels[:2] == ["샘플 파일 미리보기", "시나리오 자세히 보기"]
    assert "이 시나리오의 기준열 보기" not in labels
    order = [getattr(n, "key", None) or getattr(n, "label", None) for n in iter_nodes(at.main)]
    assert order.index("sample_btn") < order.index("샘플 파일 미리보기") < order.index("시나리오 자세히 보기")
    assert all(not e.proto.expanded for e in at.expander)          # 기본은 접힘
    # 합친 칸에 기준열 표, 중복기준 도움말, 기존 설명 글이 모두 들어 있다
    merged = next(e for e in at.expander if e.label == "시나리오 자세히 보기")
    tables = [n for n in iter_nodes(merged) if type(n).__name__ == "Dataframe"]
    assert len(tables) == 1 and "기준열" in tables[0].value.columns
    md_nodes = [n for n in iter_nodes(merged) if type(n).__name__ == "Markdown"]
    assert any(n.value.startswith("중복기준: ") and n.proto.help for n in md_nodes)
    text = " ".join(n.value for n in md_nodes)
    for phrase in ("**시나리오**는", "**샘플 파일로 바로 체험**", "**AI 매칭 체험**"):
        assert phrase in text, phrase


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


def test_sidebar_has_only_repo_link_and_author():
    at = first_screen()
    text = " ".join(m.value for m in at.sidebar.markdown) + " " + " ".join(c.value for c in at.sidebar.caption)
    for phrase in ("https://github.com/youngho930/excel-merger", "만든 사람: 신영호"):
        assert phrase in text, phrase
    assert "설계 원칙" not in text and "시나리오는 설정 파일로" not in text   # 원칙은 페이지 아래로 옮겼다
    assert not at.sidebar.subheader


def test_about_section_shows_three_principle_cards_at_the_bottom():
    at = first_screen()
    about = next(n for n in iter_nodes(at.main) if getattr(n, "key", None) == "about")
    assert "이 도구에 대해" in [n.value for n in iter_nodes(about) if type(n).__name__ == "Subheader"]
    cards = [next(n for n in iter_nodes(about) if getattr(n, "key", None) == f"card_principle_{i}")
             for i in (1, 2, 3)]
    titles = ["시나리오는 설정 파일로", "규칙으로 먼저, 남은 열만 AI", "원본은 수정하지 않고 외부로는 열 이름만 전송"]
    for card, title in zip(cards, titles):
        heads = [n.value for n in iter_nodes(card) if type(n).__name__ == "Markdown"]
        notes = [n.value for n in iter_nodes(card) if type(n).__name__ == "Caption"]
        assert len(heads) == 1 and ":material/" in heads[0] and f"**{title}**" in heads[0]   # 아이콘 + 제목
        assert len(notes) == 1 and notes[0]                                                     # 한 줄 설명
    captions = " ".join(n.value for n in iter_nodes(about) if type(n).__name__ == "Caption")
    assert "손이 많이 가고 실수가 잦습니다" in captions                                          # 만든 목적
    # 페이지 아래쪽: 업로드 단계보다 뒤, 푸터보다 앞
    keys = [getattr(n, "key", None) for n in iter_nodes(at.main)]
    assert keys.index("card_upload") < keys.index("about") < keys.index("footer")


def test_page_is_centered_and_sidebar_starts_collapsed():
    import ast

    import app
    assert "max-width: calc(1200px + 5rem)" in app.APP_CSS and "margin: 0 auto" in app.APP_CSS
    tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
    config = next(n for n in ast.walk(tree) if isinstance(n, ast.Call)
                  and isinstance(n.func, ast.Attribute) and n.func.attr == "set_page_config")
    kw = {k.arg: ast.literal_eval(k.value) for k in config.keywords}
    assert kw["layout"] == "wide" and kw["initial_sidebar_state"] == "collapsed"


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
