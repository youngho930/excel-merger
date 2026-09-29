"""디자인용 HTML/CSS 보안 원칙.

1. st.html / unsafe_allow_html 로 내보내는 것은 코드에 고정된 문자열(모듈 상수)이나
   app.SAFE_HTML_FUNCS 함수의 결과뿐이다. 그 함수들은 앱이 계산한 정수만 받는다.
2. 파일 이름·열 이름·시나리오 이름·설명·오류 메시지는 그 HTML에 들어가지 않는다 (실제 화면에서 확인).
3. 외부에서 불러오는 것은 Pretendard 글꼴 CSS 하나뿐이다.
"""

import ast
import inspect
import re
import tomllib

import pytest
from streamlit.testing.v1 import AppTest

from answer_key import ROOT
from engine import KINDS

import app

APP_PATH = ROOT / "app.py"
TREE = ast.parse(APP_PATH.read_text(encoding="utf-8"))
PRETENDARD = "https://cdn.jsdelivr.net/gh/orioncactus/pretendard@"
DESIGN_CONSTANTS = ("APP_CSS", "HERO_TEMPLATE", "HERO_TESTS_CARD", "FLOW_TEMPLATE", "_ICON")


def module_assignments() -> dict[str, list[ast.expr]]:
    """모듈 맨 바깥에서 이름에 대입한 값들 (같은 이름에 여러 번 대입하면 모두)."""
    found: dict[str, list[ast.expr]] = {}
    for node in TREE.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name):
                    found.setdefault(target.id, []).append(node.value)
    return found


def reassigned_anywhere(name: str) -> bool:
    """함수 안 등 모듈 맨 바깥이 아닌 곳에서 name 에 대입하거나 global 로 바꾸는지."""
    top = {id(n) for n in TREE.body}
    for node in ast.walk(TREE):
        if isinstance(node, ast.Global) and name in node.names:
            return True
        if isinstance(node, (ast.Assign, ast.AugAssign, ast.AnnAssign)) and id(node) not in top:
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            if any(isinstance(t, ast.Name) and t.id == name for t in targets):
                return True
    return False


def literal_constants() -> set[str]:
    """글자 그대로의 문자열 하나만 대입한 모듈 상수 (f-string·변수 조합 없음)."""
    out = set()
    for name, values in module_assignments().items():
        if len(values) == 1 and isinstance(values[0], ast.Constant) and isinstance(values[0].value, str) \
                and not reassigned_anywhere(name):
            out.add(name)
    return out


def html_calls() -> list[ast.Call]:
    calls = []
    for node in ast.walk(TREE):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        is_html = (isinstance(func, ast.Attribute) and func.attr == "html") or \
                  (isinstance(func, ast.Name) and func.id == "html")
        unsafe = any(kw.arg == "unsafe_allow_html" and not (isinstance(kw.value, ast.Constant) and kw.value.value is False)
                     for kw in node.keywords)
        if is_html or unsafe:
            calls.append(node)
    return calls


def html_argument(call: ast.Call) -> ast.expr:
    if call.args:
        return call.args[0]
    for kw in call.keywords:
        if kw.arg in ("body", "html"):
            return kw.value
    raise AssertionError(f"{call.lineno}행: HTML 인자를 찾지 못했습니다.")


def test_every_raw_html_call_uses_a_constant_or_a_safe_function():
    calls = html_calls()
    assert len(calls) >= 3, "st.html 호출(CSS·히어로·흐름 띠)을 찾지 못했습니다. 검사가 비어 있으면 안 됩니다."
    constants = literal_constants()
    assert "APP_CSS" in constants
    for call in calls:
        arg = html_argument(call)
        if isinstance(arg, ast.Name):
            assert arg.id in constants, f"{call.lineno}행: {arg.id} 는 코드에 고정된 문자열 상수가 아닙니다."
        elif isinstance(arg, ast.Call) and isinstance(arg.func, ast.Name):
            assert arg.func.id in app.SAFE_HTML_FUNCS, f"{call.lineno}행: {arg.func.id} 는 안전 함수 목록에 없습니다."
        else:
            raise AssertionError(f"{call.lineno}행: HTML 인자는 상수나 안전 함수 호출이어야 합니다: {ast.unparse(arg)}")


def test_markdown_is_never_rendered_as_html():
    # 디자인 HTML은 st.html 로만. st.markdown 등에서 unsafe_allow_html 을 켜지 않는다
    for call in html_calls():
        assert not any(kw.arg == "unsafe_allow_html" for kw in call.keywords), \
            f"{call.lineno}행: unsafe_allow_html 대신 st.html(고정 문자열)을 쓰세요."


@pytest.mark.parametrize("name", app.SAFE_HTML_FUNCS)
def test_safe_functions_only_fill_constant_templates_with_whole_numbers(name):
    func = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == name)
    assignments = module_assignments()
    body_nodes = [n for stmt in func.body for n in ast.walk(stmt)]   # 타입 표기(int | None)는 빼고 본문만
    for node in body_nodes:
        assert not isinstance(node, ast.JoinedStr), f"{name}: f-string 을 쓰지 마세요 (고정 템플릿 + 정수만)."
        if isinstance(node, ast.BinOp):
            raise AssertionError(f"{name}: 문자열을 이어 붙이지 마세요 (고정 템플릿 + 정수만).")
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) and node.func.attr == "format":
            receiver = node.func.value
            assert isinstance(receiver, ast.Name) and receiver.id in assignments \
                and receiver.id in DESIGN_CONSTANTS and not reassigned_anywhere(receiver.id), \
                f"{name}: format 은 모듈의 고정 템플릿에만 씁니다."
            assert not node.args, f"{name}: format 에는 이름 붙은 값만 넘깁니다."
            for kw in node.keywords:
                v = kw.value
                ok = (isinstance(v, ast.Call) and isinstance(v.func, ast.Name) and v.func.id == "_whole_number") \
                    or isinstance(v, ast.Name)
                assert ok, f"{name}: {kw.arg} 값은 _whole_number(...) 를 거쳐야 합니다."
    # 실제로 글자·bool·음수는 거부한다
    fn = getattr(app, name)
    n_params = len(inspect.signature(fn).parameters)
    for bad in ("<img src=x onerror=alert(1)>", "5", True, -1, 1.5):
        with pytest.raises(TypeError):
            fn(*([bad] * n_params))
    out = fn(*([7] * n_params))
    assert "7" in out and "{" not in out and "}" not in out


def test_design_constants_do_not_load_anything_external():
    for name in DESIGN_CONSTANTS:
        text = getattr(app, name)
        low = text.lower()
        assert "http:" not in low and "https:" not in low and "//" not in re.sub(r"/\*.*?\*/", "", low, flags=re.S), name
        assert "@import" not in low and "url(" not in low and "<script" not in low and "<link" not in low, name
        assert not re.search(r"\son\w+\s*=", low), f"{name}: 이벤트 속성(onclick 등)을 쓰지 마세요."


def test_only_external_resource_is_the_pretendard_css():
    cfg = tomllib.loads((ROOT / ".streamlit" / "config.toml").read_text(encoding="utf-8"))
    theme = cfg.get("theme", {})

    def urls(value):
        if isinstance(value, dict):
            return [u for v in value.values() for u in urls(v)]
        if isinstance(value, list):
            return [u for v in value for u in urls(v)]
        return re.findall(r"https?://[^\s,\"']+", value) if isinstance(value, str) else []

    found = urls(theme)
    assert len(found) == 1 and found[0].startswith(PRETENDARD), found
    # 글꼴을 못 불러와도 대체 글꼴로 넘어간다
    families = theme["font"].split(":", 1)[0]
    assert "Pretendard" in families and families.rstrip().endswith("sans-serif")
    # 기존 설정은 그대로
    assert cfg["server"]["maxUploadSize"] == 20
    assert cfg["client"]["showErrorDetails"] == "none" and cfg["client"]["toolbarMode"] == "minimal"
    assert cfg["browser"]["gatherUsageStats"] is False


def test_rendered_html_contains_no_user_or_scenario_text():
    at = AppTest.from_file(str(APP_PATH), default_timeout=120)
    at.run()
    at.selectbox(key="scenario").set_value("stock_count").run()
    at.button(key="sample_btn").click().run()
    at.button(key="run_btn").click().run()
    assert not at.exception, at.exception
    allowed = {app.APP_CSS, app.hero_html(len(KINDS), app.load_test_count()), app.flow_html(len(KINDS))}
    bodies = [e.proto.body for e in at.get("html")]
    assert len(bodies) >= 3
    assert set(bodies) <= allowed
    # 샘플 파일 이름·열 이름·시나리오 이름이 HTML 어디에도 없다
    joined = "".join(bodies)
    plans = at.session_state["plans"]
    assert plans
    for plan in plans:
        assert plan.file_name not in joined
        assert not [h for h in plan.table.headers if h and len(str(h)) > 1 and str(h) in joined]
    for s in app.list_scenarios(app.SCENARIO_DIR):
        assert s.name not in joined
