"""화면(app.py) 테스트 — streamlit.testing.v1.AppTest.

샘플 체험 -> 실행 -> 결과 요약 건수, 수정 제안값 일괄 적용 on/off, 열 매칭 수정.
파일 읽기·실행은 실제 하위 프로세스(engine.jobs)로 돈다.
"""

from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest

from answer_key import ROOT, load_answer
from engine.validate import KINDS

APP = str(ROOT / "app.py")
EXPECTED_TOTAL = {"incoming_inspection": 16, "stock_count": 15, "monthly_report": 11}
FIXED_RGB = "C6EFCE"


def start(key: str) -> AppTest:
    at = AppTest.from_file(APP, default_timeout=120)
    at.run()
    assert not at.exception
    at.selectbox(key="scenario").set_value(key).run()
    at.button(key="sample_btn").click().run()
    assert not at.exception, at.exception
    assert not at.error, [e.value for e in at.error]
    return at


def run_merge(at: AppTest) -> AppTest:
    at.button(key="run_btn").click().run()
    assert not at.exception, at.exception
    assert not at.error, [e.value for e in at.error]
    return at


def metric(at: AppTest, label: str) -> str:
    return next(m.value for m in at.metric if m.label == label)


def output(at: AppTest, applied: bool) -> Path:
    return Path(at.session_state["outputs"][applied])


def test_first_screen_has_no_errors():
    at = AppTest.from_file(APP, default_timeout=60)
    at.run()
    assert not at.exception and not at.error
    assert at.title[0].value == "엑셀 자동 취합·검증기"
    assert len(at.selectbox(key="scenario").options) == 3
    assert not at.metric   # 파일을 올리기 전에는 결과가 없다


@pytest.mark.parametrize("key", sorted(EXPECTED_TOTAL))
def test_sample_run_summary(key):
    at = start(key)
    info = load_answer()["scenarios"][key]
    assert at.success[0].value == f"샘플 파일 {len(info['files'])}개를 읽었습니다."
    assert not at.warning   # 샘플은 매칭 경고가 없다
    run_merge(at)

    assert metric(at, "파일 수") == f"{len(info['files'])}개"
    assert metric(at, "취합 행 수") == f"{sum(f['data_rows'] for f in info['files'])}행"
    assert metric(at, "오류 합계") == f"{EXPECTED_TOTAL[key]}건"
    per_kind = {k: int(metric(at, k).rstrip("건")) for k in KINDS}
    assert sum(per_kind.values()) == EXPECTED_TOTAL[key]

    # 오류 표와 종류별 거르기
    df = at.dataframe[-1].value
    assert len(df) == EXPECTED_TOTAL[key]
    at.multiselect(key="kind_filter").set_value(["중복 행"]).run()
    df = at.dataframe[-1].value
    assert len(df) == per_kind["중복 행"] and set(df["오류 종류"]) == {"중복 행"}

    # 다운로드 버튼과 결과 파일 (범례는 별도 시트, 오류목록에 이름 없는 열 없음)
    assert at.get("download_button")
    path = output(at, False)
    assert path.is_file() and path.parent.parent == Path(at.session_state["workspace"])
    assert load_workbook(path).sheetnames == ["취합결과", "오류목록", "범례"]
    errors = pd.read_excel(path, sheet_name="오류목록")
    assert len(errors) == EXPECTED_TOTAL[key]
    assert not [c for c in errors.columns if str(c).startswith("Unnamed")]


def _cell(ws, file, row, column):
    header = [c.value for c in ws[1]]
    col = header.index(column) + 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value == file and ws.cell(r, 2).value == row:
            return ws.cell(r, col)
    raise KeyError((file, row))


def test_apply_suggestions_off_then_on():
    at = run_merge(start("incoming_inspection"))
    assert at.checkbox(key="apply_suggestions").value is False   # 기본값: 꺼짐
    assert metric(at, "수정 제안 있음") == "1건"

    # 꺼짐: 원래 값 그대로, "처리" 열 없음
    off = load_workbook(output(at, False))
    cell = _cell(off["취합결과"], "수입검사_대성정공.xlsx", 14, "판정")
    assert cell.value == "합격 " and not cell.fill.fgColor.rgb.endswith(FIXED_RGB)
    assert "처리" not in [c.value for c in off["오류목록"][1]]
    assert "처리" not in at.dataframe[-1].value.columns

    # 켬: 제안값으로 바꾸고 초록색, 오류목록 "처리" 열에 "자동 수정됨"
    at.checkbox(key="apply_suggestions").check().run()
    assert not at.exception and not at.error
    on_path = output(at, True)
    assert on_path != output(at, False)
    on = load_workbook(on_path)
    cell = _cell(on["취합결과"], "수입검사_대성정공.xlsx", 14, "판정")
    assert cell.value == "합격" and cell.fill.fgColor.rgb.endswith(FIXED_RGB)
    errors = pd.read_excel(on_path, sheet_name="오류목록")
    fixed = errors[errors["처리"] == "자동 수정됨"]
    assert list(zip(fixed["파일"], fixed["행"], fixed["기준열"])) == [("수입검사_대성정공.xlsx", 14, "판정")]
    assert fixed["값"].iloc[0] == "합격 " and fixed["오류 종류"].iloc[0] == "허용값 아닌 값"
    assert len(errors) == 16   # 오류 건수는 그대로 (고친 사실만 남긴다)
    assert "자동 수정됨" in [c.value for c in on["범례"]["A"]]
    df = at.dataframe[-1].value
    assert (df["처리"] == "자동 수정됨").sum() == 1

    # 제안값이 없는 다른 오류 셀은 원래 값 그대로
    cell = _cell(on["취합결과"], "수입검사_미래부품.xlsx", 7, "판정")
    assert cell.value == "OK"

    # 다시 끄면 원래 값 파일
    at.checkbox(key="apply_suggestions").uncheck().run()
    assert at.checkbox(key="apply_suggestions").value is False
    assert "처리" not in at.dataframe[-1].value.columns


def test_matching_can_be_changed_on_screen():
    at = start("stock_count")
    load_id = at.session_state["load_id"]
    plans = at.session_state["plans"]
    fi = 0
    m = plans[fi].match
    si = next(i for i, cm in enumerate(m.matches) if cm.required)
    std = m.matches[si].standard
    first = m.matches[si].source_index
    other = next(cm.source_index for cm in m.matches if cm.standard != std and cm.source_index is not None)
    key = f"map_{load_id}_{fi}_{si}"
    assert at.selectbox(key=key).value == first

    # 필수 열의 매칭을 해제하면 누락 경고가 뜬다
    at.selectbox(key=key).set_value(None).run()
    assert not at.exception
    assert any(f"필수 열 '{std}'" in w.value for w in at.warning)
    assert at.session_state["plans"][fi].match.get(std).method == "매칭 안 됨"

    # 이미 다른 기준열에 쓰인 원본 열을 고르면 한국어 안내 후 되돌린다
    at.selectbox(key=key).set_value(other).run()
    assert any("이미" in e.value for e in at.error)
    assert at.selectbox(key=key).value is None

    # 원래 열로 되돌리면 "사용자 지정", 경고 사라짐 -> 실행 결과는 정답과 같다
    at.selectbox(key=key).set_value(first).run()
    assert at.session_state["plans"][fi].match.get(std).method == "사용자 지정"
    assert not at.warning
    run_merge(at)
    assert metric(at, "오류 합계") == "15건"


def test_changing_scenario_resets_screen():
    at = run_merge(start("monthly_report"))
    assert at.metric
    at.selectbox(key="scenario").set_value("stock_count").run()
    assert not at.exception
    assert not at.metric and at.session_state["plans"] is None


def test_upload_file_name_is_reduced_to_base_name():
    from app import safe_name
    assert safe_name("../../etc/passwd.xlsx") == "passwd.xlsx"
    assert safe_name("..\\..\\Windows\\a.xlsx") == "a.xlsx"
    assert safe_name("C:\\Users\\x\\b.xlsx") == "b.xlsx"
    assert safe_name("보고서.xlsx") == "보고서.xlsx"
    assert safe_name("..") == ".."   # 이런 이름은 load_files 가 거부한다


def test_upload_limit_matches_reader_limit():
    import tomllib
    from engine.reader import MAX_FILE_BYTES
    cfg = tomllib.loads((ROOT / ".streamlit" / "config.toml").read_text(encoding="utf-8"))
    assert cfg["server"]["maxUploadSize"] * 1024 * 1024 == MAX_FILE_BYTES
    assert cfg["client"]["showErrorDetails"] == "none"
