"""자유 양식 샘플(samples/free_form/)을 정답지(samples/free_form_expected_errors.json)와 대조한다.

기존 3개 시나리오의 정답지(expected_errors.json, 37건)와는 섞지 않는다.
"""

import re

from answer_key import ROOT, compare, free_form_scenario, load_answer, load_free_answer
from engine import execute, list_scenarios, prepare
from engine.free_form import check_dates, export_yaml, group_headers, make_plans, scan_tables

ANSWER = load_free_answer()
FOLDER = ROOT / "samples" / ANSWER["folder"]
PATHS = [FOLDER / f["file"] for f in ANSWER["files"]]


def run(paths=PATHS):
    tables = scan_tables(sorted(paths))            # 가나다순으로 읽어도 결과가 같아야 한다
    sc, groups = free_form_scenario(ANSWER["config"], tables)
    return tables, groups, sc, execute(sc, make_plans(tables, sc))


def test_free_form_matches_answer_key():
    _, _, _, result = run()
    cmp = compare(result.issues, ANSWER["errors"])
    print(cmp.report())
    assert cmp.ok, "\n" + cmp.report()
    assert cmp.matched == len(ANSWER["errors"]) == 8
    counts = result.counts()
    assert (counts["필수값 빈칸"], counts["중복 행"]) == (5, 6)       # 중복 3쌍은 두 행 모두 표시
    assert sum(counts.values()) == 11                                  # 형식·범위·허용값 오류는 없다


def test_headers_are_detected_and_grouped():
    tables, groups, sc, _ = run()
    by_name = {t.file_name: t for t in tables}
    for f in ANSWER["files"]:
        t = by_name[f["file"]]
        assert (t.header_row, t.header_confidence, len(t.rows)) == (f["header_row"], "높음", f["data_rows"])
        assert [h for h in t.headers] == f["columns"]
    auto = group_headers(tables)
    assert len(auto) == 13 and len(groups) == 8                        # 자동 13개 -> 묶은 뒤 8개
    assert {g.label for g in groups} >= {"고객명", "연락처", "제품명", "증상", "접수일"}
    date_groups = [g.label for g in groups if check_dates(tables, g).is_date]
    assert date_groups == ["접수일"]
    assert sc.column("접수일").aliases == ("등록일",)


def test_sample_values_are_fictional():
    tables, *_ = run()
    phones = [c for t in tables for r in t.rows for c in r.cells if isinstance(c, str) and c.startswith("010")]
    assert phones and all(re.fullmatch(r"010-0000-\d{4}", p) for p in phones)
    assert "가상" in ANSWER["note"]


def test_exported_yaml_runs_as_a_real_scenario(tmp_path):
    _, _, sc, result = run()
    folder = tmp_path / "scenarios"
    folder.mkdir()
    (folder / "as_intake.yaml").write_text(export_yaml(sc, "as_intake"), encoding="utf-8")
    [real] = list_scenarios(folder)
    plans = prepare(real, PATHS)
    assert [p.table.header_row for p in plans] == [f["header_row"] for f in ANSWER["files"]]
    assert all(p.match.missing_required == [] for p in plans)
    formal = execute(real, plans)
    cmp = compare(formal.issues, ANSWER["errors"])
    assert cmp.ok, "\n" + cmp.report()
    assert {(i.file, i.row, i.kind) for i in formal.issues} == {(i.file, i.row, i.kind) for i in result.issues}


def test_existing_answer_key_is_not_mixed():
    assert sum(len(s["errors"]) for s in load_answer()["scenarios"].values()) == 37
    assert "free_form" not in load_answer()["scenarios"]
