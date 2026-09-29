"""자유 양식 엔진: 열 묶기, 날짜 판단, 결과 열 검사, 메모리 안 Scenario, YAML 내보내기."""

from datetime import datetime

import pytest
from openpyxl import Workbook

from engine import execute, list_scenarios, load_scenario, prepare
from engine.free_form import (AI, AUTO, USER, ColumnChoice, FreeFormError, build_scenario, check_choices,
                              check_dates, check_file_key, export_yaml, group_headers, make_plans, merge_groups,
                              regroup, remove_merges, scan_tables, scenario_to_dict, skipped_columns)
from engine.normalize import DATE, TEXT
from engine.scenario import MAX_COLUMNS


def xlsx(path, rows):
    wb = Workbook()
    for r in rows:
        wb.active.append(list(r))
    wb.save(path)
    return path


@pytest.fixture
def three(tmp_path):
    a = xlsx(tmp_path / "대리점.xlsx", [
        ("접수번호", "접수일", "고객명", "연락처", "제품명", "담당기사"),
        ("A-1", datetime(2026, 9, 1), "가상일", "010-0000-0001", "WM-100", "기사1"),
        ("A-2", datetime(2026, 9, 2), None, "010-0000-0002", "TV-200", None),
        ("A-3", datetime(2026, 9, 3), "가상삼", "010-0000-0003", "RF-300", "기사2")])
    b = xlsx(tmp_path / "콜센터.xlsx", [
        ("2026년 9월 AS 접수 현황",), (),
        ("등록일", "성명", "전화번호", "모델", "접수번호"),
        ("2026.9.2", "가상이", "010-0000-0002", "TV-200", "A-2"),
        ("2026.9.4", "가상사", "010-0000-0004", "WM-100", "C-1")])
    c = xlsx(tmp_path / "온라인.xlsx", [
        ("접수번호", "고객 명", "연락처", "제품명", "접수일"),
        ("O-1", "가상오", "010-0000-0005", "RF-300", "2026-09-05"),
        ("O-2", "가상삼", "010-0000-0003 ", "RF-300", "2026-09-03")])
    return [a, b, c]


MERGES = [({"고객명", "성명"}, USER), ({"연락처", "전화번호"}, USER), ({"제품명", "모델"}, USER),
          ({"접수일", "등록일"}, USER)]


def by_label(groups):
    return {g.label: g for g in groups}


def test_groups_by_normalized_name_in_first_appearance_order(three):
    tables = scan_tables(three)
    assert [t.header_row for t in tables] == [1, 3, 1]
    groups = group_headers(tables)
    labels = [g.label for g in groups]
    assert labels == ["접수번호", "접수일", "고객명", "연락처", "제품명", "담당기사", "등록일", "성명", "전화번호", "모델"]
    g = by_label(groups)
    assert g["고객명"].names == ["고객명", "고객 명"] and g["고객명"].files == {0, 2}     # 공백 차이는 자동으로 묶음
    assert g["접수번호"].files == {0, 1, 2} and all(x.origin == AUTO for x in groups)


def test_blank_headers_are_skipped(tmp_path):
    p = xlsx(tmp_path / "b.xlsx", [("이름", None, "---", "코드"), ("가", 1, 2, "x")])
    tables = scan_tables([p])
    assert [g.label for g in group_headers(tables)] == ["이름", "코드"]
    assert skipped_columns(tables) == [(0, 1), (0, 2)]


def test_merge_split_and_same_file_is_rejected(three):
    tables = scan_tables(three)
    groups = group_headers(tables)
    g = by_label(groups)
    merged = merge_groups(groups, [g["고객명"].gid, g["성명"].gid], USER)
    m = by_label(merged)["고객명"]
    assert m.names == ["고객명", "성명", "고객 명"] and m.files == {0, 1, 2} and m.origin == USER
    assert len(merged) == len(groups) - 1
    # 같은 파일에 함께 있는 두 열(접수번호·고객명)은 묶을 수 없다
    with pytest.raises(FreeFormError, match="같은 파일"):
        merge_groups(groups, [g["접수번호"].gid, g["고객명"].gid], file_names=[t.file_name for t in tables])
    with pytest.raises(FreeFormError):
        merge_groups(groups, [g["고객명"].gid])
    # 풀기: 기록에서 지우고 다시 묶으면 원래대로
    merges = [({"고객명", "성명"}, AI)]
    again = regroup(tables, merges)
    assert by_label(again)["고객명"].origin == AI
    assert [x.gid for x in regroup(tables, remove_merges(merges, by_label(again)["고객명"]))] == \
        [x.gid for x in groups]


def test_gid_is_stable_when_rereading(three):
    a = [g.gid for g in regroup(scan_tables(three), MERGES)]
    b = [g.gid for g in regroup(scan_tables(three), MERGES)]
    assert a == b and len(set(a)) == len(a)


def test_date_check_ignores_blanks_and_reports_unreadable_values(tmp_path):
    p = xlsx(tmp_path / "d.xlsx", [
        ("일자", "섞임", "코드", "시각", "이름"),
        (datetime(2026, 9, 1), "2026.9.1", "20260901", datetime(2026, 9, 1, 10, 30), "가"),
        ("2026년 9월 2일", "2026.13.02", "20260902", datetime(2026, 9, 2), "나"),
        (None, "2026-09-03", "20260903", datetime(2026, 9, 3), "다"),
        ("2026/09/04", None, None, None, "라")])
    tables = scan_tables([p])
    g = by_label(group_headers(tables))
    ok = check_dates(tables, g["일자"])
    assert ok.is_date and ok.total == 3 and ok.bad == 0            # 빈칸은 빼고 판단
    mixed = check_dates(tables, g["섞임"])
    assert not mixed.is_date and mixed.mixed and mixed.bad == 1 and mixed.examples == ["2026.13.02"]
    code = check_dates(tables, g["코드"])
    assert not code.is_date and not code.mixed                     # 8자리 숫자 글자는 날짜로 보지 않음
    assert not check_dates(tables, g["시각"]).is_date               # 시간이 있는 값이 있으면 문자
    assert not check_dates(tables, g["이름"]).mixed


def choices_for(groups, spec):
    g = by_label(groups)
    return [ColumnChoice(g[label].gid, name, req, dup) for label, name, req, dup in spec]


SPEC = [("접수일", "접수일", True, True), ("고객명", "고객명", True, False), ("연락처", "연락처", True, True),
        ("제품명", "제품명", True, True), ("담당기사", "담당기사", False, False)]


def test_build_scenario_and_run_existing_pipeline(three):
    tables = scan_tables(three)
    groups = regroup(tables, MERGES)
    choices = choices_for(groups, SPEC)
    dates = {g.gid for g in groups if check_dates(tables, g).is_date}
    assert by_label(groups)["접수일"].gid in dates and by_label(groups)["연락처"].gid not in dates
    sc = build_scenario(groups, choices, dates, name="AS 접수")
    assert sc.column_names == ["접수일", "고객명", "연락처", "제품명", "담당기사"]
    assert sc.column("접수일").fmt == DATE and sc.column("고객명").fmt == TEXT
    assert sc.column("고객명").aliases == ("성명", "고객 명")
    assert all(c.range is None and c.allowed is None for c in sc.columns)
    assert sc.dup_keys == ["접수일", "연락처", "제품명"]
    result = execute(sc, make_plans(tables, sc))
    kinds = sorted((i.file, i.row, i.kind) for i in result.issues)
    # 빈칸 1건 + 날짜 표기가 달라도 같은 날로 본 중복 2쌍 (대리점 3행 = 콜센터 4행, 대리점 4행 = 온라인 3행)
    assert kinds == sorted([("대리점.xlsx", 3, "필수값 빈칸"),
                            ("대리점.xlsx", 3, "중복 행"), ("콜센터.xlsx", 4, "중복 행"),
                            ("대리점.xlsx", 4, "중복 행"), ("온라인.xlsx", 3, "중복 행")])
    # 한 파일에만 있는 열(담당기사)은 다른 파일에서 빈칸이지만 필수가 아니므로 오류가 아니다
    assert not [i for i in result.issues if i.standard == "담당기사"]


def test_required_on_column_in_one_file_gives_one_error_per_other_file(three):
    tables = scan_tables(three)
    groups = regroup(tables, MERGES)
    sc = build_scenario(groups, choices_for(groups, [("고객명", "고객명", True, False),
                                                     ("담당기사", "담당기사", True, False)]))
    missing = [i for i in execute(sc, make_plans(tables, sc)).issues if i.row is None]
    assert sorted(i.file for i in missing) == ["온라인.xlsx", "콜센터.xlsx"]


def test_dup_key_forces_required(three):
    tables = scan_tables(three)
    groups = regroup(tables, MERGES)
    sc = build_scenario(groups, choices_for(groups, [("연락처", "연락처", False, True)]))
    assert sc.column("연락처").required


@pytest.mark.parametrize("names, expect", [
    (["", "연락처"], "비어 있습니다"),
    (["고객", "고객 "], "두 번"),
    (["원래 행", "연락처"], "'원래 행' 열과 겹쳐"),
    (["출처파일", "연락처"], "'출처 파일' 열과 겹쳐"),
    (["전화번호", "성명"], "다른 묶음"),          # 결과 열 이름이 다른 묶음의 원본 이름 -> 데이터를 잘못 가져옴
    (["-*-", "연락처"], "글자나 숫자"),
    (["가" * 51, "연락처"], "너무 깁니다"),
])
def test_check_choices_blocks_bad_names(three, names, expect):
    tables = scan_tables(three)
    groups = group_headers(tables)                   # 묶지 않은 상태: 고객명·성명·연락처·전화번호가 따로
    g = by_label(groups)
    choices = [ColumnChoice(g["고객명"].gid, names[0]), ColumnChoice(g["연락처"].gid, names[1])]
    problems = check_choices(groups, choices)
    assert problems and expect in " ".join(problems), problems
    with pytest.raises(FreeFormError):
        build_scenario(groups, choices)


def test_check_choices_limits(three):
    groups = group_headers(scan_tables(three))
    assert check_choices(groups, []) == ["결과에 넣을 열을 하나 이상 골라 주세요."]
    g = groups[0]
    many = [ColumnChoice(g.gid, f"열{i}") for i in range(MAX_COLUMNS + 1)]
    assert any(f"최대 {MAX_COLUMNS}개" in p for p in check_choices(groups, many))


EVIL = ["a: b", "- x", "&anchor", "*alias", "!!python/object:os.system", "# 주석", "줄\n바꿈", "'\"", "{x: 1}",
        "[1, 2]", "?", "@at", "=SUM(A1)", "yes", "null", "1e3"]


def test_yaml_round_trip_with_tricky_names(tmp_path):
    rows = [tuple(EVIL), tuple(range(len(EVIL)))]
    tables = scan_tables([xlsx(tmp_path / "e.xlsx", rows)])
    groups = group_headers(tables)
    choices = [ColumnChoice(g.gid, f"열{i}", i == 0, i == 1) for i, g in enumerate(groups)]
    sc = build_scenario(groups, choices, name="이상한: 이름 #1")
    text = export_yaml(sc, "tricky")
    assert "!!python" not in text.replace("'!!python", "").replace('"!!python', "")   # 태그로 해석되지 않게 따옴표
    path = tmp_path / "tricky.yaml"
    path.write_text(text, encoding="utf-8")
    loaded = load_scenario(path)
    assert scenario_to_dict(loaded) == scenario_to_dict(sc)
    # 원본 이름이 동의어로 정확히 남아, 정식 시나리오로 읽어도 같은 열에 매칭된다
    plans = prepare(loaded, [tmp_path / "e.xlsx"])
    got = {m.standard: m.source_index for m in plans[0].match.matches}
    assert got == {c.name: g.members[0].col_index for c, g in zip(choices, groups)}


def test_exported_yaml_works_as_a_real_scenario(three, tmp_path):
    tables = scan_tables(three)
    groups = regroup(tables, MERGES)
    dates = {g.gid for g in groups if check_dates(tables, g).is_date}
    sc = build_scenario(groups, choices_for(groups, SPEC), dates, name="AS 접수")
    folder = tmp_path / "scenarios"
    folder.mkdir()
    (folder / "as_intake.yaml").write_text(export_yaml(sc, "as_intake"), encoding="utf-8")
    [real] = list_scenarios(folder)
    assert real.key == "as_intake" and real.name == "AS 접수"
    free = {(i.file, i.row, i.standard, i.kind) for i in execute(sc, make_plans(tables, sc)).issues}
    formal = {(i.file, i.row, i.standard, i.kind) for i in execute(real, prepare(real, three)).issues}
    assert free == formal and free


def test_export_rejects_bad_file_key_and_huge_yaml(three):
    groups = group_headers(scan_tables(three))
    sc = build_scenario(groups, [ColumnChoice(groups[0].gid, "번호")])
    for bad in ("", "../x", "한글", "a b", "x" * 61, "a\n"):
        assert check_file_key(bad)
        with pytest.raises(FreeFormError):
            export_yaml(sc, bad)
    sc.columns = [type(c)(name=c.name, aliases=("긴이름" * 40000,), required=c.required, fmt=c.fmt)
                  for c in sc.columns]
    with pytest.raises(FreeFormError, match="너무 커서"):
        export_yaml(sc, "big")
