from datetime import date, datetime

import pytest

from engine.normalize import normalize_header, parse_date, parse_int, parse_real, parse_value
from engine.scenario import ColumnSpec
from engine.validate import validate_cell


@pytest.mark.parametrize("a,b", [("Insp. Date", "inspdate"), ("P/N", "pn"), ("입고 수량", "입고수량"),
                                 ("LOT No", "lotno"), ("  Part_No ", "partno"), ("ＰＡＲＴ　ＮＯ", "partno")])
def test_normalize_header(a, b):
    assert normalize_header(a) == b


@pytest.mark.parametrize("raw,expected", [
    (datetime(2026, 9, 1), date(2026, 9, 1)),
    (date(2026, 9, 1), date(2026, 9, 1)),
    ("2026-09-03", date(2026, 9, 3)),
    ("2026.09.02", date(2026, 9, 2)),
    ("2026/09/03", date(2026, 9, 3)),
    ("2026.9.3", date(2026, 9, 3)),
    ("2026. 9. 3.", date(2026, 9, 3)),
    ("2026년 9월 27일", date(2026, 9, 27)),
    ("2026년9월27일", date(2026, 9, 27)),
    ("20260903", date(2026, 9, 3)),
    ("2026-09-03 00:00:00", date(2026, 9, 3)),
    ("2026-08", date(2026, 8, 1)),
    ("2026.8", date(2026, 8, 1)),
    ("2026년 8월", date(2026, 8, 1)),
    (" 2026-08 ", date(2026, 8, 1)),
])
def test_dates_accepted(raw, expected):
    p = parse_date(raw)
    assert p.ok and p.value == expected


@pytest.mark.parametrize("raw,reason_part", [
    ("2026.13.02", "13월"),
    ("2026.9.31", "9월 31일"),
    ("2026.2.30", "2월 30일"),
    ("2026-13", "13월"),
    ("26.9.3", "날짜로 읽을 수 없는"),
    ("2026-09/03", "날짜로 읽을 수 없는"),
    ("다음주 월요일", "날짜로 읽을 수 없는"),
    ("1800-01-01", "연도"),
])
def test_dates_rejected(raw, reason_part):
    p = parse_date(raw)
    assert not p.ok and reason_part in p.reason


def test_number_in_date_cell_is_error_with_suggestion():
    p = parse_date(20260803)
    assert not p.ok and p.suggestion == date(2026, 8, 3)
    serial = (date(2026, 8, 3) - date(1899, 12, 30)).days
    p = parse_date(serial)
    assert not p.ok and p.suggestion == date(2026, 8, 3)
    p = parse_date(12)           # 1900년 1월: 합리적 범위 밖이라 제안하지 않음
    assert not p.ok and p.suggestion is None
    p = parse_date(20261340)     # 8자리지만 존재하지 않는 날짜
    assert not p.ok and p.suggestion is None
    spec = ColumnSpec("검사일", fmt="날짜", required=True)
    r = validate_cell(20260803, spec)
    assert r.kind == "형식 오류" and r.message.endswith("2026-08-03으로 수정 제안")
    assert r.suggestion == date(2026, 8, 3)


@pytest.mark.parametrize("raw,val", [(12, 12), (12.0, 12), ("12", 12), (" 1,000 ", 1000), ("-5", -5),
                                     ("12.0", 12)])
def test_int_accepted(raw, val):
    p = parse_int(raw)
    assert p.ok and p.value == val and isinstance(p.value, int)


@pytest.mark.parametrize("raw,reason", [
    ("12개", "정수 칸에 문자('개')가 섞임"),
    ("20EA", "정수 칸에 문자('EA')가 섞임"),
    ("약 120", "정수 칸에 문자('약')가 섞임"),
    (12.5, "정수 칸에 소수(12.5)가 있음"),
    ("1,00", "정수로 읽을 수 없는 값"),
    (True, "참/거짓"),
])
def test_int_rejected(raw, reason):
    p = parse_int(raw)
    assert not p.ok and reason in p.reason


@pytest.mark.parametrize("raw,reason", [("95%", "실수 칸에 문자('%')가 섞임"),
                                        ("3건", "실수 칸에 문자('건')가 섞임"),
                                        ("약 500", "실수 칸에 문자('약')가 섞임"),
                                        (float("nan"), "실수로 읽을 수 없는")])
def test_real_rejected(raw, reason):
    p = parse_real(raw)
    assert not p.ok and reason in p.reason


def test_real_accepted():
    assert parse_real("1,234.5").value == 1234.5
    assert parse_real(3).value == 3


def test_text_keeps_spaces_and_converts_numbers():
    assert parse_value("합격 ", "문자").value == "합격 "
    assert parse_value(1001.0, "문자").value == "1001"
    assert parse_value(datetime(2026, 9, 1), "문자").value == "2026-09-01"


def test_cell_check_order_first_error_only():
    spec = ColumnSpec("수량", fmt="정수", required=True, range=(0, 10), allowed=(1, 2, 3))
    assert validate_cell(None, spec).kind == "필수값 빈칸"
    assert validate_cell("  ", spec).kind == "필수값 빈칸"
    assert validate_cell("x", spec).kind == "형식 오류"
    assert validate_cell(-1, spec).kind == "범위 밖 값"
    assert validate_cell(5, spec).kind == "허용값 아닌 값"
    assert validate_cell(2, spec).ok
    optional = ColumnSpec("비고", fmt="문자")
    assert validate_cell(None, optional).ok


def test_allowed_messages():
    spec = ColumnSpec("판정", fmt="문자", required=True, allowed=("합격", "불합격"))
    r = validate_cell(" 합격", spec)
    assert r.message == '앞뒤 공백 있음 → "합격"으로 수정 제안' and r.suggestion == "합격"
    r = validate_cell("불 합격", spec)
    assert r.message == '공백 차이 → "불합격"일 가능성' and r.suggestion == "불합격"
    r = validate_cell("OK", spec)
    assert r.message == "허용값(합격/불합격)이 아닌 'OK'" and r.suggestion is None
