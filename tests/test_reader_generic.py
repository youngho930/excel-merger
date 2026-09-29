"""자유 양식의 머리글 자동 탐지 (시나리오 어휘 없이)."""

from datetime import datetime

from openpyxl import Workbook

from engine.free_form import scan_tables
from engine.reader import detect_header_row_generic, label_key, read_table


def xlsx(path, rows):
    wb = Workbook()
    for r in rows:
        wb.active.append(list(r))
    wb.save(path)
    return path


def test_labels_skip_numbers_dates_and_blanks():
    assert label_key("고객명") == "고객명"
    assert label_key("010-0000-1234") == "010-0000-1234".replace("-", "")   # 전화번호는 글자
    for v in (None, "  ", 1200, "1,200", "2026.9.2", "2026-09-05", "2026년 9월 2일", datetime(2026, 9, 2)):
        assert label_key(v) is None, v


def test_title_rows_above_header():
    rows = [("2026년 9월 AS 접수 현황",), (), ("등록일", "성명", "전화번호", "모델"),
            ("2026.9.2", "가상일", "010-0000-0001", "WM-100"), ("2026.9.3", "가상이", "010-0000-0002", "WM-200")]
    g = detect_header_row_generic(rows)
    assert (g.row, g.confident) == (3, True)


def test_mostly_text_data_rows_tie_is_broken_by_other_files(tmp_path):
    # 데이터 행도 모두 글자라 머리글 행과 글자 칸 수가 같다 -> 다른 파일의 머리글 후보와 겹치는 행을 고른다
    rows = [("구분", "메모", "비고"),                       # 겹치지 않는 3칸 (위쪽)
            ("고객명", "연락처", "증상"),                    # 다른 파일과 겹치는 3칸 = 진짜 머리글
            ("가상일", "010-0000-0001", "전원 안 켜짐"),
            ("가상이", "010-0000-0002", "소음")]
    alone = detect_header_row_generic(rows)
    assert alone.row == 1 and not alone.confident             # 혼자면 위쪽 행, 확신 낮음
    g = detect_header_row_generic(rows, frozenset({"고객명", "연락처", "제품명"}))
    assert (g.row, g.confident) == (2, True)
    # 실제 파일 두 개로: 다른 파일의 머리글 후보가 동점을 가른다
    a = xlsx(tmp_path / "a.xlsx", rows)
    b = xlsx(tmp_path / "b.xlsx", [("고객명", "연락처", "제품명", "접수일"),
                                   ("가상삼", "010-0000-0003", "TV-1", datetime(2026, 9, 1))])
    ta, tb = scan_tables([a, b])
    assert (ta.header_row, ta.header_confidence) == (2, "높음")
    assert tb.header_row == 1
    assert [r.excel_row for r in ta.rows] == [3, 4]


def test_all_text_rows_without_peers_pick_upper_row():
    # 머리글과 데이터가 모두 글자이고 비교할 다른 파일도 없으면 위쪽 행 (머리글은 데이터보다 위에 있다)
    rows = [("고객명", "연락처", "증상"), ("가상일", "010-0000-0001", "소음"), ("가상이", "010-0000-0002", "고장")]
    g = detect_header_row_generic(rows)
    assert g.row == 1 and not g.confident


def test_user_header_row_and_top_rows_kept(tmp_path):
    p = xlsx(tmp_path / "t.xlsx", [("제목",), ("A", "B"), ("x", "y")])
    t = read_table(p, None, header_row=1)
    assert t.header_row == 1 and t.header_confidence == "지정"
    assert t.top[:2] == [("제목",), ("A", "B")]           # 미리보기용 위쪽 행
