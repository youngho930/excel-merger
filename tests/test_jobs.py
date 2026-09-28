"""하위 프로세스 + 시간 제한(engine/jobs.py) 테스트."""

import time
import zipfile

import pytest
from openpyxl import Workbook

from answer_key import ROOT
from engine import ReadError, load_scenario
from engine.jobs import JobError, TimeLimitError, run_job
from engine.reader import MAX_SHEET_XML_BYTES

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def make_no_dimension_sheet(path, target_bytes=MAX_SHEET_XML_BYTES - 2 * 1024 * 1024):
    """크기 정보(<dimension>)가 없는 큰 시트 XML 하나짜리 xlsx (devlog '남은 한계'의 재현 파일).

    zip 검사(압축을 푼 크기·압축률·시트 XML 크기)는 통과하지만, openpyxl이 여는 동안 XML 전체를 훑는다.
    """
    wb = Workbook()
    wb.active.append(["창고", "품목코드"])
    wb.save(path)
    parts = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{NS}"><sheetData>',
             '<row r="1"><c r="A1" t="inlineStr"><is><t>창고</t></is></c>'
             '<c r="B1" t="inlineStr"><is><t>품목코드</t></is></c></row>']
    size = sum(len(p) for p in parts)
    r = 2
    while size < target_bytes:
        s = f'<row r="{r}"><c r="A{r}"><v>{r % 997}</v></c><c r="B{r}"><v>{r % 9973}</v></c></row>'
        parts.append(s)
        size += len(s)
        r += 1
    parts.append("</sheetData></worksheet>")
    sheet = "".join(parts).encode("utf-8")
    with zipfile.ZipFile(path) as zf:
        items = {i.filename: zf.read(i.filename) for i in zf.infolist()}
    items["xl/worksheets/sheet1.xml"] = sheet
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in items.items():
            zf.writestr(name, data)
    return path


def test_job_is_stopped_at_time_limit():
    t = time.perf_counter()
    with pytest.raises(TimeLimitError, match=r"처리 시간 제한\(2초\)"):
        run_job("sleep", 30, timeout=2)
    assert time.perf_counter() - t < 10


def test_job_returns_value():
    assert run_job("sleep", 0, timeout=30) == 0


def test_unknown_job_rejected():
    with pytest.raises(ValueError):
        run_job("os.system", "echo")


def test_read_error_comes_back_in_korean(tmp_path):
    sc = load_scenario(ROOT / "scenarios" / "stock_count.yaml")
    bad = tmp_path / "깨짐.xlsx"
    bad.write_bytes(b"not a zip")
    with pytest.raises(ReadError, match="엑셀 파일로 열 수 없습니다"):
        run_job("prepare", sc, [bad], timeout=30)


def test_unexpected_failure_hides_details():
    # 잘못된 인자: 자식 안에서 예상하지 못한 예외(AttributeError) -> 내부 정보 없는 JobError
    with pytest.raises(JobError) as e:
        run_job("execute", None, None, str(ROOT / "output"), timeout=30)
    msg = str(e.value)
    assert "Traceback" not in msg and "AttributeError" not in msg and str(ROOT) not in msg


def test_prepare_through_job_matches_direct(tmp_path):
    from engine import prepare
    sc = load_scenario(ROOT / "scenarios" / "stock_count.yaml")
    paths = sorted((ROOT / "samples" / "stock_count").glob("*.xlsx"))
    plans = run_job("prepare", sc, paths, timeout=60)
    direct = prepare(sc, paths)
    assert [(p.file_name, p.table.header_row, len(p.table.rows)) for p in plans] == \
           [(p.file_name, p.table.header_row, len(p.table.rows)) for p in direct]


def test_big_sheet_without_dimension_is_stopped(tmp_path):
    # devlog 남은 한계: 크기 정보 없는 60MB 시트는 거부될 때까지 8.6초 -> 시간 제한으로 끊는다
    sc = load_scenario(ROOT / "scenarios" / "stock_count.yaml")
    p = make_no_dimension_sheet(tmp_path / "big.xlsx")
    t = time.perf_counter()
    with pytest.raises((TimeLimitError, ReadError)) as e:
        run_job("prepare", sc, [p], timeout=2)
    elapsed = time.perf_counter() - t
    assert elapsed < 6, elapsed
    assert e.type is TimeLimitError or "너무" in str(e.value)
