"""배포 환경별 제한값 (engine/limits.py) 과 그 적용.

- 로컬 기본값은 예전 값 그대로, EXCEL_MERGER_PROFILE=cloud 면 1GB 서버용 보수적인 값
- secrets -> 환경변수 -> 프로필 기본값 순서, 엔진 안전 상한보다 크게 할 수 없음
- 작업별 메모리 상한은 Linux 에서만 동작 (GitHub Actions 의 Ubuntu 에서 실제로 확인)
"""

import sys

import pytest
from openpyxl import Workbook
from streamlit.testing.v1 import AppTest

from answer_key import ROOT
from engine import ReadError, load_scenario, read_table
from engine import jobs
from engine.jobs import JobError, run_job
from engine.limits import CLOUD, LOCAL, PROFILE_KEY, load_limits

APP = str(ROOT / "app.py")
LINUX = sys.platform.startswith("linux")


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch):
    for name in (PROFILE_KEY, "MAX_FILES", "MAX_FILE_MB", "MAX_TOTAL_UPLOAD_MB", "MAX_ROWS_PER_FILE",
                 "MAX_TOTAL_ROWS", "TIME_LIMIT_SECONDS", "MAX_CONCURRENT_JOBS", "JOB_MEMORY_MB"):
        monkeypatch.delenv(name, raising=False)
    yield
    jobs.set_max_concurrent(jobs.MAX_CONCURRENT_JOBS)   # 앱이 바꾼 동시 작업 수를 되돌린다


# ------------------------------------------------------------------ 설정 읽기
def test_local_defaults_are_unchanged():
    limits, notes = load_limits({}, {})
    assert limits == LOCAL and notes == []
    assert (limits.max_files, limits.max_file_mb, limits.max_total_upload_mb, limits.max_rows_per_file,
            limits.max_total_rows, limits.time_limit_seconds, limits.max_concurrent_jobs,
            limits.job_memory_mb) == (50, 20, 100, 50_000, 200_000, 60, 3, 0)


def test_cloud_profile_values():
    limits, notes = load_limits({PROFILE_KEY: "cloud"}, {})
    assert limits == CLOUD and notes == []
    assert (limits.max_files, limits.max_file_mb, limits.max_total_upload_mb, limits.max_rows_per_file,
            limits.max_total_rows, limits.time_limit_seconds, limits.max_concurrent_jobs,
            limits.job_memory_mb) == (10, 5, 20, 10_000, 30_000, 45, 1, 400)


def test_secrets_win_over_env_and_single_values_override_profile():
    limits, _ = load_limits({PROFILE_KEY: "cloud", "MAX_FILES": 3},
                            {PROFILE_KEY: "local", "MAX_FILES": "7", "MAX_TOTAL_ROWS": "1234"})
    assert limits.profile == "cloud"
    assert limits.max_files == 3               # secrets 가 환경변수보다 먼저
    assert limits.max_total_rows == 1234       # secrets 에 없으면 환경변수
    assert limits.max_file_mb == CLOUD.max_file_mb


def test_cannot_exceed_engine_safety_ceilings():
    limits, notes = load_limits({"MAX_FILE_MB": 500, "MAX_ROWS_PER_FILE": 10**9, "MAX_FILES": 0}, {})
    assert limits.max_file_mb == 20 and limits.max_rows_per_file == 50_000 and limits.max_files == 1
    assert len(notes) == 3 and all("허용 범위" in n for n in notes)


def test_bad_values_fall_back_with_notes():
    limits, notes = load_limits({PROFILE_KEY: "heroku", "MAX_FILES": "열 개", "JOB_MEMORY_MB": True}, {})
    assert limits == LOCAL
    assert any("local 기본값" in n for n in notes) and sum("정수가 아니어서" in n for n in notes) == 2


# ------------------------------------------------------------------ 엔진에 적용
def test_rows_per_file_limit_reaches_reader(tmp_path):
    sc = load_scenario(ROOT / "scenarios" / "stock_count.yaml")
    wb = Workbook()
    ws = wb.active
    ws.append(["창고", "품목코드", "품목명", "단위", "전산수량", "실사수량", "실사자", "실사일"])
    for i in range(30):
        ws.append(["A", f"P{i}", "n", "EA", 1, 1, "김", "2026-09-01"])
    wb.save(tmp_path / "rows.xlsx")
    assert len(read_table(tmp_path / "rows.xlsx", sc).rows) == 30
    with pytest.raises(ReadError, match="최대 10행"):
        read_table(tmp_path / "rows.xlsx", sc, max_rows=10)
    # 엔진 상한보다 큰 값은 상한으로
    assert len(read_table(tmp_path / "rows.xlsx", sc, max_rows=10**9).rows) == 30


def test_prepare_job_passes_row_limit(tmp_path):
    sc = load_scenario(ROOT / "scenarios" / "stock_count.yaml")
    files = sorted((ROOT / "samples" / "stock_count").glob("*.xlsx"))
    with pytest.raises(ReadError, match="최대 5행"):
        run_job("prepare", sc, files, None, 5, timeout=60)


def test_set_max_concurrent():
    jobs.set_max_concurrent(1)
    assert jobs._SLOTS_SIZE == 1
    assert run_job("sleep", 0, timeout=30) == 0      # 자리 1개로도 정상 동작, 자리 반납
    assert jobs._SLOTS.acquire(timeout=0.1)
    jobs._SLOTS.release()


# ------------------------------------------------------------------ 메모리 상한 (Linux)
@pytest.mark.skipif(not LINUX, reason="작업별 메모리 상한(RLIMIT_DATA)은 Linux에서만 동작")
def test_memory_limit_stops_only_that_job():
    with pytest.raises(JobError, match="메모리가 부족합니다"):
        run_job("alloc", 800, timeout=60, memory_mb=200)
    assert run_job("alloc", 50, timeout=60, memory_mb=200) == 50


@pytest.mark.skipif(not LINUX, reason="작업별 메모리 상한(RLIMIT_DATA)은 Linux에서만 동작")
def test_cloud_memory_limit_is_enough_for_samples(tmp_path):
    # 클라우드 기본값(400MB)으로 샘플 전체를 읽고 실행·저장할 수 있어야 한다
    for key in ("incoming_inspection", "stock_count", "monthly_report"):
        sc = load_scenario(ROOT / "scenarios" / f"{key}.yaml")
        files = sorted((ROOT / "samples" / key).glob("*.xlsx"))
        plans = run_job("prepare", sc, files, None, CLOUD.max_rows_per_file,
                        timeout=60, memory_mb=CLOUD.job_memory_mb)
        result, path = run_job("execute", sc, plans, tmp_path / key, False,
                               timeout=60, memory_mb=CLOUD.job_memory_mb)
        assert path.exists() and len(result.issues) > 0


def test_memory_limit_is_ignored_off_linux():
    if LINUX:
        pytest.skip("Linux 에서는 위 테스트가 실제 상한을 확인한다")
    assert run_job("alloc", 50, timeout=60, memory_mb=10) == 50


# ------------------------------------------------------------------ 화면
def test_app_uses_cloud_limits(monkeypatch):
    monkeypatch.setenv(PROFILE_KEY, "cloud")
    at = AppTest.from_file(APP, default_timeout=120)
    at.run()
    assert not at.exception, at.exception
    caption = " ".join(c.value for c in at.caption)
    assert "한 번에 최대 10개" in caption and "파일당 5MB" in caption and "합계 20MB" in caption
    assert "파일당 10,000행" in caption and "합계 30,000행" in caption and "45초" in caption
    assert jobs._SLOTS_SIZE == 1
    # 샘플은 클라우드 제한 안이라 그대로 체험할 수 있다
    at.button(key="sample_btn").click().run()
    assert not at.exception and not at.error


def test_app_local_limits_unchanged():
    at = AppTest.from_file(APP, default_timeout=120)
    at.run()
    caption = " ".join(c.value for c in at.caption)
    assert "한 번에 최대 50개" in caption and "파일당 20MB" in caption and "60초" in caption


def test_app_shows_bad_setting_note(monkeypatch):
    monkeypatch.setenv("MAX_FILES", "많이")
    at = AppTest.from_file(APP, default_timeout=120)
    at.run()
    assert "설정 안내" in " ".join(c.value for c in at.caption)
