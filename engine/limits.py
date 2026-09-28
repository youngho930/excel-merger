"""업로드·처리 제한값. 배포 환경마다 바꿀 수 있게 설정에서 읽는다.

읽는 순서: secrets(st.secrets) -> 환경변수 -> 프로필 기본값.
- EXCEL_MERGER_PROFILE = "cloud" 이면 메모리 약 1GB인 Streamlit Community Cloud 무료 서버에 맞춘
  보수적인 기본값을 쓴다. 비어 있거나 "local" 이면 로컬 기본값.
- 항목마다 따로 덮어쓸 수 있다 (예: MAX_FILES = 5).
- 설정으로 엔진의 안전 상한(reader·merge 의 MAX_*)보다 크게 할 수는 없다. 큰 값은 상한으로 줄인다.

엔진은 streamlit 을 모른다. 화면(app.py)이 st.secrets 를 dict 로 넘긴다.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, fields, replace
from typing import Any, Mapping

from .merge import MAX_FILES as ENGINE_MAX_FILES
from .reader import MAX_FILE_BYTES as ENGINE_MAX_FILE_BYTES
from .reader import MAX_ROWS as ENGINE_MAX_ROWS

PROFILE_KEY = "EXCEL_MERGER_PROFILE"
MB = 1024 * 1024


@dataclass(frozen=True)
class Limits:
    profile: str
    max_files: int              # 한 번에 올리는 파일 수
    max_file_mb: int            # 파일당 크기
    max_total_upload_mb: int    # 올리는 파일 크기 합계
    max_rows_per_file: int      # 파일당 데이터 행 수
    max_total_rows: int         # 모든 파일의 데이터 행 합계
    time_limit_seconds: int     # 파일 읽기·실행 단계마다 처리 시간
    max_concurrent_jobs: int    # 앱 전체에서 동시에 도는 처리 작업 수
    job_memory_mb: int          # 처리 작업 하나의 메모리 상한 (Linux만, 0이면 제한 없음)

    @property
    def max_file_bytes(self) -> int:
        return self.max_file_mb * MB

    @property
    def max_total_upload_bytes(self) -> int:
        return self.max_total_upload_mb * MB


LOCAL = Limits(profile="local", max_files=50, max_file_mb=20, max_total_upload_mb=100,
               max_rows_per_file=50_000, max_total_rows=200_000, time_limit_seconds=60,
               max_concurrent_jobs=3, job_memory_mb=0)

# Streamlit Community Cloud 무료 서버(메모리 약 1GB)용
CLOUD = Limits(profile="cloud", max_files=10, max_file_mb=5, max_total_upload_mb=20,
               max_rows_per_file=10_000, max_total_rows=30_000, time_limit_seconds=45,
               max_concurrent_jobs=1, job_memory_mb=400)

PROFILES = {"local": LOCAL, "cloud": CLOUD}

# 설정 이름 -> (필드, 최솟값, 최댓값). 최댓값은 엔진의 안전 상한을 넘지 않는다.
SETTINGS: dict[str, tuple[str, int, int]] = {
    "MAX_FILES": ("max_files", 1, ENGINE_MAX_FILES),
    "MAX_FILE_MB": ("max_file_mb", 1, ENGINE_MAX_FILE_BYTES // MB),
    "MAX_TOTAL_UPLOAD_MB": ("max_total_upload_mb", 1, 1000),
    "MAX_ROWS_PER_FILE": ("max_rows_per_file", 1, ENGINE_MAX_ROWS),
    "MAX_TOTAL_ROWS": ("max_total_rows", 1, 10_000_000),
    "TIME_LIMIT_SECONDS": ("time_limit_seconds", 5, 600),
    "MAX_CONCURRENT_JOBS": ("max_concurrent_jobs", 1, 16),
    "JOB_MEMORY_MB": ("job_memory_mb", 0, 64_000),
}
SETTING_NAMES = (PROFILE_KEY, *SETTINGS)

assert {f for f, _, _ in SETTINGS.values()} == {f.name for f in fields(Limits)} - {"profile"}


def _pick(name: str, secrets: Mapping[str, Any] | None, environ: Mapping[str, str]) -> Any:
    if secrets is not None:
        try:
            v = secrets.get(name)
        except Exception:
            v = None
        if v is not None and str(v).strip() != "":
            return v
    v = environ.get(name, "")
    return v if str(v).strip() != "" else None


def load_limits(secrets: Mapping[str, Any] | None = None,
                environ: Mapping[str, str] | None = None) -> tuple[Limits, list[str]]:
    """(제한값, 설정 안내). 잘못된 설정은 무시하고 기본값을 쓰며 안내에 적는다 (앱은 멈추지 않는다)."""
    environ = os.environ if environ is None else environ
    notes: list[str] = []

    raw_profile = _pick(PROFILE_KEY, secrets, environ)
    profile = str(raw_profile).strip().lower() if raw_profile is not None else "local"
    if profile not in PROFILES:
        notes.append(f"{PROFILE_KEY} 값 '{raw_profile}'을 알 수 없어 local 기본값을 씁니다 (local 또는 cloud).")
        profile = "local"
    limits = PROFILES[profile]

    changes: dict[str, int] = {}
    for name, (field, lo, hi) in SETTINGS.items():
        raw = _pick(name, secrets, environ)
        if raw is None:
            continue
        try:
            if isinstance(raw, bool):
                raise ValueError
            value = int(str(raw).strip())
        except ValueError:
            notes.append(f"{name} 값 '{raw}'이 정수가 아니어서 기본값 {getattr(limits, field)}을 씁니다.")
            continue
        if not lo <= value <= hi:
            clamped = min(max(value, lo), hi)
            notes.append(f"{name} 값 {value}은 허용 범위({lo}~{hi})를 벗어나 {clamped}로 맞췄습니다.")
            value = clamped
        changes[field] = value
    if changes:
        limits = replace(limits, **changes)
    return limits, notes
