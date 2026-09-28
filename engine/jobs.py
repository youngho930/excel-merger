"""엔진 작업을 별도 프로세스에서 시간 제한을 두고 실행한다.

화면(app.py)은 파일 읽기(prepare)와 실행(execute + 결과 저장)을 모두 run_job()으로 돌린다.

왜 프로세스인가
- 크기 정보(<dimension>)가 없는 큰 시트는 openpyxl이 파일을 여는 단계에서 XML 전체를 훑는다
  (60MB 시트 하나에 8.6초). 이 동안에는 파이썬 코드가 끼어들 수 없어 스레드로는 멈출 수 없다.
- 그래서 작업을 `python -m engine.jobs` 하위 프로세스로 실행하고, 제한 시간을 넘기면
  프로세스를 강제로 끝낸다(subprocess.run의 timeout). Windows와 Linux에서 똑같이 동작한다.

주고받는 데이터
- 부모 -> 자식: pickle((작업 이름, 인자)). 작업 이름은 JOBS에 있는 것만 허용한다.
- 자식 -> 부모: pickle((상태, 값)). 자식은 이 저장소의 코드이므로 신뢰한다.
  업로드된 엑셀의 내용은 pickle의 "값"으로만 들어가며, pickle 명령을 만들 수 없다.
- 자식의 표준 오류(내부 경로·스택이 있을 수 있음)는 화면에 보여주지 않는다.
"""

from __future__ import annotations

import os
import pickle
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_TIMEOUT = 60   # 초. 파일 읽기, 실행 단계마다 따로 적용한다

OK, READ_ERROR, SCENARIO_ERROR, FAILED = "ok", "read_error", "scenario_error", "failed"


class TimeLimitError(Exception):
    """제한 시간을 넘겨 작업을 중단했을 때. 메시지는 한국어."""


class JobError(Exception):
    """예상하지 못한 실패. 메시지에 내부 정보를 넣지 않는다."""


# ------------------------------------------------------------------ 작업 (자식 프로세스에서 실행)
def _job_prepare(scenario, paths, header_rows=None):
    from .merge import prepare
    return prepare(scenario, paths, header_rows)


def _job_execute(scenario, plans, out_dir, apply_suggestions=False):
    from .merge import execute
    from .writer import write_result
    result = execute(scenario, plans)
    path = write_result(result, out_dir, apply_suggestions=apply_suggestions)
    return result, path


def _job_write(result, out_dir, apply_suggestions=False, excluded=()):
    from .writer import write_result
    return write_result(result, out_dir, apply_suggestions=apply_suggestions, excluded=excluded)


def _job_sleep(seconds):
    """시간 제한 확인용 (테스트)."""
    time.sleep(seconds)
    return seconds


def _job_env_names():
    """자식이 받은 환경변수 이름 (비밀값이 넘어가지 않는지 확인하는 테스트용). 값은 돌려주지 않는다."""
    return sorted(os.environ)


JOBS = {
    "prepare": _job_prepare,
    "execute": _job_execute,
    "write": _job_write,
    "sleep": _job_sleep,
    "env_names": _job_env_names,
}


# ------------------------------------------------------------------ 부모 쪽
MAX_CONCURRENT_JOBS = 3
_SLOTS = threading.BoundedSemaphore(MAX_CONCURRENT_JOBS)

# 자식 프로세스에 넘기지 않는 환경변수: 이름에 이 말이 들어가면 뺀다 (보안 검토 D-7).
# 자식은 신뢰할 수 없는 xlsx를 파싱하므로 API 키 같은 비밀값을 물려주지 않는다.
_SECRET_WORDS = ("KEY", "TOKEN", "SECRET", "PASSWORD", "PASSWD", "CREDENTIAL")


def child_env() -> dict[str, str]:
    env = {k: v for k, v in os.environ.items() if not any(w in k.upper() for w in _SECRET_WORDS)}
    env["PYTHONPATH"] = str(ROOT) + (os.pathsep + env["PYTHONPATH"] if env.get("PYTHONPATH") else "")
    return env


def run_job(name: str, *args: Any, timeout: float = DEFAULT_TIMEOUT) -> Any:
    """작업 하나를 하위 프로세스에서 실행하고 결과를 돌려준다.

    - 제한 시간을 넘기면 프로세스를 끝내고 TimeLimitError
    - 엔진이 낸 ReadError·ScenarioError는 같은 종류·같은 한국어 메시지로 다시 낸다
    - 그 밖의 실패는 JobError (내부 정보 없음)
    """
    from .reader import ReadError
    from .scenario import ScenarioError

    if name not in JOBS:
        raise ValueError(f"알 수 없는 작업: {name}")
    payload = pickle.dumps((name, args), protocol=pickle.HIGHEST_PROTOCOL)
    kwargs: dict[str, Any] = {}
    if os.name == "nt":
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    env = child_env()
    # 앱 전체에서 동시에 도는 작업 수를 제한한다. 기다리는 시간도 제한 시간에 포함한다 (보안 검토 D-8)
    started = time.monotonic()
    if not _SLOTS.acquire(timeout=timeout):
        raise TimeLimitError(
            f"지금 처리 중인 작업이 많아 {_seconds(timeout)}초 안에 시작하지 못했습니다. 잠시 후 다시 시도해 주세요.")
    try:
        left = max(timeout - (time.monotonic() - started), 0.1)
        proc = subprocess.run([sys.executable, "-m", "engine.jobs"], input=payload, capture_output=True,
                              timeout=left, cwd=str(ROOT), env=env, **kwargs)
    except subprocess.TimeoutExpired:
        # subprocess.run은 시간이 넘으면 자식 프로세스를 kill한 뒤 이 예외를 낸다
        raise TimeLimitError(
            f"처리 시간 제한({_seconds(timeout)}초)을 넘어 작업을 중단했습니다. "
            "파일 수를 줄이거나, 파일을 나눠서 다시 시도해 주세요.") from None
    except OSError:
        raise JobError("처리 프로그램을 시작하지 못했습니다. 잠시 후 다시 시도해 주세요.") from None
    finally:
        _SLOTS.release()

    try:
        status, value = pickle.loads(proc.stdout) if proc.stdout else (FAILED, None)
    except Exception:
        status, value = FAILED, None
    if proc.returncode == 0 and status == OK:
        return value
    if status == READ_ERROR:
        raise ReadError(value)
    if status == SCENARIO_ERROR:
        raise ScenarioError(value)
    if proc.returncode not in (0, 1) and status == FAILED:
        # 메모리 부족 등으로 운영체제가 프로세스를 끝낸 경우 (Linux: 음수 = 시그널)
        raise JobError("파일을 처리하는 중 작업이 강제로 끝났습니다. 파일이 너무 크거나 복잡할 수 있습니다. "
                       "파일 수를 줄이거나 파일을 나눠서 다시 시도해 주세요.")
    raise JobError("파일을 처리하는 중 예상하지 못한 문제가 생겼습니다. 파일이 손상되지 않았는지 확인해 주세요.")


def _seconds(t: float) -> str:
    return str(int(t)) if float(t).is_integer() else f"{t:g}"


# ------------------------------------------------------------------ 자식 쪽
def _child_main() -> int:
    from .reader import ReadError
    from .scenario import ScenarioError

    out = sys.stdout.buffer
    sys.stdout = sys.stderr   # 혹시 누가 print해도 결과(pickle)가 섞이지 않게
    try:
        name, args = pickle.loads(sys.stdin.buffer.read())
        reply = (OK, JOBS[name](*args))
        code = 0
    except ReadError as e:
        reply, code = (READ_ERROR, str(e)), 1
    except ScenarioError as e:
        reply, code = (SCENARIO_ERROR, str(e)), 1
    except Exception:
        import traceback
        traceback.print_exc()   # 표준 오류로만 (부모는 화면에 보여주지 않는다)
        reply, code = (FAILED, None), 1
    out.write(pickle.dumps(reply, protocol=pickle.HIGHEST_PROTOCOL))
    out.flush()
    return code


if __name__ == "__main__":
    raise SystemExit(_child_main())
