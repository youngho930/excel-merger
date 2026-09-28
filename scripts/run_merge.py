"""시나리오 이름 + 엑셀 폴더 -> 취합 결과 엑셀.

실행 예:
    python scripts/run_merge.py stock_count samples/stock_count
    python scripts/run_merge.py 경로/새시나리오.yaml 폴더 --out output

화면 없이 실행하므로 실행 전 경고는 출력만 하고 그대로 진행한다.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from engine import (ReadError, ScenarioError, execute, find_scenario, load_scenario,  # noqa: E402
                    pre_run_warnings, prepare, write_result)
from engine.reader import EXCEL_SUFFIXES  # noqa: E402
from engine.validate import KINDS  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def excel_files(folder: Path) -> list[Path]:
    """폴더 안 엑셀 파일 (엑셀 임시파일 '~$…' 제외), 파일명 순."""
    return sorted(p for p in folder.iterdir()
                  if p.is_file() and p.suffix.lower() in EXCEL_SUFFIXES and not p.name.startswith("~$"))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="엑셀 파일들을 시나리오대로 취합하고 오류를 찾습니다.")
    ap.add_argument("scenario", help="시나리오 이름(scenarios/ 안의 파일 이름) 또는 .yaml 경로")
    ap.add_argument("folder", help="취합할 엑셀 파일이 있는 폴더")
    ap.add_argument("--out", default=str(ROOT / "output"), help="결과를 저장할 폴더 (기본: output)")
    args = ap.parse_args(argv)

    try:
        if args.scenario.lower().endswith((".yaml", ".yml")):
            scenario = load_scenario(args.scenario)
        else:
            scenario = find_scenario(args.scenario, ROOT / "scenarios")
        folder = Path(args.folder)
        if not folder.is_dir():
            print(f"오류: 폴더 '{folder}'를 찾을 수 없습니다.", file=sys.stderr)
            return 2
        files = excel_files(folder)
        if not files:
            print(f"오류: '{folder}'에 엑셀(.xlsx) 파일이 없습니다.", file=sys.stderr)
            return 2
        plans = prepare(scenario, files)
    except (ScenarioError, ReadError) as e:
        print(f"오류: {e}", file=sys.stderr)
        return 2

    print(f"시나리오: {scenario.name} ({scenario.key})")
    print(f"파일 {len(plans)}개\n")
    for plan in plans:
        t, m = plan.table, plan.match
        print(f"■ {t.file_name}  (시트 '{t.sheet_name}', 머리글 {t.header_row}행[{t.header_confidence}], "
              f"데이터 {len(t.rows)}행)")
        for cm in m.matches:
            src = m.column_label(cm.source_index) if cm.source_index is not None else "-"
            req = "*" if cm.required else " "
            print(f"   {req}{cm.standard:<8} ↔ {src:<20} ↔ {cm.method}")
        unmatched = [m.column_label(i) for i in m.unmatched_sources]
        print(f"   매칭 안 된 원본 열: {', '.join(unmatched) if unmatched else '없음'}")
    print("   (* 필수 열)\n")

    warnings = pre_run_warnings(plans)
    if warnings:
        print("⚠ 실행 전 경고")
        for w in warnings:
            print(f"   - {w}")
        print("   경고가 있지만 그대로 진행합니다.\n")

    result = execute(scenario, plans)
    out = write_result(result, args.out)

    counts = result.counts()
    print(f"취합 행 수: {len(result.rows)}")
    print(f"오류 {len(result.issues)}건: " + ", ".join(f"{k} {counts[k]}" for k in KINDS))
    print(f"결과 파일: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
