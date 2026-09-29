"""자유 양식 체험용 샘플을 만든다: samples/free_form/ + 정답지(samples/free_form_expected_errors.md/.json)

기존 샘플(scripts/make_samples.py, 정답지 37건)과 별개다. 그 파일들은 건드리지 않는다.

- "대리점·채널별 AS 접수 내역" 엑셀 3개. 열 이름·열 순서·날짜 표기가 파일마다 다르다.
  고객명/성명/고객 명, 연락처/전화번호, 제품명/모델, 증상/고장내용, 접수일/등록일.
  한 파일에만 있는 열: 담당기사(대리점), 주문번호(온라인몰).
- 콜센터 파일은 위에 제목 줄이 있어 머리글이 3행이다.
- 이름·전화번호는 모두 가상 값이다 (전화번호는 010-0000-xxxx).
- 정답지는 아래 CONFIG(묶기·결과 열·필수·중복기준)로 취합했을 때의 오류다.
  빈칸 5건 + 파일 간 중복 3쌍. 중복 2쌍은 날짜 표기가 달라(날짜 셀 / 2026.9.3 / 2026-09-08) 날짜로 맞춰야 잡힌다.

실행: python scripts/make_free_form_samples.py
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "samples"
OUT = SAMPLES / "free_form"
ANSWER_MD = SAMPLES / "free_form_expected_errors.md"
ANSWER_JSON = SAMPLES / "free_form_expected_errors.json"
FICTION_NOTE = "샘플의 이름·전화번호·주문번호는 모두 가상입니다. 실제 인물과 관계없습니다."
REQUIRED, DUPLICATE = "필수값 빈칸", "중복 행"

# 자유 양식에서 정할 설정: 결과 열(순서대로), 묶을 원본 이름, 필수, 중복 판단 기준
CONFIG = {
    "name": "대리점·채널별 AS 접수 내역",
    "columns": [
        {"name": "접수일", "sources": ["접수일", "등록일"], "required": True, "dup": True},
        {"name": "접수번호", "sources": ["접수번호"], "required": False, "dup": False},
        {"name": "고객명", "sources": ["고객명", "성명", "고객 명"], "required": True, "dup": False},
        {"name": "연락처", "sources": ["연락처", "전화번호"], "required": True, "dup": True},
        {"name": "제품명", "sources": ["제품명", "모델"], "required": True, "dup": True},
        {"name": "증상", "sources": ["증상", "고장내용"], "required": False, "dup": False},
        {"name": "담당기사", "sources": ["담당기사"], "required": False, "dup": False},
        {"name": "주문번호", "sources": ["주문번호"], "required": False, "dup": False},
    ],
}
DUP_LABEL = "접수일 + 연락처 + 제품명"
DUP_SHORT = "접수일+연락처+제품명"

# 논리 열 -> 파일마다 쓰인 이름 (순서가 곧 파일의 열 순서)
# 행: (접수번호, 9월 며칠(None=빈칸), 이름, 전화 끝 4자리(None=빈칸), 제품, 증상, 파일별 추가 값)
FILES = [
    {
        "file": "AS접수_강남대리점.xlsx", "title": [], "date": "cell",
        "columns": [("no", "접수번호"), ("date", "접수일"), ("name", "고객명"), ("phone", "연락처"),
                    ("product", "제품명"), ("symptom", "증상"), ("extra", "담당기사")],
        "rows": [
            ("AS-0001", 1, "김가온", "1001", "세탁기 WM-2400", "탈수 안 됨", "기사 박"),
            ("AS-0002", 1, "이나래", "1002", "냉장고 RF-8100", "냉동실 성에", "기사 최"),
            ("AS-0003", 2, "박다솜", "1003", "에어컨 AC-3300", "찬바람 약함", None),     # 담당기사 빈칸: 필수 아님
            ("AS-0004", 2, None, "1004", "TV QN-5500", "화면 깜빡임", "기사 박"),        # 고객명 빈칸
            ("AS-0005", 3, "최라온", "1005", "건조기 DR-1200", "소음 큼", "기사 최"),
            ("AS-0006", 3, "정마루", "1006", "식기세척기 DW-600", "물 샘", "기사 정"),
            ("AS-0007", 4, "강바다", "1007", "세탁기 WM-2400", "문 안 열림", "기사 박"),
            ("AS-0008", 4, "조사랑", None, "냉장고 RF-8100", "소음", "기사 최"),        # 연락처 빈칸
            ("AS-0009", 5, "윤아름", "1009", "에어컨 AC-3300", "리모컨 불량", "기사 정"),
            ("AS-0010", 7, "장자운", "1010", "TV QN-5500", "소리 안 남", None),
            ("AS-0011", 8, "임차린", "1011", "건조기 DR-1200", "건조 안 됨", "기사 박"),
            ("AS-0012", 9, "한카라", "1012", "세탁기 WM-2400", "진동 심함", "기사 최"),
        ],
    },
    {
        "file": "AS접수_콜센터.xlsx", "title": ["2026년 9월 AS 접수 현황", "작성: 콜센터 운영팀 (가상 데이터)"],
        "date": "dotted",
        "columns": [("date", "등록일"), ("name", "성명"), ("phone", "전화번호"), ("product", "모델"),
                    ("symptom", "고장내용"), ("no", "접수번호")],
        "rows": [
            ("CC-2001", 1, "오타미", "2001", "냉장고 RF-8100", "문 고무패킹 찢어짐", None),
            ("CC-2002", 2, "서파랑", "2002", "세탁기 WM-2400", "급수 안 됨", None),
            ("AS-0003", 2, "신하늘", "2003", "TV QN-5500", "전원 안 켜짐", None),     # 접수번호가 대리점과 같음: 오류 아님
            ("CC-2004", 3, "최라온", "1005", "건조기 DR-1200", "소음이 큼", None),     # 대리점 6행과 중복
            ("CC-2005", 3, None, "2005", "에어컨 AC-3300", "물 떨어짐", None),         # 성명 빈칸
            ("CC-2006", 4, "권가람", "2006", "식기세척기 DW-600", "세척력 약함", None),
            ("CC-2007", 4, "강바다", "1007", "냉장고 RF-8100", "냉기 약함", None),     # 같은 사람·같은 날, 제품이 달라 중복 아님
            ("CC-2008", 5, "황나린", "2008", None, "원인 불명", None),                 # 모델 빈칸
            ("CC-2009", 6, "안다온", "2009", "건조기 DR-1200", "필터 경고", None),
            ("CC-2010", 8, "송라엘", "2010", "세탁기 WM-2400", "탈수 소음", None),
            ("CC-2011", 9, "류마음", "2011", "TV QN-5500", "리모컨 안 됨", None),
            ("CC-2012", 10, "홍바름", "2012", "에어컨 AC-3300", "실외기 소음", None),
        ],
    },
    {
        "file": "AS접수_온라인몰.xlsx", "title": [], "date": "dash",
        "columns": [("no", "접수번호"), ("extra", "주문번호"), ("name", "고객 명"), ("phone", "연락처"),
                    ("product", "제품명"), ("symptom", "증상"), ("date", "접수일")],
        "rows": [
            ("ON-3001", 1, "문가을", "3001", "냉장고 RF-8100", "얼음 안 나옴", "ORD-2609001"),
            ("ON-3002", 2, "배나무", "3002", "세탁기 WM-2400", "세제 투입구 막힘", "ORD-2609002"),
            ("ON-3003", 5, "윤아름", "1009", "에어컨 AC-3300", "리모컨 작동 안 함", "ORD-2609003"),  # 대리점 10행과 중복
            ("ON-3004", 6, "노을빛", "3004", "TV QN-5500", "화면 줄 생김", "ORD-2609004"),
            ("ON-3005", None, "도하람", "3005", "건조기 DR-1200", "문 잠김", "ORD-2609005"),    # 접수일 빈칸
            ("ON-3006", 8, "송라엘", "2010", "세탁기 WM-2400", "탈수할 때 소음", "ORD-2609006"),  # 콜센터 13행과 중복
            ("ON-3007", 9, "마서준", "3007", "식기세척기 DW-600", "건조 불량", "ORD-2609007"),
            ("ON-3008", 10, "변하윤", "3008", "냉장고 RF-8100", "소음", "ORD-2609008"),
            ("ON-3009", 11, "석지우", "3009", "에어컨 AC-3300", "냄새 남", "ORD-2609009"),
            ("ON-3010", 12, "옥예린", "3010", "세탁기 WM-2400", "물 샘", "ORD-2609010"),
        ],
    },
]
DATE_NOTES = {"cell": "엑셀 날짜 셀", "dotted": "문자 `YYYY.M.D`", "dash": "문자 `YYYY-MM-DD`"}
KEYS = ("no", "date", "name", "phone", "product", "symptom", "extra")


def date_value(day, style):
    if day is None:
        return None
    if style == "cell":
        return datetime(2026, 9, day)
    if style == "dotted":
        return f"2026.9.{day}"
    return f"2026-09-{day:02d}"


def cell_values(f, row):
    data = dict(zip(KEYS, row))
    data["date"] = date_value(data["date"], f["date"])
    data["phone"] = None if data["phone"] is None else f"010-0000-{data['phone']}"
    return [data[k] for k, _ in f["columns"]]


def header_row(f):
    return len(f["title"]) + 1 if f["title"] else 1


def excel_row(f, i):
    return header_row(f) + 1 + i


def result_name(source_name):
    return next(c["name"] for c in CONFIG["columns"] if source_name in c["sources"])


def save(f):
    wb = Workbook()
    ws = wb.active
    ws.title = "AS접수"
    for t in f["title"]:
        ws.append([t])
    if f["title"]:
        ws["A1"].font = Font(bold=True, size=14)
    ws.append([name for _, name in f["columns"]])
    for c in ws[header_row(f)]:
        c.font = Font(bold=True)
    for row in f["rows"]:
        ws.append(cell_values(f, row))
    for i, (key, _) in enumerate(f["columns"], start=1):
        ws.column_dimensions[get_column_letter(i)].width = {"symptom": 20, "product": 18, "phone": 15}.get(key, 12)
        if key == "date" and f["date"] == "cell":
            for r in range(header_row(f) + 1, header_row(f) + 1 + len(f["rows"])):
                ws.cell(r, i).number_format = "yyyy-mm-dd"
    OUT.mkdir(parents=True, exist_ok=True)
    wb.save(OUT / f["file"])


def blank_errors():
    errors = []
    for f in FILES:
        required = {c["name"] for c in CONFIG["columns"] if c["required"]}
        for i, row in enumerate(f["rows"]):
            values = cell_values(f, row)
            for (key, col), v in zip(f["columns"], values):
                std = result_name(col)
                if v is None and std in required:
                    errors.append({"file": f["file"], "row": excel_row(f, i), "column": col, "standard": std,
                                   "kind": REQUIRED, "value": None, "shown": "(빈칸)",
                                   "description": f"필수 열 '{std}'이(가) 비어 있음", "duplicate_of": None})
    return errors


def dup_errors():
    """중복기준(접수일·연락처·제품명)이 같은 행. 뒤에 나온 행을 오류로 적고, 먼저 나온 행을 duplicate_of 로."""
    seen: dict[tuple, tuple[str, int]] = {}
    errors = []
    for f in FILES:
        for i, row in enumerate(f["rows"]):
            data = dict(zip(KEYS, row))
            key = (data["date"], data["phone"], data["product"])
            if None in key:
                continue
            here = (f["file"], excel_row(f, i))
            if key in seen:
                first = seen[key]
                shown = f"접수일=2026-09-{key[0]:02d}, 연락처=010-0000-{key[1]}, 제품명={key[2]}"
                errors.append({"file": here[0], "row": here[1], "column": "(행 전체)", "standard": DUP_LABEL,
                               "kind": DUPLICATE, "value": None, "shown": shown,
                               "description": f"{first[0]} {first[1]}행과 중복기준 값이 같음",
                               "duplicate_of": {"file": first[0], "row": first[1]}})
            else:
                seen[key] = here
    return errors


def main():
    for f in FILES:
        save(f)
        print(f"  free_form/{f['file']}: 머리글 {header_row(f)}행, 데이터 {len(f['rows'])}행")
    order = {f["file"]: i for i, f in enumerate(FILES)}
    errors = sorted(blank_errors() + dup_errors(), key=lambda e: (order[e["file"]], e["row"]))
    data = {
        "version": 1, "note": FICTION_NOTE, "folder": "free_form", "config": CONFIG,
        "files": [{"file": f["file"], "header_row": header_row(f), "data_rows": len(f["rows"]),
                   "columns": [name for _, name in f["columns"]]} for f in FILES],
        "errors": errors,
    }
    ANSWER_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    lines = [
        "# 자유 양식 정답지: AS 접수 샘플에 심은 오류 목록", "",
        "> 이 파일은 `scripts/make_free_form_samples.py`가 샘플과 함께 자동으로 만듭니다. 직접 고치지 말고 스크립트를 고친 뒤 다시 실행하세요.", "",
        f"> {FICTION_NOTE}", "",
        "기존 3개 시나리오의 정답지(`samples/expected_errors.md`, 37건)와는 별개입니다.", "",
        "## 자유 양식 설정", "",
        "아래처럼 묶고 골랐을 때의 정답입니다. 필수값 빈칸과 중복 행만 검사합니다.", "",
        "| 순서 | 결과 열 이름 | 묶을 원본 열 이름 | 필수 | 중복 판단 기준 |", "|---|---|---|---|---|",
    ]
    for i, c in enumerate(CONFIG["columns"], start=1):
        lines.append(f"| {i} | {c['name']} | {', '.join(c['sources'])} | {'예' if c['required'] or c['dup'] else ''} "
                     f"| {'예' if c['dup'] else ''} |")
    lines += ["", "- 정규화(공백·대소문자·기호 무시)만으로 자동으로 묶이는 것: 고객명 = 고객 명, 연락처, 제품명, 증상, 접수일, 접수번호",
              "- 직접(또는 AI 추천으로) 묶어야 하는 것: 성명 → 고객명, 전화번호 → 연락처, 모델 → 제품명, 고장내용 → 증상, 등록일 → 접수일",
              "- 접수일 열은 빈칸을 뺀 값이 모두 날짜로 읽혀 **날짜로 비교**합니다. 그래서 표기가 달라도 같은 날이면 중복입니다.",
              "", "## 파일별 요약", "",
              "| 파일 | 머리글 행 | 데이터 행 수 | 날짜 표기 | 열 (순서대로) |", "|---|---|---|---|---|"]
    for f in FILES:
        lines.append(f"| `free_form/{f['file']}` | {header_row(f)}행 | {len(f['rows'])} | {DATE_NOTES[f['date']]} "
                     f"| {', '.join(n for _, n in f['columns'])} |")
    lines += ["", f"## 오류 목록 (총 **{len(errors)}건**)", "",
              "- **중복 행**은 뒤에 나온 행을 적고, 설명에 먼저 나온 행을 적었습니다. 엔진이 두 행 모두 표시해도 정답으로 봅니다.", "",
              "| 파일 | 행 | 열 이름 (결과 열) | 오류 종류 | 값 | 설명 |", "|---|---|---|---|---|---|"]
    for e in errors:
        col = f"(행 전체) — 중복기준: {e['standard']}" if e["kind"] == DUPLICATE else f"{e['column']} ({e['standard']})"
        lines.append(f"| {e['file']} | {e['row']} | {col} | {e['kind']} | {e['shown']} | {e['description']} |")
    lines += ["", "## 오류가 아닌 것 (함정)", "",
              "- 콜센터 6행 `AS-0003`: 접수번호가 대리점 4행과 같지만, 접수번호는 중복기준이 아닙니다 (채널마다 번호를 따로 매김).",
              "- 콜센터 10행: 대리점 8행과 같은 사람·같은 날이지만 제품이 달라 중복이 아닙니다.",
              "- 대리점 4행·11행의 담당기사 빈칸: 담당기사는 필수가 아닙니다. 다른 파일에는 담당기사 열이 없어 빈칸이지만 오류가 아닙니다.",
              "- 파일마다 다른 열 이름·열 순서·날짜 표기, 콜센터 파일 위쪽의 제목 줄."]
    ANSWER_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"정답지 작성: {ANSWER_MD.name}, {ANSWER_JSON.name} (총 {len(errors)}건)")


if __name__ == "__main__":
    main()
