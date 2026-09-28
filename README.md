# 엑셀 자동 취합·검증기

[![tests](https://github.com/youngho930/excel-merger/actions/workflows/tests.yml/badge.svg)](https://github.com/youngho930/excel-merger/actions/workflows/tests.yml)

**라이브 데모: 배포 예정**

부서·협력사·지점마다 양식이 제각각인 엑셀 파일을 하나로 모으고, 오류를 자동으로 찾아 표시하는 도구입니다.

> 샘플의 회사명·인명은 모두 가상입니다. 실제 회사나 인물과 관계없습니다.

## 주요 기능

- **양식이 달라도 취합**: 열 이름·열 순서·날짜 표기가 파일마다 달라도 기준열로 맞춰 모읍니다. 제목 줄 때문에 머리글이 1행이 아닌 파일도 머리글 행을 자동으로 찾습니다.
- **열 매칭 확인표**: 정확히 일치 → 동의어 사전 → (선택) AI 추천 순서로 짝을 찾고, 결과를 바로 적용하지 않고 화면에서 확인·수정한 뒤 실행합니다.
- **검증 5종**: 필수값 빈칸 / 형식 오류(숫자 칸의 문자, 날짜) / 범위 밖 값 / 허용값 아닌 값 / 중복 행.
- **결과 엑셀**: 원본 파일은 건드리지 않고 새 파일을 만듭니다.
  - `취합결과`: 출처 파일·원래 행 번호가 붙은 취합 데이터. 오류 셀은 종류별 색
  - `오류목록`: 파일, 행, 열, 오류 종류, 값, 설명, 수정 제안값, 중복 그룹, 처리
  - `제외된 행`: 중복 그룹에서 뺀 행의 원래 값 (뺀 행이 있을 때만)
  - `요약`: 오류 종류별 전체 / 처리됨 / 남은 오류
  - `범례`: 색의 뜻
- **사용자 처리**: 수정 제안값 일괄 적용(기본 꺼짐), 중복 그룹마다 남길 행 고르기. 처리한 오류는 `처리` 열에 "자동 수정됨" / "행 제외됨" / "중복 해소(이 행을 남김)"로 남고, 화면은 "남은 오류"를 가장 크게 보여줍니다.
- **바로 체험**: 시나리오 3종(협력사 수입검사, 지점·창고 재고실사, 부서별 월간 실적) 샘플과 AI 매칭 체험 파일을 버튼 한 번으로 불러옵니다.

## 실행 방법

Python 3.11 이상이 필요합니다.

```bash
git clone https://github.com/youngho930/excel-merger.git
cd excel-merger
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
```

설정 파일을 예시에서 복사합니다. **값을 비워 둬도 됩니다** (AI 키가 없으면 동의어 매칭만 쓰고, 제한값은 로컬 기본값).

```bash
# Windows (PowerShell)
Copy-Item .streamlit\secrets.toml.example .streamlit\secrets.toml
# macOS / Linux
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

`.streamlit/secrets.toml` 은 `.gitignore` 에 들어 있어 커밋되지 않습니다. 키를 넣었다면 이 파일을 다른 곳에 올리지 마세요.

```bash
streamlit run app.py
```

브라우저에서 http://localhost:8501 이 열립니다.

화면 없이 명령줄로도 실행할 수 있습니다 (실행 전 경고는 출력만 하고 그대로 진행).

```bash
python scripts/run_merge.py stock_count samples/stock_count
python scripts/run_merge.py 경로/새시나리오.yaml 엑셀폴더 --out output
```

## 테스트

```bash
pip install -r requirements.txt -r requirements-dev.txt
pytest
```

- `samples/expected_errors.md`(정답지, 37건)와 엔진 결과를 대조하는 테스트가 포함되어 있습니다.
- 화면은 `streamlit.testing.v1.AppTest` 로 시험합니다. AI 매칭 테스트는 가짜 응답을 쓰므로 실제 API를 부르지 않습니다.
- 샘플을 다시 만들려면 `python scripts/make_samples.py`(시나리오 샘플과 정답지), `python scripts/make_ai_demo.py`(AI 매칭 체험 파일).

## Gemini 키 설정 (선택)

**키가 없어도 동작합니다.** 키가 없으면 화면에 "AI 없이 동의어 매칭만 사용 중"이라고 안내하고, 동의어 사전과 드롭다운 수정만으로 매칭합니다.

AI 추천을 쓰려면 `.streamlit/secrets.toml.example` 을 `.streamlit/secrets.toml` 로 복사하고 키를 넣습니다 (이 파일은 `.gitignore` 에 들어 있어 커밋되지 않습니다). 환경변수 `GEMINI_API_KEY` 로도 됩니다.

```toml
GEMINI_API_KEY = "여기에 키"
# GEMINI_MODEL = "gemini-3.1-flash-lite"   # 비우면 기본 모델
```

Streamlit Community Cloud 에서는 앱 설정의 Secrets 칸에 같은 형식으로 넣습니다.

AI에게 보내는 내용은 기본으로 **매칭 안 된 기준열 이름과 원본 열 이름뿐**이고(파일 이름은 F1, F2 번호로 바꿈), "예시값 함께 보내기"를 켰을 때만 열마다 예시값 3개를 더 보냅니다. 화면의 "전송될 내용 미리 보기"에서 실제로 보낼 내용을 확인할 수 있습니다. AI 응답은 검증을 통과한 추천만 "AI 추천"으로 확인표에 채우며, 사용자가 확인·수정합니다.

## 시나리오 추가 방법

엔진 코드는 시나리오 이름이나 열 이름을 모릅니다. `scenarios/` 에 YAML 파일 하나를 추가하면 화면과 명령줄에 바로 나타납니다.

```yaml
name: 설비 일일점검                 # 화면에 보이는 이름
description: 공장별 설비 점검표를 모읍니다.
columns:
  - 기준명: 설비ID
    동의어: [장비번호, Equip ID]
    필수: true
    형식: 문자                      # 문자 | 정수 | 실수 | 날짜
  - 기준명: 온도
    동의어: [온도(℃), Temp]
    필수: true
    형식: 실수
    범위: [-20, 80]                 # 숫자 열만, 선택
  - 기준명: 상태
    필수: true
    형식: 문자
    허용값: [정상, 점검필요, 고장]   # 선택
중복기준: [설비ID]                   # 중복 판단에 쓸 열 (없으면 [])
```

- 그 시나리오의 샘플을 `samples/<yaml 파일 이름>/` 에 넣으면 "샘플 파일로 바로 체험" 버튼이 생깁니다.
- 실제 사례: 기준열 이름을 ERP 용어(`품번` → `품목코드`, `품명` → `품목명`)로 바꿀 때 YAML만 고쳤고, **엔진 코드 변경은 0줄**이었습니다 (`docs/devlog.md`).

## 설계 원칙

1. 취합 엔진은 하나, 시나리오는 설정 파일(YAML). 엔진에 특정 시나리오·열 이름을 하드코딩하지 않는다.
2. 열 매칭은 정확히 일치 → 동의어 → AI 추천 순서. AI 없이도 기본 동작한다.
3. 매칭 결과는 바로 적용하지 않고 사용자가 확인·수정한 뒤 적용한다.
4. 원본 파일은 절대 수정하지 않는다. 결과는 새 파일로만 만든다.
5. 머리글이 1행이 아닌 파일도 처리한다.
6. 화면과 오류 메시지는 비개발자도 이해할 수 있는 한국어로 쓴다.

폴더 구조: `app.py`(화면) / `engine/`(매칭·취합·검증·결과 저장·하위 프로세스 실행·AI 매칭) / `scenarios/` / `samples/` / `scripts/` / `tests/` / `docs/devlog.md`(개발 일지).

## 제한 사항

업로드·처리 제한은 `.streamlit/secrets.toml`(또는 환경변수)로 바꿀 수 있습니다. `EXCEL_MERGER_PROFILE = "cloud"` 한 줄을 넣으면 메모리 약 1GB인 **Streamlit Community Cloud 무료 서버용 보수적인 값**을 쓰고, 항목마다 따로 덮어쓸 수도 있습니다. 설정으로 엔진의 안전 상한(파일당 20MB, 50,000행, 50개)보다 크게 할 수는 없습니다.

| 항목 (설정 이름) | 로컬 기본값 | `cloud` 프로필 |
|---|---|---|
| 한 번에 올리는 파일 수 (`MAX_FILES`) | 50개 | 10개 |
| 파일 크기 (`MAX_FILE_MB`) | 파일당 20MB | 파일당 5MB |
| 올리는 파일 합계 (`MAX_TOTAL_UPLOAD_MB`) | 100MB | 20MB |
| 파일당 행 수 (`MAX_ROWS_PER_FILE`) | 50,000행 | 10,000행 |
| 모든 파일 행 합계 (`MAX_TOTAL_ROWS`) | 200,000행 | 30,000행 |
| 처리 시간 (`TIME_LIMIT_SECONDS`) | 단계마다 60초 | 45초 |
| 동시 처리 작업 수 (`MAX_CONCURRENT_JOBS`) | 3 | 1 |
| 작업별 메모리 상한 (`JOB_MEMORY_MB`, Linux만) | 없음 | 400MB |

```toml
# Streamlit Community Cloud 의 Secrets 칸 예시
EXCEL_MERGER_PROFILE = "cloud"
GEMINI_API_KEY = "여기에 키"   # 선택
# MAX_FILES = 5                 # 개별 값만 바꾸고 싶을 때
```

그 밖의 고정 제한:

| 항목 | 제한 |
|---|---|
| 파일 내용 검사 | 압축을 푼 크기·압축률·시트 수 등 (압축 폭탄 방지) |
| 오류목록 | 최대 20,000건 표시 |
| 중복 행 고르기 | 화면에서 최대 50개 그룹 (나머지는 모든 행을 남김) |
| AI 추천 호출 | 세션당 5회, 앱 전체 분당 10회·하루 200회. Google 쪽 일시적 오류(503·429)는 1초·3초 뒤 자동 재시도하고, 그렇게 끝난 시도는 세션 횟수에서 빼지 않습니다 |

- 파일 형식은 `.xlsx`, `.xlsm` 입니다. `.xls` 는 엑셀에서 `.xlsx` 로 다시 저장해 주세요.
- 업로드 칸의 일부 안내 문구("Upload", "20MB per file")는 Streamlit이 만드는 영어입니다. `cloud` 프로필에서도 이 칸에는 20MB로 보이지만, 실제로는 앱이 5MB 제한을 검사하고 한국어로 안내합니다.
