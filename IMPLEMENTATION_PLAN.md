# MSS SKU 매칭 및 발주 파일 변환기 — 기술 문서

> 최종 업데이트: 2026-03-29 (v2.0)

---

## 1. 프로젝트 개요

### 1.1 목적
구글 시트(또는 엑셀/CSV)에 있는 상품 발주 수량 목록을 기준으로, **시스템 업로드용** 파일과 **브랜드 송부용** 파일 2개의 `.xlsx` 파일을 자동 생성한다.

### 1.2 대상 사용자
- 비개발자 업무 담당자 (이커머스/리테일 운영팀)
- Windows PC 사용
- 매일 반복적으로 발주 데이터를 처리하는 실무자

### 1.3 핵심 제약사항

| 항목 | 요구사항 |
|------|---------|
| 보안 | 파일 데이터가 로컬 환경 밖으로 나가지 않을 것 |
| 환경 | Windows PC, Python/터미널 사용 불가 |
| 배포 | 단일 `.exe` 파일로 설치 없이 실행 |
| GUI | 비개발자가 직관적으로 사용 가능한 인터페이스 |

---

## 2. 기능 요약

### 2.1 기능 A: SKU 매칭 변환
3개의 입력 파일(발주 데이터, 바코드-UID 매칭, 옵션 정보)을 받아 브랜드별로 2개의 출력 파일을 생성.

### 2.2 기능 B: 발주 파일 통합
날짜별·브랜드별로 분산된 발주 파일들을 하나의 통합 파일로 병합.

### 2.3 v2 UX 개선 기능

| 기능 | 설명 | 구현 위치 |
|------|------|----------|
| 프로그레스 바 | 처리 진행 상태를 시각적으로 표시 | `app.py` |
| 파일 자동 검증 | 파일 선택 직후 필수 컬럼 존재 여부 확인 | `core/ui_helpers.py` |
| 프리셋 | 자주 쓰는 파일+브랜드 조합 저장/불러오기 | `core/session.py`, `app.py` |
| 세션 자동 복원 | 마지막 사용한 파일/폴더/브랜드를 앱 재시작 시 복원 | `core/session.py` |
| 작업 이력 | 과거 변환 기록을 "작업 이력" 탭에서 조회 | `core/session.py`, `app.py` |
| 매칭 실패 리포트 | 경고 상세 내역을 Excel로 내보내기 | `app.py` |
| 출력 미리보기 | 실행 전 처리 대상과 생성 파일 목록을 확인 | `app.py` |
| 덮어쓰기 보호 | 기존 출력 파일이 있으면 확인 후 진행 | `app.py` |
| 취소 기능 | 처리 중 "취소" 버튼으로 안전하게 중단 | `app.py` |
| 드래그앤드롭 | 파일 탐색기에서 끌어다 놓기로 파일 선택 | `app.py`, `hook-tkinterdnd2.py` |
| 에러 메시지 개선 | Python traceback 대신 한국어 안내 메시지 | `core/ui_helpers.py` |
| 툴팁 | 주요 입력 필드에 마우스 오버 시 설명 표시 | `core/ui_helpers.py` |
| 반응형 창 크기 | 창 크기 조절 가능, 최소 크기 720x800 | `app.py` |
| MouseWheel 스코프 | 체크박스 캔버스 위에서만 스크롤 동작 | `app.py` |
| 초기화 버튼 | 모든 입력을 초기 상태로 리셋 | `app.py` |

---

## 3. 기술 스택

| 구분 | 기술 | 용도 |
|------|------|------|
| 언어 | Python 3.9+ | |
| GUI | tkinter | 네이티브 Windows GUI |
| 드래그앤드롭 | tkinterdnd2 | 파일 DnD 지원 (exe에 번들) |
| Excel 처리 | openpyxl | .xlsx 읽기/쓰기, 템플릿 양식 유지 |
| 데이터 처리 | pandas | 필터링, 조인, 집계 |
| 문자열 매칭 | RapidFuzz | 자연어 옵션 매칭 (3단계 fallback) |
| 빌드 | PyInstaller | 단일 .exe 생성 |
| CI/CD | GitHub Actions | push 시 자동 Windows exe 빌드 |

---

## 4. 프로젝트 구조

```
MSS_SKU_matching_and_converter/
├── app.py                      # GUI 메인 앱 (tkinter, 전체 UI/UX 로직)
├── core/
│   ├── __init__.py
│   ├── loader.py               # 파일 로딩 및 전처리, 헤더 행 자동 탐색
│   ├── matcher.py              # 바코드-UID 매칭, 옵션 유무 판별, 옵션 슬롯 매핑
│   ├── generator.py            # 시스템 업로드용/브랜드 송부용 출력 파일 생성
│   ├── merger.py               # 기능 B: 발주 파일 통합 로직
│   ├── utils.py                # 텍스트 정규화 (normalize_for_matching, normalize_strict, split_product_option)
│   ├── ui_helpers.py           # [v2] ToolTip, 에러 메시지 변환, 파일 스키마 검증, FileStatusLabel
│   └── session.py              # [v2] 세션 저장/복원, 작업 이력, 프리셋 관리
├── hook-tkinterdnd2.py         # [v2] PyInstaller hook: tkinterdnd2 DLL 번들링
├── build.py                    # PyInstaller 빌드 스크립트 (로컬 Windows용)
├── requirements.txt            # Python 의존성
├── .github/
│   └── workflows/
│       └── build-windows.yml   # GitHub Actions 자동 빌드
├── docs/
│   ├── index.html              # 웹 도움말 페이지 (GitHub Pages)
│   ├── 사용_가이드.md           # 사용자 가이드 (마크다운)
│   ├── UX_REVIEW.md            # [v2] UX 리뷰 보고서
│   ├── 문제_해결_가이드.md      # 문제 해결 가이드
│   └── DEBUG_REPORT_v2.md      # 디버깅 리포트 (자빈드서울 테스트)
├── templates/                  # 출력 파일 템플릿
├── test data/                  # 테스트 데이터셋
├── sample data - data integration/  # 기능 B 샘플 데이터
├── REQUIREMENTS.md             # 요구사항 정의서
├── IMPLEMENTATION_PLAN.md      # 이 문서 (기술 문서)
└── raw_requirement.md          # 원본 요구사항
```

---

## 5. 모듈 상세

### 5.1 `core/loader.py` — 파일 로딩

| 함수 | 역할 |
|------|------|
| `load_order_data(filepath)` | 파일 1 로드. `브랜드명` 컬럼이 없으면 헤더 행 자동 탐색 |
| `load_matching_data(filepath)` | 파일 2 로드. 빈 첫 행 자동 스킵, Unnamed 컬럼 감지 시 재탐색 |
| `load_option_data(filepath)` | 파일 3 로드 |
| `get_brand_list(order_df)` | `브랜드명` 컬럼에서 고유 브랜드 목록 추출 (정렬) |
| `filter_by_brand(order_df, brand)` | 특정 브랜드로 필터링 |
| `_find_header_row(filepath, markers)` | 엑셀 첫 14행을 탐색하여 마커 컬럼이 있는 헤더 행 반환 |

### 5.2 `core/matcher.py` — 매칭 엔진

| 함수 | 역할 |
|------|------|
| `match_barcode_to_uid(order_df, matching_df)` | 88코드 → 바코드 left join으로 UID, 스타일번호 매칭. 미매칭 바코드 목록 반환 |
| `detect_option_products(merged_df, matching_df)` | UID별 바코드 수로 옵션 유무 판별. `{uid: bool}` 반환 |
| `match_option_info(merged_df, option_df, has_option, matching_df)` | 사이즈유형 매칭 + 옵션 슬롯 결정. 경고 목록 반환 |

**옵션 매핑 전략 (3단계 fallback):**

```
사이즈유형 매칭:
  1. 파일2 상품명 → 파일3 사이즈유형명 (normalize_strict 정확 매칭)
  2. 포함관계 매칭 (substring)
  3. RapidFuzz token_sort_ratio (임계값 80%)

옵션 슬롯 매칭:
  방법 A: 파일2의 UID별 바코드 순서 맵 사용
  방법 B: 옵션명 → 파일3 Size 슬롯 매칭 (정확 → RapidFuzz ratio 75%)
  방법 C: 파일1 상품명에서 # 기반 옵션 추출 (레거시 호환)
```

### 5.3 `core/generator.py` — 출력 생성

| 함수 | 역할 |
|------|------|
| `generate_system_upload(merged_df, matching_df, option_df, has_option, brand, output_path)` | 시스템 업로드용 파일 생성 (2개 시트) |
| `generate_brand_order(merged_df, has_option, brand, output_path)` | 브랜드 송부용 파일 생성 |

**시스템 업로드용 구조:**
- 시트 1 `매입상품 구매오더 업로드 양식`: Row 1-6 빈 행, Row 7 라벨, Row 8 필수 마커, Row 9 메인 헤더, Row 10 서브 헤더, Row 11+ 데이터 (UID 기준 1행)
- 시트 2 `Barcode 업로드(옵션)`: 스타일번호, UID, 사이즈, 바코드 (바코드 기준 1행)

**브랜드 송부용 구조:**
- Row 1 빈 행, Row 2 지정입고일, Row 3-5 빈 행, Row 6 헤더, Row 7 TOTAL, Row 8+ 데이터

### 5.4 `core/merger.py` — 발주 파일 통합

| 함수 | 역할 |
|------|------|
| `merge_order_files(file_paths, matching_filepath, output_path)` | 여러 발주 파일을 통합. 지정입고일 자동 추출, 매칭 파일로 상품번호 채움 |

### 5.5 `core/utils.py` — 텍스트 정규화

| 함수 | 역할 |
|------|------|
| `normalize_for_matching(text)` | `#` 주변 공백 통일, 연속 공백 제거, 전각→반각 괄호 |
| `normalize_strict(text)` | 모든 공백·특수문자 제거 + lowercase (정확 비교용) |
| `split_product_option(name)` | `#` 기호 기준으로 기본 상품명 / 옵션명 분리 |

### 5.6 `core/ui_helpers.py` — UI 헬퍼 (v2)

| 클래스/함수 | 역할 |
|------------|------|
| `ToolTip(widget, text)` | 마우스 오버 시 툴팁 표시 위젯 |
| `friendly_error(exc)` | 예외를 사용자 친화적 한국어 메시지로 변환. 7개 에러 패턴 매핑 |
| `validate_file_schema(filepath, file_num)` | 파일의 필수 컬럼 존재 여부 검증. `{valid, message, columns, row_count}` 반환 |
| `FileStatusLabel(parent)` | `[OK]`/`[!!]` 파일 상태 표시 위젯 |

**에러 패턴 매핑:**

| 원본 에러 | 변환 메시지 |
|----------|-----------|
| `'X' 컬럼을 찾을 수 없습니다` | 올바른 파일인지 확인 안내 |
| `브랜드명 컬럼이 없습니다` | 파일 1 확인 안내 |
| `Permission denied` | 파일이 열려있지 않은지 확인 |
| `FileNotFoundError` | 파일 경로 확인 |
| `BadZipFile` / `InvalidFileException` | 파일 형식 확인 |

### 5.7 `core/session.py` — 세션 관리 (v2)

저장 위치: `~/.mss_converter/` (사용자 홈 디렉토리)

| 함수 | 파일 | 역할 |
|------|------|------|
| `save_session(state)` / `load_session()` | `session.json` | 파일 경로, 저장 위치, 브랜드 선택 상태 저장/복원 |
| `add_history_entry(entry)` / `get_history()` | `history.json` | 작업 이력 저장/조회 (최대 100건) |
| `save_preset(name, config)` / `load_preset(name)` | `presets.json` | 프리셋 저장/로드/삭제 |

---

## 6. GUI 설계 (v2)

### 6.1 전체 레이아웃

```
┌──────────────────────────────────────────────────────┐
│  MSS SKU 매칭 및 발주 파일 변환기                        │
├────────────────┬──────────────────┬──────────────────┤
│ A. SKU 매칭 변환 │ B. 발주 파일 통합  │   작업 이력       │
├────────────────┴──────────────────┴──────────────────┤
│                                                      │
│  프리셋: [콤보박스 ▼] [불러오기] [저장] [삭제]           │
│                                                      │
│  ┌─ 입력 파일 ──────────────────────────────────────┐ │
│  │ 1. 총 발주 수량   [경로] [OK 1234행] [파일 선택]  │ │
│  │ 2. 매칭 데이터셋  [경로] [OK 500행]  [파일 선택]  │ │
│  │ 3. 옵션 정보      [경로] [OK 50행]   [파일 선택]  │ │
│  │  (파일을 끌어다 놓을 수도 있습니다)                │ │
│  └──────────────────────────────────────────────────┘ │
│                                                      │
│  ┌─ 브랜드 선택 ────────────────────────────────────┐ │
│  │ (●) 모든 브랜드  ( ) 브랜드 선택                   │ │
│  │ 선택됨 (3개): 브랜드A, 브랜드B, 브랜드C            │ │
│  │ 검색: [________]                                  │ │
│  │ [전체 선택] [전체 해제]                             │ │
│  │ ☑ 브랜드A  ☑ 브랜드B  ☑ 브랜드C  ☐ 브랜드D     │ │
│  │ (체크박스 180px 스크롤 영역)                       │ │
│  └──────────────────────────────────────────────────┘ │
│                                                      │
│  저장 위치: [바탕화면                 ] [폴더 선택]    │
│                                                      │
│  [변환 실행]  [취소]  [초기화]                         │
│                                                      │
│  ████████████████░░░░░░  브랜드B (3/10)              │
│                                                      │
│  ┌─ 처리 결과 ──────────────────────────────────────┐ │
│  │ (스크롤 로그 영역)                                │ │
│  └──────────────────────────────────────────────────┘ │
│                                                      │
│  [출력 폴더 열기]  [매칭 실패 리포트]                   │
└──────────────────────────────────────────────────────┘
```

### 6.2 사용자 플로우

```
앱 실행 → 세션 자동 복원
  │
  ├─ (선택) 프리셋 불러오기
  │
  ├─ 파일 1 선택 → 자동 검증 + 브랜드 로드
  ├─ 파일 2 선택 → 자동 검증
  ├─ 파일 3 선택 → 자동 검증
  │
  ├─ 브랜드 선택 (모든 브랜드 / 체크박스 선택)
  │
  ├─ "변환 실행" 클릭
  │     ├─ 미리보기 다이얼로그 (덮어쓰기 경고 포함)
  │     ├─ "확인" → 처리 시작
  │     ├─ 프로그레스 바 업데이트
  │     └─ (선택) "취소" 버튼으로 중단
  │
  ├─ 결과 확인
  │     ├─ "출력 폴더 열기"
  │     └─ (경고 발생 시) "매칭 실패 리포트" → Excel 저장
  │
  └─ 세션 자동 저장 + 이력 기록
```

---

## 7. 빌드 및 배포

### 7.1 GitHub Actions 자동 빌드

`.github/workflows/build-windows.yml`:
- **트리거**: `main` 브랜치 push 또는 수동 실행
- **환경**: `windows-latest`, Python 3.11
- **의존성**: pandas, openpyxl, rapidfuzz, pyinstaller, tkinterdnd2
- **빌드**: PyInstaller `--onefile --windowed` + `--additional-hooks-dir=.` (tkinterdnd2 DLL 번들링)
- **아티팩트**: `dist/SKU_변환기.exe` → GitHub Actions Artifacts에 업로드

### 7.2 tkinterdnd2 번들링

tkinterdnd2는 Tcl/Tk 네이티브 DLL을 포함하므로 PyInstaller에서 특별 처리가 필요:

1. `hook-tkinterdnd2.py`: `collect_data_files('tkinterdnd2')`로 DLL 포함
2. 빌드 시 `--additional-hooks-dir=.` 옵션으로 hook 활성화
3. tkinterdnd2의 `_require()` 함수가 `__file__` 기준으로 `tkdnd/win-x64/` 내 DLL을 찾음
4. `app.py`에서 import 실패 시 graceful fallback (버튼만 표시)

### 7.3 배포

```
배포물: SKU_변환기.exe (단일 파일, ~35MB)
다운로드: GitHub Actions > 최신 빌드 > Artifacts > SKU_변환기
```

1. `main` 브랜치에 push → 자동 빌드
2. GitHub Actions에서 Artifacts 다운로드
3. 사용자에게 `.exe` 파일 전달 (메일, USB, 사내 공유 등)
4. 사용자는 `.exe` 더블클릭으로 바로 사용 — 설치 과정 없음

### 7.4 로컬 빌드 (Windows에서만)

```bash
pip install -r requirements.txt
python build.py
# 또는
pyinstaller app.py --onefile --windowed --name=SKU_변환기 --add-data=core;core --additional-hooks-dir=. --hidden-import=tkinterdnd2
```

---

## 8. 데이터 흐름

### 8.1 기능 A: SKU 매칭 변환

```
파일 1 (총 발주 수량) ──┐
                        ├──→ 브랜드 필터링
파일 2 (바코드 매칭) ───┤      │
                        │      ▼
                        ├──→ 바코드 → UID 매칭
                        │      │
파일 3 (옵션 정보) ─────┤      ▼
                        ├──→ 옵션 유무 판별 (UID별 바코드 수)
                        │      │
                        │      ▼
                        └──→ 옵션 매핑 (사이즈유형 + 슬롯)
                               │
                               ▼
                    ┌──────────┴──────────┐
                    ▼                     ▼
          시스템 업로드용              브랜드 송부용
          (UID 기준 1행)             (바코드 기준 1행)
```

### 8.2 기능 B: 발주 파일 통합

```
발주 파일 1 ─┐
발주 파일 2 ─┤──→ 각 파일 로드 + 지정입고일 추출
발주 파일 N ─┘      │
                     ▼
              컬럼명 정규화 + concat
                     │
                     ▼
(선택) 매칭 파일 ──→ 88코드 → 상품번호 채움
                     │
                     ▼
              통합_발주수량.xlsx
```

---

## 9. 설정 파일 구조

### 9.1 세션 (`~/.mss_converter/session.json`)
```json
{
  "tab_a": {
    "file1": "C:/path/to/file1.xlsx",
    "file2": "C:/path/to/file2.xlsx",
    "file3": "C:/path/to/file3.xlsx",
    "save_path": "C:/Users/user/Desktop",
    "brand_mode": "select",
    "selected_brands": ["브랜드A", "브랜드B"]
  },
  "tab_b": {
    "files": ["C:/path/a.xlsx", "C:/path/b.xlsx"],
    "matching_file": "C:/path/match.xlsx",
    "save_path": "C:/Users/user/Desktop"
  },
  "last_tab": 0
}
```

### 9.2 이력 (`~/.mss_converter/history.json`)
```json
[
  {
    "timestamp": "2026-03-29 14:30:00",
    "type": "A",
    "brands": ["브랜드A"],
    "success": 1, "fail": 0, "total": 1,
    "warnings": 2,
    "files": {"file1": "발주.xlsx", "file2": "매칭.xlsx", "file3": "옵션.xlsx"},
    "output_path": "C:/Users/user/Desktop"
  }
]
```

### 9.3 프리셋 (`~/.mss_converter/presets.json`)
```json
{
  "3월 정기 발주": {
    "file1": "C:/path/to/file1.xlsx",
    "file2": "C:/path/to/file2.xlsx",
    "file3": "C:/path/to/file3.xlsx",
    "save_path": "C:/Users/user/Desktop",
    "brand_mode": "select",
    "selected_brands": ["브랜드A"],
    "updated_at": "2026-03-29 14:00:00"
  }
}
```

---

## 10. 보안 확인사항

| 항목 | 보장 방식 |
|------|----------|
| 파일 데이터 외부 전송 없음 | 로컬 전용 데스크톱 앱, 네트워크 모듈 미포함 |
| 서버 불필요 | 사용자 PC에서 독립 실행 |
| 임시 파일 관리 | 메모리 내 처리, 결과물만 사용자 지정 폴더에 저장 |
| 설정 파일 | `~/.mss_converter/`에 JSON으로 저장, 민감 데이터 미포함 (경로만 저장) |
| 소스코드 투명성 | Python 스크립트로 전체 로직 확인 가능 |

---

## 11. 문서 목록

| 문서 | 위치 | 용도 |
|------|------|------|
| 기술 문서 (이 파일) | `IMPLEMENTATION_PLAN.md` | 아키텍처, 모듈 구조, 빌드/배포 |
| 요구사항 정의서 | `REQUIREMENTS.md` | 입출력 명세, 처리 로직 상세 |
| 사용자 가이드 | `docs/사용_가이드.md` | 최종 사용자용 사용법 (마크다운) |
| 웹 도움말 | `docs/index.html` | 최종 사용자용 사용법 (GitHub Pages 웹페이지) |
| UX 리뷰 보고서 | `docs/UX_REVIEW.md` | v2 UX 분석 및 개선 아이디어 |
| 문제 해결 가이드 | `docs/문제_해결_가이드.md` | 트러블슈팅 (웹 도움말에도 포함) |
| 디버깅 리포트 | `docs/DEBUG_REPORT_v2.md` | 자빈드서울 테스트 시 발견된 문제 및 수정 기록 |
| 원본 요구사항 | `raw_requirement.md` | 고객 원문 요구사항 |

---

## 12. 변경 이력

| 버전 | 날짜 | 변경 내용 |
|------|------|----------|
| v1.0 | 2026-03-28 | 초기 구현 — 기능 A (SKU 매칭 변환) |
| v1.1 | 2026-03-28 | 기능 B (발주 파일 통합) 추가 |
| v1.2 | 2026-03-28 | 브랜드 선택 UX 개선 (체크박스 기반), 헤더 행 자동 탐색 |
| v2.0 | 2026-03-29 | UX 전면 개선 — 15개 개선사항 반영 (프로그레스 바, 프리셋, 세션 복원, 파일 검증, 이력, 리포트, 미리보기, 덮어쓰기 보호, 취소, DnD, 에러 메시지, 툴팁, 반응형 창, MouseWheel 수정, 초기화) |
