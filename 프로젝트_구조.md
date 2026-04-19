# 프로젝트 구조 설명서

> 마지막 업데이트: 2026-04-19 (eco_input_gui 레이아웃 개선 — 파일추가/DnD/모드별 활성화)

---

## 1. 프로젝트 개요

측정인.kr 사이트와 로컬 엑셀 성적서 간의 데이터 **자동 입력 / 검토 / 보고서 생성**을 수행하는 자동화 도구 모음.

- **대상 사이트:** 측정인.kr (대기측정 / 수질측정)
- **대상 엑셀:** NAS에 저장된 성적서 `.xlsm` 파일들
- **실행 환경:** Windows + Python 3.x + Chrome + Office (COM)

---

## 2. 전체 아키텍처

```
┌─────────────────────────────────────────────────────────────────┐
│                         사용자 실행                              │
│   .bat 파일 → GUI (tkinter) → 핵심 로직 (Selenium + openpyxl)   │
└─────────────────────────────────────────────────────────────────┘

┌──────────────┐    ┌──────────────┐    ┌──────────────────────┐
│  GUI 계층     │    │  핵심 로직    │    │  공통 유틸리티 계층    │
│              │    │              │    │                      │
│ eco_input_gui│───→│ eco_input    │───→│ measin_utils         │
│ eco_check_gui│───→│ eco_check    │    │ selenium_utils       │
│ report_check │───→│ report_check │    │ realgrid_utils       │
│   _gui       │    │              │    │ select2_utils        │
│              │    │ receipt      │    │ data_utils           │
│ (dash.py     │    │ dash         │    │ excel_utils          │
│  receipt.py  │    │ Vehicle_log  │    │ excel_com_utils      │
│  Vehicle_log │    │ tab4_pdf_gui │    │ file_utils           │
│  = GUI+로직  │    │              │    │ format_utils         │
│   통합형)    │    │              │    │ pdf_utils            │
│              │    │              │    │ gui_common           │
│              │    │              │    │ worker_utils         │
│              │    │              │    │ log_utils            │
└──────────────┘    └──────────────┘    └──────────────────────┘

┌────────────────────────────────────────────────────────────┐
│                     분리 모듈 (신규)                         │
│                                                            │
│  backdata_utils.py   ← eco_input에서 백데이터 분리    │
│  tab4_utils.py       ← eco_input에서 탭4 로직 분리      │
│  measin_constants.py ← 공통 상수 (input+check 공유)  │
└────────────────────────────────────────────────────────────┘
```

---

## 3. 실행 흐름별 파일 관계

### 3-1. 측정인 자동입력 (대기/수질)

```
1.측정인 자동입력.bat
  └─ eco_input_gui.py        ← GUI (fake_input 패턴)
       └─ eco_input.py       ← 핵심 로직
            ├─ backdata_utils.py ← 백데이터(수분/THC) 추출
            ├─ tab4_utils.py     ← 탭4 입력/탐색/저장
            ├─ measin_constants.py ← 공통 상수
            ├─ measin_utils.py    (로그인/검색/상세페이지)
            ├─ selenium_utils.py  (브라우저 조작)
            ├─ realgrid_utils.py  (리얼그리드 입력)
            ├─ select2_utils.py   (드롭다운 조작)
            ├─ excel_utils.py     (성적서 파싱)
            ├─ excel_com_utils.py (Excel COM 관리)
            ├─ pdf_utils.py       (PDF 생성/병합)
            ├─ data_utils.py      (시료번호 파싱)
            ├─ format_utils.py    (숫자/시간 서식)
            ├─ file_utils.py      (NAS 파일 검색)
            └─ log_utils.py       (에러 로그)
```

### 3-2. 측정인 검토

```
2.측정인 검토.bat
  └─ eco_check_gui.py        ← GUI
       └─ eco_check.py       ← 핵심 로직
            ├─ measin_constants.py ← 공통 상수
            ├─ measin_utils.py
            ├─ selenium_utils.py
            ├─ realgrid_utils.py  (API로 사이트값 읽기)
            ├─ excel_utils.py     (엑셀 기대값 생성)
            ├─ data_utils.py
            ├─ format_utils.py
            ├─ file_utils.py
            └─ log_utils.py
```

### 3-3. 발송대장 검토

```
3.발송대장 검토.bat
  └─ receipt.py               ← GUI + 로직 통합
       ├─ file_utils.py
       ├─ data_utils.py
       └─ log_utils.py
```

### 3-4. 종합 검토 (대시보드)

```
4.종합 검토.bat
  └─ dash.py                  ← GUI + 로직 통합
       ├─ data_utils.py
       ├─ format_utils.py
       └─ log_utils.py
```

### 3-5. 성적서 검토

```
5.성적서 검토.bat
  └─ report_check_gui.py      ← GUI (드래그앤드롭)
       └─ report_check.py     ← 핵심 로직
            ├─ excel_utils.py
            ├─ data_utils.py
            ├─ format_utils.py
            └─ log_utils.py
```

### 3-6. 차량운행일지 검토

```
6.차량운행일지 검토.bat
  └─ Vehicle_operation_log.py  ← GUI + 로직 통합
       ├─ data_utils.py
       └─ log_utils.py
```

### 3-7. PDF 생성

```
7.PDF 생성.bat
  └─ tab4_pdf_final_gui.py    ← GUI + 로직
       ├─ pdf_utils.py
       ├─ excel_com_utils.py
       └─ log_utils.py
```

---

## 4. 파일 분류표

### 4-1. GUI 계층 (사용자 인터페이스)

| 파일 | 유형 | 설명 |
|------|------|------|
| `eco_input_gui.py` | GUI 래퍼 | 자동입력 GUI (fake_input, DnD, 파일추가) |
| `eco_check_gui.py` | GUI 래퍼 | 검토 GUI (fake_input 패턴) |
| `report_check_gui.py` | GUI 래퍼 | 성적서 검토 GUI (드래그앤드롭) |
| `receipt.py` | GUI+로직 통합 | 발송대장 검토 |
| `dash.py` | GUI+로직 통합 | 종합 대시보드 |
| `Vehicle_operation_log.py` | GUI+로직 통합 | 차량운행일지 검토 |
| `tab4_pdf_final_gui.py` | GUI+로직 통합 | 탭4 PDF 수동 생성 |

### 4-2. 핵심 로직 계층

| 파일 | 줄 수 | 설명 |
|------|------|------|
| `eco_input.py` | ~1,700 | 측정인 사이트 자동 입력 (대기/수질) |
| `eco_check.py` | ~1,170 | 사이트 vs 엑셀 비교 검토 |
| `report_check.py` | ~1,480 | 성적서 내용 자체 검사 |
| `backdata_utils.py` | ~440 | 백데이터(수분/THC) 추출 (신규 분리) |
| `tab4_utils.py` | ~443 | 탭4 입력/탐색/저장 (신규 분리) |
| `measin_constants.py` | ~61 | 공통 상수 + CSS 셀렉터 (input+check 공유) (신규) |

### 4-3. 공통 유틸리티 계층

| 파일 | 줄 수 | 카테고리 | 설명 |
|------|------|----------|------|
| `measin_utils.py` | ~250 | 사이트 | 로그인/페이지 이동/시료검색 |
| `selenium_utils.py` | ~174 | 사이트 | Chrome 드라이버/클릭/대기 + wait() |
| `realgrid_utils.py` | ~291 | 사이트 | RealGrid 읽기/쓰기 |
| `select2_utils.py` | ~60 | 사이트 | Select2 드롭다운 조작 |
| `excel_utils.py` | ~130 | 엑셀 | openpyxl 시트찾기/파싱 |
| `excel_com_utils.py` | ~60 | 엑셀 | Excel COM 전역 관리 |
| `pdf_utils.py` | ~75 | 엑셀 | PDF 생성/병합 |
| `file_utils.py` | ~130 | 파일 | NAS 파일 검색 |
| `data_utils.py` | ~233 | 데이터 | 시료번호/시간/업체 정규화 + clean_leading_mark |
| `format_utils.py` | ~180 | 데이터 | 숫자/날짜 서식 변환 |
| `gui_common.py` | ~200 | GUI | Tkinter 공통 컴포넌트 |
| `worker_utils.py` | ~120 | GUI | 백그라운드 작업/input 가로채기 |
| `log_utils.py` | ~45 | 기타 | 에러 로그 파일 기록 |

---

## 5. 주요 디자인 패턴

### fake_input 패턴
`eco_input.py`/`eco_check.py`는 콘솔용 `input()`으로 사용자 입력을 받음.
GUI에서는 `worker_utils.InputInterceptor`로 `input()`을 가로채서 미리 준비한 답변 배열을 순서대로 주입.

> ⚠ `input()` 호출 순서가 바뀌면 GUI의 answers 배열도 반드시 같이 수정 필요

### Excel COM 전역 재사용
`excel_com_utils.get_excel_app()`으로 Excel 프로세스를 1개만 띄워 전역 재사용.
프로그램 종료 시 `atexit`로 자동 정리.

### RealGrid API 직접 조작
사이트의 RealGrid는 일반 HTML이 아니므로 JavaScript를 직접 실행하여 DataProvider/GridView API로 데이터를 읽고 씀.
(`realgrid_utils.py`의 `_CORE_REALGRID_JS`)

---

## 6. 데이터 흐름 요약

```
NAS 엑셀 성적서 (.xlsm)
    ↓ openpyxl / Excel COM
파이썬 딕셔너리 (per_item, data 등)
    ↓ Selenium + JS API
측정인.kr 사이트 RealGrid / 폼 필드
    ↓ 비교 (eco_check)
결과 엑셀 (.xlsx)
```

---

## 7. 시료번호 규칙

- 형식: `{매체}{YYMMDD}{팀번호}-{순번}`
- 예: `A2604181-03` = **대기**(A), **2026-04-18**, **1팀**, **3번째**
- 매체: A=대기, W=수질
- 팀번호: 1~5
- 순번: 01~99
