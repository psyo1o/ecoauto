# 측정인 자동화 도구 안내서

> 마지막 업데이트: 2026-06-02 (폴더 구조: `3.py`)
> 대상 독자: 처음 사용하는 담당자, 인수인계 받는 사용자, 운영 담당자

---

## 1. 이 프로그램이 하는 일

이 폴더는 `측정인.kr` 사이트 업무를 빠르게 처리하기 위한 Windows 전용 자동화 도구 모음이다.

주요 목적은 다음과 같다.

- NAS에 저장된 성적서 엑셀 파일을 읽어 측정인.kr에 자동 입력 (대기·수질)
- 사이트에 입력된 값과 엑셀 값을 비교 검토
- 발송대장, 차량운행일지, 성적서 상태를 점검해 결과 보고서 생성
- 탭4 PDF를 자동 또는 반자동으로 생성하고 최종본을 병합
- 수분량/THC 백데이터를 규칙에 맞게 추출하여 NAS에 저장

이 프로그램은 단일 실행 파일 하나가 아니라, 업무 종류별로 나뉜 7개의 도구로 구성되어 있다.

---

## 2. 처음 사용할 때

### 2-1. 실행 환경

이 프로그램은 다음 환경을 전제로 한다.

- Windows PC
- Google Chrome 설치
- Python 3.x 설치
- Microsoft Excel 사용 가능 환경
- NAS 경로 `\\192.168.10.163\측정팀` 접근 가능 (IP 등은 config.ini에서 변경 가능)

Windows 전용인 이유는 다음 기능을 사용하기 때문이다.

- Excel COM 자동화 (`win32com`, `pywin32`)
- 파일 열기 창 자동 제어 (`pywinauto`)
- Windows 레지스트리 수정으로 보안 경고 완화
- 통합 런처(`2.검토 및 입력프로그램.pyw`) 더블클릭 실행 (콘솔 창 없음)

### 2-2. 첫 설치 순서

처음 실행하는 PC에서는 `0.처음사용시` 폴더를 먼저 사용한다.

권장 순서는 다음과 같다.

1. `0.처음사용시\드라이버+파이썬패키지+보안경고 통합설치.bat` 실행
2. 설치가 끝나면 프로그램 폴더의 **`2.검토 및 입력프로그램.pyw`** 를 더블클릭해 사용

통합 설치 파일이 하는 일은 다음과 같다.

- Python 실행 명령 확인 (`py -3` 또는 `python`)
- Excel 신뢰 위치에 NAS 등록
- Windows 로컬 인트라넷에 NAS 등록
- Excel의 인터넷 차단 보안 설정 완화
- Chrome 버전 확인 후 ChromeDriver 다운로드 및 PATH 등록
- 필요한 Python 패키지 설치

설치에 포함된 주요 패키지:

- `selenium`
- `webdriver-manager`
- `openpyxl`
- `pandas`
- `pywin32`
- `requests`
- `pywinauto`
- `pyperclip`
- `PyPDF2`
- `pypdf`
- `tkinterdnd2`
- `pillow`

### 2-3. 설치 관련 보조 파일

`0.처음사용시` 폴더에는 아래 파일들이 있다.

| 파일 | 용도 |
|------|------|
| `드라이버+파이썬패키지+보안경고 통합설치.bat` | 처음 사용하는 PC에서 가장 먼저 실행하는 통합 설치 파일 |
| `드라이버 자동설치(크롬).bat` | ChromeDriver만 따로 설치할 때 사용 |
| `파이썬 패키지.bat` | Python 패키지만 다시 설치하거나 누락분만 설치할 때 사용 |
| `보안경고시 실행.REG` | Excel 보안 경고 관련 레지스트리만 따로 적용할 때 사용 |

---

## 3. 어떤 파일을 실행하면 되는지

**실행 진입점은 하나뿐이다.** NAS 프로그램 폴더에서 **`2.검토 및 입력프로그램.pyw`** 를 더블클릭하면 통합 런처 GUI가 열린다. 원하는 도구 버튼을 누르면 해당 Python GUI가 백그라운드로 실행된다.

| 런처 버튼 | 실행되는 Python 파일 | 용도 |
|------|------|------|
| 1. 측정인 자동입력 | `3.py\eco_input_gui.py` | 측정인.kr 자동 입력, 백데이터, 탭4 PDF 처리 |
| 2. 측정인 검토 | `3.py\eco_check_gui.py` | 사이트 값과 엑셀 값 비교 검토 |
| 3. 발송대장 검토 | `3.py\receipt.py` | 발송대장 기준 상태 점검 |
| 4. 종합 검토 | `3.py\dash.py` | 여러 검토 결과를 종합한 대시보드 생성 |
| 5. 성적서 검토 | `3.py\report_check_gui.py` | 성적서 파일 자체 무결성 검사 |
| 6. 차량운행일지 검토 | `3.py\Vehicle_operation_log.py` | 차량운행일지/근무시간 검토 |
| 7. PDF 생성 | `3.py\tab4_pdf_final_gui.py` | 탭4 PDF 최종본 수동 생성 |

런처 동작 요약:

- `.pyw` 확장자로 더블클릭 시 **콘솔 창이 뜨지 않음**
- 하위 프로그램은 `pythonw` + `CREATE_NO_WINDOW` 로 실행
- 작업 디렉터리(`cwd`)는 **`3.py` 폴더**로 고정 (import·UNC 경로 이탈 방지)
- Python 소스는 루트가 아니라 **`3.py`** 폴더에 모음 (`config.ini`는 루트에 유지)

> 예전에 쓰던 `1.측정인 자동입력 V2.9.bat` ~ `7.PDF 생성 V1.2.bat` 은 런처로 대체되어 **삭제**했다.  
> `0.처음사용시` 폴더의 설치용 `.bat` 만 그대로 둔다.

---

## 4. 업무별로 어떤 도구를 써야 하는지

### 4-1. 측정인 자동입력

사용 상황:

- NAS 성적서를 기준으로 측정인.kr에 자동 입력해야 할 때
- 대기 탭1/탭2/탭4 입력이 필요할 때
- 수질 탭2·탭4·PDF 입력이 필요할 때
- 백데이터와 PDF 생성까지 한 번에 처리할 때

핵심 파일:

- `eco_input_gui.py`
- `eco_input.py`
- `backdata_utils.py`
- `tab4_utils.py` (대기 탭4)
- `water_input_utils.py` (수질 탭2·탭4)

GUI 요약 (`eco_input_gui.py`):

- 매체: **대기(1)** / **수질(2)** — 선택에 따라 입력 영역 전환
- **대기** 옵션(2열 3행): 탭1 | 백데이터 / 탭2 | 탭2 PDF / 탭4 | 탭4 PDF
- 대기 탭1: 성적서 **입력** 시트 **F10** → `#edit_meas_purpose` (1=자가측정용 `SELF`, 2=참고용 `CF`)  
  - 탭1 **인력·차량·장비**는 엑셀에 ` / ` 로 붙어 있어도 사이트 Select2에는 아래 규칙으로만 등록 (공백은 모두 제거) — 상세는 **§4-1-2**
  - 탭2 끄면 탭2 PDF 자동 OFF, 탭4 끄면 탭4 PDF 자동 OFF  
  - 작업 `2`(백데이터만)이면 사이트 입력·PDF 체크는 모두 OFF, 백데이터만 ON
- **수질** 옵션(한 줄): 시험의뢰정보 (탭2) / 탭4 / 탭4 PDF (탭4 끄면 PDF OFF)
- 수질 파일 추가·NAS 확인 기본 경로: `14.수질성적서\0.입력중` (`WATER_REPORT_INPUT`)
- 실행 시 `main()`·`input()` 가로채기 없이 `_main_air()` / `_main_water()`에 옵션을 **키워드 인자**로 직접 전달 (GUI 무반응 방지)

최근 반영된 동작:

- 상세 진입 실패·세션 만료 시 재로그인 후 같은 시료 재시도 — **대기**는 날짜 검색 복구, **수질**은 날짜 검색 없이 시료번호만 검색
- 탭2 저장이 성공했고 PDF 업로드를 켠 경우에는 PDF 삭제까지 끝난 시료만 완료로 보고 이후 사이트 입력을 건너뜀(탭4만 남았으면 탭4만 진행)
- 탭2 저장에 실패했거나 PDF 업로드·삭제가 끝나지 않은 시료는 완료로 표시하지 않아 세션 복구 후 같은 시료를 다시 시도
- 수질 시료번호 `W2606300-01` 형식 지원 (`data_utils.extract_sample_from_name`)
- 수질 탭4: PDF 없으면 임시저장(`#btnTempSave`), PDF 있으면 분석완료(`#btnCompSave`, 확인 2회)
- 수질 탭4만·탭2 미완료 시 탭2 입력완료 선행 후 `#ui-id-4` 진입
- 수질 PDF 업로드: `#anzeFile1`(시험분석일지) / `#anzeFile2`(측정기록부) — 대기 탭4 `#newFile1/2` 와 **id 다름**
- 수질 병합 PDF·대행기록부 단독 PDF 저장 경로는 **§4-1-1**, `config.ini` 참고
- 대기 탭1 인력·차량·장비 슬래시 규칙은 **§4-1-2**

### 4-1-1. 수질 자동입력 상세 (유지보수·수정 시)

사이트·검색:

| 항목 | 값 | 수정 위치 |
|------|-----|-----------|
| URL | `https://측정인.kr/ms/field_water.do` | `config.ini` → `FIELD_URL_WATER` |
| 성적서 NAS 검색 | `...\14.수질성적서\0.입력중` | `WATER_REPORT_INPUT` |
| 시료번호 형식 | `W` + YYMMDD + 팀 + `-` + 순번 | `data_utils.extract_sample_from_name` |
| **목록 검색** | **시료번호만** (`#search_meas_mgmt_no` → `#btnSearch` → 더블클릭) | `_main_water` — **날짜 범위 검색(`search_date`) 안 함** (대기와 다름) |
| 세션 복구 | 재로그인 후 시료번호만 재검색 | `open_detail_with_session_recovery(..., date_str=None)` |

GUI → `_main_water()` 인자:

| 체크박스 | 인자 | 비고 |
|----------|------|------|
| 시험의뢰정보 (탭2) | `do_tab2` | PE / 용량·개수 1 / 측정위치 `폭기조` → `#btnSaveMsFieldDoc` (확인 2회) |
| 탭4 | `do_tab4` | OFF면 PDF도 OFF |
| 탭4 PDF | `do_pdf_final` | 분석일지+대행기록부 생성·업로드·병합 저장 |

탭·버튼 셀렉터 (`water_input_utils.py` / `eco_input.py`):

| 용도 | 셀렉터 |
|------|--------|
| 탭2 | `a#ui-id-2` |
| 탭4 | `a#ui-id-4` |
| 탭2 입력완료 | `#btnSaveMsFieldDoc` |
| 탭4 임시저장 (PDF OFF) | `#btnTempSave` |
| 탭4 분석완료 (PDF ON) | `#btnCompSave` → 확인 alert (탭2와 동일 3단계) |
| 탭4 **재수정·분석완료** 시 수정 사유 | 확인 처리 후 `#update_reason` 뜨면 입력·`#btnSaveupdateReason` → 확인 재처리 (`tab4_after_comp_save_confirm`). **임시저장에는 안 뜸** |
| PDF 파일1 (시험분석일지) | `#anzeFile1` ← `WATER_TAB4_FILE1` / `WATER_FILE_BTN1` |
| PDF 파일2 (측정기록부) | `#anzeFile2` ← `WATER_TAB4_FILE2` / `WATER_FILE_BTN2` |

매크로 엑셀 (`WATER_TAB4_MACRO_FILE`):

| 항목 | 위치 |
|------|------|
| 파일 | `\\...\2.성적서\측정인 측정분석 입력 26.05(수질).xlsm` |
| 시료번호 | A6 |
| 시료접수일시 / 분석시작 / 분석종료 | H6 / H7 / J7 |
| RealGrid 데이터 | 시트 `00. 측정분석결과 입력샘플` — 오픈 시 **해당 시트로 Activate** 후 A6·9~54행 읽기, **8행=헤더**, **A~M** |
| 붙여넣기 | 그리드 **4열(측정항목)** 부터 엑셀 **A~M** 전체 (`water_input_utils`) |

RealGrid (`WATER_TAB4_GRID_ROOT`):

| 항목 | 기본값 |
|------|--------|
| 루트 id | `#gridAnalySampAnzeDataAirItemList1` (바깥 래퍼 `#gridMain1` 과 별도) |
| 항목명 매칭 | 8행 헤더와 RealGrid 헤더 정규화 후 매칭 (`_norm_rg`) |

PDF 저장 (성적서 엑셀 = NAS에서 찾은 `.xlsm` / `.xls`):

| 종류 | 경로 규칙 | config 키 | 코드 |
|------|-----------|-----------|------|
| **병합본** (분석일지+대행기록부) | `{PDF_WATER}\{연도}년\{대행기록부 D3, 공백제거}\{엑셀파일명}.pdf` | `PDF_WATER` | `build_final_path_water()`, `read_water_record_d3()` |
| **대행기록부만** | `{WATER_RECORD_PDF_DIR}\{엑셀파일명}.pdf` (폴더 없으면 생성) | `WATER_RECORD_PDF_DIR` | `build_water_record_copy_path()` |

예시 (IP 192.168.10.163):

- 병합: `\\192.168.10.163\측정팀\2.성적서\0.PDF\1.수질pdf\2026년\문의휴게소\W2606300-01 문의휴게소.pdf` (D3 `문의 휴 게소` → `문의휴게소`)
- 대행만: `\\192.168.10.163\scan\PDF\W2606300-01 문의휴게소.pdf`

PDF 생성·업로드 순서 (`make_tab4_pdfs_water` → `upload_tab4_pdfs`):

1. 성적서 엑셀에서 `분석일지` / `대행기록부` 시트 각각 PDF export
2. 두 PDF 병합 → `PDF_WATER` 경로 저장
3. 대행기록부 PDF만 → `scan\PDF` 복사
4. 탭4 재확인(`open_water_tab4`) 후 `#anzeFile1`·`#anzeFile2`에 `send_keys` 업로드
5. `#btnCompSave` + 확인 2회

수정 시 같이 볼 파일:

- `eco_input.py` — `_main_water()`, `make_tab4_pdfs_water()`, `read_water_record_d3()`
- `water_input_utils.py` — 탭2/탭4·RealGrid
- `eco_input_gui.py` — `_build_answers_water()`
- `config.ini` / `config.py` — 아래 PATHS·URLS 수질 키

### 4-1-2. 대기 탭1 인력·차량·장비 (슬래시 정규화)

성적서 **입력** 시트에 `이름 / 팀 / 직급`처럼 슬래시로 구분된 문자열이 있어도, **자동입력·검토가 같은 규칙**으로 잘라서 사이트·비교에 쓴다. (`format_utils.py` → `excel_utils.parse_measuring_record` → `eco_input.fill_tab1` / `eco_check.compare_list`)

| 구분 | 엑셀 예시 (AC·AD / AE / AH~AQ) | 사이트·비교에 쓰는 값 | 규칙 |
|------|-------------------------------|----------------------|------|
| **인력** | `임종희 / 대기팀 / 기사` | `임종희` | 첫 `/` **앞**, 공백 제거 |
| **차량** | `그랜드 스타렉스 / 81주6787` | `81주6787` | 첫 `/` **뒤**, 공백 제거 (`/` 없으면 전체에서 공백만 제거) |
| **장비** | `대기배출가스측정기5 / Optima 7 / 322913` | `대기배출가스측정기5` | 첫 `/` **앞**, 공백 제거 |

예: `81 주 6787` → `81주6787`, `대기 배출 가스측정기5 / …` → `대기배출가스측정기5`

**측정인 검토**에서는 사이트·엑셀 **양쪽**을 위 규칙으로 맞춘 뒤 목록(set) 비교한다. 사이트에 `그랜드 스타렉스 / 81주6787`처럼 긴 표기가 남아 있어도 엑셀 `81주6787`과 일치로 본다.

수정 시 같이 볼 파일: `format_utils.py`, `excel_utils.py`, `eco_input.py` (`fill_tab1`), `eco_check.py` (`compare_list`)

### 4-2. 측정인 검토

사용 상황:

- 사이트에 입력된 값과 성적서 엑셀 값이 일치하는지 확인할 때
- 자동입력 후 검증이 필요할 때

핵심 파일:

- `eco_check_gui.py`
- `eco_check.py`

최근 반영된 동작:

- **측정목적 검증**: 상세페이지 탭1의 측정목적(자가측정용/참고용)과 엑셀 '입력' 시트 F10 셀의 값(1/2)이 일치하는지 자동으로 확인 (시료번호 일치 여부와 함께 출력)
- **탭1 인력·차량·장비**: 자동입력과 동일한 슬래시·공백 정규화 후 비교 (`readme.md` **§4-1-2**)
- **하드코딩 제거**: 모든 NAS 경로가 `config.ini`를 통해 중앙 관리되도록 개선

### 4-3. 발송대장 검토

사용 상황:

- 시료번호별로 성적서, 수분량, THC, PDF, 백데이터 상태를 한 번에 확인할 때
- 발송대장 시간과 성적서 시간을 비교할 때

최근 반영된 기준:

- 발송대장 A열 값이 시료번호 형식이 아니면 검사/결과 출력에서 제외

핵심 파일:

- `receipt.py`

최근 반영된 동작:

- **진행률 표시**: GUI 하단에 현재 처리 중인 시료 수와 총 시료 수(진행: X / Y)를 표시하여 대규모 작업 시 진행 상황 확인 가능
- **하드코딩 제거**: 모든 NAS 경로가 `config.ini`를 통해 중앙 관리되도록 개선

### 4-4. 종합 검토

사용 상황:

- 검토 결과 파일을 모아 대시보드 형태로 정리할 때
- 인력, 차량, 장비 중복이나 시간 충돌을 종합적으로 보고 싶을 때

핵심 파일:

- `dash.py`

### 4-5. 성적서 검토

사용 상황:

- 성적서 파일 자체의 구성, 값, 조건부 규칙, 시간 정합성을 점검할 때

핵심 파일:

- `report_check_gui.py`
- `report_check.py`

최근 반영된 검토 기준:

- 입자상 채취 시간은 수분량자동측정기, 가스자동측정기 시간과 겹치면 안 됨
- 수분량자동측정기와 가스자동측정기는 입자상보다 먼저 끝나야 함
- 가스자동측정기 시간은 `황산화물`, `질소산화물`, `일산화탄소` 측정시간으로 판단
- 여러 항목이 함께 측정된 경우 세 항목 시간은 서로 일치해야 함
- sample-2 시트에 발송대장 기준 가스미터/VOC 전후값을 자동 채움
- 파일 추가 버튼 기본 탐색 경로는 `\\192.168.10.163\측정팀\2.성적서\0 2.검토중`
- 입자상 채취기준 검사(sample-3에 `[채취기준미달]`로 표시)
  - 먼지: 채취시간 ≥ 40분 OR 시료채취량 ≥ 0.4 OR 무게(C−D) ≥ 0.005 중 하나라도 충족되면 정상
  - 중금속(8종): 시료채취량 ≥ 1 이면 정상
  - 벤조(a)피렌: 시료채취량 ≥ 3 이면 정상
  - 시료채취량 열은 입력(분석값) 1행 헤더(`시료채취량`) 기준으로 자동 검색,
    측정시간은 입력(분석값) 측정시작/측정 종료(I/K열)로 계산

### 4-6. 차량운행일지 검토

사용 상황:

- 차량운행일지, 발송대장, 기술인력 정보 간 정합성을 볼 때
- 주 52시간 관련 검토를 할 때

핵심 파일:

- `Vehicle_operation_log.py`

### 4-7. PDF 생성

사용 상황:

- 자동입력과 별개로 탭4 PDF 최종본만 따로 생성하고 싶을 때

핵심 파일:

- `tab4_pdf_final_gui.py`

---

## 5. 권장 업무 순서

일반적인 대기 업무 흐름은 다음 순서로 이해하면 된다.

1. 성적서 파일 준비
2. 런처에서 **측정인 자동입력**으로 사이트 입력, 백데이터, PDF 처리
3. 런처에서 **측정인 검토**로 사이트와 엑셀 비교
4. 런처에서 **성적서 검토**로 파일 자체 검토
5. 필요 시 **PDF 생성**으로 최종 PDF 생성
6. 일자 단위 마감 검토 시 **발송대장 검토**
7. 전체 현황 요약이 필요하면 **종합 검토**

수질은 런처 **측정인 자동입력** GUI에서 매체 **수질**을 선택해 처리한다. 상세 표·셀렉터는 **§4-1-1** 참고.

- 사이트: [현장측정분석(수질)](https://측정인.kr/ms/field_water.do)
- 성적서 검색·파일 추가: `WATER_REPORT_INPUT` (`...\14.수질성적서\0.입력중`)
- 탭2: PE / 용량·개수 1 / 폭기조 → `#btnSaveMsFieldDoc`
- 탭4: 매크로 엑셀 → 날짜 H6/H7/J7, RealGrid 4열부터 A~M, PDF `#anzeFile1/2` 업로드
- 병합 PDF: `PDF_WATER\{연도}년\{대행기록부 D3 공백제거}\{엑셀파일명}.pdf`
- 대행기록부만: `WATER_RECORD_PDF_DIR` (기본 `\\...\scan\PDF`)

---

## 6. 전체 구조 한눈에 보기

```text
[프로그램 루트]
  ├─ 2.검토 및 입력프로그램.pyw   ← 더블클릭 진입점
  ├─ config.ini                  ← 경로·URL 설정 (선택)
  └─ 3.py\                       ← Python 소스 전부
        ├─ config.py, eco_input.py, …
        └─ *_gui.py, *_utils.py

[사용자 실행]
  └─ 2.검토 및 입력프로그램.pyw (통합 런처)
      └─ 3.py\ 안의 선택 도구 GUI 실행
          └─ 핵심 로직 호출
              ├─ Selenium으로 측정인.kr 조작
              ├─ openpyxl/Excel COM으로 엑셀 읽기
              ├─ NAS에서 파일 검색/저장
              └─ 결과 엑셀/PDF/백데이터 생성
```

구조를 계층으로 나누면 다음과 같다.

### 6-1. 실행·GUI 계층

- `2.검토 및 입력프로그램.pyw` — 7개 도구 통합 런처 (프로그램 루트)
- `3.py\` — 아래 GUI·로직·유틸 Python 파일 전부
- `eco_input_gui.py` (3.py 안)
- `eco_check_gui.py`
- `report_check_gui.py`
- `receipt.py`
- `dash.py`
- `Vehicle_operation_log.py`
- `tab4_pdf_final_gui.py`

이 중 다음 파일은 GUI와 로직이 한 파일에 합쳐져 있다.

- `receipt.py`
- `dash.py`
- `Vehicle_operation_log.py`
- `tab4_pdf_final_gui.py`

### 6-2. 핵심 로직 계층

- `eco_input.py`
- `eco_check.py`
- `report_check.py`
- `backdata_utils.py`
- `tab4_utils.py`
- `water_input_utils.py`
- `measin_constants.py`

### 6-3. 공통 유틸리티 계층

- 사이트 자동화: `measin_utils.py`, `selenium_utils.py`, `realgrid_utils.py`, `select2_utils.py`
- 엑셀/PDF 처리: `excel_utils.py`, `excel_com_utils.py`, `pdf_utils.py`
- 데이터/형식 변환: `data_utils.py`, `format_utils.py`
- 파일/NAS 검색: `file_utils.py`
- GUI 공통: `gui_common.py`, `worker_utils.py`
- 로깅: `log_utils.py`

---

## 7. 주요 데이터 흐름

### 7-1. 대기 자동입력 흐름

```text
NAS 성적서(.xlsm)
  ↓
eco_input.py
  ├─ 탭1 현장측정정보 입력 (F10 측정용도 + 항목/인력/차량/장비 — §4-1-2 슬래시 정규화)
  ├─ 탭2 시료채취/측정정보 입력
  ├─ 백데이터 생성(수분/THC)
  ├─ 탭2 PDF 업로드
  └─ 탭4 측정분석결과 입력
      ↓
측정인.kr 반영
      ↓
최종 PDF 저장
```

### 7-2. 수질 자동입력 흐름

```text
NAS 수질 성적서 (0.입력중)              매크로: 측정인 측정분석 입력 26.05(수질).xlsm
  ↓                                              ↓
eco_input._main_water() + water_input_utils.py
  ├─ 로그인 → 시료번호만 목록 검색 (field_water.do, 날짜 범위 검색 없음)
  ├─ [옵션] 탭2: PE/1/1/폭기조 → #btnSaveMsFieldDoc
  └─ [옵션] 탭4: a#ui-id-4
      ├─ 매크로 읽기 (A6, H6/H7/J7, 8행 헤더·9~54행 A~M)
      ├─ 날짜 필드 입력 (#smpl_rcpt_dt, #anze_start_dt, #anze_end_dt)
      ├─ RealGrid: 4열(측정항목)부터 A~M 붙여넣기 (#gridAnalySampAnzeDataAirItemList1)
      ├─ [PDF ON] 성적서 시트 → PDF 2종 생성
      │     ├─ 병합 → PDF_WATER\{연도}년\{대행기록부 D3 공백제거}\{엑셀명}.pdf
      │     ├─ 대행만 → WATER_RECORD_PDF_DIR\{엑셀명}.pdf  (scan\PDF 등)
      │     ├─ 업로드 #anzeFile1, #anzeFile2
      │     └─ #btnCompSave (확인 2회)
      └─ [PDF OFF] #btnTempSave
  ↓
목록 복귀 → 다음 시료
```

### 7-3. 백데이터 생성 흐름

```text
성적서 엑셀
  ↓
backdata_utils.py
  ├─ 수분량 시트 → CSV
  └─ THC 시트 → CSV 또는 FID
      ↓
NAS 저장
```

중요한 동작 규칙:

- 수분량 백데이터는 `입력` 시트 `B23` 시작시간 기준으로 종료 예상 시점이 지나야 생성된다.
- THC 백데이터는 `입력(분석값)` 시트 `K64` 기준 종료 시점이 지나야 생성된다.
- 아직 종료 전이면 파일을 만들지 않고 다음과 같은 로그가 나온다.
  - `⚠ 수분 백데이터 생성 스킵 ...`
  - `⚠ THC 백데이터 생성 스킵 ...`

이 동작은 실제 측정 종료 전 백데이터가 먼저 만들어지는 문제를 방지하기 위한 것이다.

---

## 8. 주요 경로와 저장 위치

프로그램은 `config.ini` 파일에 설정된 NAS 경로를 사용한다. 따라서 서버 IP나 폴더 구조가 바뀌면 `config.ini` 파일만 수정하면 모든 프로그램에 즉시 반영된다.

**주의: `config.ini` 파일은 보안상 깃허브에 올라가지 않으므로, 프로그램 폴더에 해당 파일이 없다면 `config.py`의 기본값을 확인하거나 새로 작성해야 한다.**

대표 경로 예시(config.ini 기본값 기준):

- 성적서 기본 NAS: `\\192.168.10.163\측정팀\2.성적서` (REPORT_BASE)
- 수질 성적서 NAS: `\\192.168.10.163\측정팀\2.성적서\14.수질성적서` (WATER_REPORT_BASE)
- 수질 입력중(자동입력·파일추가): `...\14.수질성적서\0.입력중` (WATER_REPORT_INPUT)
- 대기 PDF 저장 기본 경로: `\\192.168.10.163\측정팀\2.성적서\0.PDF\2.대기pdf` (PDF_AIR)
- 수질 PDF 병합 저장: `\\192.168.10.163\측정팀\2.성적서\0.PDF\1.수질pdf` (PDF_WATER) — 하위 `{연도}년\{대행기록부 D3 공백제거}\`
- 수질 대행기록부만: `\\192.168.10.163\scan\PDF` (WATER_RECORD_PDF_DIR, 없으면 자동 생성)
- 수질 매크로 엑셀: `WATER_TAB4_MACRO_FILE`
- 최종완료 경로: `\\192.168.10.163\측정팀\2.성적서\0 5.최종완료` (REPORT_DONE)
- 수분량 저장 경로: `\\192.168.10.163\측정팀\2.성적서\0.수분량` (MOISTURE_ROOT)
- THC 저장 경로: `\\192.168.10.163\측정팀\2.성적서\0.THC` (THC_ROOT)
- 공용 오류 로그: `\\192.168.10.163\측정팀\10.검토\_logs\error_log.txt` (LOG_DIR)

---

## 9. 시료번호 규칙

시료번호는 여러 모듈에서 공통으로 사용하는 핵심 키다.

형식:

- `{매체}{YYMMDD}{팀번호}-{순번}`

예시:

- `A2604181-03` (대기)
- `W2606300-01` (수질)

의미:

- `A`: 대기
- `26`: 2026년
- `0418`: 4월 18일
- `1`: 1팀
- `03`: 세 번째 시료

매체 구분:

- `A`: 대기
- `W`: 수질

이 규칙은 다음 모듈들이 공통으로 사용한다.

- `data_utils.py`
- `file_utils.py`
- `excel_utils.py`
- `eco_input.py`
- `eco_check.py`
- `receipt.py`

---

## 10. 초보자가 자주 헷갈리는 점

### 10-1. 자동입력 GUI와 콘솔 실행 방식이 다르다

**측정인 자동입력** (`eco_input_gui.py` → `eco_input.py`):

- GUI는 `_build_answers_air()` / `_build_answers_water()`가 만든 **dict**를 `_main_air()` / `_main_water()`에 **키워드 인자**로 넘긴다. (`fake_input` / `input()` 가로채기 없음)
- GUI 종료 시 `input("엔터...")` 는 호출하지 않음 (`login_id` 전달 = gui_mode). 콘솔에서만 엔터 대기.
- 콘솔에서 `eco_input.py`만 직접 실행하면 예전처럼 `input()`·`ask_yesno()` 순서로 질문한다.
- GUI에서 체크박스 순서를 바꿔도 로직에는 **bool 이름**만 전달되므로, 콘솔 질문 순서와는 무관하다. 옵션을 새로 추가·삭제할 때는 GUI dict와 `_main_*()` 인자를 함께 맞춘다.

**측정인 검토** (`eco_check_gui.py` → `eco_check.py`):

- 아직 `input()` 가로채기(fake input) 패턴을 쓴다. `eco_check.py` 질문 순서가 바뀌면 GUI answer 배열도 같이 수정해야 한다.

### 10-2. 드래그앤드롭은 선택 기능이다

`eco_input_gui.py`와 `report_check_gui.py`는 `tkinterdnd2`가 설치되어 있으면 드래그앤드롭을 지원한다.
설치되지 않았거나 DLL 충돌이 있으면 기본 Tk 모드로 돌아간다.

즉, 드래그앤드롭이 안 돼도 프로그램 자체가 망가진 것은 아니다.

### 10-3. Excel이 숨어서 동작할 수 있다

일부 작업은 `win32com`으로 Excel을 백그라운드에서 띄운다.
PDF 생성, 표시값 복사, 백데이터 생성이 대표적이다.

그래서 Excel 파일이 잠겨 있거나 다른 경고창이 떠 있으면 실패할 수 있다.

### 10-4. RealGrid는 일반 HTML 표가 아니다

측정인.kr의 표는 일반 테이블이 아니라 RealGrid라서 Selenium으로 단순 입력이 안 되는 구간이 있다.
이 때문에 `realgrid_utils.py`에서 JavaScript API를 직접 호출한다.

---

## 11. 문제 생기면 먼저 볼 곳

### 11-1. 공용 오류 로그

- `\\192.168.10.163\측정팀\10.검토\_logs\error_log.txt`

예외가 잡히면 `log_utils.py`를 통해 여기에 남는다.

### 11-2. GUI 안의 로그 창

대부분의 GUI는 화면 아래 또는 오른쪽에 로그 창이 있다.
콘솔 출력이 로그 창으로 전달되므로, 실제 실행 순서를 보려면 이 로그를 보면 된다.

### 11-3. 자주 발생하는 원인

- Chrome 또는 ChromeDriver 버전 문제
- Python 패키지 누락
- NAS 경로 접근 실패
- Excel 보안 경고 또는 잠금
- 시트명 변경
- 사이트 CSS 셀렉터 변경

---

## 12. 경로나 사이트가 바뀌었을 때 어디를 고쳐야 하는지

### NAS 경로가 바뀌었을 때

**먼저** 프로그램 루트 `config.ini` 의 `[PATHS]` 를 수정한다. (없으면 `3.py/config.py` 기본값)

그다음 참고:

- `eco_input.py`, `eco_check.py`, `backdata_utils.py`, `tab4_utils.py`, `water_input_utils.py`
- `receipt.py`, `dash.py`

수질 전용 키: `WATER_REPORT_*`, `PDF_WATER`, `WATER_RECORD_PDF_DIR`, `WATER_TAB4_*`

### 측정인.kr 사이트 구조가 바뀌었을 때

주로 아래 파일을 본다.

- `measin_constants.py`
- `eco_input.py`, `water_input_utils.py` (수질 탭·PDF input id)
- `eco_check.py`
- `tab4_utils.py`
- `selenium_utils.py`

수질 PDF input id: `config.ini` → `WATER_TAB4_FILE1` / `WATER_TAB4_FILE2` 및 `eco_input.py` `WATER_FILE_BTN1/2`

### 엑셀 시트명이 바뀌었을 때

주로 아래 파일을 본다.

- `excel_utils.py`
- `eco_input.py`
- `backdata_utils.py`
- `tab4_pdf_final_gui.py`

---

## 13. 문서별 역할

이 폴더의 주요 문서는 역할이 다르다.

| 문서 | 용도 |
|------|------|
| `readme.md` | 처음 사용하는 사람을 위한 전체 안내 |
| `파일별_설명.md` | 개발자/유지보수 담당자를 위한 파일별 상세 설명 |
| `모듈화_분석.md` | 리팩터링 및 모듈화 작업 기록 |
| `버전노트.txt` | 버전별 변경 이력 (도구별·런처별) |

### 버전노트는 이렇게 쓴다

예전에는 `.bat` 파일 이름에 버전(`V2.9` 등)을 붙였다. 런처 통합 이후에는 아래처럼 나누어 `버전노트.txt`에만 기록한다.

| 구분 | 기록 예시 | 언제 쓰나 |
|------|-----------|-----------|
| **통합 런처** | `0.검토 및 입력프로그램` | 런처 UI, 실행 방식, 진입점 변경 |
| **개별 도구** | `1.측정인 자동입력 V2.10` | `eco_input.py` 등 해당 도구 로직 변경 |
| **설치** | `0.처음사용시` (필요 시) | 패키지·드라이버·보안 설정 변경 |

- **파일 이름에는 버전을 붙이지 않는다.** (`eco_input_gui.py` 그대로)
- **버전 번호는 `버전노트.txt` 한 곳에서만 올린다.**
- 한 번에 여러 도구를 고쳤으면 도구마다 줄을 나눠 적는다.
- 런처만 고친 날은 개별 도구 버전을 올릴 필요 없다.

---

## 14. 빠르게 기억해야 할 핵심만 정리

- 처음 PC면 `0.처음사용시`의 통합 설치부터 실행
- 실행은 **`2.검토 및 입력프로그램.pyw`** 더블클릭 → 원하는 버튼 선택
- 자동입력·성적서 검토 등은 런처 버튼으로 실행
- 경로/IP 변경은 루트의 `config.ini` 파일에서 한 번에 수정 가능
- 에러는 먼저 공용 로그와 GUI 로그창 확인
- 백데이터는 측정 종료 전이면 일부러 생성되지 않음
