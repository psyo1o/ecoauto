# -*- coding: utf-8 -*-
"""
측정인.kr 자동 입력 – 팀별 시료번호 자동 처리 / 비산먼지 전용 로직 포함(최종)
요청된 부분만 수정하고 기존 안정동작 부분은 전혀 손대지 않음.
"""
import os
import time
import re
import traceback
import pyperclip
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
import pythoncom
import win32com.client as win32
import win32gui, win32con
from pywinauto.application import Application
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
from selenium.common.exceptions import (
    UnexpectedAlertPresentException,
    NoAlertPresentException,
    TimeoutException,
)
import datetime as dt
from excel_utils import find_sheet_by_candidates, parse_measuring_record

#===================경고 제거========================
# ✅ openpyxl 조건부서식 경고 숨김 (오류 아님, 콘솔 정리용)
import warnings
warnings.filterwarnings(
    "ignore",
    category=UserWarning,
    message="Conditional Formatting extension is not supported and will be removed"
)

# ============================================================
# 공통 유틸 모듈 (모듈화)
# ============================================================
from selenium_utils import (
    init_driver, safe_click, wait_el, set_date_js,
    accept_all_alerts as _accept_all_alerts,
    accept_all_alerts as _accept_alerts_base,
    close_popup
)
from measin_utils import (
    login, search_date, wait_grid_loaded,
    open_sample_detail as _open_sample_detail,
    go_back_to_list as _go_back_to_list,
    collect_samples_from_files,
    LOGIN_URL, FIELD_URL, NAS_BASE, NAS_DIRS
)
from format_utils import format_time as trim_hm, to_f1, to_f2
from data_utils import sample_to_datestr
from realgrid_utils import (
    rg_dump_headers, rg_find_col, rg_find_cols, rg_get_body, rg_api_write_data, 
    rg_scroll_top, rg_find_tr_by_item, rg_paste_to_tr, rg_paste_to_tr_tab4, rg_set_cell_by_keys
)
from select2_utils import Select2Handler as _Select2Handler
from pdf_utils import merge_pdfs as _merge_pdfs, PDFExporter as _PDFExporter
from excel_utils import find_sheet_by_candidates as _find_sheet_by_candidates
from file_utils import find_excel_for_sample as _find_excel_util
from excel_com_utils import get_excel_app, kill_excel_app
from log_utils import log_error

# =====================================================================
# 설정
# =====================================================================


# PDF 임시 생성 폴더(로컬 권장: 업로드 창에서 경로 인식/권한 문제 줄어듦)
PDF_TMP_DIR = r"C:\measin_upload_tmp"

# 백데이터 저장 루트
MOIST_ROOT = r"\\192.168.10.163\측정팀\2.성적서\0.수분량"
THC_ROOT   = r"\\192.168.10.163\측정팀\2.성적서\0.THC"

ANZE_XLSM = r"\\192.168.10.163\측정팀\2.성적서\측정인 측정분석 입력 26.01.xlsm"
ANZE_SHEET = "00. 측정분석결과 입력샘플"

TAB4_SELECTOR = "#ui-id-4"
TAB4_GRID_ROOT = "#gridAnalySampAnzeDataAirItemList1"

TAB4_FILE_BTN1 = "#newFile1"   # 시험 분석일지(PDF) 업로드 버튼(열기창)
TAB4_FILE_BTN2 = "#newFile2"   # 측정기록부(PDF) 업로드 버튼(열기창)

FINAL_DONE_DIR = r"\\192.168.10.163\측정팀\2.성적서\0 5.최종완료"
PDF_BASE_DIR   = r"\\192.168.10.163\측정팀\2.성적서\0.PDF\2.대기pdf"

# ── 수질 ──────────────────────────────────────────────────
FIELD_URL_WATER    = "https://측정인.kr/ms/field_outwater.do"
# 수질 성적서: NAS_BASE_WATER\YYYY년\업체명\파일.xlsm
NAS_BASE_WATER     = r"\\192.168.10.163\측정팀\2.성적서\14.수질성적서"
# 수질 PDF:    PDF_BASE_DIR_WATER\YYYY년\업체명\파일.pdf
PDF_BASE_DIR_WATER = r"\\192.168.10.163\측정팀\2.성적서\0.PDF\1.수질pdf"
# 수질 탭4 PDF 업로드 셀렉터
WATER_FILE_BTN1    = "#anzeFile1"   # 분석일지
WATER_FILE_BTN2    = "#anzeFile2"   # 대행기록부
# 수질 엑셀 시트명
WATER_SHEET_ANALY  = "분석일지"
WATER_SHEET_RECORD = "대행기록부"

# ======================== 공통 import ========================
import atexit
import win32clipboard  # 기존 로직 유지용

# ======================== Excel COM 재사용(전역) ========================
_EXCEL_APP = None

# Excel 상수(Win32 상수값)
XL_UP = -4162
XL_TOLEFT = -4159



# ========================측정분석 입력 엑셀 ========================
def wait_until_sheet_updates(ws, timeout=10.0):
    """
    A6 입력 후 표가 갱신될 때까지 대기.
    기준: A9 텍스트가 빈칸이 아니고, 이전 값과 달라지는 시점.
    """
    end = time.time() + timeout
    before = str(ws.Range("A9").Text).strip()

    while time.time() < end:
        try:
            ws.Parent.Application.Calculate()
        except:
            pass

        cur = str(ws.Range("A9").Text).strip()
        if cur and cur != before:
            return True
        time.sleep(0.2)

    return False

def _cell_text(ws, addr: str) -> str:
    try:
        return str(ws.Range(addr).Text).strip()
    except:
        v = ws.Range(addr).Value
        return "" if v is None else str(v).strip()


def read_tab4_from_macro_xlsm(sample_no: str) -> dict:
    excel = get_excel_app()
    wb = None
    try:
        # ✅ 이벤트/경고 설정 (Calculation은 건드리지 말자 - 너 에러났음)
        excel.DisplayAlerts = False
        excel.EnableEvents = True
        try:
            excel.AutomationSecurity = 1  # msoAutomationSecurityLow
        except:
            pass

        wb = excel.Workbooks.Open(ANZE_XLSM, ReadOnly=True, UpdateLinks=0)
        ws = wb.Worksheets(ANZE_SHEET)

        # ✅ A6 입력
        ws.Range("A6").Value = sample_no

        # ✅ 계산/갱신 트리거
        try:
            excel.CalculateFullRebuild()
        except:
            try:
                excel.Calculate()
            except:
                pass

        # ✅ 표 갱신 대기 (없으면 rows가 0으로 나옴)
        ok = wait_until_sheet_updates(ws, timeout=10.0)

        # 날짜/시간 (Text로)
        rcpt_dt = str(ws.Range("I6").Text).strip()
        start_src = str(ws.Range("I7").Text).strip()
        end_dt = str(ws.Range("K7").Text).strip()

        # ✅ A9:M54 읽기 - 숨김행 스킵
        rows = []
        cols = "ABCDEFGHIJKLM"

        for excel_row in range(9, 55):
            try:
                if ws.Rows(excel_row).Hidden:
                    continue
            except:
                pass

            row_vals = [_cell_text(ws, f"{c}{excel_row}") for c in cols]  # ✅ A..M 모두 Text

            item = row_vals[0].strip()
            if not item:
                continue

            rows.append(row_vals)


        # 탭4 엑셀 행 수만 로그
        print(f"[TAB4-EXCEL] rows={len(rows)}")
        return {"rcpt_dt": rcpt_dt, "start_src": start_src, "end_dt": end_dt, "rows": rows}


    finally:
        try: wb.Close(SaveChanges=False)
        except: pass
        ws_out = None
        wb = None


# ======================== 클립보드 유틸(공통) ========================
# ======================== 클립보드 유틸(공통) 및 파일 저장 안전망 ========================

def _clipboard_clear():
    """클립보드 충돌 방지를 위한 재시도 로직 적용"""
    for _ in range(4):
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
            return
        except:
            try: win32clipboard.CloseClipboard()
            except: pass
            time.sleep(0.15)

def _clipboard_get_unicode_text():
    """클립보드 읽기 재시도 로직 적용"""
    for _ in range(4):
        try:
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                    return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_TEXT):
                    b = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
                    return b.decode("cp949", errors="ignore")
                return ""
            finally:
                win32clipboard.CloseClipboard()
        except:
            time.sleep(0.15)
    return ""

def _copy_range_and_get_text(excel_app, rng, wait_max=3.0, retries=4):
    """✅ Range.Copy() → 클립보드 텍스트를 폴링해서 가져옴(재시도/에러 무시 로직 대폭 강화)"""
    for attempt in range(retries):
        try: excel_app.CutCopyMode = False
        except: pass
        _clipboard_clear()

        # 복사 시도
        try:
            rng.Copy()
        except:
            time.sleep(0.3)
            try: rng.Copy()
            except: pass

        # 클립보드 폴링
        end = time.time() + float(wait_max)
        txt = ""
        while time.time() < end:
            time.sleep(0.1)
            txt = _clipboard_get_unicode_text()
            if txt and txt.strip():
                break

        # 성공했으면 복사모드 해제 후 리턴
        if txt and txt.strip():
            try: excel_app.CutCopyMode = False
            except: pass
            _clipboard_clear()
            return txt

        # 실패 시 약간 대기 후 다음 시도(attempt)
        time.sleep(0.5)

    try: excel_app.CutCopyMode = False
    except: pass
    _clipboard_clear()
    return ""

def _safe_write_file(path, text, encoding="utf-8-sig", newline="", retries=4):
    """네트워크 드라이브(NAS) 접근 지연 및 권한 오류 방지용 안전 쓰기"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    for attempt in range(retries):
        try:
            with open(path, "w", encoding=encoding, newline=newline) as f:
                f.write(text)
            return True
        except Exception as e:
            if attempt == retries - 1:
                raise RuntimeError(f"파일 저장 4회 시도 실패 ({path}): {e}")
            time.sleep(0.5)  # NAS 지연 대기 후 재시도
    return False

def _year_month_folder_from_sample(sample_no: str):
    """
    sample_no: AYYMMDDT-XX 에서
    YYYY, 'M월' 폴더명 반환
    """
    ds = sample_to_datestr(sample_no)  # 너 코드에 이미 있음
    if not ds:
        return None, None
    yyyy, mm, _ = ds.split("-")
    mm_int = int(mm)
    return yyyy, f"{mm_int}월"

# ======================== 수분 CSV(표시값 그대로) ========================
def _export_moist_csv_from_open_ws(excel_app, ws, out_csv_path: str, max_rows=None):
    """
    ✅ '열린 ws(엑셀 COM)'에서 표시값을 복사해서 CSV로 저장
    - 날짜/시간/소수점 등 표시 서식 유지
    - 탭 → 콤마
    """
    os.makedirs(os.path.dirname(out_csv_path), exist_ok=True)

    # 마지막 행/열(값 기준, 빠름)
    last_row = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    last_col = ws.Cells(1, ws.Columns.Count).End(XL_TOLEFT).Column

    if max_rows is not None:
        last_row = min(int(max_rows), int(last_row))

    if last_row < 1:
        last_row = 1
    if last_col < 1:
        last_col = 1

    rng = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, last_col))
    txt = _copy_range_and_get_text(excel_app, rng, wait_max=3.0)

    if not txt or not txt.strip():
        raise RuntimeError("클립보드에서 수분 데이터를 복사해오지 못했습니다. (엑셀 응답 없음)")
        
    # 탭 → 콤마
    txt = txt.replace("\t", ",")
    # 줄바꿈 정리(윈도우 CRLF 유지)
    txt = txt.replace("\r\n", "\n").replace("\r", "\n")
    txt = "\r\n".join(txt.split("\n"))

    _safe_write_file(out_csv_path, txt, encoding="utf-8-sig", newline="")

    return out_csv_path


def export_csv_display_as_is(excel_path: str, sheet_name: str, out_csv_path: str, max_rows=None):
    """
    ✅ (외부에서도 쓸 수 있게 유지) 엑셀 파일 열어서 표시값 CSV 저장
    - 내부적으로는 Excel App 재사용
    """
    excel = get_excel_app()
    wb = None
    try:
        wb = excel.Workbooks.Open(excel_path, ReadOnly=True, UpdateLinks=0)
        ws = wb.Worksheets(sheet_name)
        return _export_moist_csv_from_open_ws(excel, ws, out_csv_path, max_rows=max_rows)
    finally:
        try:
            if wb is not None:
                try:
                    excel.CutCopyMode = False
                except:
                    pass
                wb.Close(False)
        except:
            pass


# ======================== THC PF → FID(150행 고정, 복사 기반) ========================
def _export_pf_fid_from_open_ws(excel_app, ws, out_fid_path: str, fixed_rows=150):
    """
    ✅ 열린 ws(엑셀 COM)에서 150행 고정으로 복사한 텍스트를 .FID로 저장
    """
    os.makedirs(os.path.dirname(out_fid_path), exist_ok=True)

    last_col = ws.UsedRange.Columns.Count
    if not last_col or last_col < 1:
        last_col = 1

    rng = ws.Range(ws.Cells(1, 1), ws.Cells(int(fixed_rows), int(last_col)))
    txt = _copy_range_and_get_text(excel_app, rng, wait_max=3.0)

    if not txt or not txt.strip():
        raise RuntimeError("클립보드에서 THC(PF) 데이터를 복사해오지 못했습니다.")

    # 안전 저장 유틸 사용
    _safe_write_file(out_fid_path, txt, encoding="utf-8-sig", newline="")

    return out_fid_path


def export_fid_by_excel_copy(excel_path: str, sheet_name: str, out_fid_path: str):
    """
    ✅ (기존 시그니처 유지) 엑셀 열어서 PF 시트 복사→FID 저장
    - 내부적으로 Excel App 재사용
    """
    excel = get_excel_app()
    wb = None
    try:
        wb = excel.Workbooks.Open(excel_path, ReadOnly=True, UpdateLinks=0)
        ws = wb.Worksheets(sheet_name)
        return _export_pf_fid_from_open_ws(excel, ws, out_fid_path, fixed_rows=150)
    finally:
        try:
            if wb is not None:
                try:
                    excel.CutCopyMode = False
                except:
                    pass
                wb.Close(False)
        except:
            pass


# ======================== THC CSV(openpyxl) - 너 기존 그대로 ========================
def _safe_str(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    return str(v)


def _used_range_bounds(ws):
    max_row = 1
    max_col = 1
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            if cell.value is not None and str(cell.value).strip() != "":
                if cell.row > max_row:
                    max_row = cell.row
                if cell.column > max_col:
                    max_col = cell.column
    return max_row, max_col


def _export_ws_to_csv(ws, out_csv_path, encoding="utf-8-sig"):
    last_r, last_c = _used_range_bounds(ws)
    os.makedirs(os.path.dirname(out_csv_path), exist_ok=True)

    def q(s):
        s = _safe_str(s)
        if any(ch in s for ch in [",", '"', "\n", "\r"]):
            s = s.replace('"', '""')
            return f'"{s}"'
        return s

    lines = []
    for r in range(1, last_r + 1):
        row_vals = [q(ws.cell(r, c).value) for c in range(1, last_c + 1)]
        while row_vals and row_vals[-1] == "":
            row_vals.pop()
        lines.append(",".join(row_vals).rstrip())

    while lines and lines[-1].strip() == "":
        lines.pop()

    text_data = "\n".join(lines)
    _safe_write_file(out_csv_path, text_data, encoding=encoding, newline="\n")


# ======================== export_backdata_moist_thc (최적화 적용) ========================
def export_backdata_moist_thc(excel_path: str, sample_no: str):
    """
    ✅ 규칙 그대로 + 최적화:
    - openpyxl로 조건 판단
    - 수분/THC PF(FID) 둘 중 하나라도 필요하면
      Excel 파일을 "한 번만 Open → 필요한 시트들 처리 → Close"
    - Excel 앱은 전역 1개 재사용(프로그램 끝에서 close_excel_app()로 닫기)
    """
    # ✅ 비산먼지 스킵
    if ("비산먼지" in os.path.basename(excel_path)) or ("비산" in os.path.basename(excel_path)):
        print("▶ 비산먼지 성적서 → 백데이터 스킵")
        return

    yyyy, mm_folder = _year_month_folder_from_sample(sample_no)  # 너 기존 함수 그대로 사용
    if not yyyy:
        print(f"⚠ 백데이터: 시료번호 날짜 파싱 실패 → {sample_no}")
        return
    print(f"▶백데이터: 시료번호 {sample_no} 진행중")
    # ---- 1) openpyxl로 조건 판단(빠름) ----
    wb = load_workbook(excel_path, data_only=True)
    
    # 수분 조건
    need_moist = False
    try:
        ws_in = wb["입력"] if "입력" in wb.sheetnames else None
        if ws_in is not None:
            b18 = "" if ws_in["B18"].value is None else str(ws_in["B18"].value).strip()
            need_moist = (b18 == "사용") and ("수분량자동측정" in wb.sheetnames)
    except:
        need_moist = False

    # THC 조건
    ws_an = None
    if "입력(분석값)" in wb.sheetnames:
        ws_an = wb["입력(분석값)"]
    else:
        for n in wb.sheetnames:
            if n.replace(" ", "") == "입력(분석값)".replace(" ", ""):
                ws_an = wb[n]
                break

    need_thc = False
    is_fid_mode = False
    if ws_an is not None:
        a64 = "" if ws_an["A64"].value is None else str(ws_an["A64"].value).strip()
        if a64 != "":
            need_thc = True
            c65 = ws_an["C65"].value
            try:
                is_fid_mode = (int(float(c65)) == 1)
            except:
                is_fid_mode = (str(c65).strip() == "1")

    # ---- 2) 엑셀 COM은 "필요할 때만" 파일 1회 Open/Close ----
    excel = None
    wb_xl = None
    try:
        if need_moist or (need_thc and not is_fid_mode):
            excel = get_excel_app()
            wb_xl = excel.Workbooks.Open(excel_path, ReadOnly=True, UpdateLinks=0)

        # -----------------------
        # (A) 수분 CSV (표시값 그대로)
        # -----------------------
        try:
            if need_moist:
                out_dir = os.path.join(MOIST_ROOT, yyyy, mm_folder)   # 너 기존 상수 그대로
                out_csv = os.path.join(out_dir, f"{sample_no}.csv")

                ws_m_xl = wb_xl.Worksheets("수분량자동측정")
                _export_moist_csv_from_open_ws(excel, ws_m_xl, out_csv, max_rows=6)  # max_rows=6 유지
                print(f"✅ 수분 CSV 저장: {out_csv}")
            else:
                # 기존 로그 스타일 유지
                if ws_in is None:
                    print("⚠ 수분: '입력' 시트 없음 → 스킵")
                else:
                    b18 = "" if ws_in["B18"].value is None else str(ws_in["B18"].value).strip()
                    if b18 != "사용":
                        print("▶ 수분: 입력!B18 != '사용' → 스킵")
                    elif "수분량자동측정" not in wb.sheetnames:
                        print("⚠ 수분: '수분량자동측정' 시트 없음 → 스킵")
        except Exception as e:
            print(f"❌ 수분 백데이터 실패: {e}")

        # -----------------------
        # (B) THC
        # -----------------------
        try:
            if not need_thc:
                if ws_an is None:
                    print("⚠ THC: '입력(분석값)' 시트 없음 → 스킵")
                else:
                    print("▶ THC: 입력(분석값)!A64 빈칸 → 스킵")
                return

            out_dir = os.path.join(THC_ROOT, yyyy, mm_folder)  # 너 기존 상수 그대로
            os.makedirs(out_dir, exist_ok=True)

            if is_fid_mode:
                # C65==1 → THC 측정값(FID) → CSV (openpyxl 그대로)
                sheet_name = "THC 측정값(FID)"
                if sheet_name not in wb.sheetnames:
                    cand = [n for n in wb.sheetnames if n.replace(" ", "") == sheet_name.replace(" ", "")]
                    if cand:
                        sheet_name = cand[0]
                    else:
                        print(f"⚠ THC: 시트 없음({sheet_name}) → 스킵")
                        return

                ws_t = wb[sheet_name]
                out_csv = os.path.join(out_dir, f"{sample_no}.csv")
                _export_ws_to_csv(ws_t, out_csv)
                print(f"✅ THC(FID모드) CSV 저장: {out_csv}")

            else:
                # C65!=1 → THC 측정값(PF) → .FID (Excel 복사 기반, 150행 고정)
                sheet_name = "THC 측정값(PF)"
                # 시트명 보정(최소)
                try:
                    ws_pf_xl = wb_xl.Worksheets(sheet_name)
                except:
                    # 공백 제거 매칭
                    ws_pf_xl = None
                    for sh in wb_xl.Worksheets:
                        if str(sh.Name).replace(" ", "") == sheet_name.replace(" ", ""):
                            ws_pf_xl = sh
                            break
                    if ws_pf_xl is None:
                        print(f"⚠ THC: 시트 없음({sheet_name}) → 스킵")
                        return

                out_fid = os.path.join(out_dir, f"{sample_no}.FID")
                _export_pf_fid_from_open_ws(excel, ws_pf_xl, out_fid, fixed_rows=150)
                print(f"✅ THC(PF모드) FID 저장: {out_fid}")

        except Exception as e:
            print(f"❌ THC 백데이터 실패: {e}")

    finally:
        # ✅ 시료 1개당 workbook은 닫는다 (한번만 열고 닫자)
        try:
            if wb_xl is not None:
                try:
                    excel.CutCopyMode = False
                except:
                    pass
                wb_xl.Close(False)
        except:
            pass

    # ✅ Excel 앱은 전역 재사용이므로 여기서 Quit 안 함
    # 프로그램 끝에서 close_excel_app() 한 번 호출(또는 atexit 자동)



# =====================================================================
# RealGrid 예외 규칙(전역)
# =====================================================================

SKIP_VOL_AND_SPEED = {
    "매연", "황산화물", "질소산화물", "일산화탄소", "총탄화수소"
}

SKIP_SPEED_ONLY = {
    "먼지", "비소화합물", "수은화합물", "구리화합물", "아연화합물", "카드뮴화합물",
    "납화합물", "크로뮴화합물", "니켈화합물", "베릴륨화합물", "벤조(a)피렌"
}

def make_dust_per_item(excel_path: str) -> dict:
    """
    비산먼지용 per_item 생성
    - 항목은 그리드/엑셀에 무엇으로 있든, 결국 '비산먼지' 1개만 넣는 방식(가장 안전)
    """
    dv = read_dust_realgird_values(excel_path)  # 너가 이미 만든 함수
    return {
        "비산먼지": {
            "시작시간": dv.get("시작시간", ""),
            "종료시간": dv.get("종료시간", ""),
            "시료채취량": dv.get("시료채취량", ""),
            "흡인속도": dv.get("흡인속도", ""),
            "채취량단위": dv.get("채취량단위", "L"),
            "흡인속도단위": dv.get("흡인속도단위", "L/min"),
        }
    }


def rg2_fill_measure_grid_api(driver, sample_no: str, per_item: dict):
    """RealGrid API로 직접 입력 (js 모듈화 완료)"""
    date_str = sample_to_datestr(sample_no)
    if not date_str:
        raise RuntimeError("시료번호에서 날짜 파싱 실패")

    payload = []
    for item_name, v in per_item.items():
        st = v.get("시작시간", "")
        et = v.get("종료시간", "")
        vol = v.get("시료채취량", "")
        spd = v.get("흡인속도", "")
        vol_u = v.get("채취량단위", "L")
        spd_u = v.get("흡인속도단위", "L/min")

        if item_name in SKIP_VOL_AND_SPEED:
            vol, vol_u, spd, spd_u = "", "", "", ""
        elif item_name in SKIP_SPEED_ONLY:
            spd, spd_u = "", ""

        payload.append({
            "item": item_name, "sd": date_str, "st": st, "ed": date_str, "et": et,
            "vol": vol, "vol_u": vol_u, "spd": spd, "spd_u": spd_u,
        })

    # 길었던 JS 블록 삭제! 대신 유틸리티 함수 호출
    result = rg_api_write_data(driver, payload)
    
    print("▶ RealGrid API 입력 결과:", result)
    missing = result.get("missing")
    if missing:
        print("⚠ RealGrid에 없어서 미적용된 항목들:")
        for x in missing:
            print("   -", x)
    return result

def fill_tab2_realgird_dust_only(driver, excel_path: str, sample_no: str, grid_root_css: str):
    """
    ✅ 비산먼지 전용: RealGrid에서 '비산먼지' 행만 찾아서
       시료채취량/단위/흡인속도/단위/날짜/시간을 셀 하나씩 입력.
    - '먼지'로 대체하지 않음. 없으면 즉시 에러.
    """
    date_str = sample_to_datestr(sample_no)
    if not date_str:
        raise RuntimeError("시료번호에서 날짜 파싱 실패")

    dv = read_dust_realgird_values(excel_path)
    vol   = dv.get("시료채취량", "")
    vol_u = dv.get("채취량단위", "L")
    spd   = dv.get("흡인속도", "")
    spd_u = dv.get("흡인속도단위", "L/min")
    st    = dv.get("시작시간", "")
    et    = dv.get("종료시간", "")

    # 컬럼 찾기
    c_item = rg_find_col(driver, grid_root_css, ["측정항목"])
    c_vol  = rg_find_col(driver, grid_root_css, ["시료채취량"])
    c_spd  = rg_find_col(driver, grid_root_css, ["흡인속도"])
    c_sd   = rg_find_col(driver, grid_root_css, ["측정일(시작)"])
    c_st   = rg_find_col(driver, grid_root_css, ["시작시간"])
    c_ed   = rg_find_col(driver, grid_root_css, ["측정일(종료)"])
    c_et   = rg_find_col(driver, grid_root_css, ["종료시간"])

    unit_cols = rg_find_cols(driver, grid_root_css, "단위")
    c_vol_u = unit_cols[0] if len(unit_cols) >= 1 else (c_vol + 1 if c_vol else None)
    c_spd_u = unit_cols[1] if len(unit_cols) >= 2 else (c_spd + 1 if c_spd else None)

    need = [c_item, c_vol, c_vol_u, c_spd, c_spd_u, c_sd, c_st, c_ed, c_et]
    if any(x is None for x in need):
        raise RuntimeError(f"[비산먼지] 컬럼 탐색 실패: {rg_dump_headers(driver, grid_root_css)}")

    # ✅ '비산먼지' 행만 찾기 (fallback 없음)
    tr = rg_find_tr_by_item(driver, grid_root_css, c_item, "비산먼지")
    if not tr:
        # 디버그용: 현재 화면에 보이는 항목들만이라도 출력
        try:
            visible = rg_list_items(driver, grid_root_css, c_item)
        except:
            visible = "rg_list_items 실패"
        raise RuntimeError(f"[비산먼지] 그리드에서 '비산먼지' 행을 못 찾음. visible={visible}")

    # ✅ 셀 하나씩 입력 (커밋 이벤트 확실)
    rg_set_cell_by_keys(driver, tr, c_sd, date_str)
    rg_set_cell_by_keys(driver, tr, c_st, st)
    rg_set_cell_by_keys(driver, tr, c_ed, date_str)
    rg_set_cell_by_keys(driver, tr, c_et, et)

    #rg_set_cell_by_keys(driver, tr, c_vol, vol)
    #rg_set_cell_by_keys(driver, tr, c_vol_u, vol_u)

    print("▶ 비산먼지 RealGrid 입력 완료")


# eco_input.py 내의 기존 _norm_rg를 더 강력하게 수정
def _norm_rg(s: str) -> str:
    if s is None: return ""
    # 별표(*) 제거 및 모든 종류의 공백(nbsp 포함) 정리
    s = str(s).replace("*", "").replace("\xa0", " ").replace("\n", " ").strip()
    return " ".join(s.split())


def tab4_find_tr_by_item(driver, item_name: str, max_steps=600):
    """
    탭4 RealGrid 탐색(ArrowDown/ArrowUp 다중 이동):
    - 현재 화면의 보이는 행에서 먼저 탐색
    - 없으면 ArrowDown을 10칸씩 이동하며 탐색
    - 아래 끝 도달 시 ArrowUp으로 10칸씩 올라가며 재탐색
    """
    grid = driver.find_element(By.CSS_SELECTOR, TAB4_GRID_ROOT)
    body = grid.find_element(By.CSS_SELECTOR, "div.rg-body")
    target = _norm_rg(item_name)

    # 전체 페이지 기준으로 그리드를 먼저 중앙에 맞춤
    # (off-screen 상태에서 클릭 시 1칸 밀림 완화)
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'nearest'});", grid)
        time.sleep(0.08)
    except:
        pass

    def _scan_visible():
        trs = body.find_elements(By.CSS_SELECTOR, "table > tbody > tr")
        first_txt, last_txt = "", ""
        for idx, tr in enumerate(trs):
            try:
                cell = tr.find_element(By.CSS_SELECTOR, "td:nth-child(3) > div")
                txt = _norm_rg(cell.text)
                if idx == 0:
                    first_txt = txt
                last_txt = txt
                if txt == target:
                    return tr, first_txt, last_txt
            except:
                continue
        return None, first_txt, last_txt

    def _send_arrows(key, count):
        try:
            for _ in range(count):
                driver.switch_to.active_element.send_keys(key)
                time.sleep(0.02)
            return True
        except:
            try:
                for _ in range(count):
                    ActionChains(driver).send_keys(key).perform()
                    time.sleep(0.02)
                return True
            except:
                return False

    # 키 입력 포커스 확보
    try:
        first_tr = body.find_element(By.CSS_SELECTOR, "table > tbody > tr:first-child")
        first_cell = first_tr.find_element(By.CSS_SELECTOR, "td:nth-child(3) > div")
        ActionChains(driver).move_to_element(first_cell).click(first_cell).perform()
        time.sleep(0.05)
    except:
        try:
            ActionChains(driver).move_to_element(body).click(body).perform()
            time.sleep(0.05)
        except:
            pass

    # 0) 현재 화면 먼저 확인
    tr_now, _, _ = _scan_visible()
    if tr_now:
        return tr_now

    # 화면 변화가 일정 횟수 이상 없으면 끝으로 판단
    visible_rows = len(body.find_elements(By.CSS_SELECTOR, "table > tbody > tr"))
    stable_limit = max(8, visible_rows // 2)

    # 1) ArrowDown 15칸씩 탐색
    _, prev_first, prev_last = _scan_visible()
    down_stable = 0
    for _ in range(max_steps):
        if not _send_arrows(Keys.ARROW_DOWN, 15):
            break
        time.sleep(0.3)

        tr_found, cur_first, cur_last = _scan_visible()
        if tr_found:
            time.sleep(0.08)
            return tr_found

        if cur_first == prev_first and cur_last == prev_last:
            down_stable += 1
            if down_stable >= stable_limit:
                break
        else:
            down_stable = 0
            prev_first, prev_last = cur_first, cur_last

    # 2) ArrowUp 15칸씩 재탐색
    _, prev_first, prev_last = _scan_visible()
    up_stable = 0
    for _ in range(max_steps):
        if not _send_arrows(Keys.ARROW_UP, 15):
            break
        time.sleep(0.3)

        tr_found, cur_first, cur_last = _scan_visible()
        if tr_found:
            time.sleep(0.08)
            return tr_found

        if cur_first == prev_first and cur_last == prev_last:
            up_stable += 1
            if up_stable >= stable_limit:
                break
        else:
            up_stable = 0
            prev_first, prev_last = cur_first, cur_last

    return None

def tab4_paste_row_using_tab2(driver, tr, values_A_to_M):
    """
    values_A_to_M: 엑셀 A~M (13개)
    Tab4 전용 - rg_paste_to_tr_tab4 사용 (move_to_element 없음)
    """
    # 1차: 3열(측정항목)부터
    try:
        rg_paste_to_tr_tab4(driver, tr, start_col=3, values=values_A_to_M)
        return True
    except:
        pass

    # 2차: 4열부터(3열이 뭔가 막혔을 때 대비)
    try:
        rg_paste_to_tr_tab4(driver, tr, start_col=4, values=values_A_to_M)
        return True
    except:
        return False


def tab4_get_api_items(driver):
    """
    탭4 RealGrid API에서 전체 항목명 목록 추출.
    - 가능하면 DataSource 전체 row를 읽고
    - 실패 시 현재 화면 렌더 행만 fallback으로 수집
    """
    try:
        data = driver.execute_script(
            """
            const root = document.querySelector(arguments[0]);
            const gv = window.measGridViews && window.measGridViews.gridView1 ? window.measGridViews.gridView1 : null;
            if (!gv) return {count: 0, items: []};

            let items = [];
            try {
                const ds = (typeof gv.getDataSource === 'function') ? gv.getDataSource() : null;
                if (ds && typeof ds.getRowCount === 'function') {
                    const cnt = ds.getRowCount();
                    for (let i = 0; i < cnt; i++) {
                        let v = "";
                        try {
                            if (typeof ds.getValue === 'function') v = ds.getValue(i, 'anze_item');
                        } catch (e) {}
                        try {
                            if ((v === null || v === undefined || v === "") && typeof gv.getValue === 'function') {
                                v = gv.getValue(i, 'anze_item');
                            }
                        } catch (e) {}
                        items.push(v == null ? "" : String(v));
                    }
                    return {count: cnt, items};
                }
            } catch (e) {}

            // fallback: 현재 화면에 보이는 항목만
            try {
                if (!root) return {count: 0, items: []};
                const cells = root.querySelectorAll('div.rg-body table > tbody > tr td:nth-child(3) > div');
                const arr = Array.from(cells).map(el => (el.innerText || '').trim());
                return {count: arr.length, items: arr};
            } catch (e) {
                return {count: 0, items: []};
            }
            """,
            TAB4_GRID_ROOT,
        )
        if isinstance(data, dict):
            return data
    except:
        pass
    return {"count": 0, "items": []}



def fill_tab4_grid_only(driver, table_rows):
    """
    table_rows: 엑셀에서 읽은 [A..M] 리스트들
    - A: 항목
    - B~M: 탭4에 붙여넣을 데이터(열 순서가 탭4와 동일하게 맞춰져 있음)
    - 숨김행은 이미 read_tab4...에서 걸러진 상태라고 가정
    """
    # 엑셀 항목 정리
    rows_to_input = []
    for r in table_rows:
        item = _norm_rg(r[0])
        if not item:
            continue

        values = [("" if v is None else str(v).strip()) for v in r]  # ✅ A~M
        if all(v == "" for v in values):
            continue

        rows_to_input.append((item, values))

    # API 항목 수집
    api_data = tab4_get_api_items(driver)
    api_items_raw = api_data.get("items", []) if isinstance(api_data, dict) else []
    api_count = int(api_data.get("count", len(api_items_raw))) if isinstance(api_data, dict) else len(api_items_raw)
    excel_count = len(rows_to_input)

    print(f"[탭4-카운트] API 항목 {api_count}개 / 엑셀 항목 {excel_count}개")

    ok_count = 0
    failed = []
    fail_reason = []

    for idx, (item, values) in enumerate(rows_to_input, start=1):
        paste_ok = False
        last_error = None

        # 최대 3회 재시도
        for attempt in range(3):
            tr = tab4_find_tr_by_item(driver, item, max_steps=2000)
            if not tr:
                last_error = "행탐색실패"
                continue

            # 찾은 TR을 직접 클릭해서 선택
            cell = None
            try:
                cell = tr.find_element(By.CSS_SELECTOR, "td:nth-child(3) > div")
                ActionChains(driver).move_to_element(cell).click(cell).perform()
                time.sleep(0.15)
            except:
                last_error = "클릭실패"
                continue

            # 붙여넣기 시도
            ok = tab4_paste_row_using_tab2(driver, tr, values)
            if not ok:
                last_error = "붙여넣기실패"
                continue

            # 성공
            ok_count += 1
            print(f"   [입력 {idx}/{excel_count}] {item}")
            paste_ok = True
            break

        # 3회 모두 실패한 경우
        if not paste_ok:
            failed.append(item)
            if last_error:
                fail_reason.append((item, last_error))
                print(f"   [미적용 {idx}/{excel_count}] {item} ({last_error})")
            else:
                fail_reason.append((item, "미상"))
                print(f"   [미적용 {idx}/{excel_count}] {item}")

    print(f"[탭4-결과] 성공 {ok_count}개 / 실패 {len(failed)}개")
    if failed:
        rs = {}
        for _, reason in fail_reason:
            rs[reason] = rs.get(reason, 0) + 1
        print("   ↳ 실패 원인: " + ", ".join([f"{k} {v}개" for k, v in rs.items()]))



# =====================================================================
# 공통 유틸
# =====================================================================

def parse_sample_input(s: str):
    """
    "A2512313-01, A2512313-02" / 공백 / 줄바꿈 섞여도 여러개 파싱
    """
    if not s:
        return []
    parts = re.split(r"[,\s]+", s.strip())
    out, seen = [], set()
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


def find_sample_file_in_nas(sample_no: str):
    """대기 NAS에서 엑셀 파일 검색"""
    return _find_excel_util(sample_no, nas_base=NAS_BASE, nas_dirs=NAS_DIRS, strict=False)


def find_sample_file_in_water_nas(sample_no: str):
    """수질 NAS에서 엑셀 파일 검색.
    경로 구조: NAS_BASE_WATER\YYYY년\업체명\파일.xlsm  (하위 전체 재귀)
    """
    return _find_excel_util(sample_no, nas_base=NAS_BASE_WATER, nas_dirs=[""], strict=False)


def ask_yesno(msg, default_yes=True):
    """
    예/아니오 입력 받기.
    - yes: 예, y, yes, ㅇ, 1
    - no : 아니오, n, no, ㄴ, 0
    엔터만 치면 default_yes 적용
    """
    while True:
        s = input(msg).strip().lower()
        if s == "":
            return default_yes
        if s in ("예", "y", "yes", "ㅇ", "1"):
            return True
        if s in ("아니오", "n", "no", "ㄴ", "0"):
            return False
        print("⚠ 입력은 예/아니오(또는 y/n)로 해줘.")


def wait(sec): time.sleep(sec)

# clean: eco_input 전용 (앞표시 제거)
def clean(s):
    if not s: return ""
    s = str(s).strip()
    if s.startswith(("×","✕",",")): return s[1:].strip()
    return s

# accept_any_alert: eco_input 전용 (단일 알림 수락 + 텍스트 반환)
def accept_any_alert(driver, timeout=2):
    end = time.time() + timeout
    while time.time() < end:
        try:
            a = driver.switch_to.alert
            txt = a.text
            a.accept()
            time.sleep(0.2)
            return txt
        except:
            time.sleep(0.1)
    return None

# =====================================================================
# 시료번호 목록 자동 생성
# =====================================================================

def extract_samples_from_nas(team_no, date_str):
    """
    파일명에서 시작부분 AYYMMDDT-XX 추출.
    team_no, 날짜(YYMMDD) 일치하는 파일만 목록에 추가.
    파일명에 '비산먼지' 포함 시 dust=True
    """
    sample_list=[]
    pat=re.compile(r"^(A\d{6}\d-\d{2})")  # A2512313-01 패턴

    date_key = date_str.replace("-", "")[2:]   # '2025-12-31' → '251231'

    for d in NAS_DIRS:
        folder=os.path.join(NAS_BASE,d)
        if not os.path.isdir(folder): continue

        for root,dirs,files in os.walk(folder):
            for f in files:
                if f.startswith("~$"): continue
                m=pat.match(f)
                if not m: continue

                sample_no=m.group(1)

                # 날짜 필터 ↓↓↓
                if sample_no[1:7] != date_key:
                    continue

                # 팀 필터 ↓↓↓
                if sample_no[7] != str(team_no):
                    continue

                dust = ("비산먼지" in f)

                sample_list.append({
                    "sample": sample_no,
                    "path": os.path.join(root,f),
                    "dust": dust
                })

    return sample_list

# =====================================================================
# 사이트 상세 진입 – 시료번호 검색 방식
# =====================================================================

def open_detail_by_sample(d, sample_no):
    """measin_utils.open_sample_detail 래퍼 (eco_input 호환성 유지)"""
    return _open_sample_detail(d, sample_no)

def back_to_list(d, btn_selector="#btnMsFieldDocCancel"):
    """measin_utils.go_back_to_list 래퍼 (eco_input 호환성 유지)"""
    selectors = [btn_selector]
    if btn_selector != "#btnGoList":
        selectors.append("#btnGoList")
    # accept_any_alert 먼저 처리
    accept_any_alert(d, timeout=2)
    result = _go_back_to_list(d, btn_selectors=selectors)
    accept_any_alert(d, timeout=2)
    return result

# =====================================================================
# 탭1 입력
# =====================================================================

def set_dust_radio(d, is_dust):
    try:
        if is_dust:
            sel = "#rowAirRpt > section:nth-child(1) > div > label:nth-child(1)"
        else:
            sel = "#rowAirRpt > section:nth-child(1) > div > label.radio.ml-1 > i"
        safe_click(d, sel)
        print("  → 비산먼지 설정 완료:", is_dust)
    except:
        print("⚠ 비산먼지 설정 실패")

def clear_multi(d, ul_sel):
    _Select2Handler(d).clear(ul_sel)

def add_multi(d, search_sel, val):
    _Select2Handler(d).add(search_sel, val)

def fill_multi(d, ul_sel, search_sel, values):
    _Select2Handler(d).fill(ul_sel, search_sel, values)

def fill_tab1(d, data, is_dust):
    print("▶ 탭1 입력 시작")
    set_dust_radio(d, is_dust)

    # --------------------------------------
    # ① 비산먼지인 경우
    # --------------------------------------
    if is_dust:
        print("▶ 비산먼지 → 기존 선택 삭제 + 전용 셀렉터로 입력")


        # 측정항목
        fill_multi(
            d,
            "#inairTargetItem > div:nth-child(2) > div > span > span.selection > span > ul",
            "#inairTargetItem > div:nth-child(2) > div > span > span.selection > span > ul > li.select2-search.select2-search--inline > input",
            data["측정항목"]
        )

        # 인력
        fill_multi(
            d,
            "#wid-id-4 > div > div.widget-body.no-padding > div > fieldset > div.row.input-full > section:nth-child(2) > span > span.selection > span > ul",
            "#wid-id-4 > div > div.widget-body.no-padding > div > fieldset > div.row.input-full > section:nth-child(2) > span > span.selection > span > ul > li.select2-search.select2-search--inline > input",
            data["인력"]
        )

        # 차량
        fill_multi(
            d,
            "#carSection > div > span > span.selection > span > ul",
            "#carSection > div > span > span.selection > span > ul > li.select2-search.select2-search--inline > input",
            data["차량"]
        )

        # 장비
        fill_multi(
            d,
            "#machineDiv > div > span > span.selection > span > ul",
            "#machineDiv > div > span > span.selection > span > ul > li.select2-search.select2-search--inline > input",
            data["장비"]
        )


        print("▶ 비산먼지 탭1 입력 완료")

    # --------------------------------------
    # ② 비산먼지 아닌 경우 (원래 잘 되던 로직 그대로)
    # --------------------------------------
    else:
        # 측정항목
        fill_multi(
            d,
            "#inairTargetItem > div:nth-child(2) > div > span > span.selection > span > ul",
            "#inairTargetItem input[type='search']",
            data["측정항목"]
        )

        # 인력
        fill_multi(
            d,
            "#edit_emp_id + span ul.select2-selection__rendered",
            "#edit_emp_id + span input.select2-search__field",
            data["인력"]
        )

        # 차량
        fill_multi(
            d,
            "#carSection > div > span > span.selection > span > ul",
            "#carSection input[type='search']",
            data["차량"]
        )

        # 장비
        fill_multi(
            d,
            "#machineDiv > div > span > span.selection > span > ul",
            "#machineDiv input[type='search']",
            data["장비"]
        )

    print("▶ 탭1 입력 완료")

    # 탭1 저장 버튼 (비산먼지/일반 공통)
    try:
        safe_click(d, "#updateFieldPlanBtn")
        try:
            d.switch_to.alert.accept(); wait(0.5)
        except:
            pass
        try:
            d.switch_to.alert.accept(); wait(0.5)
        except:
            pass
        print("✅ 탭1 저장 완료")
    except:
        print("⚠ 탭1 저장 실패")


# =====================================================================
# 탭2 입력
# =====================================================================

def set_wind(d, txt):
    if not txt: return
    mp={"북":0,"북-북동":1,"북동":2,"동-북동":3,"동":4,"동-남동":5,"남동":6,
        "남-남동":7,"남":8,"남-남서":9,"남서":10,"서-남서":11,"서":12,
        "서-북서":13,"북서":14,"북-북서":15,"정온":99}
    v=mp.get(txt)
    if v is None: return
    try:
        el=WebDriverWait(d,10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,"select.meas_wdir")))
        Select(el).select_by_value(str(v))
        d.execute_script(
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", el)
    except: pass

def fill_tab2(d, data, is_dust):
    print("▶ 탭2 입력 시작")

    def sv(sel,val):
        try:
            e=d.find_element(By.CSS_SELECTOR,sel)
            e.clear(); e.send_keys(str(val)); wait(0.1)
        except: pass

    # 공통
    sv("input[name='meas_wthr']", data["기상"])
    sv("input[name='meas_temper']", data["기온"])
    sv("input.meas_humd", data["습도"])
    sv("input.meas_atoms", data["기압"])
    set_wind(d, data["풍향"])
    sv("input.meas_wspd", data["풍속"])

    sv("#meas_start_time", data["채취시작"])
    sv("#meas_end_time", data["채취끝"])

    if not is_dust:
        sv("#basis_o2c", data["표준산소농도"])
        sv("#meas_o2c", data["실측산소농도"])
        sv("#meas_gas_fvol", data["배출가스유량전"])
        sv("#meas_gas_fvol_o2_aft", data["배출가스유량후"])
        sv("#meas_humd_vol", data["수분량"])
        sv("#gas_meter_temper", data["배출가스온도"])
        sv("#meas_fspd", data["배출가스유속"])

    print("✅ 탭2 입력 완료")



# ------------------------------------------------------------
# 🔽🔽 여기부터 새 기능 추가 (기존 코드 절대 수정 없음, 아래만 수정)
# ------------------------------------------------------------
def parse_facility_from_excel(path):
    """
    엑셀의 '대기측정기록부'에서 26~29행 D,F,H,J,L,N 가져오기
    (시설/가동상황 표)
    """
    wb = load_workbook(path, data_only=True)

    if "대기측정기록부" in wb.sheetnames:
        ws = wb["대기측정기록부"]
    else:
        ws = wb[wb.sheetnames[0]]

    rows = []
    for r in range(26, 30):  # 26,27,28,29
        row = {
            "fuel": ws[f"D{r}"].value,        # 연료사용량
            "prod": ws[f"F{r}"].value,        # 제품생산량
            "inc": ws[f"H{r}"].value,         # 소각량
            "raw": ws[f"J{r}"].value,         # 원료투입량
            "kind": ws[f"L{r}"].value,        # 종류
            "unit": ws[f"N{r}"].value,        # 단위
        }
        rows.append(row)
    return rows


def select_first_prev_facility(driver, row_idx):
    """
    방지시설 Select2:
    첫 번째 옵션 클릭 후 ENTER로 선택 확정
    """
    try:
        search_input = f"#prev_fac_no_by_emis_fac_{row_idx} + span input.select2-search__field"

        # 검색창 클릭(열기)
        inp = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, search_input))
        )
        inp.click()
        time.sleep(0.2)

        # 옵션 목록 가져오기
        options = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "ul.select2-results__options li")
            )
        )

        if not options:
            print(f"⚠ row{row_idx}: 방지시설 옵션 없음")
            return

        first = options[0]

        # 1차 클릭(선택 후보)
        driver.execute_script("arguments[0].click();", first)
        time.sleep(0.2)

        # ENTER로 선택 확정
        inp.send_keys(Keys.ENTER)
        time.sleep(0.2)

        print(f" - row{row_idx}: 방지시설 첫 항목 ENTER로 최종 선택 완료")

    except Exception as e:
        print(f"⚠ row{row_idx}: 방지시설 선택 실패 → {e}")

def fill_facility_rows(driver, fac_rows):
    """
    rowfac0 ~ rowfac3 까지 4행 자동 입력
    fac_rows: parse_facility_from_excel 결과 리스트
    """

    print("▶ 배출시설 가동상황 자동 입력")

    for idx, row in enumerate(fac_rows):
        try:
            fuel = "" if row["fuel"] is None else str(row["fuel"]).strip()
            prod = "" if row["prod"] is None else str(row["prod"]).strip()
            inc  = "" if row["inc"]  is None else str(row["inc"]).strip()
            raw  = "" if row["raw"]  is None else str(row["raw"]).strip()
            kind = "" if row["kind"] is None else str(row["kind"]).strip()
            unit = "" if row["unit"] is None else str(row["unit"]).strip()

            # 각 칸 입력 (기존 기능 유지)
            driver.find_element(By.CSS_SELECTOR, f"#fuel_used_{idx}").clear()
            driver.find_element(By.CSS_SELECTOR, f"#fuel_used_{idx}").send_keys(fuel)

            driver.find_element(By.CSS_SELECTOR, f"#production_{idx}").clear()
            driver.find_element(By.CSS_SELECTOR, f"#production_{idx}").send_keys(prod)

            driver.find_element(By.CSS_SELECTOR, f"#incinerate_{idx}").clear()
            driver.find_element(By.CSS_SELECTOR, f"#incinerate_{idx}").send_keys(inc)

            driver.find_element(By.CSS_SELECTOR, f"#raw_material_{idx}").clear()
            driver.find_element(By.CSS_SELECTOR, f"#raw_material_{idx}").send_keys(raw)

            driver.find_element(By.CSS_SELECTOR, f"#kind_{idx}").clear()
            driver.find_element(By.CSS_SELECTOR, f"#kind_{idx}").send_keys(kind)

            driver.find_element(By.CSS_SELECTOR, f"#unit_{idx}").clear()
            driver.find_element(By.CSS_SELECTOR, f"#unit_{idx}").send_keys(unit)

            # ✅ 연료사용량이 있으면 → 방지시설 첫 번째 항목 선택
            if fuel != "":
                select_first_prev_facility(driver, idx)
            else:
                print(f" - row{idx}: 연료사용량 없음 → 방지시설 선택 안 함")

        except Exception as e:
            print(f"⚠ 배출시설 row{idx} 입력 실패:", e)

    print("✅ 배출시설 가동상황 입력 완료")


def write_sampler_comment(driver, text="특이사항 없음"):
    """
    채취자 의견
    """
    try:
        opin = driver.find_element(By.CSS_SELECTOR, "#smpl_ctor_opin")
        opin.clear()
        opin.send_keys(text)
        time.sleep(0.3)
        print("✅ 채취자 의견 입력 완료")

    except Exception as e:
        print("⚠ 채취자 의견 입력 실패:", e)


# ------------------------------------------------------------
# 🔽 탭4 입력
# ------------------------------------------------------------
def fill_tab4_dates(driver, sample_no: str, tab4_meta: dict):
    sample_date = sample_to_datestr(sample_no)  # 기존 함수 (YYYY-MM-DD)

    rcpt_dt = tab4_meta.get("rcpt_dt", "").strip()
    start_src = tab4_meta.get("start_src", "").strip()
    end_dt = tab4_meta.get("end_dt", "").strip()

    # 1) 시료접수일시
    if rcpt_dt:
        set_date_js(driver, "#smpl_rcpt_dt", rcpt_dt)

    # 2) 분석기간 시작
    if start_src:
        set_date_js(driver, "#anze_start_dt", start_src)
    
    # 3) 분석기간 끝
    if end_dt:
        set_date_js(driver, "#anze_end_dt", end_dt)


# ------------------------------------------------------------
# 🔽 탭2 저장 (입력완료→확인2번)
# ------------------------------------------------------------

def _should_draft_by_sampling_end(date_str: str, end_hm: str) -> bool:
    """
    현재시간이 '시료채취(해당 날짜의) 끝 시간'보다 이르면 True(임시저장)
    - date_str: 'YYYY-MM-DD'
    - end_hm: 'HH:MM' 또는 'HH:MM:SS'
    """

    if not date_str or not end_hm:
        return False  # 값 없으면 최종저장 쪽(원하면 True로 바꿔도 됨)

    date_str = str(date_str).strip()
    end_hm = str(end_hm).strip()

    # 날짜 파싱
    try:
        meas_date = datetime.strptime(date_str, "%Y-%m-%d").date()
    except Exception:
        return False

    # 시간 파싱
    try:
        parts = end_hm.split(":")
        hh = int(parts[0]); mm = int(parts[1]); ss = int(parts[2]) if len(parts) >= 3 else 0
        end_t = dt.time(hh, mm, ss)
    except Exception:
        return False

    now = datetime.now()

    # ✅ 날짜가 과거면 이미 끝난걸로 보고 "최종저장" 가능 (임시저장 X)
    if meas_date < now.date():
        return False

    # ✅ 날짜가 미래면 아직 측정일도 안됨 → 임시저장 강제
    if meas_date > now.date():
        return True

    # ✅ 날짜가 오늘이면 시간 비교
    end_dt = datetime.combine(meas_date, end_t)
    return now < end_dt



def save_tab2(driver, use_final_save=False):
    """
    use_final_save=False  → 임시저장(#btnDraftMsFieldDoc)
    use_final_save=True   → 입력완료(#btnSaveMsFieldDoc)  (PDF 업로드 동반 시)
    """
    btn_sel = "#btnSaveMsFieldDoc" if use_final_save else "#btnDraftMsFieldDoc"
    print(f"▶ 탭2 저장 시작 ({'입력완료' if use_final_save else '임시저장'})")

    try:
        btn = driver.find_element(By.CSS_SELECTOR, btn_sel)
        driver.execute_script("arguments[0].click();", btn)

        # ✅ 클릭 직후 바로 뜨는 알럿들 처리
        _accept_all_alerts(driver, total_wait=2.5, poll=0.15, label="1차")

        # ✅ 업로드/서버 처리 대기(입력완료면 더 길게)
        #    이 구간에도 알럿이 늦게 뜰 수 있으니 '기다리기+알럿처리'로 통합
        wait_time = 1.5 if not use_final_save else 7.0
        _accept_all_alerts(driver, total_wait=wait_time, poll=0.2, label="대기중")

        # ✅ 마무리로 한 번 더(잔여 알럿 방지)
        _accept_all_alerts(driver, total_wait=2.0, poll=0.2, label="마무리")

        print("▶ 탭2 저장 완료")
        return True

    except Exception as e:
        print("⚠ 탭2 저장 실패:", e)
        return False



        
        
# ============================================================
# PDF 생성(엑셀 COM) + 사이트 업로드(파일창 자동) + 업로드 후 PDF 삭제
# ============================================================

def _find_sheet_name(wb, prefer_names):
    names = [sh.Name for sh in wb.Worksheets]
    for pn in prefer_names:
        if pn in names:
            return pn
    for pn in prefer_names:
        p = pn.replace(" ", "")
        for n in names:
            if p in n.replace(" ", ""):
                return n
    return None


def export_pdf_from_excel(excel_path: str, sample_no: str, is_dust: bool) -> str:
    out_dir = PDF_TMP_DIR
    os.makedirs(out_dir, exist_ok=True)
    out_pdf = os.path.join(out_dir, f"{sample_no}.pdf")

    if is_dust:
        sheet_a_candidates = ["대기시료채취 및 분석일지", "대기시료채취및분석일지"]
        sheet_b_candidates = ["교정"]
    else:
        sheet_a_candidates = ["대기시료채취 및 분석일지", "대기시료채취및분석일지"]
        sheet_b_candidates = ["먼지시료채취기록지"]

    excel = None
    wb = None
    try:
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(excel_path, ReadOnly=True)

        sh_a = _find_sheet_name(wb, sheet_a_candidates)
        sh_b = _find_sheet_name(wb, sheet_b_candidates)

        if not sh_a:
            raise RuntimeError(f"필수 시트 없음: {sheet_a_candidates}")
        if not sh_b:
            raise RuntimeError(f"필수 시트 없음: {sheet_b_candidates}")

        # ✅ VARIANT 없이 "시트 2개" 멀티선택 (안정적)
        wb.Worksheets(sh_a).Select()          # 첫 시트 선택
        wb.Worksheets(sh_b).Select(False)     # Replace=False → 선택에 추가

        # 0 = xlTypePDF
        wb.ActiveSheet.ExportAsFixedFormat(0, out_pdf)


        return out_pdf

    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except:
            pass
        try:
            pythoncom.CoUninitialize()
        except:
            pass


def _win32_is_file_open_dialog(hwnd) -> bool:
    """#32770 공용 대화상자 중 '열기/Open' 버튼 + Edit(파일명 입력칸) 있는지로 판별"""
    try:
        if not win32gui.IsWindowVisible(hwnd):
            return False
        if win32gui.GetClassName(hwnd) != "#32770":
            return False

        found_open_btn = False
        found_edit = False

        def enum_child(ch, _):
            nonlocal found_open_btn, found_edit
            cls = win32gui.GetClassName(ch)
            txt = win32gui.GetWindowText(ch)

            if cls == "Button" and (("열기" in txt) or ("Open" in txt)):
                found_open_btn = True

            # 파일이름 입력칸은 Edit가 직접 있거나 ComboBox 안에 숨어있음
            if cls == "Edit":
                found_edit = True

        win32gui.EnumChildWindows(hwnd, enum_child, None)

        # Edit가 바로 없으면 ComboBox/ComboBoxEx32 안의 Edit도 탐색
        if not found_edit:
            def enum_child2(ch, _):
                nonlocal found_edit
                cls = win32gui.GetClassName(ch)
                if cls in ("ComboBoxEx32", "ComboBox"):
                    def enum_grand(gch, __):
                        nonlocal found_edit
                        if win32gui.GetClassName(gch) == "Edit":
                            found_edit = True
                    win32gui.EnumChildWindows(ch, enum_grand, None)
            win32gui.EnumChildWindows(hwnd, enum_child2, None)

        return found_open_btn and found_edit
    except:
        return False


def _find_open_dialog(timeout=30):
    """
    ✅ 제목이 비어 있어도 찾도록 개선:
    1) Win32 EnumWindows로 #32770 중 '열기/Open 버튼 + 파일명 입력칸' 가진 창을 찾음
    2) 실패 시 UIA로 '열기/Open 버튼' 가진 창을 찾음(제목 무시)
    """
    end = time.time() + timeout

    while time.time() < end:
        # 1) Win32: 최우선(가장 빠르고 확실)
        try:
            found = []
            def enum_top(hwnd, _):
                if _win32_is_file_open_dialog(hwnd):
                    found.append(hwnd)
            win32gui.EnumWindows(enum_top, None)

            if found:
                # 가장 최근에 뜬 애를 잡기(보통 마지막이 최신)
                hwnd = found[-1]
                # pywinauto win32 wrapper 반환
                return Desktop(backend="win32").window(handle=hwnd)
        except:
            pass

        # 2) UIA: 제목이 없어도 '열기/Open' 버튼이 있는 창 찾기
        try:
            for w in Desktop(backend="uia").windows():
                try:
                    # 버튼 이름으로 찾기(열기/Open)
                    btns = w.descendants(control_type="Button")
                    if any(("열기" in b.window_text()) or ("Open" in b.window_text()) for b in btns):
                        return w
                except:
                    continue
        except:
            pass

        time.sleep(0.1)

    raise RuntimeError("열기(Open) 창을 찾지 못함(timeout). (제목이 비거나 다른 구조일 가능성)")



def pick_file_in_open_dialog(file_path: str, timeout=30):
    """
    ✅ 가장 안정적인 입력: Alt+N → Ctrl+V → Enter
    (UIA Edit set_text는 PC마다 안 먹는 경우가 많아서 키보드 입력으로 고정)
    """
    dlg = _find_open_dialog(timeout=timeout)

    # 전면/포커스 강제
    try:
        hwnd = dlg.handle
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except:
        pass

    try:
        dlg.set_focus()
    except:
        pass

    time.sleep(0.2)

    # 파일 이름 칸 포커스(Alt+N) → 붙여넣기 → Enter
    pyperclip.copy(file_path)
    send_keys("%n")                 # 파일 이름(N)
    time.sleep(0.05)
    send_keys("^a{BACKSPACE}")      # 지우기
    time.sleep(0.05)
    send_keys("^v")                 # 붙여넣기
    time.sleep(0.1)
    send_keys("{ENTER}")            # 열기

    # 창이 닫히면 성공
    try:
        dlg.wait_not("visible", timeout=6)
        return True
    except:
        raise RuntimeError(f"파일 선택 실패(열기창이 닫히지 않음). 경로/권한 확인: {file_path}")



def trigger_file_dialog(driver, timeout=6):
    """
    '파일추가' 버튼( onclick="control.openFileDialog();" )로 파일 선택창을 띄운다.
    1) 진짜 클릭(ActionChains)
    2) 안 되면 JS로 control.openFileDialog() 직접 호출
    """
    btn_sel = "#fileArea input.btn.btnView[value='파일추가']"

    # 1) 진짜 클릭
    try:
        btn = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, btn_sel))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        time.sleep(0.2)
        ActionChains(driver).move_to_element(btn).pause(0.05).click(btn).perform()
        time.sleep(0.4)
        return True
    except:
        pass

    # 2) JS 직접 호출(버튼 클릭이 막힐 때)
    try:
        driver.execute_script("""
            try {
                if (window.control && typeof window.control.openFileDialog === 'function') {
                    window.control.openFileDialog();
                    return true;
                }
                if (typeof control !== 'undefined' && control && typeof control.openFileDialog === 'function') {
                    control.openFileDialog();
                    return true;
                }
            } catch(e) {}
            return false;
        """)
        time.sleep(0.4)
        return True
    except:
        return False
        
        
def read_input_h7(excel_path: str) -> str:
    wb = load_workbook(excel_path, data_only=True)
    ws = wb["입력"] if "입력" in wb.sheetnames else wb[wb.sheetnames[0]]
    v = ws["H7"].value
    return "" if v is None else str(v).strip()

# ------------------------------------------------------------
# 🔽 PDF 최종본 저장
# ------------------------------------------------------------

def export_sheet_pdf(excel_path: str, sheet_name: str, out_pdf_path: str) -> bool:
    """엑셀 시트를 PDF로 내보내기 (pdf_utils.PDFExporter 사용)"""
    exporter = _PDFExporter(excel_path)
    try:
        return exporter.export_sheet(sheet_name, out_pdf_path)
    finally:
        exporter.close()

def merge_pdfs(pdf_list, out_pdf):
    """PDF 파일 병합 (pdf_utils.merge_pdfs 사용)"""
    _merge_pdfs(pdf_list, out_pdf)

def build_final_paths(excel_path: str, sample_no: str):
    yyyy = sample_no[1:3]  # 네 SN 규칙에 맞게 이미 함수가 있을 수도 있음
    yyyy_full = f"20{yyyy}"  # A26....이면 2026 같은

    h7 = read_input_h7(excel_path)
    p1_dir = FINAL_DONE_DIR
    p2_dir = os.path.join(PDF_BASE_DIR, yyyy_full, h7)

    os.makedirs(p1_dir, exist_ok=True)
    os.makedirs(p2_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    out_name = f"{base_name}.pdf"
    return os.path.join(p1_dir, out_name), os.path.join(p2_dir, out_name)


def build_final_path_water(excel_path: str, sample_no: str) -> str:
    """수질 PDF 최종 저장 경로.
    → PDF_BASE_DIR_WATER\YYYY년\업체명\파일명.pdf
    업체명은 엑셀 H7 셀 값(read_input_h7) 사용.
    """
    yyyy_full = f"20{sample_no[1:3]}"        # "26" → "2026"
    company   = read_input_h7(excel_path)    # 엑셀 H7 = 업체명
    out_dir   = os.path.join(PDF_BASE_DIR_WATER, f"{yyyy_full}년", company)
    os.makedirs(out_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    return os.path.join(out_dir, f"{base_name}.pdf")


def make_tab4_pdfs_water(excel_path: str, sample_no: str) -> dict:
    """수질 탭4 PDF 생성 + 최종본 저장.
    업1(anzeFile1): 분석일지 단독
    업2(anzeFile2): 대행기록부 단독
    최종본: 두 시트 병합 → build_final_path_water()
    """
    tmp_dir = PDF_TMP_DIR
    os.makedirs(tmp_dir, exist_ok=True)

    pdf_analy  = os.path.join(tmp_dir, f"{sample_no}__분석일지.pdf")
    pdf_record = os.path.join(tmp_dir, f"{sample_no}__대행기록부.pdf")

    export_sheet_pdf(excel_path, WATER_SHEET_ANALY,  pdf_analy)
    export_sheet_pdf(excel_path, WATER_SHEET_RECORD, pdf_record)

    # 최종본: 분석일지 + 대행기록부 병합
    merge_src = [p for p in (pdf_analy, pdf_record) if os.path.isfile(p)]
    final_tmp = os.path.join(tmp_dir, f"{sample_no}__FINAL.pdf")
    merge_pdfs(merge_src, final_tmp)

    import shutil
    p_out = build_final_path_water(excel_path, sample_no)
    shutil.copy2(final_tmp, p_out)

    return {
        "final_tmp" : final_tmp,
        "pdf_analy" : pdf_analy,    # #anzeFile1 업로드용
        "pdf_record": pdf_record,   # #anzeFile2 업로드용
        "p_out"     : p_out,
    }


def make_tab4_pdfs(excel_path: str, sample_no: str):
    tmp_dir = PDF_TMP_DIR
    os.makedirs(tmp_dir, exist_ok=True)

    # 업로드용 원본 시트 PDF
    pdf_analy = os.path.join(tmp_dir, f"{sample_no}__대기시료채취및분석일지.pdf")
    pdf_record = os.path.join(tmp_dir, f"{sample_no}__대기측정기록부.pdf")

    export_sheet_pdf(excel_path, "대기시료채취 및 분석일지", pdf_analy)
    export_sheet_pdf(excel_path, "대기측정기록부", pdf_record)

    # 커버(시험항목날짜서명) PDF
    pdf_cover = os.path.join(tmp_dir, f"{sample_no}__시험항목날짜서명.pdf")
    export_sheet_pdf(excel_path, "시험항목날짜서명", pdf_cover)

    # ✅ 업1만: 커버 + 분석일지 병합해서 새 파일 생성
    upload1 = os.path.join(tmp_dir, f"{sample_no}__분석일지.pdf")  # newFile1에 올릴 파일
    if os.path.isfile(pdf_cover) and os.path.isfile(pdf_analy):
        merge_pdfs([pdf_cover, pdf_analy], upload1)
    else:
        # 커버가 없거나 분석일지가 없으면 fallback (원래 파일이라도 올릴 수 있게)
        upload1 = pdf_analy

    # ✅ 업2는 "대기측정기록부 단독" 그대로
    upload2 = pdf_record

    # 최종 병합본(3시트) = 기존 그대로
    merge_src = [p for p in (pdf_cover, pdf_analy, pdf_record) if os.path.isfile(p)]
    final_tmp = os.path.join(tmp_dir, f"{sample_no}__FINAL.pdf")
    merge_pdfs(merge_src, final_tmp)

    # 두 경로에 복사(이 부분은 기존 유지)
    p1, p2 = build_final_paths(excel_path, sample_no)
    import shutil
    shutil.copy2(final_tmp, p1)
    shutil.copy2(final_tmp, p2)

    # ✅ 리턴에서 업로드용은 upload1/upload2로
    return {
        "final_tmp": final_tmp,
        "pdf_analy": upload1,     # 업1
        "pdf_record": upload2,    # 업2
        "p1": p1, "p2": p2
    }

def cleanup_tmp_pdfs(tmp_dir: str, sample_no: str, retries=6, delay=0.25):
    import glob
    pattern = os.path.join(tmp_dir, f"{sample_no}*.pdf")
    targets = glob.glob(pattern)

    for p in targets:
        ok = False
        for _ in range(retries):
            try:
                if os.path.isfile(p):
                    os.remove(p)
                ok = True
                break
            except PermissionError:
                time.sleep(delay)
            except Exception:
                time.sleep(delay)

        if not ok and os.path.isfile(p):
            print(f"⚠ tmp 삭제 실패(잠김): {p}")



def wait_file_selected(driver, css_sel, timeout=5.0):
    end = time.time() + timeout
    while time.time() < end:
        v = driver.execute_script("return arguments[0].value;", driver.find_element(By.CSS_SELECTOR, css_sel))
        if v and str(v).strip():
            return True
        time.sleep(0.1)
    return False


def upload_tab4_pdfs(driver, pdf_analy_path: str, pdf_record_path: str,
                     f1_sel: str = "#newFile1", f2_sel: str = "#newFile2"):
    """탭4 PDF 업로드.  대기: #newFile1/2,  수질: #anzeFile1/2"""

    f1 = driver.find_element(By.CSS_SELECTOR, f1_sel)
    f2 = driver.find_element(By.CSS_SELECTOR, f2_sel)

    driver.execute_script("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", f1)
    driver.execute_script("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", f2)

    # 1) 파일 선택(트리거: onchange)
    f1.send_keys(pdf_analy_path)
    if not wait_file_selected(driver, f1_sel, timeout=5.0):
        raise RuntimeError("⚠탭4 파일1 선택 확인 실패(newFile1)")

    f2.send_keys(pdf_record_path)
    if not wait_file_selected(driver, f2_sel, timeout=5.0):
        raise RuntimeError("⚠탭4 파일2 선택 확인 실패(newFile2)")

    # 2) setFile 처리 시간(서버 업로드/검증이 있을 수 있어 최소 대기)
    time.sleep(1.0)


def tab4_temp_save(driver):
    # 임시저장
    safe_click(driver, "#btnTempSave")
    _accept_all_alerts(driver, total_wait=6.0, poll=0.2, label="탭4임시저장")
    
def tab4_comp_save(driver):
    # 입력완료
    safe_click(driver, "#btnCompSave")
    _accept_all_alerts(driver, total_wait=6.0, poll=0.2, label="탭4저장완료")    


#==============================탭2 리얼그리드 입력==============================

def read_dust_realgird_values(excel_path: str):
    """
    비산먼지일 때:
    - 시료채취량 = 평균(B19,B20) * B2 * 1000
    - 흡인속도 = 평균(B19,B20) * 1000
    - 시작시간 = I38, 종료시간 = J38
    - 단위: L, L/min
    """
    wb = load_workbook(excel_path, data_only=True)
    ws = _find_sheet_by_candidates(wb, ["입력", "입력(분석값)"])
    if ws is None:
        raise RuntimeError("엑셀에서 '입력' 시트를 찾지 못함")

    def f(v):
        if v is None or v == "":
            return None
        try:
            return float(v)
        except:
            return None

    b19 = f(ws["B19"].value)
    b20 = f(ws["B20"].value)
    b2  = f(ws["B2"].value)

    if b19 is None or b20 is None or b2 is None:
        raise RuntimeError("비산먼지 계산용 셀(B19,B20,B2) 값이 비어있음")

    avg = (b19 + b20) / 2.0

    samp_vol = avg * b2 * 1000.0     # L
    suction  = avg * 1000.0          # L/min

    st = ws["I38"].value
    ed = ws["J38"].value

    st = trim_hm(st)  # 너가 이미 가진 함수
    ed = trim_hm(ed)

    return {
        "시료채취량": f"{samp_vol:.0f}",   # 보통 정수로 들어가길래 0자리
        "채취량단위": "L",
        "흡인속도": f"{suction:.0f}",
        "흡인속도단위": "L/min",
        "시작시간": st,
        "종료시간": ed,
    }


def read_realgird_headers(driver, grid_root_css: str):
    """
    RealGrid 컨테이너(grid_root_css)에서 헤더 텍스트를 읽어
    [{'idx': 1, 'text': '대분류'}, ...] 형태로 반환
    idx는 tbody td:nth-child(idx) 와 맞추는 용도
    """
    js = r"""
    const root = document.querySelector(arguments[0]);
    if (!root) return null;

    const tries = [
      "div.rg-header table thead tr:last-child th",
      "div.rg-header table thead tr:last-child td",
      "table.rg-header-table thead tr:last-child th",
      "table.rg-header-table thead tr:last-child td",
      "thead tr:last-child th",
      "thead tr:last-child td"
    ];

    for (const sel of tries) {
      const els = root.querySelectorAll(sel);
      if (els && els.length) {
        return Array.from(els).map((el, i) => ({
          idx: i + 1,
          text: (el.innerText || "").trim()
        }));
      }
    }

    // fallback: 헤더를 못 찾으면 전체 thead 긁기
    const els = root.querySelectorAll("thead th, thead td");
    if (els && els.length) {
      return Array.from(els).map((el, i) => ({
        idx: i + 1,
        text: (el.innerText || "").trim()
      }));
    }
    return [];
    """
    headers = driver.execute_script(js, grid_root_css)
    if not headers:
        return []
    # 빈 텍스트 제거(가끔 공백 헤더가 섞임)
    return [h for h in headers if h.get("text")]


def build_header_map(headers):
    """
    headers(list) -> dict로 변환
    '단위'가 2번 나오므로
      - 시료채취량단위
      - 흡인속도단위
    로 분리해서 맵핑해줌
    """
    def find_idx(text):
        for h in headers:
            if h["text"] == text:
                return h["idx"]
        return None

    def find_next_idx_after(text, after_idx):
        for h in headers:
            if h["idx"] > after_idx and h["text"] == text:
                return h["idx"]
        return None

    m = {}
    # 기본 컬럼들
    for key in ["대분류","중분류","측정항목","시료채취량","흡인속도",
                "측정일(시작)","시작시간","측정일(종료)","종료시간","비고"]:
        m[key] = find_idx(key)

    # 단위 2개 분리
    vol_idx = m.get("시료채취량")
    spd_idx = m.get("흡인속도")

    if vol_idx:
        m["시료채취량단위"] = find_next_idx_after("단위", vol_idx)
    if spd_idx:
        m["흡인속도단위"] = find_next_idx_after("단위", spd_idx)

    return m


def rg_find_row_by_item(driver, grid_root_css, item_name):
    """
    tbody의 td div.rg-renderer 텍스트 중 item_name과 같은 행 찾기 (1-based row index)
    """
    js = r"""
    const root = document.querySelector(arguments[0]);
    const name = arguments[1];
    if (!root) return null;
    const body = root.querySelector('.rg-body table tbody');
    if (!body) return null;
    const rows = body.querySelectorAll('tr');
    for (let i=0;i<rows.length;i++){
      const tds = rows[i].querySelectorAll('td div.rg-renderer');
      for (const d of tds){
        if ((d.innerText||'').trim() === name) return i+1;
      }
    }
    return null;
    """
    return driver.execute_script(js, grid_root_css, item_name)


def rg_set_cell(driver, grid_root_css, row_idx, col_idx, value):
    """
    셀 클릭 → Ctrl+A → 값 → Enter
    """
    # td 선택 후 클릭 (RealGrid는 td 안에 div.rg-renderer가 보통 있음)
    cell_css = f"{grid_root_css} .rg-body table tbody tr:nth-child({row_idx}) td:nth-child({col_idx})"
    el = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, cell_css)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.1)
    ActionChains(driver).move_to_element(el).click(el).pause(0.05).perform()
    time.sleep(0.05)

    # 입력
    ActionChains(driver)\
        .key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL)\
        .send_keys(Keys.BACKSPACE)\
        .send_keys(str(value))\
        .send_keys(Keys.ENTER)\
        .perform()
    time.sleep(0.05)

def read_realgird_values(excel_path):
    wb = load_workbook(excel_path, data_only=True)
    if "입력(분석값)" in wb.sheetnames:
        ws = wb["입력(분석값)"]
    elif "입력" in wb.sheetnames:
        ws = wb["입력"]
    else:
        ws = wb[wb.sheetnames[0]]

    # 1행에서 필요한 컬럼 찾기
    header_map = {}
    for col in range(1, 60):
        v = ws.cell(row=1, column=col).value
        if not v:
            continue
        t = str(v).strip()
        header_map[t] = col

    def col_of(name):
        return header_map.get(name)

    c_start = col_of("측정시작")
    c_end   = col_of("측정 종료") or col_of("측정종료")
    c_spd   = col_of("시료흡인속도")
    c_vol   = col_of("시료채취량")

    if not (c_start and c_end and c_spd and c_vol):
        raise RuntimeError(f"입력(분석값) 1행 헤더 매칭 실패: {list(header_map.keys())}")

    per_item = {}

    # B2~B64에 항목명
    for r in range(2, 65):
        
        # ✅ A열이 비어있으면(측정 안한 항목) 패스
        a_val = ws.cell(row=r, column=1).value
        if a_val is None or str(a_val).strip() == "":
            continue
        
        item = ws.cell(row=r, column=2).value
        if not item:
            continue
        item = str(item).strip()
        if not item:
            continue

        st = ws.cell(row=r, column=c_start).value
        et = ws.cell(row=r, column=c_end).value
        spd = ws.cell(row=r, column=c_spd).value
        vol = ws.cell(row=r, column=c_vol).value

        per_item[item] = {
            "시작시간": "" if st is None else str(st).strip(),
            "종료시간": "" if et is None else str(et).strip(),
            "흡인속도": "" if spd is None else str(spd).strip(),
            "시료채취량": "" if vol is None else str(vol).strip(),
            "채취량단위": "L",
            "흡인속도단위": "L/min",
        }

    # 2) 비산먼지 계산값이 존재하면 "비산먼지" 항목에 값만 덮어쓰기
    #    (구분 로직이 아니라, 값 소스만 자동 적용)
    def to_float(x):
        try:
            return float(x)
        except:
            return None

    b2  = to_float(ws["B2"].value)   # 필요하다 했던 값
    b19 = to_float(ws["B19"].value)
    b20 = to_float(ws["B20"].value)
    i38 = ws["I38"].value
    j38 = ws["J38"].value

    if (b2 is not None) and (b19 is not None) and (b20 is not None):
        avg = (b19 + b20) / 2.0
        vol_calc = avg * b2 * 1000.0
        spd_calc = avg * 1000.0
        st_calc = "" if i38 is None else str(i38).strip()
        et_calc = "" if j38 is None else str(j38).strip()

        # 덮어쓸 후보 항목명들(사이트/엑셀 표현 차이 대비)
        dust_names = ["비산먼지"]
        for dn in dust_names:
            if dn in per_item:
                per_item[dn].update({
                    "시작시간": st_calc,
                    "종료시간": et_calc,
                    "시료채취량": f"{vol_calc:.0f}",
                    "흡인속도": f"{spd_calc:.0f}",
                    "채취량단위": "L",
                    "흡인속도단위": "L/min",
                })

    return per_item

def fill_tab2_realgird(driver, excel_path, sample_no, grid_root_css):
    date_str = sample_to_datestr(sample_no)
    if not date_str:
        raise RuntimeError("시료번호에서 날짜 파싱 실패")

    # 컬럼 찾기
    c_item = rg_find_col(driver, grid_root_css, ["측정항목"])
    c_vol  = rg_find_col(driver, grid_root_css, ["시료채취량"])
    c_spd  = rg_find_col(driver, grid_root_css, ["흡인속도"])
    c_sd   = rg_find_col(driver, grid_root_css, ["측정일(시작)"])
    c_st   = rg_find_col(driver, grid_root_css, ["시작시간"])
    c_ed   = rg_find_col(driver, grid_root_css, ["측정일(종료)"])
    c_et   = rg_find_col(driver, grid_root_css, ["종료시간"])

    unit_cols = rg_find_cols(driver, grid_root_css, "단위")
    if len(unit_cols) >= 2:
        c_vol_u = unit_cols[0]
        c_spd_u = unit_cols[1]
    else:
        c_vol_u = unit_cols[0] if unit_cols else (c_vol + 1 if c_vol is not None else None)
        c_spd_u = (c_spd + 1) if c_spd is not None else None

    need = [c_item, c_vol, c_vol_u, c_spd, c_spd_u, c_sd, c_st, c_ed, c_et]
    if any(x is None for x in need):
        raise RuntimeError(f"RealGrid 컬럼 탐색 실패: {rg_dump_headers(driver, grid_root_css)}")

    # 엑셀 -> 항목별 값
    per_item = read_realgird_values(excel_path)

    # 엑셀에 있는 항목만 처리(빠름)
    for item_name, v in per_item.items():
        tr = rg_find_tr_by_item(driver, grid_root_css, c_item, item_name)
        if not tr:
            continue

        st = v.get("시작시간", "")
        et = v.get("종료시간", "")
        vol = v.get("시료채취량", "")
        vol_u = v.get("채취량단위", "L")
        spd = v.get("흡인속도", "")
        spd_u = v.get("흡인속도단위", "L/min")

        # 예외별 붙여넣기 구성
        if item_name in SKIP_VOL_AND_SPEED:
            start_col = c_sd
            values = [date_str, st, date_str, et]   # 날짜/시간만
        elif item_name in SKIP_SPEED_ONLY:
            start_col = c_vol
            values = [vol, vol_u, "", "", date_str, st, date_str, et]  # 흡인 2칸 비움
        else:
            start_col = c_vol
            values = [vol, vol_u, spd, spd_u, date_str, st, date_str, et]

        rg_paste_to_tr(driver, tr, start_col, values)

    print("▶ RealGrid 붙여넣기 입력 완료")


#===============================팝업닫기========================

POPUP_CLOSE_SEL = "body > div > div > div.modal-body > div.modal-footer.row > form > input"




# =====================================================================
# 메인
# =====================================================================

def main():
    """매체(대기/수질) 선택 후 해당 흐름으로 분기."""
    print("=== 자동 입력 시작 ===")
    media = input("매체 선택 (1=대기 / 2=수질) [기본:1]: ").strip() or "1"
    if media == "2":
        _main_water()
    else:
        _main_air()


# ══════════════════════════════════════════════════════════════
# 수질 흐름  ─  탭4(측정분석결과)만 입력
# ══════════════════════════════════════════════════════════════
def _main_water():
    """수질 흐름: 탭4 진입 → PDF 생성/업로드 → 저장만 수행 (그리드 입력 없음)"""
    login_id = input("측정인 아이디: ").strip()
    login_pw = input("측정인 비밀번호: ").strip()

    raw = input("처리할 수질 시료번호 (쉼표/공백/줄바꿈 구분):\n").strip()
    target_samples = parse_sample_input(raw)
    if not target_samples:
        print("⚠ 시료번호 없음 → 종료"); return

    # 수질 NAS에서 파일 찾기 (NAS_BASE_WATER\YYYY년\업체명\파일.xlsm)
    items = []
    for sn in target_samples:
        p = find_sample_file_in_water_nas(sn)
        if not p:
            print(f"⚠ 수질 NAS에서 파일 못 찾음: {sn} → 스킵")
            continue
        items.append({"sample": sn, "path": p})

    if not items:
        print("⚠ 처리할 시료 없음 → 종료"); return

    print(f"총 {len(items)}개 수질 시료 처리 예정")

    do_pdf_final = ask_yesno("PDF 생성/업로드까지 할래? (예/아니오) [기본: 예]: ", default_yes=True)

    # 드라이버 시작 + 수질 사이트 로그인
    driver = init_driver()
    login(driver, login_id, login_pw, field_url=FIELD_URL_WATER)

    for item in items:
        sample = item["sample"]
        path   = item["path"]

        print(f"\n{'='*51}")
        print(f"▶ 수질 시료: {sample}")
        print(f"{'='*51}")

        if not open_detail_by_sample(driver, sample):
            print("❌ 상세페이지 진입 실패 → 다음 시료로")
            continue

        # 탭4 클릭 + 로딩 대기
        safe_click(driver, TAB4_SELECTOR)
        wait_el(driver, "#smpl_rcpt_dt", timeout=10)

        pdfs = None
        if do_pdf_final:
            print("▶ PDF 생성/업로드중")
            pdfs = make_tab4_pdfs_water(path, sample)
            upload_tab4_pdfs(
                driver,
                pdfs["pdf_analy"],
                pdfs["pdf_record"],
                f1_sel=WATER_FILE_BTN1,
                f2_sel=WATER_FILE_BTN2,
            )
            tab4_comp_save(driver)
            print(f"✅ 완료  →  {pdfs['p_out']}")
        else:
            tab4_temp_save(driver)
            print("  임시저장")

        if pdfs:
            cleanup_tmp_pdfs(PDF_TMP_DIR, sample)

        back_to_list(driver)

    print("\n=== 수질 처리 완료 ===")
    try:
        input("엔터 누르면 종료...")
    except EOFError:
        pass


# ══════════════════════════════════════════════════════════════
# 대기 흐름  ─  기존 그대로
# ══════════════════════════════════════════════════════════════
def _main_air():
    job = input("작업 선택 (1=측정인 입력/저장/PDF/백데이터 / 2=백데이터만(로그인X)) [기본:1]: ").strip() or "1"

    login_id = ""
    login_pw = ""
    if job == "1":
        login_id = input("측정인 아이디: ").strip()
        login_pw = input("측정인 비밀번호: ").strip()

    mode = input("모드 선택 (1=시료번호 직접입력 / 2=팀+날짜 자동추출) [기본:2]: ").strip() or "2"

    do_tab1       = ask_yesno("현장측정정보(탭1) 입력할래? (예/아니오) [기본: 예]: ",          default_yes=True)
    do_tab2       = ask_yesno("시료채취/측정정보(탭2) 입력할래? (예/아니오) [기본: 예]: ",      default_yes=True)
    do_pdf_upload = ask_yesno("탭2 PDF 생성/업로드까지 할래? (예/아니오) [기본: 예]: ",        default_yes=True)
    do_backdata   = ask_yesno("백데이터(수분 CSV + THC CSV/FID)? (예/아니오) [기본: 예]: ",   default_yes=True)
    do_tab4       = ask_yesno("측정분석결과(탭4) 입력할래? (예/아니오) [기본: 아니오]: ",      default_yes=False)
    do_pdf_final  = ask_yesno("PDF최종본 생성(탭4용) 할래? (예/아니오) [기본: 아니오]: ",      default_yes=False)

    if job == "2":
        do_tab1 = do_tab2 = do_pdf_upload = do_tab4 = do_pdf_final = False
        do_backdata = True

    if not any([do_tab1, do_tab2, do_backdata, do_tab4]):
        print("⚠ 실행할 항목이 없어 종료"); return

    # 시료 목록 생성
    date_groups = {}

    if mode == "1":
        raw = input("처리할 시료번호 입력 (쉼표/공백/줄바꿈 구분):\n").strip()
        target_samples = parse_sample_input(raw)
        if not target_samples:
            print("⚠ 시료번호 없음 → 종료"); return

        for sn in target_samples:
            p = find_sample_file_in_nas(sn)
            if not p:
                print(f"⚠ NAS에서 파일 못 찾음: {sn} → 스킵"); continue
            dust = ("비산먼지" in os.path.basename(p)) or ("비산" in os.path.basename(p))
            ds = sample_to_datestr(sn)
            if not ds:
                print(f"⚠ 날짜 파싱 실패: {sn} → 스킵"); continue
            date_groups.setdefault(ds, []).append({"sample": sn, "path": p, "dust": dust})

        if not date_groups:
            print("⚠ 처리 가능한 시료 없음 → 종료"); return
        total = sum(len(v) for v in date_groups.values())
        print(f"총 {total}개 시료(날짜 {len(date_groups)}개 그룹) 처리 예정")

    else:
        team_no  = input("팀 번호(1~5): ").strip()
        date_str = input("날짜(YYYY-MM-DD): ").strip()
        if not team_no.isdigit():
            print("팀 번호 오류"); return
        samples = extract_samples_from_nas(int(team_no), date_str)
        print(f"총 {len(samples)}개 시료 자동 입력 예정")
        date_groups = {date_str: samples}

    # 드라이버 / 로그인
    driver = None
    if job == "1":
        driver = init_driver()
        login(driver, login_id, login_pw)
    else:
        print("▶ 백데이터만 모드 → 브라우저 스킵")

    # 날짜 그룹별 처리
    for date_str, day_samples in date_groups.items():
        print(f"\n{'='*51}")
        print(f"▶ 날짜 검색: {date_str}  | 대상 시료: {len(day_samples)}개")
        print(f"{'='*51}")

        if job == "1":
            set_date_js(driver, "#search_meas_start_dt", date_str)
            set_date_js(driver, "#search_meas_end_dt",   date_str)
            safe_click(driver, "#btnSearch")
            wait(2)
        else:
            print("▶ 백데이터만 모드 → 날짜검색 스킵")

        for item in day_samples:
            sample  = item["sample"]
            path    = item["path"]
            is_dust = item["dust"]

            print(f"\n대기 시료: {sample}  | 비산먼지: {is_dust}")

            # 백데이터
            if do_backdata and not is_dust:
                try:
                    export_backdata_moist_thc(path, sample)
                except Exception as e:
                    print(f"⚠ 백데이터 실패: {e}")
            elif do_backdata and is_dust:
                print("▶ 비산먼지 → 백데이터 스킵")

            if job != "1":
                continue

            if not open_detail_by_sample(driver, sample):
                print("❌ 상세페이지 실패 → 다음 시료로"); continue

            excel = parse_measuring_record(path, sample)

            # 탭1
            if do_tab1:
                safe_click(driver, "#ui-id-1")
                fill_tab1(driver, excel, is_dust)
            else:
                print("▶ 탭1 스킵")

            # 탭2
            if do_tab2:
                safe_click(driver, "#ui-id-2")
                fill_tab2(driver, excel, is_dust)
                set_date_js(driver, "#meas_end_dt", date_str)

                if not is_dust:
                    fill_facility_rows(driver, parse_facility_from_excel(path))
                else:
                    print("▶ 비산먼지 → 배출시설 스킵")

                write_sampler_comment(driver)

                GRID_ROOT = "#measGridAnalySampAnzeDataAirItemList1"
                if is_dust:
                    fill_tab2_realgird_dust_only(driver, path, sample, GRID_ROOT)
                else:
                    rg2_fill_measure_grid_api(driver, sample, read_realgird_values(path))

                # PDF 업로드
                did_pdf  = False
                pdf_path = None
                if do_pdf_upload:
                    try:
                        pdf_path = export_pdf_from_excel(path, sample, is_dust)
                        print(f"▶ PDF 생성 완료: {pdf_path}")
                        if not trigger_file_dialog(driver, timeout=8):
                            raise RuntimeError("파일추가 트리거 실패")
                        time.sleep(0.8)
                        pick_file_in_open_dialog(pdf_path, timeout=30)
                        print("✅ 파일 선택 완료")
                        did_pdf = True
                    except Exception as e:
                        print(f"❌ PDF 실패: {e}")
                else:
                    print("▶ PDF 업로드 스킵")
                time.sleep(1)

                draft_by_time  = _should_draft_by_sampling_end(date_str, excel.get("채취끝", ""))
                use_final_save = did_pdf and not draft_by_time
                print(f"▶ 탭2 {'✅입력완료' if use_final_save else '⚠임시저장'}")
                ok = save_tab2(driver, use_final_save=use_final_save)

                if ok and did_pdf and pdf_path and os.path.isfile(pdf_path):
                    try:
                        os.remove(pdf_path)
                        print(f"🗑 PDF 삭제: {pdf_path}")
                    except Exception as e:
                        print(f"⚠ PDF 삭제 실패: {e}")
            else:
                print("▶ 탭2 스킵")

            # 탭4
            if do_tab4:
                safe_click(driver, TAB4_SELECTOR)
                wait_el(driver, "#smpl_rcpt_dt", timeout=10)

                print("▶ 탭4 데이터 읽는중")
                tab4_meta = read_tab4_from_macro_xlsm(sample)
                print("✅ 탭4 데이터 읽기 완료")

                fill_tab4_dates(driver, sample, tab4_meta)
                fill_tab4_grid_only(driver, tab4_meta["rows"])

                pdfs = None
                if do_pdf_final:
                    print("▶ PDF 탭4 생성중")
                    pdfs = make_tab4_pdfs(path, sample)
                    upload_tab4_pdfs(driver, pdfs["pdf_analy"], pdfs["pdf_record"])
                    print("✅ PDF 탭4 완료")

                if do_pdf_final:
                    tab4_comp_save(driver)
                    print("✅ 탭4 입력완료")
                else:
                    tab4_temp_save(driver)
                    print("⚠ 탭4임시저장")

                if pdfs:
                    cleanup_tmp_pdfs(PDF_TMP_DIR, sample)

            back_to_list(driver)

    print("\n=== 대기 처리 완료 ===")
    try:
        input("엔터 누르면 종료...")
    except EOFError:
        pass


if __name__=="__main__":
    try:
        main()
    except Exception as e:
        # 공용 에러 로그에 남기고, 원래대로 예외는 다시 올려서 콘솔/GUI에서도 확인 가능하게 유지
        log_error("eco_input.main", e)
        raise