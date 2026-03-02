# -*- coding: utf-8 -*-
"""
측정인.kr 공통 자동화 유틸리티
- eco_check.py / eco_input.py 양쪽에서 쓰이는 공통 로직을 추출
- 로그인, 날짜검색, 상세페이지 진입, 목록 복귀, Grid 대기 등
"""

import os
import re
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from selenium_utils import safe_click, set_date_js, close_popup, wait_el


# ======================================================================
# URL 상수 (공유)
# ======================================================================
LOGIN_URL = "https://측정인.kr/init.go"
FIELD_URL = "https://측정인.kr/ms/field_outair.do"

NAS_BASE = r"\\192.168.10.163\측정팀\2.성적서"
NAS_DIRS = ["0 0.입력중", "0 1.완료", "0 2.검토중", "0 3.검토완료",
            "0 4.출력완료&에코랩입력중", "0 5.최종완료"]
# ======================================================================
# 로그인
# ======================================================================

def login(driver, login_id: str, login_pw: str,
          login_url: str = LOGIN_URL,
          field_url: str = FIELD_URL):
    """
    측정인.kr 로그인 후 현장측정분석(대기) 페이지로 이동.
    ID/PW 자동 입력 실패 시 사용자에게 직접 입력 요청.
    """
    print("[1] 로그인 페이지 이동")
    driver.get(login_url)
    time.sleep(2)

    try:
        driver.find_element(By.CSS_SELECTOR, "#user_email").send_keys(login_id)
        driver.find_element(By.CSS_SELECTOR, "#login_pwd_confirm").send_keys(login_pw)
    except Exception:
        input("ID/PW 직접 입력 후 엔터")

    try:
        driver.find_element(By.CSS_SELECTOR, "#login").click()
    except Exception:
        input("로그인 후 엔터")

    time.sleep(3)
    close_popup(driver)

    print("[2] 현장측정분석(대기) 이동")
    driver.get(field_url)
    time.sleep(2)


# ======================================================================
# 날짜 검색
# ======================================================================

def search_date(driver, start_date: str, end_date: str,
                start_sel: str = "#search_meas_start_dt",
                end_sel: str = "#search_meas_end_dt",
                btn_sel: str = "#btnSearch",
                wait_sec: float = 2.0):
    """
    날짜 범위로 목록 검색.
    start_date / end_date: 'YYYY-MM-DD' 형식
    """
    print(f"[3] 날짜 검색: {start_date} ~ {end_date}")
    set_date_js(driver, start_sel, start_date)
    set_date_js(driver, end_sel, end_date)
    safe_click(driver, btn_sel)
    time.sleep(wait_sec)


# ======================================================================
# RealGrid 로딩 대기 / 현재 페이지 시료번호 수집
# ======================================================================

def wait_grid_loaded(driver, timeout: float = 10, warn_msg: str = "⚠ RealGrid 느림"):
    """RealGrid 행이 렌더될 때까지 대기."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR, ".rg-renderer")) > 0
        )
        time.sleep(0.3)
    except Exception:
        print(warn_msg)


def get_samples_current_page(driver, prefix: str = "A") -> list:
    """
    현재 페이지 RealGrid에서 시료번호(prefix + 숫자-숫자 형태) 추출.
    prefix 기본값 'A' (A25xxxx-xx 등)
    """
    wait_grid_loaded(driver)
    cells = driver.find_elements(By.CSS_SELECTOR, ".rg-renderer")
    arr = []
    for c in cells:
        t = c.text.strip()
        if t.startswith(prefix) and "-" in t:
            arr.append(t)
    return list(dict.fromkeys(arr))


# ======================================================================
# 상세 페이지 진입
# ======================================================================

def _try_find_sample_and_open(driver, sample_no: str, max_pages: int = 5) -> bool:
    """
    현재 페이지부터 max_pages까지 이동하며 sample_no를 더블클릭 → 상세 진입.
    찾으면 True, 못 찾으면 False 반환.
    (내부 헬퍼 - 직접 호출보다 open_sample_detail 사용 권장)
    """
    xp = f"//div[contains(@class,'rg-renderer') and normalize-space()='{sample_no}']"

    for _ in range(max_pages):
        try:
            cell = WebDriverWait(driver, 1).until(
                EC.element_to_be_clickable((By.XPATH, xp))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", cell)
            time.sleep(0.15)
            ActionChains(driver).double_click(cell).perform()
            wait_el(driver, "#machineDiv", timeout=10)
            time.sleep(0.3)
            return True
        except Exception:
            pass

        # 다음 페이지로
        try:
            nxt_i = driver.find_element(By.CSS_SELECTOR, "i.fa.fa-chevron-right")
            nxt_a = nxt_i.find_element(By.XPATH, "./ancestor::a[1]")
        except Exception:
            return False

        nxt_a.click()
        time.sleep(1)

    return False


def open_sample_detail(driver, sample_no: str,
                       search_box: str = "#search_meas_mgmt_no",
                       max_pages: int = 5) -> bool:
    """
    시료번호로 검색 후 상세페이지 진입 (최대 2회 시도).
    성공하면 True, 실패하면 False.
    """
    print(f"\n[상세 진입] {sample_no}")
    wait_grid_loaded(driver)

    # 시료번호 검색창 입력
    try:
        inp = driver.find_element(By.CSS_SELECTOR, search_box)
        inp.clear()
        inp.send_keys(sample_no)
        time.sleep(0.3)
    except Exception:
        print("❌ 검색창 못 찾음")
        return False

    safe_click(driver, "#btnSearch")
    time.sleep(1.5)

    # 1차 시도
    if _try_find_sample_and_open(driver, sample_no, max_pages):
        return True

    print(" → 1차 실패: 재검색 후 재시도")
    try:
        safe_click(driver, "#btnSearch")
        time.sleep(1.5)
    except Exception:
        pass

    # 2차 시도
    if _try_find_sample_and_open(driver, sample_no, max_pages):
        return True

    print(f" → 2차 실패: {sample_no} 없음")
    return False


# ======================================================================
# 목록으로 복귀
# ======================================================================

def go_back_to_list(driver,
                    btn_selectors: list = None,
                    grid_wait_sel: str = ".rg-renderer",
                    timeout: float = 8.0) -> bool:
    """
    상세 페이지에서 목록으로 복귀.
    btn_selectors: 시도할 버튼 CSS 셀렉터 리스트 (순서대로 fallback)
    기본값: eco_check 스타일(취소버튼) + eco_input 스타일(목록/완료버튼) 모두 커버
    """
    if btn_selectors is None:
        btn_selectors = [
            "#btnMsFieldDocCancel",                                  # eco_input 기본
            "#btnGoList",                                            # eco_input 탭4
            "#t3 > div:nth-child(2) > div > button.btn.btnCancel",  # eco_check 스타일
        ]

    btn = None
    for sel in btn_selectors:
        try:
            btn = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, sel))
            )
            break
        except Exception:
            continue

    if btn is None:
        print("⚠ 목록 복귀 버튼 없음 (이미 목록일 수 있음)")
        return True

    try:
        driver.execute_script("arguments[0].scrollIntoView(true);", btn)
        time.sleep(0.2)
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(0.8)
    except Exception as e:
        print(f"⚠ 목록 버튼 클릭 실패: {e}")
        return False

    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, grid_wait_sel))
        )
        time.sleep(0.5)
        print("✅ 목록 복귀 완료")
        return True
    except Exception:
        print("⚠ 목록 화면 로딩 실패")
        return False


# ======================================================================
# NAS 파일 기반 시료번호 수집
# ======================================================================

def collect_samples_from_files(start_date: str,
                               nas_base: str = None,
                               nas_dirs: list = None) -> list:
    """
    날짜(YYYY-MM-DD)에 해당하는 NAS 파일명에서 시료번호 목록 추출.
    정확한 패턴: A + YYMMDD(6자리) + 팀번호(1자리) + '-' + 2자리 측정번호
    """
    yyyymmdd = start_date.replace("-", "")
    yymmdd = yyyymmdd[2:]  # '2025-12-06' → '251206'
    pattern = re.compile(rf"^(A{yymmdd}\d-\d{{2}})", re.IGNORECASE)

    result = []

    for d in (nas_dirs or []):
        folder = os.path.join(nas_base or "", d)
        if not os.path.isdir(folder):
            continue

        for root, dirs, files in os.walk(folder):
            for f in files:
                if f.startswith("~$"):
                    continue
                low = f.lower()
                if not (low.endswith(".xlsx") or low.endswith(".xlsm")):
                    continue
                m = pattern.match(f)
                if m:
                    sn = m.group(1)
                    if sn not in result:
                        result.append(sn)

    return sorted(result)
