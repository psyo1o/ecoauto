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

from selenium_utils import safe_click, set_date_js, close_popup, wait_el, accept_all_alerts
from config import REPORT_BASE, REPORT_WORKFLOW_DIRS, LOGIN_URL, FIELD_URL

NAS_BASE = REPORT_BASE
NAS_DIRS = REPORT_WORKFLOW_DIRS

# 로그인 페이지 ID 입력란 — 보이면 로그아웃(세션 만료) 상태
LOGIN_ID_SEL = "#user_email"
LOGIN_PW_SEL = "#login_pwd_confirm"
LOGIN_BTN_SEL = "#login"

# ======================================================================
# 로그인 / 로그아웃 감지
# ======================================================================

def _clear_and_fill_input(driver, selector: str, value: str, timeout: float = 10.0):
    """로그인 등 — 기존 값 제거 후 재입력 (세션 만료 후 재로그인용)."""
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
    )
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    except Exception:
        pass
    try:
        el.clear()
    except Exception:
        pass
    try:
        driver.execute_script(
            """
            arguments[0].value = '';
            arguments[0].dispatchEvent(new Event('input', {bubbles: true}));
            arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
            """,
            el,
        )
    except Exception:
        pass
    if value:
        el.send_keys(str(value))


def is_logged_out(driver, login_id_sel: str = LOGIN_ID_SEL) -> bool:
    """ID 입력란이 보이면 로그아웃(세션 만료) 상태."""
    try:
        el = driver.find_element(By.CSS_SELECTOR, login_id_sel)
        return el.is_displayed()
    except Exception:
        return False


def dismiss_logout_alerts(driver, rounds: int = 3):
    """로그아웃 시 연속으로 뜨는 확인창(2~3회) 처리."""
    for i in range(1, rounds + 1):
        accept_all_alerts(
            driver,
            total_wait=3.0,
            poll=0.2,
            max_accept=8,
            label=f"로그아웃확인{i}",
        )
        close_popup(driver)
        time.sleep(0.25)


def ensure_logged_in_or_recover(
    driver,
    login_id: str,
    login_pw: str,
    *,
    start_date: str | None = None,
    end_date: str | None = None,
    date_str: str | None = None,
    field_url: str = FIELD_URL,
    max_recover: int = 5,
) -> bool:
    """로그아웃 상태면 확인창 처리 후 재로그인. 로그인 유지 중이면 True."""
    recover_count = 0
    while is_logged_out(driver):
        if recover_count >= max_recover:
            print("❌ 로그아웃 복구 반복 한도 초과")
            return False
        recover_count += 1
        print(f"▶ 로그아웃 감지 (ID 입력란) → 재로그인 ({recover_count}/{max_recover})")
        recover_site_session(
            driver,
            login_id,
            login_pw,
            start_date=start_date,
            end_date=end_date,
            date_str=date_str,
            field_url=field_url,
        )
    return True


def login(driver, login_id: str, login_pw: str,
          login_url: str = LOGIN_URL,
          field_url: str = FIELD_URL):
    """
    측정인.kr 로그인 후 현장측정분석(대기) 페이지로 이동.
    ID/PW 자동 입력 실패 시 사용자에게 직접 입력 요청.
    """
    dismiss_logout_alerts(driver)

    print("[1] 로그인 페이지 이동")
    if not is_logged_out(driver):
        driver.get(login_url)
        time.sleep(2)

    try:
        _clear_and_fill_input(driver, LOGIN_ID_SEL, login_id)
        _clear_and_fill_input(driver, LOGIN_PW_SEL, login_pw)
    except Exception:
        input("ID/PW 직접 입력 후 엔터")

    try:
        driver.find_element(By.CSS_SELECTOR, LOGIN_BTN_SEL).click()
    except Exception:
        input("로그인 후 엔터")

    time.sleep(3)
    dismiss_logout_alerts(driver, rounds=2)
    close_popup(driver)

    media = "수질" if "field_water" in (field_url or "") else "대기"
    print(f"[2] 현장측정분석({media}) 이동")
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


def is_field_list_ready(driver, search_box: str = "#search_meas_mgmt_no") -> bool:
    """현장측정분석 목록 화면(시료번호 검색창)이 보이면 True."""
    if is_logged_out(driver):
        return False
    try:
        return driver.find_element(By.CSS_SELECTOR, search_box).is_displayed()
    except Exception:
        return False


def is_session_lost(driver) -> bool:
    """로그아웃됐거나 목록 화면이 아니면 세션 복구 필요."""
    if is_logged_out(driver):
        return True
    return not is_field_list_ready(driver)


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


# 시료 상세 진입 실패 시 — 목록 복귀 후 시료번호 검색부터 재시도 (eco_input / eco_check 공통)
MAX_SAMPLE_DETAIL_RETRY = 3

# ======================================================================
# 상세 페이지 진입
# ======================================================================

def _is_list_search_visible(driver, search_box: str = "#search_meas_mgmt_no") -> bool:
    """목록 화면 시료번호 검색창이 보이면 True."""
    try:
        return driver.find_element(By.CSS_SELECTOR, search_box).is_displayed()
    except Exception:
        return False


def verify_detail_page_opened(
    driver,
    detail_wait_sel: str = "#machineDiv",
    search_box: str = "#search_meas_mgmt_no",
    tab1_sel: str = "a#ui-id-1",
    timeout: float = 10.0,
) -> bool:
    """
    상세 페이지 진입 여부 검증.
    - 목록 검색창이 보이지 않을 것
    - detail_wait_sel 또는 tab1_sel 이 표시·클릭 가능할 것
    """
    markers = []
    for sel in (detail_wait_sel, tab1_sel):
        if sel and sel not in markers:
            markers.append(sel)

    deadline = time.time() + timeout
    while time.time() < deadline:
        if _is_list_search_visible(driver, search_box):
            time.sleep(0.2)
            continue
        for sel in markers:
            try:
                el = WebDriverWait(driver, 0.5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, sel))
                )
                if el.is_displayed():
                    time.sleep(0.3)
                    return True
            except Exception:
                continue
        time.sleep(0.2)
    return False


def ensure_detail_page_for_tab1(
    driver,
    detail_wait_sel: str = "#machineDiv",
    search_box: str = "#search_meas_mgmt_no",
    tab1_sel: str = "a#ui-id-1",
    timeout: float = 8.0,
) -> bool:
    """탭1 입력 직전 — 목록 화면이 아닌지·상세 탭1 요소가 준비됐는지 재확인."""
    ok = verify_detail_page_opened(
        driver,
        detail_wait_sel=detail_wait_sel,
        search_box=search_box,
        tab1_sel=tab1_sel,
        timeout=timeout,
    )
    if not ok:
        print("❌ 상세 페이지 확인 실패 (목록 화면이거나 탭1 요소 없음)")
    return ok


def _try_find_sample_and_open(
    driver,
    sample_no: str,
    max_pages: int = 5,
    detail_wait_sel: str = "#machineDiv",
    search_box: str = "#search_meas_mgmt_no",
) -> bool:
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
            if verify_detail_page_opened(
                driver, detail_wait_sel, search_box, timeout=10.0
            ):
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


def open_sample_detail(
    driver,
    sample_no: str,
    search_box: str = "#search_meas_mgmt_no",
    max_pages: int = 5,
    detail_wait_sel: str = "#machineDiv",
) -> bool:
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
    if _try_find_sample_and_open(
        driver, sample_no, max_pages, detail_wait_sel, search_box
    ):
        return True

    print(" → 1차 실패: 재검색 후 재시도")
    try:
        safe_click(driver, "#btnSearch")
        time.sleep(1.5)
    except Exception:
        pass

    # 2차 시도
    if _try_find_sample_and_open(
        driver, sample_no, max_pages, detail_wait_sel, search_box
    ):
        return True

    print(f" → 2차 실패: {sample_no} 없음")
    return False


def reopen_sample_from_search(
    driver,
    sample_no: str,
    detail_wait_sel: str = "#machineDiv",
    search_box: str = "#search_meas_mgmt_no",
    login_id: str | None = None,
    login_pw: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    date_str: str | None = None,
    field_url: str = FIELD_URL,
) -> bool:
    """목록 복귀 후 시료번호 검색·상세 진입 (시료 처리 재시도용)."""
    print(f"▶ {sample_no} — 목록 복귀 후 시료 검색부터 재시도")

    if login_id and login_pw:
        if not ensure_logged_in_or_recover(
            driver,
            login_id,
            login_pw,
            start_date=start_date,
            end_date=end_date,
            date_str=date_str,
            field_url=field_url,
        ):
            return False

    if not is_field_list_ready(driver):
        go_back_to_list(driver)
        time.sleep(0.8)
    if not is_field_list_ready(driver):
        if login_id and login_pw and is_logged_out(driver):
            if not ensure_logged_in_or_recover(
                driver,
                login_id,
                login_pw,
                start_date=start_date,
                end_date=end_date,
                date_str=date_str,
                field_url=field_url,
            ):
                return False
        if not is_field_list_ready(driver):
            print("❌ 목록 화면 복귀 실패")
            return False
    return open_sample_detail(
        driver,
        sample_no,
        search_box=search_box,
        detail_wait_sel=detail_wait_sel,
    )


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
# 세션 만료 복구 (eco_input / eco_check 공통)
# ======================================================================

def recover_site_session(
    driver,
    login_id: str,
    login_pw: str,
    start_date: str | None = None,
    end_date: str | None = None,
    date_str: str | None = None,
    field_url: str = FIELD_URL,
):
    """세션 만료(로그아웃) 시 확인창 처리 후 로그인·날짜 검색까지 복구."""
    if is_logged_out(driver):
        print("▶ 로그아웃 감지 (ID 입력란 표시) → 확인창 처리 후 재로그인")
    else:
        print("▶ 로그인 세션 만료 → 확인창 처리 후 재로그인")

    dismiss_logout_alerts(driver, rounds=3)

    login(driver, login_id, login_pw, field_url=field_url)
    sd = start_date or date_str
    ed = end_date or date_str or start_date
    if sd and ed:
        search_date(driver, sd, ed)


def open_detail_with_session_recovery(
    driver,
    sample_no: str,
    login_id: str,
    login_pw: str,
    start_date: str | None = None,
    end_date: str | None = None,
    date_str: str | None = None,
    field_url: str = FIELD_URL,
    max_recover: int = 5,
    detail_wait_sel: str = "#machineDiv",
) -> bool:
    """상세 진입 실패 시 로그아웃이면 재로그인 후 같은 시료부터 재시도."""
    recover_count = 0
    while True:
        if not ensure_logged_in_or_recover(
            driver,
            login_id,
            login_pw,
            start_date=start_date,
            end_date=end_date,
            date_str=date_str,
            field_url=field_url,
            max_recover=max_recover,
        ):
            return False

        if open_sample_detail(driver, sample_no, detail_wait_sel=detail_wait_sel):
            return True

        if is_field_list_ready(driver):
            return False

        if recover_count >= max_recover:
            print("❌ 세션 복구 반복 한도 초과 → 상세 진입 중단")
            return False

        recover_count += 1
        print(f"▶ 상세 진입 실패 → 세션 복구 후 재시도 ({recover_count}/{max_recover})")
        recover_site_session(
            driver,
            login_id,
            login_pw,
            start_date=start_date,
            end_date=end_date,
            date_str=date_str,
            field_url=field_url,
        )


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
