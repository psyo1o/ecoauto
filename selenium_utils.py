# -*- coding: utf-8 -*-
"""
Selenium 웹 자동화 통합
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoAlertPresentException
import time


def init_driver():
    """Selenium 드라이버 초기화"""
    opt = webdriver.ChromeOptions()
    opt.add_experimental_option("detach", True)
    opt.add_experimental_option(
        "prefs",
        {
            "profile.default_content_setting_values.popups": 2,
            "profile.default_content_setting_values.notifications": 2,
        },
    )
    opt.add_argument("--disable-notifications")
    d = webdriver.Chrome(options=opt)
    d.maximize_window()
    return d


def safe_click(driver, selector, timeout=10):
    """안전한 클릭"""
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", el)
        time.sleep(0.2)
        driver.execute_script("arguments[0].click();", el)
        time.sleep(0.3)
        return True
    except:
        return False


def wait_el(driver, selector, timeout=10):
    """엘리먼트 대기"""
    try:
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
        )
    except:
        return None


def accept_all_alerts(driver, total_wait=8.0, poll=0.2, max_accept=10, label=""):
    """모든 alert 자동 수락"""
    end = time.time() + float(total_wait)
    accepted = 0

    while time.time() < end and accepted < max_accept:
        try:
            alert = driver.switch_to.alert
            alert.accept()
            accepted += 1
            time.sleep(0.25)
            end = time.time() + float(total_wait)
        except NoAlertPresentException:
            time.sleep(poll)
        except:
            time.sleep(poll)

    if accepted and label:
        print(f"   ↳ {accepted}회 확인({label})")
    return accepted


def set_date_js(driver, selector, value):
    """JS로 날짜 입력"""
    try:
        js = f"document.querySelector('{selector}').value='{value}';"
        driver.execute_script(js)
        time.sleep(0.1)
    except Exception as e:
        print("⚠ 날짜 입력 실패:", selector, e)


def fill_select_option(driver, selector, value):
    """Select 엘리먼트에서 옵션 선택"""
    try:
        element = driver.find_element(By.CSS_SELECTOR, selector)
        Select(element).select_by_value(value)
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", element
        )
        return True
    except:
        return False


def close_popup(driver):
    """팝업 닫기"""
    POPUP_CLOSE_SEL = "body > div > div > div.modal-body > div.modal-footer.row > form > input"
    
    if not hasattr(driver, "_main_handle"):
        try:
            driver._main_handle = driver.current_window_handle
        except:
            driver._main_handle = None

    main_handle = getattr(driver, "_main_handle", None)

    try:
        el = WebDriverWait(driver, 0.3).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, POPUP_CLOSE_SEL))
        )
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        except:
            pass

        try:
            el.click()
            return
        except:
            pass
        try:
            driver.execute_script("arguments[0].click();", el)
            return
        except:
            pass
    except TimeoutException:
        pass
    except:
        pass

    try:
        handles = list(driver.window_handles)
        if main_handle and main_handle in handles:
            for h in handles:
                if h != main_handle:
                    try:
                        driver.switch_to.window(h)
                        try:
                            el = WebDriverWait(driver, 0.3).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, POPUP_CLOSE_SEL))
                            )
                            try:
                                driver.execute_script("arguments[0].click();", el)
                            except:
                                pass
                        except:
                            pass

                        try:
                            driver.close()
                        except:
                            pass
                    except:
                        pass

            try:
                driver.switch_to.window(main_handle)
            except:
                pass
    except:
        pass