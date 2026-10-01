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
import os
import re
import subprocess
import time
import urllib.request

# eco_input / eco_check 가 같은 크롬 창을 이어 쓰기 위한 원격 디버깅 포트·전용 프로필
REUSE_DEBUG_PORT = 9322
REUSE_PROFILE_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"),
    "ecoauto",
    "chrome_profile",
)
_LAST_LOGIN_FILE = os.path.join(os.path.dirname(REUSE_PROFILE_DIR), "last_login_id.txt")


def wait(sec):
    """공통 대기 — eco_input, eco_check 등에서 time.sleep 대신 사용"""
    time.sleep(sec)


def _chrome_options():
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
    return opt


def _raise_policy_block(cause: BaseException) -> None:
    """WinError 4551 등 앱 제어 정책으로 chromedriver.exe 실행이 막힐 때."""
    raise OSError(
        "ChromeDriver 실행이 Windows '애플리케이션 제어 정책'에 막혔습니다.\n"
        "(대기·수질 모두 같은 Chrome 드라이버를 사용합니다.)\n\n"
        "조치:\n"
        "1) PC에서 0.처음사용시\\드라이버 자동설치(크롬).py 실행\n"
        "2) 새 터미널·프로그램 재실행 또는 PC 재로그인\n"
        "3) 계속되면 IT에 chromedriver.exe / Google Chrome 실행 허용 요청\n\n"
        f"원인: {cause}"
    ) from cause


def get_last_login_id() -> str:
    try:
        with open(_LAST_LOGIN_FILE, encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return ""


def save_last_login_id(login_id: str) -> None:
    try:
        os.makedirs(os.path.dirname(_LAST_LOGIN_FILE), exist_ok=True)
        with open(_LAST_LOGIN_FILE, "w", encoding="utf-8") as f:
            f.write(str(login_id or "").strip())
    except Exception:
        pass


def _debug_port_alive(port: int = REUSE_DEBUG_PORT, timeout: float = 1.0) -> bool:
    try:
        with urllib.request.urlopen(
            f"http://127.0.0.1:{port}/json/version", timeout=timeout
        ) as r:
            return r.status == 200
    except Exception:
        return False


def _find_chrome_exe() -> str | None:
    candidates = []
    for env in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        base = os.environ.get(env)
        if base:
            candidates.append(os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"))
    for p in candidates:
        if os.path.isfile(p):
            return p
    try:
        import winreg

        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(
                    root, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"
                ) as k:
                    p = winreg.QueryValue(k, None)
                    if p and os.path.isfile(p):
                        return p
            except OSError:
                continue
    except Exception:
        pass
    return None


def _launch_debug_chrome(port: int = REUSE_DEBUG_PORT) -> bool:
    """원격 디버깅 포트를 연 크롬을 GUI 종료 후에도 남도록 독립 프로세스로 실행."""
    exe = _find_chrome_exe()
    if not exe:
        print("   ⚠ chrome.exe 위치를 찾지 못해 브라우저 재사용 불가", flush=True)
        return False
    os.makedirs(REUSE_PROFILE_DIR, exist_ok=True)
    args = [
        exe,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={REUSE_PROFILE_DIR}",
        "--disable-notifications",
        "--no-first-run",
        "--no-default-browser-check",
        "--start-maximized",
        "about:blank",
    ]
    flags = 0
    if os.name == "nt":
        flags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
    try:
        subprocess.Popen(
            args,
            creationflags=flags,
            close_fds=True,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as e:
        print(f"   ⚠ 크롬 실행 실패: {e}", flush=True)
        return False
    t0 = time.time()
    while time.time() - t0 < 20:
        if _debug_port_alive(port):
            return True
        time.sleep(0.5)
    print("   ⚠ 크롬 원격 디버깅 포트 응답 없음", flush=True)
    return False


def _tidy_reused_windows(d) -> None:
    """이어 쓰는 창: 알림창 닫고 탭 하나만 남김."""
    try:
        d.switch_to.alert.accept()
    except Exception:
        pass
    try:
        handles = list(d.window_handles)
        if not handles:
            d.switch_to.new_window("tab")
            return
        keep = handles[0]
        for h in handles[1:]:
            try:
                d.switch_to.window(h)
                d.close()
            except Exception:
                pass
        d.switch_to.window(keep)
    except Exception:
        pass


def _init_reused_driver():
    """열려 있는 재사용 크롬에 붙고, 없으면 새로 띄운 뒤 붙는다. 실패 시 None."""
    reused = _debug_port_alive()
    if reused:
        print("   ↳ 기존 브라우저에 연결합니다 (이어서 작업)", flush=True)
    elif not _launch_debug_chrome():
        return None

    opt = webdriver.ChromeOptions()
    opt.debugger_address = f"127.0.0.1:{REUSE_DEBUG_PORT}"
    try:
        d = _start_webdriver(opt)
    except Exception as e:
        print(f"   ⚠ 재사용 브라우저 연결 실패 → 새 브라우저로 진행: {e}", flush=True)
        return None
    d._eco_reused = reused
    _tidy_reused_windows(d)
    try:
        d.maximize_window()
    except Exception:
        pass
    return d


def init_driver():
    """Selenium Chrome 드라이버 초기화.

    전용 프로필 크롬(원격 디버깅 포트)을 재사용해 이전 작업 창에서 이어서 진행한다.
    재사용이 안 되면 예전처럼 새 크롬을 띄운다.
    """
    d = _init_reused_driver()
    if d is not None:
        return d
    d = _start_webdriver(_chrome_options())
    d._eco_reused = False
    d.maximize_window()
    return d


def _start_webdriver(opt):
    """PATH의 chromedriver(네트워크·C:\\chromedriver 등)가 정책에 막히는 PC는
    webdriver-manager가 사용자 폴더(.wdm)에 받은 드라이버를 우선 사용한다.
    """
    last_err = None

    try:
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager

        driver_path = ChromeDriverManager().install()
        print(f"   ↳ ChromeDriver: {driver_path}", flush=True)
        return webdriver.Chrome(service=Service(driver_path), options=opt)
    except ImportError:
        pass
    except OSError as e:
        if getattr(e, "winerror", None) == 4551:
            _raise_policy_block(e)
        last_err = e
        print(f"   ⚠ webdriver-manager 경로 실패: {e}", flush=True)
    except Exception as e:
        last_err = e
        print(f"   ⚠ webdriver-manager 경로 실패: {e}", flush=True)

    try:
        return webdriver.Chrome(options=opt)
    except OSError as e:
        if getattr(e, "winerror", None) == 4551:
            _raise_policy_block(e)
        raise
    except Exception as e:
        if last_err:
            raise RuntimeError(
                f"Chrome 드라이버를 시작하지 못했습니다.\n"
                f"webdriver-manager: {last_err}\n"
                f"Selenium 기본: {e}"
            ) from e
        raise


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


_ALERT_LOG: list = []

# 확인 질문('?' 포함)은 제외하고, 이 표현이 들어간 alert는 저장 실패로 본다
SAVE_ERROR_RE = re.compile(
    r"입력하세요|입력해\s*주|입력하십시오|선택하세요|선택해\s*주|선택하십시오"
    r"|필수|오류|에러|실패|올바르지|잘못|불가|초과|누락"
)


def reset_alert_log() -> None:
    _ALERT_LOG.clear()


def save_alert_errors() -> list:
    """reset_alert_log() 이후 수락한 alert 중 오류로 보이는 문구."""
    return [t for t in _ALERT_LOG if t and "?" not in t and SAVE_ERROR_RE.search(t)]


def check_save_alerts(label: str) -> bool:
    errs = save_alert_errors()
    for t in errs:
        print(f"❌ {label} 저장 오류 확인창: {t}")
    return not errs


def accept_all_alerts(driver, total_wait=8.0, poll=0.2, max_accept=10, label=""):
    """모든 alert 자동 수락 (문구는 _ALERT_LOG에 기록)"""
    end = time.time() + float(total_wait)
    accepted = 0

    while time.time() < end and accepted < max_accept:
        try:
            alert = driver.switch_to.alert
            try:
                text = (alert.text or "").strip()
            except Exception:
                text = ""
            if text:
                _ALERT_LOG.append(text)
                print(f"   ↳ 확인창: {text.replace(chr(10), ' ')[:120]}")
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


# 탭4 재저장 시 페이지 내 「수정 사유」 (팝업창 아님, 동일 브라우저 탭)
TAB4_UPDATE_REASON_SEL = "#update_reason"
TAB4_UPDATE_REASON_SAVE_SEL = "#btnSaveupdateReason"
TAB4_UPDATE_REASON_DEFAULT = "오기 수정합니다."


def fill_tab4_update_reason_if_present(
    driver,
    reason: str = TAB4_UPDATE_REASON_DEFAULT,
    textarea_sel: str = TAB4_UPDATE_REASON_SEL,
    save_sel: str = TAB4_UPDATE_REASON_SAVE_SEL,
    wait_sec: float = 5.0,
) -> bool:
    """수정 사유 textarea가 보이면 입력 후 #btnSaveupdateReason 클릭. 없으면 False."""
    end = time.time() + float(wait_sec)
    while time.time() < end:
        try:
            ta = None
            for el in driver.find_elements(By.CSS_SELECTOR, textarea_sel):
                try:
                    if el.is_displayed():
                        ta = el
                        break
                except Exception:
                    continue
            if ta is not None:
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});", ta
                )
                time.sleep(0.15)
                driver.execute_script(
                    """
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input', {bubbles: true}));
                    arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
                    """,
                    ta,
                    reason,
                )
                time.sleep(0.2)
                btn = WebDriverWait(driver, 4).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, save_sel))
                )
                driver.execute_script("arguments[0].click();", btn)
                print(f"   ↳ 수정 사유 입력·저장: {reason}")
                time.sleep(0.35)
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False


def tab4_alerts_after_save(driver, label: str = "탭4"):
    """저장 클릭 직후 확인 alert만 (임시저장 등)."""
    accept_all_alerts(driver, total_wait=2.5, poll=0.15, label=f"{label}-1차")
    accept_all_alerts(driver, total_wait=7.0, poll=0.2, label=f"{label}-대기")
    accept_all_alerts(driver, total_wait=2.0, poll=0.2, label=f"{label}-마무리")


def tab4_after_comp_save_confirm(driver, label: str = "탭4"):
    """분석완료(#btnCompSave) 후 처리.
    1) 확인창 처리(탭2와 동일 3단계)
    2) 수정 사유 창이 뜨면 입력/저장
    3) 저장 후 확인창 재처리(탭2와 동일 3단계)
    """
    # 1) 먼저 분석완료 확인창을 처리
    tab4_alerts_after_save(driver, f"{label}-초기")

    # 2) 재수정 케이스면 수정 사유 입력/저장
    wrote_reason = fill_tab4_update_reason_if_present(driver, wait_sec=1.8)

    # 3) 수정 사유 저장으로 추가 alert가 뜰 수 있어 동일 패턴으로 다시 처리
    if wrote_reason:
        tab4_alerts_after_save(driver, f"{label}-사유저장")


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