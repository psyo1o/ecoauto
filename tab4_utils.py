# -*- coding: utf-8 -*-
"""
탭4(측정분석결과) 입력 유틸
- eco_input.py에서 분리
- 탭4 RealGrid 탐색/입력, 날짜 입력, 저장
- 측정분석 엑셀(ANZE_XLSM) 읽기
"""
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium_utils import safe_click, set_date_js
from selenium_utils import accept_all_alerts as _accept_all_alerts
from realgrid_utils import rg_paste_to_tr_tab4
from excel_com_utils import get_excel_app
from data_utils import sample_to_datestr

# =====================================================================
# 상수
# =====================================================================
ANZE_XLSM  = r"\\192.168.10.163\측정팀\2.성적서\측정인 측정분석 입력 26.01.xlsm"
ANZE_SHEET = "00. 측정분석결과 입력샘플"

TAB4_SELECTOR  = "#ui-id-4"
TAB4_GRID_ROOT = "#gridAnalySampAnzeDataAirItemList1"


# =====================================================================
# 측정분석 엑셀 헬퍼
# =====================================================================
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


# =====================================================================
# RealGrid 탐색·입력
# =====================================================================
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
# 탭4 날짜 입력
# =====================================================================
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


# =====================================================================
# 탭4 저장
# =====================================================================
def tab4_temp_save(driver):
    # 임시저장
    safe_click(driver, "#btnTempSave")
    _accept_all_alerts(driver, total_wait=6.0, poll=0.2, label="탭4임시저장")
    
def tab4_comp_save(driver):
    # 입력완료
    safe_click(driver, "#btnCompSave")
    _accept_all_alerts(driver, total_wait=6.0, poll=0.2, label="탭4저장완료")
