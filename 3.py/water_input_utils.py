# -*- coding: utf-8 -*-
"""
수질 현장측정분석 자동입력 유틸 (탭2·탭4)
- eco_input._main_water() 에서 호출
- 측정분석 매크로 엑셀(26.05 수질) → 사이트 field_water.do 반영

탭2: 시료용기/용량/개수/측정위치 고정값 + 입력완료(#btnSaveMsFieldDoc, 확인 2회)
탭4: 날짜(H6/H7/J7) + RealGrid 4열(측정항목)부터 엑셀 A~M + PDF(#anzeFile1/#anzeFile2) + 분석완료
"""
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium_utils import safe_click, set_date_js, wait_el, tab4_alerts_after_save, tab4_after_comp_save_confirm
from selenium_utils import accept_all_alerts as _accept_all_alerts
from excel_com_utils import get_excel_app
from realgrid_utils import rg_dump_headers
from tab4_utils import _norm_rg, tab4_find_tr_by_item, tab4_paste_row_using_tab2
from config import WATER_TAB4_MACRO_FILE, WATER_TAB4_GRID_ROOT

# ── 사이트 셀렉터 (수질 field_water.do) ─────────────────────
WATER_TAB2_SELECTOR = "a#ui-id-2"
WATER_TAB4_TAB_ID = "ui-id-4"       # 측정분석결과 탭 (탭2 입력완료 후 활성)
WATER_TAB4_SELECTOR = "a#ui-id-4"
WATER_TAB4_READY_SELECTORS = (
    "#gridAnalySampAnzeDataAirItemList1 .rg-body",
    "#meas_start_dtm",
    "#anze_end_dt",
    "#anze_start_dt",
    "#smpl_rcpt_dt",
)
WATER_DETAIL_ENTRY_SEL = "a#ui-id-2"        # 상세 진입 검증용 (탭 링크, 표시됨)
WATER_DETAIL_WAIT_SEL = "#samp_vesl_code"   # 탭2 전환 후 필드 대기

# 탭2 고정 입력값 (요청 사양)
WATER_TAB2_VESSEL = "PE"
WATER_TAB2_LITRE = "1"
WATER_TAB2_CNT = "1"
WATER_TAB2_LOC = "폭기조"

# 매크로 시트 (대기 TAB4_MACRO_FILE 과 동일 시트명·구조)
WATER_ANZE_SHEET = "00. 측정분석결과 입력샘플"
WATER_ANZE_SHEET_CANDIDATES = [
    WATER_ANZE_SHEET,
    "측정분석결과 입력샘플",
    "00.측정분석결과 입력샘플",
]

# 엑셀 ↔ RealGrid: 8행=헤더(매칭), 9~54행 A~M (대기 탭4와 동일, 항목=A열 우선)
WATER_HEADER_ROW = 8
WATER_DATA_FIRST_ROW = 9
WATER_DATA_LAST_ROW = 54
WATER_EXCEL_COLS = "ABCDEFGHIJKLM"   # A~M
WATER_GRID_PASTE_START_COL = 5
WATER_TAB4_ITEM_GRID_COL = 4     # 대분류1·중분류2·동시분석3·측정항목4


def _excel_header_index_map(excel_headers: list) -> dict:
    """엑셀 8행 헤더 → vals 배열 인덱스 (* 없음, _norm_rg 로 정규화)."""
    m = {}
    for i, eh in enumerate(excel_headers):
        k = _norm_rg(eh)
        if k and k not in m:
            m[k] = i
    return m


def _find_excel_idx_for_grid_header(gkey: str, excel_map: dict) -> int | None:
    """RealGrid 헤더(별표 제거 후)에 대응하는 엑셀 열 인덱스."""
    if not gkey:
        return None
    if gkey in excel_map:
        return excel_map[gkey]
    # 대기 realgrid_utils findByHeader 와 같이 부분 일치 허용
    for ek, ei in excel_map.items():
        if gkey in ek or ek in gkey:
            return ei
    return None


def build_water_paste_plan(grid_headers: list, excel_headers: list) -> tuple[int, list]:
    """
    RealGrid(헤더에 * 있음) ↔ 엑셀 8행( * 없음) 열 매칭.
    rg_dump_headers 는 그리드 쪽 * 를 이미 제거함.

    반환: (붙여넣기 시작 grid col idx, grid 열 순서대로 값 배열용 excel 인덱스 또는 None)
    """
    excel_map = _excel_header_index_map(excel_headers)
    paste_headers = sorted(
        [h for h in grid_headers if h.get("idx", 0) >= WATER_GRID_PASTE_START_COL],
        key=lambda h: h["idx"],
    )
    if not paste_headers:
        return WATER_GRID_PASTE_START_COL, []

    plan = []  # (grid_idx, excel_idx|None)
    for gh in paste_headers:
        gkey = _norm_rg(gh.get("text", ""))
        plan.append((gh["idx"], _find_excel_idx_for_grid_header(gkey, excel_map)))

    start_col = plan[0][0]
    return start_col, plan


def _values_for_grid_row(vals: list, plan: list) -> list:
    """열 매칭 plan 기준으로 RealGrid 한 행 분량 TSV 값 생성."""
    out = []
    for _gidx, excel_i in plan:
        if excel_i is None or excel_i >= len(vals):
            out.append("")
        else:
            v = vals[excel_i]
            out.append("" if v is None else str(v).strip())
    return out


def wait_until_water_sheet_updates(ws, timeout=15.0) -> bool:
    """A6 시료번호 입력 후 표 갱신 대기 (대기 탭4와 같이 A9·B9 모두 확인)."""
    end = time.time() + timeout
    before_a = _cell_text(ws, "A9")
    before_b = _cell_text(ws, "B9")

    while time.time() < end:
        try:
            ws.Parent.Application.Calculate()
        except Exception:
            pass
        a9 = _cell_text(ws, "A9")
        b9 = _cell_text(ws, "B9")
        if (a9 and a9 != before_a) or (b9 and b9 != before_b):
            return True
        if a9 or b9:
            return True
        time.sleep(0.2)
    return False


def _cell_text(ws, addr: str) -> str:
    try:
        return str(ws.Range(addr).Text).strip()
    except Exception:
        v = ws.Range(addr).Value
        return "" if v is None else str(v).strip()


def read_water_tab4_from_macro_xlsm(sample_no: str) -> dict:
    """
    측정인 측정분석 입력 26.05(수질).xlsm — 대기 tab4_utils 와 동일 COM·매크로 흐름.

    - Open 시 EnableEvents=True → Workbook_Open 등 매크로 실행 (대기와 동일)
    - A6 시료번호 → Calculate → B9 갱신 대기
    - 8행 B~ : RealGrid 헤더 매칭용 (엑셀엔 * 없음, 그리드 * 는 입력 시 제거)
    - 9~54행 B~ : 붙여넣기 데이터 (항목명=B열)

    반환: rcpt_dt(H6), start_src(H7), end_dt(J7), headers, rows
    """
    excel = get_excel_app()
    wb = None
    try:
        # 대기 read_tab4_from_macro_xlsm 과 동일 (Calculation 모드는 건드리지 않음)
        excel.DisplayAlerts = False
        excel.EnableEvents = True
        try:
            excel.AutomationSecurity = 1  # msoAutomationSecurityLow
        except Exception:
            pass

        wb = excel.Workbooks.Open(WATER_TAB4_MACRO_FILE, ReadOnly=True, UpdateLinks=0)

        ws = None
        for name in WATER_ANZE_SHEET_CANDIDATES:
            try:
                ws = wb.Worksheets(name)
                break
            except Exception:
                continue
        if ws is None:
            ws = wb.Worksheets(WATER_ANZE_SHEET)

        # 활성 시트가 다른 곳이어도 작업 시트를 강제로 맞춘다.
        try:
            ws.Activate()
        except Exception:
            pass

        ws.Range("A6").Value = sample_no

        try:
            excel.CalculateFullRebuild()
        except Exception:
            try:
                excel.Calculate()
            except Exception:
                pass

        ok_wait = wait_until_water_sheet_updates(ws, timeout=15.0)

        rcpt_dt = _cell_text(ws, "H6") or _cell_text(ws, "I6")
        start_src = _cell_text(ws, "H7") or _cell_text(ws, "I7")
        end_dt = _cell_text(ws, "J7") or _cell_text(ws, "K7")

        # 8행 헤더 A~M (RealGrid 헤더 매칭·로그용)
        headers = [
            _cell_text(ws, f"{c}{WATER_HEADER_ROW}") for c in WATER_EXCEL_COLS
        ]

        # 9~54행 A~M (대기 read_tab4_from_macro_xlsm 와 동일)
        rows = []
        for excel_row in range(WATER_DATA_FIRST_ROW, WATER_DATA_LAST_ROW + 1):
            try:
                if ws.Rows(excel_row).Hidden:
                    continue
            except Exception:
                pass

            row_vals = [
                _cell_text(ws, f"{c}{excel_row}") for c in WATER_EXCEL_COLS
            ]
            item = _norm_rg(row_vals[0]) if row_vals else ""
            if not item and len(row_vals) > 1:
                item = _norm_rg(row_vals[1])
            if not item:
                continue
            rows.append(row_vals)

        a9 = _cell_text(ws, "A9")
        b9 = _cell_text(ws, "B9")
        print(
            f"[수질-TAB4-EXCEL] wait={ok_wait} A9='{a9}' B9='{b9}' "
            f"header_row={WATER_HEADER_ROW} rows={len(rows)}"
        )
        return {
            "rcpt_dt": rcpt_dt,
            "start_src": start_src,
            "end_dt": end_dt,
            "headers": headers,
            "rows": rows,
        }
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass


def _col_letter(n: int) -> str:
    """1-based column index → Excel column letter."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _discover_water_grid_id(driver) -> str | None:
    """측정분석 탭 RealGrid id — .rg-root 우선 (gridMain1 등), 탭4 패널·날짜필드 근처 우선."""
    try:
        return driver.execute_script(
            r"""
            function vis(el) {
              if (!el) return false;
              const r = el.getBoundingClientRect();
              return r.width > 30 && r.height > 30;
            }
            function hasItemHeader(root) {
              const ths = root.querySelectorAll('.rg-header th');
              for (const th of ths) {
                const t = (th.innerText || '').replace(/\*/g, '').trim();
                if (t.includes('측정항목') || t === '항목' || t.includes('분석항목')) return true;
              }
              return false;
            }
            function tab4Scope() {
              const tab = document.querySelector('a#ui-id-4');
              if (tab) {
                const href = tab.getAttribute('href');
                if (href && href.charAt(0) === '#') {
                  const p = document.querySelector(href);
                  if (p) return p;
                }
              }
              return document.querySelector('.ui-tabs-panel.ui-tabs-active')
                || document.querySelector('.ui-tabs-panel[aria-hidden="false"]')
                || document.body;
            }
            const scope = tab4Scope();
            const anchor = document.querySelector('#anze_end_dt, #anze_start_dt, #smpl_rcpt_dt');
            const cands = [];
            const roots = document.querySelectorAll('.rg-root[id], [id].rg-root');
            for (const el of roots) {
              if (!el.id || !el.querySelector('.rg-header') || !vis(el)) continue;
              const inScope = scope.contains(el);
              const near = anchor && (el.compareDocumentPosition(anchor) & 16);
              const area = el.getBoundingClientRect().width * el.getBoundingClientRect().height;
              let score = area / 1000;
              if (inScope) score += 50;
              if (near) score += 30;
              if (/water|Water|Analy|anze/i.test(el.id)) score += 40;
              if (/^gridMain/i.test(el.id)) score += 35;
              if (hasItemHeader(el)) score += 60;
              cands.push({ id: el.id, score });
            }
            if (!cands.length) {
              for (const el of scope.querySelectorAll('[id]')) {
                if (!/grid|Grid|measGrid|gridMain/i.test(el.id)) continue;
                if (!el.querySelector('.rg-header') || !vis(el)) continue;
                cands.push({ id: el.id, score: 10 });
              }
            }
            cands.sort((a, b) => b.score - a.score);
            return cands.length ? cands[0].id : null;
            """
        )
    except Exception:
        return None


def _list_visible_rg_roots(driver) -> list:
    """오류 시 로그용 — 보이는 RealGrid id 목록."""
    try:
        ids = driver.execute_script(
            r"""
            const out = [];
            for (const el of document.querySelectorAll('.rg-root[id]')) {
              const r = el.getBoundingClientRect();
              if (r.width > 20 && r.height > 20) out.push(el.id);
            }
            return out;
            """
        )
        return ids or []
    except Exception:
        return []


def water_tab4_is_ready(driver) -> bool:
    """측정분석 탭 패널이 열려 날짜 필드가 보이는지."""
    for sel in WATER_TAB4_READY_SELECTORS:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            if el.is_displayed():
                return True
        except Exception:
            continue
    return False


def water_tab4_tab_clickable(driver) -> bool:
    """측정분석결과(a#ui-id-4) 탭이 클릭 가능한지 (탭2 미완료 시 비활성)."""
    try:
        el = driver.find_element(By.CSS_SELECTOR, WATER_TAB4_SELECTOR)
        if not el.is_displayed():
            return False
        if (el.get_attribute("aria-disabled") or "").lower() == "true":
            return False
        try:
            li = el.find_element(By.XPATH, "./ancestor::li[1]")
            cls = li.get_attribute("class") or ""
            if "ui-state-disabled" in cls or "disabled" in cls:
                return False
        except Exception:
            pass
        return True
    except Exception:
        return False


def open_water_tab4(driver) -> bool:
    """
    상세 화면 → 측정분석결과 탭(a#ui-id-4) 전환.
    사이트 규칙: 탭2 입력완료(#btnSaveMsFieldDoc) 전에는 탭4가 비활성·미표시일 수 있음.
    """
    if water_tab4_is_ready(driver):
        print("▶ 수질 탭4(측정분석결과) 이미 열림")
        return True

    if not water_tab4_tab_clickable(driver):
        print(
            "❌ 수질 탭4(측정분석결과, #ui-id-4) 탭이 비활성입니다.\n"
            "   ↳ 탭2 시료채취/측정정보를 입력 후 「입력완료」(#btnSaveMsFieldDoc) 저장이 필요합니다."
        )
        return False

    try:
        el = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, WATER_TAB4_SELECTOR))
        )
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
            el,
        )
        time.sleep(0.5)
        if water_tab4_is_ready(driver):
            print(f"▶ 수질 탭4 전환 완료 ({WATER_TAB4_SELECTOR})")
            return True
        print(f"   ⚠ {WATER_TAB4_SELECTOR} 클릭 후 탭4 필드(#anze_end_dt 등) 없음")
    except Exception as e:
        print(f"   ⚠ 수질 탭4 클릭 실패: {e}")

    print(
        "❌ 수질 탭4 전환 실패 — 탭1에 머물렀을 수 있습니다.\n"
        "   ↳ 탭2 입력완료 여부를 확인하세요."
    )
    return False


def resolve_water_tab4_grid_root(driver, timeout=15.0) -> str:
    """수질 탭4 RealGrid root — #gridAnalySampAnzeDataAirItemList1 (gridMain1 은 래퍼)."""
    static = [
        WATER_TAB4_GRID_ROOT,
        "#gridAnalySampAnzeDataAirItemList1",
        "#gridAnalySampAnzeDataWaterItemList1",
        "#gridMain1",
        "#gridMain2",
        "#measGridAnalySampAnzeDataWaterItemList1",
        "#gridAnalySampAnzeDataOutWaterItemList1",
        "#measGridAnalySampAnzeDataOutWaterItemList1",
    ]
    end = time.time() + timeout
    while time.time() < end:
        for sel in static:
            if not sel:
                continue
            try:
                els = driver.find_elements(By.CSS_SELECTOR, sel)
                if els and els[0].is_displayed():
                    hdr = rg_dump_headers(driver, sel)
                    if hdr:
                        print(f"   ↳ 수질 RealGrid: {sel}")
                        return sel
            except Exception:
                pass

        gid = _discover_water_grid_id(driver)
        if gid:
            sel = f'[id="{gid}"]'
            hdr = rg_dump_headers(driver, sel)
            if hdr:
                print(f"   ↳ 수질 RealGrid(자동탐색): {sel}")
                return sel
        time.sleep(0.35)

    visible = _list_visible_rg_roots(driver)
    hint = ", ".join(visible[:8]) if visible else "(화면에 .rg-root 없음)"
    raise RuntimeError(
        "수질 탭4 RealGrid root를 찾지 못았습니다.\n"
        f"보이는 그리드 id: {hint}\n"
        "config.ini 에 WATER_TAB4_GRID_ROOT = #gridMain1 등 실제 id 를 넣을 수 있습니다."
    )


def water_tab4_item_col(driver, grid_root: str) -> int:
    """RealGrid 헤더에서 '측정항목' 열 번호 (1-based)."""
    headers = rg_dump_headers(driver, grid_root)
    for h in headers:
        t = _norm_rg(h.get("text", ""))
        if not t:
            continue
        if t in ("측정항목", "분석항목", "항목") or "측정항목" in t:
            print(f"   ↳ 수질 그리드 측정항목 열: {h['idx']} ({t})")
            return h["idx"]
    print(f"   ↳ 수질 그리드 측정항목 열(기본): {WATER_TAB4_ITEM_GRID_COL}")
    return WATER_TAB4_ITEM_GRID_COL


def water_tab4_list_visible_items(driver, grid_root: str, item_col: int) -> list:
    """디버그: 현재 화면에 보이는 측정항목 목록."""
    try:
        grid = driver.find_element(By.CSS_SELECTOR, grid_root)
        body = grid.find_element(By.CSS_SELECTOR, "div.rg-body")
        trs = body.find_elements(By.CSS_SELECTOR, "table > tbody > tr")
        items = []
        for tr in trs:
            try:
                cell = tr.find_element(
                    By.CSS_SELECTOR, f"td:nth-child({item_col}) > div"
                )
                t = _norm_rg(cell.text)
                if t:
                    items.append(t)
            except Exception:
                continue
        return items
    except Exception:
        return []


def fill_water_tab4_dates(driver, tab4_meta: dict):
    if not water_tab4_is_ready(driver):
        print("⚠ 수질 탭4 날짜 입력 스킵 — 측정분석 탭이 열려 있지 않음")
        return
    rcpt = tab4_meta.get("rcpt_dt", "").strip()
    start = tab4_meta.get("start_src", "").strip()
    end = tab4_meta.get("end_dt", "").strip()
    if rcpt:
        set_date_js(driver, "#smpl_rcpt_dt", rcpt)
    if start:
        set_date_js(driver, "#anze_start_dt", start)
    if end:
        set_date_js(driver, "#anze_end_dt", end)


def fill_water_tab4_grid(driver, tab4_meta: dict):
    """
    엑셀 A~M → RealGrid 측정항목 열부터 붙여넣기 (A열=그리드 측정항목 열과 같은 시작).
    """
    if not water_tab4_is_ready(driver):
        raise RuntimeError(
            "수질 탭4 그리드 입력 불가 — 측정분석 탭이 열려 있지 않습니다. "
            "탭2 스킵 후 탭4 전환이 실패했을 수 있습니다."
        )
    grid_root = resolve_water_tab4_grid_root(driver)
    item_col = water_tab4_item_col(driver, grid_root)
    paste_col = item_col
    print(f"   ↳ 수질 그리드 붙여넣기: {paste_col}열부터 엑셀 A~M")

    rows = tab4_meta.get("rows") or []
    if not rows:
        print("⚠ 수질 탭4 엑셀 행 없음 (매크로 A6·A9 확인)")
        return

    rows_to_input = []
    for r in rows:
        item = _norm_rg(r[0]) if r else ""
        if not item and len(r) > 1:
            item = _norm_rg(r[1])
        if not item:
            continue
        values = [("" if v is None else str(v).strip()) for v in r]
        if all(v == "" for v in values):
            continue
        rows_to_input.append((item, values))

    visible = water_tab4_list_visible_items(driver, grid_root, item_col)
    print(
        f"[수질-탭4] 엑셀 {len(rows_to_input)}항목 / "
        f"화면 측정항목 샘플: {visible[:5]}{'…' if len(visible) > 5 else ''}"
    )

    ok = 0
    failed = []
    for idx, (item, values) in enumerate(rows_to_input, start=1):
        tr = tab4_find_tr_by_item(
            driver, item, max_steps=2000, grid_root=grid_root, item_col=item_col
        )
        if not tr:
            failed.append(item)
            print(f"   [미적용 {idx}] {item} (행탐색실패)")
            continue
        try:
            cell = tr.find_element(
                By.CSS_SELECTOR, f"td:nth-child({item_col}) > div"
            )
            ActionChains(driver).move_to_element(cell).click(cell).perform()
            time.sleep(0.12)
        except Exception:
            failed.append(item)
            print(f"   [미적용 {idx}] {item} (클릭실패)")
            continue

        paste_vals = values
        if not tab4_paste_row_using_tab2(driver, tr, paste_vals, start_col=paste_col):
            failed.append(item)
            print(f"   [미적용 {idx}] {item} (붙여넣기실패)")
            continue
        ok += 1
        print(f"   [입력 {idx}] {item}")

    print(f"[수질-탭4-결과] 성공 {ok}개 / 실패 {len(failed)}개")
    if failed and visible:
        print(f"   ↳ 그리드에 보이는 항목 예: {', '.join(visible[:8])}")


def fill_water_tab2(driver):
    """탭2: 용기 PE, 용량/개수 1, 측정위치 폭기조 → 입력완료."""
    try:
        el = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, WATER_TAB2_SELECTOR))
        )
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
            el,
        )
        time.sleep(0.3)
    except Exception:
        safe_click(driver, WATER_TAB2_SELECTOR)
    wait_el(driver, WATER_DETAIL_WAIT_SEL, timeout=10)

    Select(driver.find_element(By.CSS_SELECTOR, "#samp_vesl_code")).select_by_value(
        WATER_TAB2_VESSEL
    )
    for sel, val in (
        ("#samp_vesl_litre", WATER_TAB2_LITRE),
        ("#samp_vesl_cnt", WATER_TAB2_CNT),
        ("#edit_meas_loc_desc_1", WATER_TAB2_LOC),
    ):
        el = driver.find_element(By.CSS_SELECTOR, sel)
        el.clear()
        el.send_keys(val)

    print("▶ 수질 탭2 필드 입력 완료")


def save_water_tab2(driver) -> bool:
    """입력완료(#btnSaveMsFieldDoc) + 확인창 2회 (대기 탭2와 동일 패턴)."""
    print("▶ 수질 탭2 입력완료 저장")
    try:
        btn = driver.find_element(By.CSS_SELECTOR, "#btnSaveMsFieldDoc")
        driver.execute_script("arguments[0].click();", btn)
        _accept_all_alerts(driver, total_wait=2.5, poll=0.15, label="수질탭2-1차")
        _accept_all_alerts(driver, total_wait=7.0, poll=0.2, label="수질탭2-대기")
        _accept_all_alerts(driver, total_wait=2.0, poll=0.2, label="수질탭2-마무리")
        print("▶ 수질 탭2 저장 완료")
        return True
    except Exception as e:
        print(f"⚠ 수질 탭2 저장 실패: {e}")
        return False


def save_water_tab4_temp(driver):
    """임시저장(#btnTempSave) — 탭4만 선택·PDF 미업로드 시."""
    print("▶ 수질 탭4 임시저장 (#btnTempSave)")
    try:
        safe_click(driver, "#btnTempSave")
        tab4_alerts_after_save(driver, label="수질탭4임시")
        print("✅ 수질 탭4 임시저장 완료")
        return True
    except Exception as e:
        print(f"⚠ 수질 탭4 임시저장 실패: {e}")
        return False


def save_water_tab4_complete(driver):
    """분석완료(#btnCompSave) — 탭4 PDF 업로드 후만, 확인창 2회 (대기 탭2 저장 패턴)."""
    print("▶ 수질 탭4 분석완료 (#btnCompSave)")
    try:
        safe_click(driver, "#btnCompSave")
        tab4_after_comp_save_confirm(driver, label="수질탭4완료")
        print("✅ 수질 탭4 분석완료")
        return True
    except Exception as e:
        print(f"⚠ 수질 탭4 분석완료 실패: {e}")
        return False
