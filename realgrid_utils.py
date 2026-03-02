# -*- coding: utf-8 -*-
"""
RealGrid 조작 통합 함수
"""

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import pyperclip
from pywinauto.keyboard import send_keys


def rg_dump_headers(driver, grid_root_css: str):
    """RealGrid 헤더 정보 조회 (별표 제거 로직 추가)"""
    js = r"""
    const root = document.querySelector(arguments[0]);
    if (!root) return [];
    const head = root.querySelector('.rg-header table thead');
    if (!head) return [];
    const ths = Array.from(head.querySelectorAll('th'));
    const out = [];
    for (let i=0; i<ths.length; i++){
        // 텍스트에서 별표(*)를 제거하고 앞뒤 공백을 정리(trim)합니다.
        let t = (ths[i].innerText || '').replace(/\*/g, '').trim();
        out.push({idx: i+1, text: t});
    }
    return out;
    """
    return driver.execute_script(js, grid_root_css)


def rg_find_col(driver, grid_root_css: str, names):
    """헤더 텍스트와 일치하는 열 idx 찾기"""
    headers = rg_dump_headers(driver, grid_root_css)
    name_set = set([n.strip() for n in names])
    for h in headers:
        if h["text"] in name_set:
            return h["idx"]
    return None


def rg_find_cols(driver, grid_root_css: str, name: str):
    """특정 헤더 텍스트와 일치하는 열 idx들 모두 반환"""
    headers = rg_dump_headers(driver, grid_root_css)
    return [h["idx"] for h in headers if h["text"] == name]


def rg_find_tr_by_item(driver, grid_root_css: str, c_item: int, item_name: str,
                       step=700, max_tries=140):
    """항목명으로 행(tr) 찾기"""
    root = driver.find_element(By.CSS_SELECTOR, grid_root_css)
    body = rg_get_body(driver, grid_root_css)
    
    rg_scroll_top(driver, grid_root_css)
    
    last_top = -1
    for _ in range(max_tries):
        xp = (
            f".//div[contains(@class,'rg-body')]//tbody/tr"
            f"/td[{c_item}]//div[contains(@class,'rg-renderer') and normalize-space()='{item_name}']"
        )
        cells = root.find_elements(By.XPATH, xp)
        if cells:
            return cells[0].find_element(By.XPATH, "./ancestor::tr[1]")
        
        driver.execute_script(
            "arguments[0].scrollTop = arguments[0].scrollTop + arguments[1];",
            body, step
        )
        time.sleep(0.05)
        
        cur_top = driver.execute_script("return arguments[0].scrollTop;", body)
        if cur_top == last_top:
            break
        last_top = cur_top
    
    return None


def rg_get_body(driver, grid_root_css: str):
    """RealGrid body 엘리먼트 가져오기"""
    root = driver.find_element(By.CSS_SELECTOR, grid_root_css)
    return root.find_element(By.CSS_SELECTOR, "div.rg-body")


def rg_scroll_top(driver, grid_root_css: str):
    """RealGrid 맨 위로 스크롤"""
    body = rg_get_body(driver, grid_root_css)
    driver.execute_script("arguments[0].scrollTop = 0;", body)
    time.sleep(0.05)


def rg_list_items(driver, grid_root_css: str, c_item: int):
    """현재 표시되는 항목들 목록"""
    root = driver.find_element(By.CSS_SELECTOR, grid_root_css)
    body = root.find_element(By.CSS_SELECTOR, ".rg-body table tbody")
    rows = body.find_elements(By.CSS_SELECTOR, "tr")
    items = []
    for row in rows:
        try:
            cell = row.find_element(By.XPATH, f"./td[{c_item}] > div")
            items.append(cell.text.strip())
        except:
            pass
    return items


def rg_paste_to_tr(driver, tr, start_col: int, values):
    """한 행에서 start_col부터 값들을 탭으로 붙여넣기"""
    tsv = "\t".join("" if v is None else str(v) for v in values)
    
    td = tr.find_element(By.XPATH, f"./td[{start_col}]")
    ActionChains(driver).move_to_element(td).click(td).pause(0.03).double_click(td).pause(0.03).perform()
    time.sleep(0.03)
    
    pyperclip.copy(tsv)
    try:
        ae = driver.switch_to.active_element
        ae.send_keys(Keys.CONTROL, "v")
        time.sleep(0.03)
        ae.send_keys(Keys.ENTER)
    except:
        send_keys("^v")
        time.sleep(0.03)
        send_keys("{ENTER}")
    
    time.sleep(0.03)


def rg_set_cell_by_keys(driver, tr, col_idx: int, text: str):
    """셀을 클릭해서 편집 → text 입력 → Enter로 커밋"""
    td = tr.find_element(By.XPATH, f"./td[{col_idx}]")
    ActionChains(driver).move_to_element(td).click(td).pause(0.03).perform()
    time.sleep(0.03)
    
    ae = driver.switch_to.active_element
    
    try:
        ae.send_keys(Keys.ALT, Keys.ARROW_DOWN)
        time.sleep(0.02)
    except:
        pass
    
    ae.send_keys(Keys.CONTROL, "a")
    ae.send_keys(text)
    ae.send_keys(Keys.ENTER)
    time.sleep(0.03)
    
    
# =====================================================================
# RealGrid API 공통 자바스크립트 로직 (그리드 탐색 및 헤더 매핑)
# =====================================================================
_CORE_REALGRID_JS = r"""
const candidates = [];
if (window.measGridViews){
  for (const k of Object.keys(window.measGridViews)){
    try{
      const v = window.measGridViews[k];
      if (v && typeof v.getColumns === "function") candidates.push([`measGridViews.${k}`, v]);
    }catch(e){}
  }
}
function scoreGridView(g){
  try{
    const cols = g.getColumns();
    const getHeader = (c) => (c.header && (c.header.text || c.header.label)) || "";
    let score = 0;
    for (const c of cols){
      const h = getHeader(c);
      if (h.includes("측정항목")) score += 10;
      if (h.includes("시료채취량")) score += 2;
      if (h.includes("시작시간")) score += 1;
    }
    return score;
  }catch(e){ return -1; }
}
let best = {score:-1, name:null, gv:null};
for (const [name, g] of candidates){
  const s = scoreGridView(g);
  if (s > best.score) best = {score:s, name, gv:g};
}
let gv = best.gv || (typeof window.getGridViewAir === "function" ? window.getGridViewAir() : null);
if (!gv) throw new Error("RealGrid gridView 못 찾음");
const dp = gv.getDataSource ? gv.getDataSource() : gv._dataProvider;
if (!dp) throw new Error("RealGrid dataProvider 못 찾음");

const cols = gv.getColumns();
const getHeader = (c) => (c.header && (c.header.text || c.header.label)) || "";
const findByHeader = (name) => {
  for (const c of cols) {
    // 별표(*) 제거 및 공백 정리
    const h = (getHeader(c) || "").replace(/\*/g, "").trim();
    if (h.indexOf(name) !== -1) return c.fieldName || c.name;
  }
  return null;
};

// 공통 컬럼 매핑
const f_item = findByHeader("측정항목") || "anze_item";
const f_vol  = findByHeader("시료채취량");
const f_spd  = findByHeader("흡인속도");
const f_sd   = findByHeader("측정일(시작)");
const f_st   = findByHeader("시작시간");
const f_ed   = findByHeader("측정일(종료)");
const f_et   = findByHeader("종료시간");

const unitCols = cols.filter(c => getHeader(c).includes("단위")).map(c => (c.fieldName || c.name));
const f_vol_u = unitCols.length >= 1 ? unitCols[0] : null;
const f_spd_u = unitCols.length >= 2 ? unitCols[1] : null;

const n = dp.getRowCount();
"""

# =====================================================================
# API 호출 함수
# =====================================================================

def rg_api_read_data(driver, grid_root_css: str) -> dict:
    """[eco_check용] RealGrid의 데이터를 DataProvider API를 통해 전부 읽어옵니다."""
    # 공통 JS 뒤에 '읽기' 전용 로직만 덧붙임
    js = _CORE_REALGRID_JS + r"""
    const out = {};
    for (let r=0; r<n; r++){
      const item = (dp.getValue(r, f_item) ?? "").toString().trim();
      if (!item) continue;
      out[item] = {
        sd: dp.getValue(r, f_sd), st: dp.getValue(r, f_st),
        ed: dp.getValue(r, f_ed), et: dp.getValue(r, f_et),
        vol: dp.getValue(r, f_vol), vol_u: dp.getValue(r, f_vol_u),
        spd: dp.getValue(r, f_spd), spd_u: dp.getValue(r, f_spd_u),
      };
    }
    return {data: out};
    """
    res = driver.execute_script(js)
    return res.get("data", {}) if isinstance(res, dict) else {}


def rg_api_write_data(driver, payload: list) -> dict:
    """[eco_input용] RealGrid에 DataProvider API를 통해 데이터를 일괄 입력합니다."""
    # 인자(payload)를 받고 공통 JS 뒤에 '쓰기' 전용 로직만 덧붙임
    js = r"const rows = arguments[0];" + "\n" + _CORE_REALGRID_JS + r"""
    const indexByItem = new Map();
    for (let r=0; r<n; r++){
      const name = (dp.getValue(r, f_item) ?? "").toString().trim();
      if (name) indexByItem.set(name, r);
    }

    let applied = 0;
    const missing = [];
    for (const row of rows){
      const r = indexByItem.get(row.item);
      if (r === undefined){ missing.push(row.item); continue; }
      
      dp.setValue(r, f_sd, row.sd); dp.setValue(r, f_st, row.st);
      dp.setValue(r, f_ed, row.ed); dp.setValue(r, f_et, row.et);
      
      if (row.vol !== "" && row.vol !== null) { dp.setValue(r, f_vol, row.vol); dp.setValue(r, f_vol_u, row.vol_u); } 
      else { dp.setValue(r, f_vol, ""); dp.setValue(r, f_vol_u, ""); }
      
      if (row.spd !== "" && row.spd !== null) { dp.setValue(r, f_spd, row.spd); dp.setValue(r, f_spd_u, row.spd_u); } 
      else { dp.setValue(r, f_spd, ""); dp.setValue(r, f_spd_u, ""); }
      
      applied++;
    }
    if (gv.commit) gv.commit(true);
    return { applied, missingCount: missing.length, missing };
    """
    return driver.execute_script(js, payload)