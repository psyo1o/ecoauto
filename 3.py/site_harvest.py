# -*- coding: utf-8 -*-
"""
측정인.kr 화면 요소 수집 (읽기 전용 — 저장·입력 버튼은 누르지 않음)

재사용 크롬(selenium_utils.init_driver)에 붙어 목록 화면과 임의 시료 상세의 탭별
입력칸·버튼·Select2·RealGrid 정보를 JSON으로 저장하고, 코드에 쓰인 #id 선택자와 대조한다.

사용: py -3.11 3.py\\site_harvest.py [air|water|all] [시료번호(선택)]
"""

import datetime as _dt
import json
import os
import re
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from selenium.webdriver.common.by import By

import measin_utils as m
import selenium_utils as s
from config import FIELD_URL_AIR, FIELD_URL_WATER, LOGIN_URL

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "site_snapshot")

SAMPLE_RE = re.compile(r"^[A-Z]\d{7}-\d{2}$")

_DUMP_JS = r"""
const root = arguments[0] || document;
function vis(el){ const r = el.getBoundingClientRect(); const st = getComputedStyle(el);
  return !!(r.width || r.height) && st.visibility !== 'hidden' && st.display !== 'none'; }
function labelOf(el){
  if (el.id){ const l = document.querySelector('label[for="'+CSS.escape(el.id)+'"]'); if (l && l.innerText.trim()) return l.innerText.trim(); }
  const sec = el.closest('section, td, th, .form-group, label');
  if (sec){ const lab = sec.querySelector('label.label, label, .label, th');
    if (lab && lab !== el && lab.innerText.trim()) return lab.innerText.trim().slice(0,40); }
  const tr = el.closest('tr'); if (tr){ const th = tr.querySelector('th'); if (th) return th.innerText.trim().slice(0,40); }
  return '';
}
const out = [];
root.querySelectorAll('input, select, textarea, button, a.btn, a[id^="btn"], [id^="btn"]').forEach(el => {
  const tag = el.tagName.toLowerCase();
  const o = { tag, id: el.id || '', name: el.getAttribute('name') || '', type: el.getAttribute('type') || '',
    cls: (el.className && el.className.baseVal === undefined ? el.className : '').toString().slice(0,80),
    visible: vis(el), disabled: !!el.disabled, label: labelOf(el) };
  if (tag === 'select'){
    o.multiple = el.multiple;
    o.select2 = !!(el.nextElementSibling && el.nextElementSibling.classList && el.nextElementSibling.classList.contains('select2'));
    o.options = Array.from(el.options).slice(0,25).map(op => [op.value, op.text.trim().slice(0,30)]);
    o.n_options = el.options.length;
  } else if (tag === 'button' || tag === 'a' || o.type === 'button' || o.type === 'submit'){
    o.text = (el.innerText || el.value || '').trim().slice(0,40);
  } else if (o.type !== 'password'){
    o.value = (el.value || '').toString().slice(0,40);
    o.placeholder = el.getAttribute('placeholder') || '';
  }
  out.push(o);
});
const grids = [];
root.querySelectorAll('.rg-root').forEach(g => {
  let host = g.id ? g : g.closest('[id]');
  const heads = Array.from(g.querySelectorAll('.rg-header table thead th, .rg-header td')).map(t => t.innerText.trim()).filter(Boolean);
  grids.push({ id: host ? host.id : '', visible: vis(g), headers: Array.from(new Set(heads)).slice(0,60) });
});
return {elements: out, grids};
"""

_GRIDVIEWS_JS = r"""
const res = {};
const mgv = window.measGridViews || {};
Object.keys(mgv).forEach(k => {
  const gv = mgv[k];
  try {
    const cols = (gv.getColumns ? gv.getColumns() : []).map(c => ({
      field: c.fieldName || c.name || '', header: (c.header && (c.header.text || c.header.label)) || '' }));
    let rows = null;
    try { const dp = gv.getDataSource ? gv.getDataSource() : gv._dataProvider; rows = dp ? dp.getRowCount() : null; } catch(e){}
    res[k] = {columns: cols, rows};
  } catch(e){ res[k] = {error: String(e)}; }
});
return {measGridViews: res, hasGetGridViewAir: typeof window.getGridViewAir === 'function'};
"""

_TABS_JS = r"""
return Array.from(document.querySelectorAll('.ui-tabs-nav li a, ul[role="tablist"] a')).map(a => ({
  id: a.id || '', href: a.getAttribute('href') || '', text: a.innerText.trim(),
  disabled: !!(a.closest('li') && a.closest('li').classList.contains('ui-state-disabled')) }));
"""


def _alerts(d):
    texts = []
    for _ in range(5):
        try:
            a = d.switch_to.alert
            texts.append(a.text)
            a.accept()
            time.sleep(0.3)
        except Exception:
            break
    return texts


def wait_login(d, timeout=900):
    if m.is_field_list_ready(d) or (not m.is_logged_out(d) and "/init.go" not in d.current_url):
        return True
    d.get(LOGIN_URL)
    print("⏳ 전용 크롬 창에서 로그인해 주세요...", flush=True)
    t0 = time.time()
    while time.time() - t0 < timeout:
        time.sleep(3)
        try:
            if not m.is_logged_out(d) and "/init.go" not in d.current_url:
                return True
        except Exception:
            pass
    return False


def dump(d, root_el=None):
    return d.execute_script(_DUMP_JS, root_el)


def pick_sample(d):
    m.wait_grid_loaded(d, timeout=10)
    for c in d.find_elements(By.CSS_SELECTOR, ".rg-renderer"):
        t = (c.text or "").strip()
        if SAMPLE_RE.match(t):
            return t
    return None


def harvest_media(d, media, sample=None):
    url = FIELD_URL_AIR if media == "air" else FIELD_URL_WATER
    d.get(url)
    time.sleep(2.5)
    res = {"media": media, "url": url, "alerts": _alerts(d)}

    if media == "air" and not sample:
        today = _dt.date.today()
        s.set_date_js(d, "#search_meas_start_dt", (today - _dt.timedelta(days=30)).isoformat())
        s.set_date_js(d, "#search_meas_end_dt", today.isoformat())
        s.safe_click(d, "#btnSearch")
        time.sleep(2.5)

    res["list_page"] = dump(d)
    res["list_page"]["title"] = d.title
    try:
        res["list_page"]["gridviews"] = d.execute_script(_GRIDVIEWS_JS)
    except Exception as e:
        res["list_page"]["gridviews"] = str(e)

    sample = sample or pick_sample(d)
    res["sample"] = sample
    if not sample:
        print(f"⚠ [{media}] 목록에서 시료번호를 찾지 못함", flush=True)
        return res

    print(f"▶ [{media}] 시료 {sample} 상세 진입", flush=True)
    wait_sel = "#machineDiv" if media == "air" else "a#ui-id-2"
    ok = m.open_sample_detail(d, sample, detail_wait_sel=wait_sel)
    res["detail_opened"] = bool(ok)
    if not ok:
        return res
    time.sleep(2)
    res["detail_alerts"] = _alerts(d)
    res["detail_url"] = d.current_url
    res["tabs"] = d.execute_script(_TABS_JS)
    res["detail_all"] = dump(d)

    panels = {}
    for tab in res["tabs"]:
        if not tab["id"] or tab["disabled"]:
            panels[tab["id"] or tab["text"]] = {"skipped": "disabled/no id", **tab}
            continue
        try:
            a = d.find_element(By.ID, tab["id"])
            d.execute_script("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", a)
            time.sleep(2.5)
            alerts = _alerts(d)
            panel_el = None
            href = tab["href"]
            if href.startswith("#"):
                try:
                    panel_el = d.find_element(By.CSS_SELECTOR, href)
                except Exception:
                    panel_el = None
            info = dump(d, panel_el)
            info["tab"] = tab
            info["alerts"] = alerts
            try:
                info["gridviews"] = d.execute_script(_GRIDVIEWS_JS)
            except Exception as e:
                info["gridviews"] = str(e)
            panels[tab["id"]] = info
            print(f"   탭 {tab['id']} '{tab['text']}' 요소 {len(info['elements'])}개, 그리드 {len(info['grids'])}개", flush=True)
        except Exception as e:
            panels[tab["id"]] = {"error": str(e), **tab}
    res["panels"] = panels

    d.get(url)
    time.sleep(2)
    _alerts(d)
    return res


def code_selectors():
    """코드 문자열 안의 #id 토큰 수집 → {id: [파일:줄, ...]}"""
    ids = {}
    pat = re.compile(r"#([A-Za-z_][\w\-]*)")
    for fn in os.listdir(HERE):
        if not fn.endswith(".py") or fn.startswith("_tmp") or fn == os.path.basename(__file__):
            continue
        with open(os.path.join(HERE, fn), encoding="utf-8", errors="ignore") as f:
            for no, line in enumerate(f, 1):
                code = line.split("#", 1)[0] if line.lstrip().startswith("#") else line
                for q in re.findall(r"[\"']([^\"']*#[^\"']*)[\"']", code):
                    for mid in pat.findall(q):
                        if re.fullmatch(r"[0-9a-fA-F]{3,8}", mid) or mid in ("x27",):
                            continue
                        ids.setdefault(mid, []).append(f"{fn}:{no}")
    return ids


def site_ids(results):
    found = {}
    for r in results:
        media = r["media"]
        blocks = [("목록", r.get("list_page"))] + [("상세", r.get("detail_all"))]
        blocks += [(k, v) for k, v in (r.get("panels") or {}).items()]
        for where, b in blocks:
            if not isinstance(b, dict):
                continue
            for e in b.get("elements", []):
                if e.get("id"):
                    found.setdefault(e["id"], set()).add(f"{media}:{where}")
            for g in b.get("grids", []):
                if g.get("id"):
                    found.setdefault(g["id"], set()).add(f"{media}:{where}(grid)")
        for t in r.get("tabs") or []:
            if t.get("id"):
                found.setdefault(t["id"], set()).add(f"{media}:탭")
    return found


def page_has_ids(d, ids):
    """폼 요소 외(div·section·td·img 등) id 존재 여부 — 현재 페이지 기준."""
    return d.execute_script(
        "return arguments[0].filter(i => document.getElementById(i));", list(ids)
    )


def main():
    which = (sys.argv[1] if len(sys.argv) > 1 else "all").lower()
    sample_arg = sys.argv[2] if len(sys.argv) > 2 else None
    medias = ["air", "water"] if which == "all" else [which]

    d = s.init_driver()
    if not wait_login(d):
        print("❌ 로그인 대기 시간 초과", flush=True)
        return

    os.makedirs(OUT_DIR, exist_ok=True)
    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    results = []
    extra_found = {}
    for media in medias:
        r = harvest_media(d, media, sample_arg)
        results.append(r)
        path = os.path.join(OUT_DIR, f"site_{media}_{stamp}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(r, f, ensure_ascii=False, indent=1)
        print(f"💾 {path}", flush=True)

    code_ids = code_selectors()
    found = site_ids(results)
    missing = sorted(i for i in code_ids if i not in found)

    # 폼 요소가 아닌 id(div/td/img 등)는 탭별로 다시 열어 확인
    for r in results:
        if not r.get("detail_opened"):
            continue
        url = r["url"]
        d.get(url)
        time.sleep(2)
        extra = set(page_has_ids(d, missing))
        wait_sel = "#machineDiv" if r["media"] == "air" else "a#ui-id-2"
        if m.open_sample_detail(d, r["sample"], detail_wait_sel=wait_sel):
            time.sleep(2)
            for t in r.get("tabs") or []:
                if not t.get("id") or t.get("disabled"):
                    continue
                try:
                    d.execute_script("arguments[0].click();", d.find_element(By.ID, t["id"]))
                    time.sleep(2)
                    _alerts(d)
                    for i in page_has_ids(d, missing):
                        extra_found.setdefault(i, set()).add(f"{r['media']}:{t['id']}")
                except Exception:
                    pass
            d.get(url)
            time.sleep(1.5)
            _alerts(d)
        for i in extra:
            extra_found.setdefault(i, set()).add(f"{r['media']}:목록")

    report = {
        "stamp": stamp,
        "samples": {r["media"]: r.get("sample") for r in results},
        "code_ids_total": len(code_ids),
        "found_as_form_element": sorted(i for i in code_ids if i in found),
        "found_as_other_element": {i: sorted(v) for i, v in extra_found.items()},
        "not_found_on_site": {
            i: code_ids[i] for i in missing if i not in extra_found
        },
    }
    path = os.path.join(OUT_DIR, f"selector_check_{stamp}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=1)
    print(f"💾 {path}", flush=True)
    print(f"코드 #id {len(code_ids)}개 / 사이트에서 못 찾음 {len(report['not_found_on_site'])}개", flush=True)


if __name__ == "__main__":
    main()
