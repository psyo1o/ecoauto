# -*- coding: utf-8 -*-
"""
COM 기반 성적서 자동 복사 – (조건부서식 + 대기시료채취 및 분석일지 + 벤조(a)피렌 처리 + 실패시트)
+ sample-3(요약/검토) 시트(이상 있을 때만 생성)
+ sample-4(THC) : THC는 "고정 범위(A1:G150)"를 "1행부터(A1)" 복사 (위치 흔들림 방지)

[유지]
- sample-1: 대기측정기록부 그림 + 입력 A23:B23(B18='사용')
- sample-2: 범위 복사 + 대기시료채취 및 분석일지 숨김 그대로 복사
- -3: D/H(30% 파랑, 초과 빨강), J14 vs D18:D21, D37/H37 역전, 시료번호 불일치
- -3: 비산먼지면 O1, 그 외 P1
- 클립보드/대용량 팝업 방지
- 숨김 행은 검토에서 PASS
- 전체측정시간(E5~F5) "엄격" 범위(경계 포함 오류), 초 제거, HH:MM만 출력

[이번 반영(핵심)]
- ✅ VOC는 A/B/C 괄호 그룹이 각각 1회 채취:
  - 그룹 내 항목 시간이 서로 다르면 [채취시간불일치]
  - 그룹은 1개 이벤트로 카운트
- ✅ 중금속 8종도 1회 채취:
  - 8종 시간 불일치면 [채취시간불일치]
- ✅ 폼알데하이드/아세트알데하이드는 1회 채취:
  - 둘 다 있으면 시간 불일치면 [채취시간불일치]
- ✅ 가스상 장비 대수 gas_cnt 기준:
  - 가스상2(항목별 1회) + VOC(A/B/C 그룹 이벤트) 동시겹침이 gas_cnt 초과면 [장비초과]
  - 끝==다음 시작도 겹침으로 간주
- ✅ 입자상 장비는 무조건 1대:
  - 먼지 / 중금속 / 수은 / 벤조피렌 이벤트가 서로 겹치면 [장비초과]
  - 끝==다음 시작도 겹침으로 간주

[THC 변경(중요)]
- ✅ "총탄화수소" 체크된 경우에만:
  - sample-4 생성 후, PF/FID 중 선택된 시트에서
    "A1:G150" 을 sample-4의 "A1"부터 값+서식으로 고정 복사
  - sample-3에서 C65(1이면 FID, 아니면 PF) 기준 기대값과 실제 복사 시트가 일치하는지 검증(불일치면 이상)

[정리]
- ❌ PrintArea/UsedRange/Find 기반 THC 유틸 제거 (원인: 파일마다 '실사용범위'가 달라 붙여넣기 위치/크기 흔들림)
- ✅ THC는 고정 범위 복사로 안정화
"""

import os
import re
import time
from datetime import datetime, timedelta

import win32com.client as win32
import win32clipboard

# ============================================================
# 공통 유틸 모듈 (모듈화)
# ============================================================
from format_utils import (
    to_float_if_pure_number, extract_numbers_from_cell,
    to_datetime_if_possible, is_excel_base_date, align_to_date,
    strip_seconds, fmt_hhmm, fmt_range
)
from file_utils import find_best_matching_file as _find_best_file_util
from excel_utils import find_sheet_by_candidates
from excel_com_utils import get_excel_app

SRC_DIR = r"\\192.168.10.163\측정팀\2.성적서\0 2.검토중"


# ============================================================
# 공용: 로그 출력 헬퍼
# ============================================================
def log(msg):
    print(msg)


# ============================================================
# 클립보드 비우기(팝업 방지)
# ============================================================
def clear_clipboard(excel_app=None):
    try:
        if excel_app is not None:
            excel_app.CutCopyMode = False
    except:
        pass
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
    except:
        pass


# ============================================================
# 엑셀 파일 찾기
# ============================================================
def find_excel(sample, folder):
    """단일 폴더에서 sample 과 정확 형식 일치 엑셀 파일 검색 (strict)"""
    result = _find_best_file_util(
        sample,
        nas_base=folder,
        nas_dirs=[""],
        extensions=(".xlsm", ".xlsx", ".xls"),
        strict=True,
    )
    if result:
        log(f" → 발견: {result}")
    else:
        log(f"⚠ 파일 없음: {sample}")
    return result


# ============================================================
# 값 + 서식 복사 (링크 제거) + 클립보드 정리
# ============================================================
def safe_copy_no_link(src_range, dst_range):
    excel_app = None
    try:
        excel_app = src_range.Application
    except:
        pass

    try:
        src_range.Copy()
        dst_range.PasteSpecial(Paste=-4163)   # 값만
        dst_range.PasteSpecial(Paste=-4122)   # 서식
    except Exception as e:
        log(f"⚠ PasteSpecial 실패: {e}")
    finally:
        clear_clipboard(excel_app)


# ============================================================
# 값+서식 복사 + 숨김(행/열 Hidden) 상태 복제
# ============================================================
def copy_range_with_hidden(src_ws, src_addr, dst_ws, dst_addr):
    src_rng = src_ws.Range(src_addr)
    dst_rng = dst_ws.Range(dst_addr)

    safe_copy_no_link(src_rng, dst_rng)

    # 행 숨김 복제
    try:
        sr1 = src_rng.Row
        sr2 = src_rng.Row + src_rng.Rows.Count - 1
        dr1 = dst_rng.Row
        for i, r in enumerate(range(sr1, sr2 + 1)):
            dst_ws.Rows(dr1 + i).Hidden = src_ws.Rows(r).Hidden
    except:
        pass

    # 열 숨김 복제
    try:
        sc1 = src_rng.Column
        sc2 = src_rng.Column + src_rng.Columns.Count - 1
        dc1 = dst_rng.Column
        for i, c in enumerate(range(sc1, sc2 + 1)):
            dst_ws.Columns(dc1 + i).Hidden = src_ws.Columns(c).Hidden
    except:
        pass


# ============================================================
# THC: 고정 범위 복사 유틸 (핵심)
# ============================================================
def is_sheet_visible(ws):
    """Excel COM: Visible == -1 이면 보임"""
    try:
        return int(ws.Visible) == -1
    except:
        return False


def copy_fixed_block(src_ws, src_addr, dst_ws, dst_addr):
    """
    ✅ THC는 여기로만 복사한다.
    - src_addr: 예) "A1:G150"
    - dst_addr: 예) "A1"
    """
    src_rng = src_ws.Range(src_addr)
    dst_rng = dst_ws.Range(dst_addr)
    safe_copy_no_link(src_rng, dst_rng)
    return int(src_rng.Rows.Count), int(src_rng.Columns.Count)


def c65_is_fid(v) -> bool:
    """
    C65가 1 / 1.0 / "1" / "1.0" 등인 경우 True(FID6000)로 판정
    그 외는 False(PF)로 판정
    """
    if v is None:
        return False
    try:
        return float(v) == 1.0
    except:
        s = str(v).strip()
        return s in ("1", "1.0", "1.00", "1,0")



# ============================================================
# 대기측정기록부를 그림으로 붙여넣기 (클립보드 오류 재시도)
# - 추가 요청: 인쇄영역만 적용
# ============================================================
def get_print_area_range(ws):
    """
    ws.PageSetup.PrintArea가 있으면 그 영역(여러개면 bounding rectangle)을 Range로 반환.
    없거나 실패하면 None 반환.
    """
    try:
        pa = ws.PageSetup.PrintArea
    except:
        pa = None

    if not pa or str(pa).strip() == "":
        return None

    try:
        parts = [p.strip() for p in str(pa).split(",") if p.strip()]
        if not parts:
            return None

        min_r, min_c = 10**9, 10**9
        max_r, max_c = 0, 0

        for addr in parts:
            rng = ws.Range(addr)
            r1 = int(rng.Row)
            c1 = int(rng.Column)
            r2 = r1 + int(rng.Rows.Count) - 1
            c2 = c1 + int(rng.Columns.Count) - 1
            min_r = min(min_r, r1)
            min_c = min(min_c, c1)
            max_r = max(max_r, r2)
            max_c = max(max_c, c2)

        if max_r <= 0 or max_c <= 0 or min_r > max_r or min_c > max_c:
            return None

        return ws.Range(ws.Cells(min_r, min_c), ws.Cells(max_r, max_c))
    except:
        return None


def copy_sheet_as_picture(main_ws, dst_ws, anchor_addr="A3", retries=5):
    """
    ✅ 대기측정기록부: 인쇄영역(PrintArea) 우선 CopyPicture
    - PrintArea 없으면 UsedRange fallback
    - 클립보드 오류 완화를 위해 재시도
    """
    excel = None
    try:
        excel = main_ws.Application
    except:
        pass

    last_err = None

    for i in range(retries):
        try:
            clear_clipboard(excel)
            try:
                if excel is not None:
                    excel.CutCopyMode = False
            except:
                pass

            # 🔥 [핵심 해결 포인트] 그림을 찍기 전에 원본 시트를 강제로 활성화하고 0.5초 숨 고르기!
            try:
                main_ws.Activate()
            except:
                pass
            import time
            time.sleep(0.5)

            # ✅ 인쇄영역 우선
            pa_rng = get_print_area_range(main_ws)
            if pa_rng is not None:
                src_rng = pa_rng
                print(f"  [PIC] 인쇄영역으로 복사: {main_ws.PageSetup.PrintArea}")
            else:
                src_rng = main_ws.UsedRange
                print("  [PIC] 인쇄영역 없음 → UsedRange로 복사")

            # CopyPicture (화면 기준, 사진 포맷)
            src_rng.CopyPicture(Appearance=1, Format=2)

            time.sleep(0.25 + 0.15 * i)

            # 붙여넣기
            try:
                dst_ws.Activate()
            except:
                pass
            try:
                if excel is not None:
                    excel.Goto(dst_ws.Range(anchor_addr))
            except:
                try:
                    dst_ws.Range(anchor_addr).Select()
                except:
                    pass

            dst_ws.Paste()
            clear_clipboard(excel)

            print("  ✔ 대기측정기록부 완료")
            return True

        except Exception as e:
            last_err = e
            clear_clipboard(excel)
            time.sleep(0.25 + 0.2 * i)

    print("⚠ 대기측정기록부 실패(재시도 후):", last_err)
    return False



# ============================================================
# 실패 시트 추가
# ============================================================
def add_fail_sheet(wb_out, fail_list):
    if not fail_list:
        return
    try:
        dummy = wb_out.Worksheets("Dummy_End")
        ws_fail = wb_out.Worksheets.Add(dummy)
    except:
        ws_fail = wb_out.Worksheets.Add()

    ws_fail.Name = "복사 실패"
    ws_fail.Range("A1").Value = "복사 실패한 시료번호"
    row = 2
    for sample in fail_list:
        ws_fail.Range(f"A{row}").Value = sample
        row += 1

# ============================================================
# 분류(부분포함 방식으로 안정화)
# ============================================================
def classify_item(item):
    s = (item or "").strip()
    if not s:
        return None
    if "비산먼지" in s:
        return "비산먼지"

    if ("벤조" in s) and ("피렌" in s):
        return "입자상"

    particle_kw = ["먼지", "수은", "비소", "구리", "아연", "카드뮴", "납", "크롬", "크로뮴", "니켈", "베릴륨"]
    if any(k in s for k in particle_kw):
        return "입자상"

    if any(k in s for k in ["황산화물", "질소산화물", "일산화탄소"]):
        return "가스상1"
    if "총탄화수소" in s:
        return "가스상1"

    gas2_kw = ["염화수소", "염소", "플루오린", "황화수소", "사이안화수소", "페놀", "브로민",
               "하이드라진", "벤지딘", "폼알데하이드", "아세트알데하이드", "이황화탄소"]
    if any(k in s for k in gas2_kw):
        return "가스상2"

    voc_kw = ["다이클로로메테인", "벤젠", "트라이클로로에틸렌", "1,3-뷰타다이엔", "아크릴로나이트릴",
              "1,2-다이클로로에테인", "클로로폼", "테트라클로로에틸렌", "스타이렌", "에틸벤젠",
              "사염화탄소", "아닐린", "염화바이닐", "이황화메틸", "에틸렌옥사이드", "프로필렌옥사이드"]
    if any(k in s for k in voc_kw):
        return "VOC"

    return None


# ===========================
# ✅ 채취 묶음 정의 (VOC A/B/C, 중금속 8종, 폼/아세, 입자상(먼지/중금속/수은/벤조피렌))
# ===========================
VOC_GROUP_A = {
    "다이클로로메테인","벤젠","트라이클로로에틸렌","1,3-뷰타다이엔","아크릴로나이트릴",
    "1,2-다이클로로에테인","클로로폼","테트라클로로에틸렌","스타이렌","에틸벤젠",
    "사염화탄소","아닐린","염화바이닐"
}
VOC_GROUP_B = {"이황화메틸"}
VOC_GROUP_C = {"에틸렌옥사이드","프로필렌옥사이드"}

HEAVY_METALS = {
    "비소화합물","크로뮴화합물","납화합물","니켈화합물","베릴륨화합물",
    "구리화합물","아연화합물","카드뮴화합물"
}

def sampling_group_id(item_name: str, cat: str):
    s = (item_name or "").strip()
    if not s:
        return None

    if cat == "VOC":
        if s in VOC_GROUP_A: return "VOC-A"
        if s in VOC_GROUP_B: return "VOC-B"
        if s in VOC_GROUP_C: return "VOC-C"
        return None

    if cat == "가스상2":
        if s in ("폼알데하이드", "아세트알데하이드"):
            return "GAS2-FA"
        return None

    if cat == "입자상":
        if s in HEAVY_METALS:
            return "PM-METALS"
        if "수은" in s:
            return "PM-HG"
        if "먼지" in s:
            return "PM-DUST"
        if ("벤조" in s) and ("피렌" in s):
            return "PM-BaP"
        return f"PM-OTHER:{s}"

    return None


def check_group_time_alignment(rows, group_id, group_label):
    items = [d for d in rows if d.get("group_id") == group_id]
    if len(items) <= 1:
        return []

    base = items[0]
    bs, be = base.get("start_dt"), base.get("end_dt")
    if not bs or not be:
        return [f"[채취시간불일치] {group_label}: 기준시간 비어있음"]

    issues = []
    for d in items[1:]:
        s, e = d.get("start_dt"), d.get("end_dt")
        if not s or not e:
            issues.append(f"[채취시간불일치] {group_label}: {d.get('item')} 시간 비어있음")
            continue
        if (strip_seconds(s), strip_seconds(e)) != (strip_seconds(bs), strip_seconds(be)):
            issues.append(
                f"[채취시간불일치] {group_label}: "
                f"{base.get('item')}({fmt_range(bs,be)}) vs {d.get('item')}({fmt_range(s,e)})"
            )
    return issues


def build_sampling_events(rows):
    buckets = {}
    for d in rows:
        s, e = d.get("start_dt"), d.get("end_dt")
        if not s or not e:
            continue
        if e < s:
            continue

        gid = d.get("group_id")
        item = str(d.get("item") or "").strip() or "(항목명없음)"

        if gid:
            key = ("GROUP", gid)
            label = gid
        else:
            key = ("ITEM", item)
            label = item

        if key not in buckets:
            buckets[key] = {"s": strip_seconds(s), "e": strip_seconds(e), "label": label, "items": [item]}
        else:
            buckets[key]["items"].append(item)

    events = []
    for b in buckets.values():
        events.append((b["s"], b["e"], b["label"], b["items"]))
    return events


def check_overlap_events(events, allow, title):
    allow = max(1, int(allow or 1))
    evs = []
    for (s, e, label, items) in events:
        if not s or not e or e < s:
            continue
        s2 = strip_seconds(s)
        e2 = strip_seconds(e)
        evs.append((s2, +1, label, items))
        evs.append((e2, -1, label, items))

    if not evs:
        return []

    evs.sort(key=lambda x: (x[0], -x[1]))

    active = []
    issues = []
    in_vio = False
    vio_start = None
    vio_snapshot = []

    def snapshot_desc(snap):
        out = []
        for lab, its in snap:
            uniq = []
            for x in its:
                x = str(x).strip()
                if x and x not in uniq:
                    uniq.append(x)

            if len(uniq) == 1 and str(lab).strip() == uniq[0]:
                out.append(f"- {uniq[0]}")
            else:
                out.append(f"- {lab}: {', '.join(uniq)}")
        return "\n".join(out)

    for t, delta, label, items in evs:
        if delta == +1:
            active.append((label, items))
        else:
            for i in range(len(active)-1, -1, -1):
                if active[i][0] == label:
                    active.pop(i)
                    break

        cnt = len(active)

        if (not in_vio) and cnt > allow:
            in_vio = True
            vio_start = t
            vio_snapshot = list(active)

        if in_vio and cnt <= allow:
            vio_end = t
            if vio_start and vio_end and vio_end > vio_start:
                issues.append(
                    f"[장비초과] {title} 동시채취 {fmt_hhmm(vio_start)}~{fmt_hhmm(vio_end)} "
                    f"(장비 {allow}대, 동시채취 {len(vio_snapshot)}건)\n{snapshot_desc(vio_snapshot)}"
                )
            in_vio = False
            vio_start = None
            vio_snapshot = []

    return issues


# ============================================================
# 입력(분석값) 읽기 (숨김 PASS)
# ============================================================
def read_analysis_items(ws_analysis):
    if ws_analysis is None:
        return []
    rows = []
    for r in range(2, 65):
        try:
            if ws_analysis.Rows(r).Hidden or ws_analysis.Rows(r).RowHeight == 0:
                continue
        except:
            pass

        flag = ws_analysis.Cells(r, 1).Value  # A
        item = ws_analysis.Cells(r, 2).Value  # B
        if item is None:
            continue
        if flag is None or str(flag).strip() == "":
            continue

        item_s = str(item).strip()
        conc = ws_analysis.Cells(r, 5).Value    # E
        limitv = ws_analysis.Cells(r, 8).Value  # H
        st_raw = ws_analysis.Cells(r, 9).Value  # I
        ed_raw = ws_analysis.Cells(r, 11).Value # K

        rows.append({
            "src": "입력(분석값)",
            "row": r,
            "item": item_s,
            "cat": classify_item(item_s),
            "conc": to_float_if_pure_number(conc),
            "limit": to_float_if_pure_number(limitv),
            "start_dt": to_datetime_if_possible(st_raw),
            "end_dt": to_datetime_if_possible(ed_raw),
            "group_id": None,
        })
    return rows

# ============================================================
# 농도 vs 기준 체크(중복 제거)
# ============================================================
def build_conc_limit_checks(rows):
    out = []
    seen = set()
    for d in rows:
        conc = d.get("conc")
        limitv = d.get("limit")
        item = d.get("item", "")
        if conc is None or limitv is None or limitv == 0:
            continue

        if conc > limitv:
            label, color = "초과 빨간색", "RED"
        elif conc >= limitv * 0.3:
            label, color = "30% 이상 파란색", "BLUE"
        else:
            continue

        key = (label, str(item).strip(), float(limitv), float(conc), color)
        if key in seen:
            continue
        seen.add(key)
        out.append((label, item, limitv, conc, color))
    return out


# ============================================================
# 전체측정시간(E5~F5) 범위 체크 (엄격히 안쪽, 초 제거)
# ============================================================
def check_total_window(all_rows, overall_start_dt, overall_end_dt):
    issues = []
    if overall_start_dt is None or overall_end_dt is None:
        issues.append("[오류] 입력 시트 E5/F5(전체측정시작/끝) 시간이 비어있거나 해석 불가")
        return issues

    os_ = strip_seconds(overall_start_dt)
    oe_ = strip_seconds(overall_end_dt)

    for d in all_rows:
        s = d.get("start_dt")
        e = d.get("end_dt")
        if s is None or e is None:
            continue

        s2 = strip_seconds(align_to_date(os_, s))
        e2 = strip_seconds(align_to_date(os_, e))

        if e2 < s2:
            issues.append(f"[시간역전] {d.get('src')} {d.get('item')} {fmt_range(s2, e2)}")
            continue

        # 엄격: 경계 포함 오류
        if s2 <= os_ or e2 >= oe_:
            issues.append(
                f"[전체시간범위이탈] {d.get('src')} {d.get('item')} {fmt_range(s2, e2)} "
                f"(전체:{fmt_hhmm(os_)}~{fmt_hhmm(oe_)})"
            )

        d["start_dt"] = s2
        d["end_dt"] = e2

    return issues


# ============================================================
# 장비표(입력 시트 AB9:AQ13, 팀=C1) 읽기
# ============================================================
def get_devices_from_input_sheet(wb_src):
    phrases = {
        "gas_sampler": "굴뚝시료채취장치(가스상)",
        "particle_sampler": "굴뚝시료채취장치(입자상)",
        "gas_meter": "대기배출가스측정기",
        "thc_meter": "대기배출가스(THC)측정기",
        "fugitive_dust": "비산먼지",
        "moisture_auto": "수분량자동측정기",
    }

    def norm_device_id(s):
        s = str(s).strip()
        s = re.sub(r"\s+", " ", s)
        parts = [p.strip() for p in s.split("/") if p.strip()]
        if parts:
            return parts[-1].upper()
        return s.upper()

    gas_ids = set()
    particle_ids = set()
    has_gas_meter = False
    has_thc_meter = False
    has_fugitive_dust = False

    # ✅ 추가: 수분량자동측정기 유무 + B18 값
    has_moisture_auto = False
    b18_value = ""

    try:
        ws = None
        for sh in wb_src.Worksheets:
            if sh.Name == "입력":
                ws = sh
                break
        if ws is None:
            # 반환값 8개로 통일
            return 0, 0, False, False, False, False, False, ""

        # ✅ B18 값 읽기 (사용/공란/기타)
        try:
            b18_value = str(ws.Range("B18").Value or "").strip()
        except:
            b18_value = ""

        team_raw = ws.Range("C1").Value
        m = re.search(r"\d+", str(team_raw) if team_raw is not None else "")
        if not m:
            return 0, 0, False, False, False, False, False, b18_value

        team = int(m.group())
        if team < 1 or team > 5:
            return 0, 0, False, False, False, False, False, b18_value

        row = 8 + team  # 1팀=9행
        rng = ws.Range(f"AB{row}:AQ{row}")

        for cell in rng:
            v = cell.Value
            if v is None:
                continue
            s = str(v).strip()
            if not s:
                continue

            if phrases["gas_sampler"] in s:
                gas_ids.add(norm_device_id(s))
            if phrases["particle_sampler"] in s:
                particle_ids.add(norm_device_id(s))

            if phrases["gas_meter"] in s and phrases["thc_meter"] not in s:
                has_gas_meter = True
            if phrases["thc_meter"] in s:
                has_thc_meter = True
            if phrases["fugitive_dust"] in s:
                has_fugitive_dust = True

            # ✅ 수분량자동측정기 포함 여부
            if phrases["moisture_auto"] in s:
                has_moisture_auto = True

        gas_cnt = 2 if len(gas_ids) >= 2 else (1 if len(gas_ids) == 1 else 0)
        particle_cnt = 2 if len(particle_ids) >= 2 else (1 if len(particle_ids) == 1 else 0)
        has_particle_sampler = (len(particle_ids) > 0)

        # ✅ 반환 8개
        return (
            gas_cnt, particle_cnt,
            has_gas_meter, has_thc_meter,
            has_fugitive_dust, has_particle_sampler,
            has_moisture_auto, b18_value
        )

    except:
        return 0, 0, False, False, False, False, False, b18_value



# ============================================================
# 장비 누락인데 측정 체크된 경우
# ============================================================
def check_device_missing_but_measured(analysis_rows, fugitive_rows,
                                      gas_cnt, particle_cnt,
                                      has_gas_meter, has_thc_meter,
                                      has_fugitive_dust, has_particle_sampler,
                                      has_moisture_auto, b18_value,
                                      is_fugitive):
    measured = list(analysis_rows) + list(fugitive_rows)

    need_gas_sampler = any(d.get("cat") in ("가스상2", "VOC") for d in measured)
    need_particle = any(d.get("cat") == "입자상" for d in measured)
    need_fugitive = len(fugitive_rows) > 0

    gas1_3 = {"황산화물", "질소산화물", "일산화탄소"}
    need_gas_meter = any((d.get("item") in gas1_3) for d in measured)

    # 총탄화수소 체크 여부
    need_thc_meter = any((d.get("item") == "총탄화수소") for d in measured)

    issues = []
    if need_gas_sampler and (gas_cnt or 0) == 0:
        issues.append("[장비누락] 가스상2/VOC 측정인데 입력 시트에 굴뚝시료채취장치(가스상) 장비가 없음")
    if need_particle and not has_particle_sampler:
        issues.append("[장비누락] 입자상 측정인데 입력 시트에 굴뚝시료채취장치(입자상) 장비 표기가 없음")
    if need_fugitive and not has_fugitive_dust:
        issues.append("[장비누락] 비산먼지 측정인데 입력 시트 장비표에 비산먼지 전용 장비 표기가 없음(장비표 셀에 '비산먼지' 포함 필요)")
    if need_gas_meter and not has_gas_meter:
        issues.append("[장비누락] 가스상1(황/질/일산화) 측정인데 입력 시트에 대기배출가스측정기 장비가 없음")
    if need_thc_meter and not has_thc_meter:
        issues.append("[장비누락] 총탄화수소 측정인데 입력 시트에 대기배출가스(THC)측정기 장비가 없음")

    # ✅ 수분량자동측정기 ↔ 입력!B18(사용) 정합성 체크
    #    비산먼지(샘플2가 비산먼지 양식)면 PASS
    if not is_fugitive:
        b18_is_use = (str(b18_value or "").strip() == "사용")

        if has_moisture_auto and (not b18_is_use):
            issues.append("[장비불일치] 수분량자동측정기 장비 표기 있음 수분량자동측정기 '사용'이 아님")
        if (not has_moisture_auto) and b18_is_use:
            issues.append("[장비불일치] 수분량자동측정기 '사용'인데 입력 장비표에 수분량자동측정기 표기가 없음")
    # ✅ [추가] 입력!B18이 "사용"이 아니면(=미사용 포함) 가스상 샘플러가 필요
        if (not b18_is_use) and (gas_cnt or 0) == 0:
            issues.append("[장비누락] 입력!B18이 '미사용'인데 굴뚝시료채취장치(가스상) 장비가 없음")


    return issues


def set_row_height_cap(ws, row_idx, cap=80):
    try:
        ws.Rows(row_idx).WrapText = True
    except:
        pass
    try:
        ws.Rows(row_idx).RowHeight = cap
    except:
        pass


# ============================================================
# sample-2 마지막행 계산 유틸(유지)
# ============================================================
def row_count_of_addr(ws, addr: str) -> int:
    try:
        return int(ws.Range(addr).Rows.Count)
    except:
        return 0

def update_bottom(bottom_row: int, dst_start_row: int, rows_copied: int) -> int:
    if rows_copied <= 0:
        return bottom_row
    return max(bottom_row, dst_start_row + rows_copied - 1)


# ============================================================
# 파일 저장: 이미 있으면 -1, -2...
# ============================================================
def uniquify_path(path):
    base, ext = os.path.splitext(path)
    if not os.path.exists(path):
        return path

    i = 1
    while True:
        cand = f"{base}-{i}{ext}"
        if not os.path.exists(cand):
            return cand
        i += 1


# ============================================================
# 메인
# ============================================================
def main(sample_list, user_name):
    log("=== 성적서 자동 복사 시작 ===")

    fail_list = []
    if not sample_list:
        log("❌ 시료번호 입력 없음")
        return

    today = datetime.now().strftime("%Y%m%d")
    output_dir = rf"\\192.168.10.163\개인업무\{user_name}"
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, f"{today} 성적서 추출결과.xlsx")

    excel = get_excel_app()

    try:
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False        # 매크로/이벤트 실행 차단
        excel.Interactive = False         # 사용자 상호작용 차단 (성능 향상)
        excel.Cursor = 2
    except:
        pass

    try:
        wb_out = excel.Workbooks.Add()
        while wb_out.Worksheets.Count > 1:
            wb_out.Worksheets(1).Delete()

        dummy_ws = wb_out.Worksheets(1)
        dummy_ws.Name = "Dummy_End"

        for item in sample_list:
            # GUI에서 넘어온 경우(튜플)와 단독 실행(문자열) 모두 호환되도록 처리
            if isinstance(item, tuple):
                sample, src_path = item
            else:
                sample = item
                # 튜플이 아니면 어쩔 수 없이 NAS를 검색 (기존 로직)
                src_path = find_excel(sample, SRC_DIR)

            log("\n-------------------------------")
            log(f" 처리: {sample}")
            
            if src_path:
                log(f" → 발견: {src_path}")
            else:
                log(f"⚠ 파일 없음: {sample}")
                fail_list.append(sample)
                continue

            try:
                wb_src = excel.Workbooks.Open(
                    src_path, 
                    UpdateLinks=0, 
                    ReadOnly=True, 
                    IgnoreReadOnlyRecommended=True
                )
            except:
                log(f"❌ 파일 열기 실패: {src_path}")
                fail_list.append(sample)
                continue

            try:
                # -----------------------------
                # 시트 핸들 수집
                # -----------------------------
                main_ws = None
                extract_ws = None
                input_ws = None        # 입력(동정압)
                input2_ws = None       # 입력
                analysis_ws = None     # 입력(분석값)

                thc_pf_ws = None
                thc_fid_ws = None

                for sh in wb_src.Worksheets:
                    if sh.Name == "대기측정기록부":
                        main_ws = sh
                    elif sh.Name == "대기시료채취 및 분석일지":
                        extract_ws = sh
                    elif sh.Name == "입력(동정압)":
                        input_ws = sh
                    elif sh.Name == "입력":
                        input2_ws = sh
                    elif sh.Name == "입력(분석값)":
                        analysis_ws = sh
                    elif sh.Name == "THC 측정값(PF)":
                        thc_pf_ws = sh
                    elif sh.Name == "THC 측정값(FID)":
                        thc_fid_ws = sh

                if main_ws is None:
                    main_ws = wb_src.Worksheets(1)

                # 계산/캐시(유지)
                #try:
                #    excel.Calculate()
                #except:
                #    pass

                # ============================================================
                # (A) 체크된 항목에서 총탄화수소 유무 + C65(FID/PF) 판정 준비
                # ============================================================
                analysis_rows_cached = read_analysis_items(analysis_ws)
                has_thc_item = any((d.get("item") == "총탄화수소") for d in analysis_rows_cached)

                c65_val = None
                is_fid_expect = False
                try:
                    c65_val = analysis_ws.Range("C65").Value if analysis_ws else None
                    is_fid_expect = c65_is_fid(c65_val)
                except:
                    c65_val = None
                    is_fid_expect = False

                # sample-4에서 실제로 복사된 THC 시트명
                thc_copied = None

                # ============================================================
                # (1) sample-1
                # ============================================================
                try:
                    # 1. 보호 해제 (그림 복사를 위해 필요)
                    try: main_ws.Unprotect("ydenv")
                    except: pass

                    ws1 = wb_out.Worksheets.Add(dummy_ws)
                    ws1.Name = f"{sample}-1"

                    log("  [1] sample-1: 대기측정기록부 시작")
                    
                    # 2. 그림 복사 (화면 갱신 잠깐 허용)
                    try: excel.ScreenUpdating = True
                    except: pass
                    
                    copy_sheet_as_picture(main_ws, ws1, "A3")
                    
                    try: excel.ScreenUpdating = False
                    except: pass

                    # 3. 입력 시트 데이터 복사
                    if input2_ws and input2_ws.Name == "입력":
                        try: input2_ws.Unprotect("ydenv")
                        except: pass
                        
                        b18_value = str(input2_ws.Range("B18").Value or "").strip()
                        if b18_value == "사용":
                            safe_copy_no_link(input2_ws.Range("A23:B23"), ws1.Range("A1"))
                            ws1.Columns("A:B").AutoFit()
                    
                    log(f" ▶ sample-1 완료")

                except Exception as e:
                    log(f"❌ sample-1 복사 실패: {e}")
                    fail_list.append(sample)
                    continue

                # ============================================================
                # (2) sample-2
                # ============================================================
                try:
                    ws2 = wb_out.Worksheets.Add(dummy_ws)
                    ws2.Name = f"{sample}-2"

                    bottom_row = 0

                    b31_raw = main_ws.Range("B31").Value
                    b31 = str(b31_raw).strip() if b31_raw is not None else ""

                    log(f"  [2] sample-2: 유형(B31)={b31}")

                    if b31 == "비산먼지":
                        if extract_ws:
                            copy_range_with_hidden(extract_ws, "A2:AC47", ws2, "A1")
                            bottom_row = update_bottom(bottom_row, 1, row_count_of_addr(extract_ws, "A2:AC47"))
                        ws2.Columns("A:AF").AutoFit()
                        log(f" ▶ sample-2 완료(비산먼지)")
                    else:
                        if input_ws:
                            try:
                                input_ws.Unprotect("ydenv")
                            except:
                                pass

                            if b31 == "벤조(a)피렌":
                                safe_copy_no_link(input_ws.Range("R113:AB117"), ws2.Range("A1"))
                                bottom_row = update_bottom(bottom_row, 1, row_count_of_addr(input_ws, "R113:AB117"))
                            else:
                                safe_copy_no_link(input_ws.Range("A139:K144"), ws2.Range("A1"))
                                bottom_row = update_bottom(bottom_row, 1, row_count_of_addr(input_ws, "A139:K144"))
                                
                            if extract_ws:
                                copy_range_with_hidden(extract_ws, "B5:AG51", ws2, "A8")
                                bottom_row = update_bottom(bottom_row, 8, row_count_of_addr(extract_ws, "B5:AG51"))
                            ws2.Columns("A:AF").AutoFit()
                            log(f" ▶ sample-2 완료")
                            try:
                                input_ws.Protect("ydenv")
                            except:
                                pass

                except Exception as e:
                    log(f"❌ sample-2 복사 실패: {e}")
                    fail_list.append(sample)
                    continue

                # ============================================================
                # (4) sample-4 (THC 고정 복사)  ★ sample-2/1과 분리해서 위치 흔들림 원천 차단
                # ============================================================
                try:
                    if not has_thc_item:
                        log("  총탄화수소 없음 → sample-4 생성/복사 스킵")
                    else:
                        ws4 = wb_out.Worksheets.Add(dummy_ws)
                        ws4.Name = f"{sample}-4"

                        pf_vis = is_sheet_visible(thc_pf_ws) if thc_pf_ws else False
                        fid_vis = is_sheet_visible(thc_fid_ws) if thc_fid_ws else False

                        chosen = None
                        if pf_vis and not fid_vis:
                            chosen = thc_pf_ws
                        elif fid_vis and not pf_vis:
                            chosen = thc_fid_ws
                        elif pf_vis and fid_vis:
                            chosen = thc_fid_ws if is_fid_expect else thc_pf_ws
                            log(f"  PF/FID 둘 다 보임 → C65={c65_val} 기준 선택: {chosen.Name if chosen else 'None'}")
                        else:
                            chosen = None

                        if chosen is None:
                            log("  PF/FID 시트 없음 또는 둘 다 숨김 → 복사 불가(샘플-4는 빈 시트로 남김)")
                        else:
                            # ✅ 고정 범위 복사: A1:G150 -> A1
                            r_cnt, c_cnt = copy_fixed_block(chosen, "A1:G150", ws4, "A1")
                            thc_copied = str(chosen.Name)
                            ws4.Columns("A:G").AutoFit()
                            log(f" sample-4 : {thc_copied}")

                        log(f" ▶ sample-4 완료(THC)")

                except Exception as e:
                    log(f"  sample-4 처리 중 오류(무시): {e}")

                # ============================================================
                # (3) sample-3 (이상 있을 때만)
                # ============================================================
                try:
                    b31_type = str(main_ws.Range("B31").Value or "").strip()
                    is_fugitive = (b31_type == "비산먼지")

                    # (A) 대기측정기록부 D/H 체크
                    summary_rows = []
                    for r in range(31, 119):
                        try:
                            if main_ws.Rows(r).Hidden or main_ws.Rows(r).RowHeight == 0:
                                continue
                        except:
                            pass

                        item_b = main_ws.Cells(r, 2).Value
                        d_num = to_float_if_pure_number(main_ws.Cells(r, 4).Value)
                        h_num = to_float_if_pure_number(main_ws.Cells(r, 8).Value)

                        if d_num is None or h_num is None or d_num == 0:
                            continue

                        if h_num > d_num:
                            summary_rows.append(("초과 빨간색", item_b, d_num, h_num, "RED"))
                        elif h_num >= d_num * 0.3:
                            summary_rows.append(("30% 이상 파란색", item_b, d_num, h_num, "BLUE"))

                    # (B) J14 vs D18:D21
                    j14_num = to_float_if_pure_number(main_ws.Range("J14").Value)
                    nums = []
                    for rr in range(18, 22):
                        nums.extend(extract_numbers_from_cell(main_ws.Range(f"D{rr}").Value))
                    max_d = max(nums) if nums else None
                    j14_flag = (j14_num is not None and max_d is not None and j14_num > max_d)

                    # (C) D37/H37 날짜 역전
                    start_raw = main_ws.Range("D37").Value
                    end_raw = main_ws.Range("H37").Value
                    sd = to_datetime_if_possible(start_raw)
                    ed = to_datetime_if_possible(end_raw)
                    date_flag = (sd is not None and ed is not None and ed < sd)

                    # (D) 시료번호 불일치: 비산먼지면 O1, 아니면 P1
                    def _norm_sn(x):
                        if x is None:
                            return ""
                        s = str(x).strip().upper()
                        s = re.sub(r"\s+", "", s)
                        s = s.replace("–", "-").replace("—", "-").replace("−", "-")
                        m = re.match(r"^([A-Z]\d{7,8})-?(\d{1,2})$", s)
                        if m:
                            head = m.group(1)
                            idx = int(m.group(2))
                            return f"{head}-{idx:02d}"
                        return re.sub(r"[^A-Z0-9]", "", s)

                    sn_cell = "O1" if is_fugitive else "P1"
                    sn_sheet = str(main_ws.Range(sn_cell).Value or "").strip()
                    sn_flag = (_norm_sn(sn_sheet) != _norm_sn(sample))

                    # (E) 입력(분석값) + 비산먼지(비산먼지일 때만)
                    analysis_rows = list(analysis_rows_cached)
                    fugitive_rows = []

                    # 장비
                    gas_cnt, particle_cnt, has_gas_meter, has_thc_meter, has_fugitive_dust, has_particle_sampler, has_moisture_auto, b18_value = get_devices_from_input_sheet(wb_src)


                    # 전체측정시간(E5~F5) + 엄격 범위 체크(초 제거, HH:MM)
                    overall_start_dt = to_datetime_if_possible(input2_ws.Range("E5").Value) if input2_ws else None
                    overall_end_dt = to_datetime_if_possible(input2_ws.Range("F5").Value) if input2_ws else None
                    total_window_issues = check_total_window(
                        all_rows=(analysis_rows + fugitive_rows),
                        overall_start_dt=overall_start_dt,
                        overall_end_dt=overall_end_dt
                    )

                    # group_id 부여
                    for d in analysis_rows:
                        d["group_id"] = sampling_group_id(d.get("item"), d.get("cat"))
                    for d in fugitive_rows:
                        d["group_id"] = None

                    # 입력(분석값) 농도/기준
                    analysis_conc_checks = build_conc_limit_checks(analysis_rows)
                    fugitive_conc_checks = build_conc_limit_checks(fugitive_rows) if is_fugitive else []

                    # 대기측정기록부와 완전히 같은 항목이면 입력(분석값)에서 제거
                    summary_key = set()
                    for label, item_b, d_num, h_num, color in summary_rows:
                        item_s = "" if item_b is None else str(item_b).strip()
                        try:
                            summary_key.add((label, item_s, float(d_num), float(h_num), color))
                        except:
                            pass

                    filtered = []
                    for label, item, limitv, concv, color in analysis_conc_checks:
                        key = (label, str(item).strip(), float(limitv), float(concv), color)
                        if key in summary_key:
                            continue
                        filtered.append((label, item, limitv, concv, color))
                    analysis_conc_checks = filtered

                    # 장비 누락인데 측정 체크된 경우
                    device_issues = check_device_missing_but_measured(
                        analysis_rows, fugitive_rows,
                        gas_cnt, particle_cnt,
                        has_gas_meter, has_thc_meter,
                        has_fugitive_dust, has_particle_sampler,
                        has_moisture_auto, b18_value,
                        is_fugitive
                    )

                    # 시간/채취 검토
                    time_issues = []
                    time_issues += check_group_time_alignment(analysis_rows, "VOC-A", "VOC(괄호1)")
                    time_issues += check_group_time_alignment(analysis_rows, "VOC-B", "VOC(이황화메틸)")
                    time_issues += check_group_time_alignment(analysis_rows, "VOC-C", "VOC(EO/PO)")
                    time_issues += check_group_time_alignment(analysis_rows, "PM-METALS", "중금속(8종)")
                    time_issues += check_group_time_alignment(analysis_rows, "GAS2-FA", "폼알데하이드/아세트알데하이드(1회채취)")
                    time_issues += total_window_issues

                    gas_rows = [d for d in analysis_rows if d.get("cat") in ("가스상2", "VOC")]
                    gas_events = build_sampling_events(gas_rows)
                    time_issues += check_overlap_events(gas_events, allow=max(1, gas_cnt), title="가스상(굴뚝)")

                    particle_rows_only = [d for d in analysis_rows if d.get("cat") == "입자상"]
                    particle_events = build_sampling_events(particle_rows_only)
                    time_issues += check_overlap_events(particle_events, allow=1, title="입자상(1대)")

                    # (THC 검증)
                    thc_issues = []
                    expected = ""
                    if has_thc_item:
                        expected = "THC 측정값(FID)" if is_fid_expect else "THC 측정값(PF)"
                        copied = thc_copied or "(미복사)"
                        if copied != expected:
                            thc_issues.append(f"[THC시트불일치] C65={c65_val} → 기대:{expected}, 복사됨:{copied}")

                    # 이상 여부 판단
                    has_any_issue = (
                        bool(summary_rows) or j14_flag or date_flag or sn_flag or
                        bool(analysis_conc_checks) or
                        (is_fugitive and bool(fugitive_conc_checks)) or
                        bool(time_issues) or bool(device_issues) or
                        bool(thc_issues)
                    )

                    if has_any_issue:
                        ws3 = wb_out.Worksheets.Add(dummy_ws)
                        ws3.Name = f"{sample}-3"
                        
                        ws3.Range("A1").Value = "배출가스유량 비교"
                        ws3.Range("A2").Value = "항목"
                        ws3.Range("B2").Value = "값(m3/min)"
                        ws3.Range("A3").Value = "배출가스유량(보정전)"
                        ws3.Range("B3").Value = j14_num if j14_num is not None else ""
                        ws3.Range("A4").Value = "방지시설 최대"
                        ws3.Range("B4").Value = max_d if max_d is not None else ""
                        if j14_flag:
                            ws3.Range("B3").Interior.ColorIndex = 3
                            ws3.Range("B4").Interior.ColorIndex = 3

                        ws3.Range("A5").Value = "분석시작일(D37)"
                        ws3.Range("B5").Value = start_raw if start_raw is not None else ""
                        ws3.Range("A6").Value = "분석완료일(H37)"
                        ws3.Range("B6").Value = end_raw if end_raw is not None else ""
                        if date_flag:
                            ws3.Range("B6").Interior.ColorIndex = 3

                        ws3.Range("A7").Value = "시료번호(파일명/리스트)"
                        ws3.Range("B7").Value = sample
                        ws3.Range("A8").Value = f"시료번호({sn_cell})"
                        ws3.Range("B8").Value = sn_sheet
                        if sn_flag:
                            ws3.Range("B8").Interior.ColorIndex = 3

                        ws3.Range("A9").Value = "전체측정시간(입력!E5~F5)"
                        ws3.Range("B9").Value = fmt_hhmm(strip_seconds(overall_start_dt)) if overall_start_dt else ""
                        ws3.Range("C9").Value = fmt_hhmm(strip_seconds(overall_end_dt)) if overall_end_dt else ""

                        # (1) 대기측정기록부 D/H 이상 목록
                        start_row = 11
                        ws3.Range(f"A{start_row}").Value = "구분"
                        ws3.Range(f"B{start_row}").Value = "항목(B)"
                        ws3.Range(f"C{start_row}").Value = "배출허용기준(D)"
                        ws3.Range(f"D{start_row}").Value = "측정값(H)"

                        row = start_row + 1
                        for label, item_b, d_num, h_num, color in summary_rows:
                            ws3.Range(f"A{row}").Value = label
                            ws3.Range(f"B{row}").Value = "" if item_b is None else str(item_b)
                            ws3.Range(f"C{row}").Value = d_num
                            ws3.Range(f"D{row}").Value = h_num
                            if color == "RED":
                                ws3.Range(f"D{row}").Interior.ColorIndex = 3
                            elif color == "BLUE":
                                ws3.Range(f"D{row}").Interior.ColorIndex = 5
                            row += 1

                        # (2) 입력(분석값) 농도/기준
                        rr = row + 2
                        if analysis_conc_checks:
                            ws3.Range(f"A{rr}").Value = "입력(분석값) 농도/기준 검토"
                            rr += 1
                            ws3.Range(f"A{rr}").Value = "구분"
                            ws3.Range(f"B{rr}").Value = "항목"
                            ws3.Range(f"C{rr}").Value = "배출허용기준(H)"
                            ws3.Range(f"D{rr}").Value = "측정농도(E)"
                            rr += 1

                            for label, item, limitv, concv, color in analysis_conc_checks:
                                ws3.Range(f"A{rr}").Value = label
                                ws3.Range(f"B{rr}").Value = item
                                ws3.Range(f"C{rr}").Value = limitv
                                ws3.Range(f"D{rr}").Value = concv
                                ws3.Range(f"D{rr}").Interior.ColorIndex = 3 if color == "RED" else 5
                                rr += 1

                        # (3) 입력(비산먼지) 농도/기준
                        if is_fugitive and fugitive_conc_checks:
                            rr += 1
                            ws3.Range(f"A{rr}").Value = "입력(비산먼지) 농도/기준 검토"
                            rr += 1
                            ws3.Range(f"A{rr}").Value = "구분"
                            ws3.Range(f"B{rr}").Value = "항목"
                            ws3.Range(f"C{rr}").Value = "배출허용기준(G)"
                            ws3.Range(f"D{rr}").Value = "측정농도(E)"
                            rr += 1

                            for label, item, limitv, concv, color in fugitive_conc_checks:
                                ws3.Range(f"A{rr}").Value = label
                                ws3.Range(f"B{rr}").Value = item
                                ws3.Range(f"C{rr}").Value = limitv
                                ws3.Range(f"D{rr}").Value = concv
                                ws3.Range(f"D{rr}").Interior.ColorIndex = 3 if color == "RED" else 5
                                rr += 1

                        # (4) 채취시간/장비 검토
                        rr += 2
                        ws3.Range(f"A{rr}").Value = "채취시간/장비 검토"
                        rr += 1
                        equipment_info = [
                            "[장비요약]",
                            f"가스상(굴뚝)={gas_cnt}대",
                            f"입자상(굴뚝표기)={particle_cnt}대",
                            f"입자상장비={('있음' if has_particle_sampler else '없음')}",
                            f"배출가스측정기={('있음' if has_gas_meter else '없음')}",
                            f"THC측정기={('있음' if has_thc_meter else '없음')}",
                            f"비산먼지장비={('있음' if has_fugitive_dust else '없음')}",
                            f"수분량자동={('있음' if has_moisture_auto else '없음')}",
                            f"수분자동입력={b18_value}"
                        ]
                        for c_idx, val in enumerate(equipment_info, start=1):
                            ws3.Cells(rr, c_idx).Value = val

                        rr += 1

                        # (4) 에러 메시지 출력 방식 변경 (A열: 태그, B열: 메인 내용, C열~: 줄바꿈 상세항목)
                        for msg in device_issues:
                            parts = msg.split("] ", 1)
                            tag = parts[0] + "]" if len(parts) > 1 else "[장비누락]"
                            desc = parts[1] if len(parts) > 1 else msg
                            
                            ws3.Cells(rr, 1).Value = tag
                            ws3.Cells(rr, 2).Value = desc
                            ws3.Range(ws3.Cells(rr, 1), ws3.Cells(rr, 2)).Interior.ColorIndex = 3
                            rr += 1

                        for msg in time_issues:
                            parts = msg.split("] ", 1)
                            tag = parts[0] + "]" if len(parts) > 1 else "[시간오류]"
                            rest = parts[1] if len(parts) > 1 else msg
                            
                            is_err = ("장비초과" in msg) or ("채취시간불일치" in msg) or ("전체시간범위이탈" in msg) or ("시간역전" in msg) or ("[오류]" in msg)
                            color = 3 if is_err else -4142 # 3=빨강, -4142=투명
                            
                            ws3.Cells(rr, 1).Value = tag
                            ws3.Cells(rr, 1).Interior.ColorIndex = color
                            
                            # \n 이 있는 경우 (예: 장비초과 시 동시채취 목록) C열, D열 등으로 분산
                            desc_parts = rest.split("\n")
                            for i, text_part in enumerate(desc_parts):
                                ws3.Cells(rr, 2 + i).Value = text_part.strip("- ") # 앞에 붙은 '-' 기호 제거
                                ws3.Cells(rr, 2 + i).Interior.ColorIndex = color
                            rr += 1

                        # THC 검증 출력
                        if has_thc_item:
                            rr += 1
                            ws3.Range(f"A{rr}").Value = "THC 시트(PF/FID) 복사 검증"
                            rr += 1
                            thc_info = [
                                f"THC장비={c65_val}",
                                f"기대={'THC 측정값(FID)' if is_fid_expect else 'THC 측정값(PF)'}",
                                f"복사됨={thc_copied or '(미복사)'}"
                            ]
                            for c_idx, val in enumerate(thc_info, start=1):
                                ws3.Cells(rr, c_idx).Value = val
                            rr += 1
                            
                            for msg in thc_issues:
                                parts = msg.split("] ", 1)
                                tag = parts[0] + "]" if len(parts) > 1 else "[THC오류]"
                                desc = parts[1] if len(parts) > 1 else msg
                                
                                ws3.Cells(rr, 1).Value = tag
                                ws3.Cells(rr, 2).Value = desc
                                ws3.Range(ws3.Cells(rr, 1), ws3.Cells(rr, 2)).Interior.ColorIndex = 3
                                rr += 1

                        # ✅ [추가] sample-3 시트 데이터 작성이 모두 끝난 후 열 너비 자동 맞춤
                        ws3.Columns("A:I").AutoFit()

                        log(" ▶ sample-3 요약 시트 생성(이상 있음)")
                    else:
                        log(" ▶ sample-3 요약 시트 미생성(이상 없음)")

                except Exception as e:
                    log(f"⚠ sample-3 처리 중 오류(무시): {e}")

            finally:
                # 1. 엑셀 설정 원상복구 (개별 파일 처리 직후)
                try:
                    if excel:
                        excel.ScreenUpdating = True
                        excel.DisplayAlerts = True
                        excel.EnableEvents = True
                        excel.Interactive = True
                        # ✅ 반드시 자동 계산으로 다시 돌려놔야 함!
                        excel.Calculation = -4105  # xlCalculationAutomatic
                except:
                    pass
                
                # 2. 클립보드 비우기
                try:
                    clear_clipboard(excel)
                except:
                    pass
                
                # 3. 원본 파일 닫기 (변수가 존재할 때만 닫도록 안전 처리)
                try:
                    # 'wb_src'가 정의되어 있고, None이 아닐 때만 Close 실행
                    if 'wb_src' in locals() and wb_src is not None:
                        wb_src.Close(SaveChanges=False)
                except:
                    pass

        # --- 모든 샘플 작업 종료 후 ---
        add_fail_sheet(wb_out, fail_list)
        
        try:
            excel.DisplayAlerts = False
            wb_out.Worksheets("Dummy_End").Delete()
            excel.DisplayAlerts = True
        except:
            pass

        try:
            wb_out.Worksheets(1).Activate()
        except:
            pass

        save_path = uniquify_path(out_path)
        wb_out.SaveAs(save_path, FileFormat=51)  # xlsx
        wb_out.Close(SaveChanges=False)

        log("\n=== 완료 ===")
        log(f"생성 파일: {save_path}")

    except Exception as e:
        log(f"\n❌ 실행 중 오류 발생: {e}")

    finally:
        # 4. 최종 앱 설정 복구 (혹시 모르니 한 번 더 실행)
        try:
            if excel:
                excel.ScreenUpdating = True
                excel.DisplayAlerts = True
                excel.Calculation = -4105
        except:
            pass
        
        try:
            clear_clipboard(excel)
        except:
            pass


# ============================================================
# (선택) 단독 실행용
# ============================================================
if __name__ == "__main__":
    # 예시:
    # sample_list = ["A2512314-02", "A2512105-01"]
    # user_name = "수영"
    #
    # 실제 사용 환경에 맞게 입력부/GUI에서 main() 호출하면 됨.
    sample_list = []
    user_name = ""
    main(sample_list, user_name)
