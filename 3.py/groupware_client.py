# -*- coding: utf-8 -*-
"""
그룹웨어(FastAPI) 연동 — 탭4 입력완료 후 측정 데이터·PDF 전송, 로컬 정리 엑셀.

API (그룹웨어 측 구현 완료 가정):
  POST /api/external/report_data   — JSON
  POST /api/reports/sync           — multipart: data(JSON) + file(PDF)  [권장]
  POST /api/external/report_pdf    — PDF만 (report_data 선행 시)
"""
from __future__ import annotations

import json
import os
import re
import time
from datetime import datetime

import requests
from openpyxl import Workbook, load_workbook

from config import (
    GROUPWARE_API_TOKEN,
    GROUPWARE_BASE_URL,
    GROUPWARE_ENABLED,
    GROUPWARE_EXCEL_OUTPUT_DIR,
    GROUPWARE_MAX_RETRIES,
    GROUPWARE_REPORT_DATA_PATH,
    GROUPWARE_REPORT_PDF_PATH,
    GROUPWARE_REPORT_SYNC_PATH,
    GROUPWARE_TIMEOUT_SEC,
)
from data_utils import normalize_company, sample_to_datestr
from excel_utils import autofit_columns, parse_measuring_record
from format_utils import format_time
from log_utils import log_error, log_message
from tab4_utils import _norm_rg


def _clean_cell(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _normalize_match_key(value) -> str:
    """
    시설 매칭키 정규화 (E2/E4, 그룹웨어 DB와 동일 규칙으로 비교).
    - 앞뒤·내부 공백 제거, nbsp 제거
    - 시설 보조키(E4) 등 '문의 휴 게소' → '문의휴게소' 형태로 맞춤
    """
    if value is None:
        return ""
    s = str(value).replace("\xa0", " ").strip()
    if not s:
        return ""
    return re.sub(r"\s+", "", s)


def read_input_sheet_keys(excel_path: str) -> dict:
    """
    **대기 전용** — 성적서 「입력」 시트 거래처·시설 매칭키.
    수질 성적서에는 E2/E4 시설 매칭 규칙을 적용하지 않음.

    | 셀 | 필드 | 용도 |
    |----|------|------|
    | H7 | company_name | 업소명 (거래처 1차 식별) |
    | E2 | sems_serial | SEMS 일련번호 (시설 1차 매칭) |
    | E4 | facility_alt | 시설 2차 매칭 (보조키) |
    """
    out = {
        "company_name": "",
        "sems_serial": "",
        "facility_alt": "",
    }
    try:
        wb = load_workbook(excel_path, data_only=True, read_only=True)
    except Exception:
        return out

    try:
        ws = None
        if "입력" in wb.sheetnames:
            ws = wb["입력"]
        else:
            for name in wb.sheetnames:
                if name.replace(" ", "") == "입력":
                    ws = wb[name]
                    break
        if ws is None:
            return out

        out["company_name"] = _clean_cell(ws["H7"].value)
        out["sems_serial"] = _clean_cell(ws["E2"].value)
        out["facility_alt"] = _clean_cell(ws["E4"].value)
    except Exception as e:
        log_error("groupware.read_input_sheet_keys", e)
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return out


def facility_match_keys(keys: dict) -> dict:
    """API payload용 시설 매칭 필드 (정규화 후 전송). E1(굴뚝번호) 미사용."""
    sems_serial = _normalize_match_key(keys.get("sems_serial"))
    facility_alt = _normalize_match_key(keys.get("facility_alt"))
    company_raw = _clean_cell(keys.get("company_name"))
    # facility_name: 거래처카드 시설명 문자열(E4). E2 숫자코드만 있으면 비움(sems_serial로 매칭).
    if facility_alt:
        facility_name = facility_alt
    elif sems_serial and not re.fullmatch(r"\d+", sems_serial):
        facility_name = sems_serial
    else:
        facility_name = ""
    return {
        "company_name": company_raw,
        "company_name_norm": normalize_company(company_raw),
        "sems_serial": sems_serial,
        "facility_alt": facility_alt,
        "facility_match_order": ["sems_serial", "facility_alt"],
        "facility_name": facility_name,
    }


# ---------------------------------------------------------------------------
# 세션 누적 (로컬 정리 엑셀용)
# ---------------------------------------------------------------------------
class GroupwareRunLog:
    """한 번의 자동입력 실행 동안 그룹웨어 전송 기록 + 결과 추적."""

    def __init__(self):
        self.rows: list[dict] = []
        # 시료별 전송 결과: {sample_no: {"data_ok", "pdf_ok", "pdf_attempted", "error", "payload"}}
        self._results: dict[str, dict] = {}

    def add_rows(self, rows: list[dict]):
        self.rows.extend(rows)

    def record_result(self, sample_no: str, result: dict, payload: dict | None = None):
        """전송 결과 기록 (재시도 시 덮어씀)."""
        self._results[sample_no] = {
            "data_ok": result.get("data_ok", False),
            "pdf_ok": result.get("pdf_ok", False),
            "pdf_attempted": result.get("pdf_attempted", False),
            "error": result.get("error", ""),
            "warnings": list(result.get("warnings") or []),
            "verify_key": result.get("verify_key", ""),
            "status": result.get("status", ""),
            "payload": payload or result.get("_payload") or {},
        }

    @property
    def count(self) -> int:
        return len(self.rows)

    def failed_samples(self) -> list[dict]:
        """데이터/PDF 실패 또는 collected_at·measure_date 등 필드 경고가 있는 시료."""
        out = []
        for sno, r in self._results.items():
            data_fail = not r["data_ok"]
            pdf_fail = r["pdf_attempted"] and not r["pdf_ok"]
            warns = r.get("warnings") or []
            field_warn = _has_critical_field_warnings(warns)
            if data_fail or pdf_fail or field_warn:
                out.append({
                    "sample_no": sno,
                    "data_ok": r["data_ok"],
                    "pdf_ok": r["pdf_ok"],
                    "pdf_attempted": r["pdf_attempted"],
                    "error": r["error"],
                    "warnings": warns,
                    "verify_key": r.get("verify_key", ""),
                    "payload": r["payload"],
                })
        return out

    def summary_text(self) -> str:
        """전송 결과 요약 문자열."""
        total = len(self._results)
        if total == 0:
            return "(그룹웨어 전송 없음)"
        ok = sum(
            1 for r in self._results.values()
            if r["data_ok"]
            and (r["pdf_ok"] or not r["pdf_attempted"])
            and not _has_critical_field_warnings(r.get("warnings") or [])
        )
        fail = total - ok
        lines = [f"그룹웨어 전송 결과: 성공 {ok} / 실패·경고 {fail} / 전체 {total}"]
        for sno, r in self._results.items():
            d = "✅" if r["data_ok"] else "❌"
            p = ("✅" if r["pdf_ok"] else ("－" if not r["pdf_attempted"] else "❌"))
            err = f"  ({r['error'][:60]})" if r["error"] else ""
            warns = r.get("warnings") or []
            wtxt = f"  ⚠{warns}" if warns else ""
            vk = f"  verify_key={r.get('verify_key')}" if r.get("verify_key") else ""
            lines.append(f"  {d}데이터 {p}PDF  {sno}{err}{wtxt}{vk}")
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# 설정
# ---------------------------------------------------------------------------
def is_groupware_enabled(override: bool | None = None) -> bool:
    """override가 있으면 그 값, 없으면 config.ini GROUPWARE.ENABLED (기본 ON)."""
    if override is not None:
        return bool(override)
    return GROUPWARE_ENABLED


def _auth_headers() -> dict:
    h = {"Accept": "application/json"}
    token = (GROUPWARE_API_TOKEN or "").strip()
    if token:
        h["Authorization"] = f"Bearer {token}"
    return h


def _url(path: str) -> str:
    base = (GROUPWARE_BASE_URL or "").rstrip("/")
    p = path if path.startswith("/") else f"/{path}"
    return f"{base}{p}"


# ---------------------------------------------------------------------------
# 추출 유틸
# ---------------------------------------------------------------------------
def format_collected_at(start_hm: str, end_hm: str) -> str:
    """채취시간 `13:10~15:30` 형식 (그룹웨어 DB·진위확인 권장)."""
    s = format_time(start_hm) if start_hm else ""
    e = format_time(end_hm) if end_hm else ""
    if s and e:
        return f"{s}~{e}"
    return s or e or ""


def _normalize_measure_date(raw, sample_no: str) -> str:
    """측정일을 항상 YYYY-MM-DD로. 엑셀값 우선, 없으면 시료번호 YYMMDD."""
    s = _clean_cell(raw)
    if s:
        head = s.replace("/", "-").replace(".", "-")
        for fmt, n in (("%Y-%m-%d %H:%M:%S", 19), ("%Y-%m-%d", 10)):
            try:
                return datetime.strptime(head[:n], fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        if len(head) >= 10 and head[4] == "-" and head[7] == "-":
            return head[:10]
    return (sample_to_datestr(sample_no) or "").strip()


def _resolve_air_meta(
    sample_no: str,
    excel_path: str,
    excel_meta: dict | None,
) -> dict:
    """
    excel_meta에 채취시간·날짜가 없으면 엑셀(대기측정기록부)에서 직접 채움.
    meta 없이 재전송해도 collected_at이 비지 않도록.
    """
    meta = dict(excel_meta or {})
    need_start = not _clean_cell(meta.get("채취시작"))
    need_end = not _clean_cell(meta.get("채취끝"))
    need_date = not _clean_cell(meta.get("날짜"))
    need_company = not _clean_cell(meta.get("업소명"))
    if not (need_start or need_end or need_date or need_company):
        return meta
    if not excel_path or not os.path.isfile(excel_path):
        return meta
    try:
        parsed = parse_measuring_record(excel_path, sample_no)
    except Exception as e:
        log_error("groupware._resolve_air_meta", e)
        return meta
    if need_start and parsed.get("채취시작"):
        meta["채취시작"] = parsed["채취시작"]
    if need_end and parsed.get("채취끝"):
        meta["채취끝"] = parsed["채취끝"]
    if need_date and parsed.get("날짜"):
        meta["날짜"] = parsed["날짜"]
    if need_company and parsed.get("업소명"):
        meta["업소명"] = parsed["업소명"]
    return meta


def _has_critical_field_warnings(warnings: list) -> bool:
    """API warnings 중 채취시간·측정일·시설명 누락/실패."""
    for w in warnings or []:
        s = str(w).lower()
        if any(
            k in s
            for k in (
                "collected_at",
                "measure_date",
                "채취시간",
                "측정일",
                "미전송",
                "시설명",
                "facility",
            )
        ):
            return True
    return False


def _parse_api_body(resp: requests.Response) -> dict:
    """응답 JSON에서 status / warnings / verify_key 추출."""
    info = {
        "status": "",
        "warnings": [],
        "verify_key": "",
        "action": "",
        "message": "",
        "raw": "",
    }
    try:
        data = resp.json()
    except Exception:
        info["raw"] = (resp.text or "")[:500]
        return info
    if not isinstance(data, dict):
        info["raw"] = str(data)[:500]
        return info
    info["status"] = str(data.get("status") or "")
    info["action"] = str(data.get("action") or "")
    info["verify_key"] = str(data.get("verify_key") or "")
    info["message"] = str(data.get("message") or data.get("detail") or "")
    warns = data.get("warnings")
    if isinstance(warns, list):
        info["warnings"] = [str(w) for w in warns if w is not None]
    elif warns:
        info["warnings"] = [str(warns)]
    return info


def _print_api_feedback(sample_no: str, info: dict, label: str) -> None:
    status = info.get("status") or info.get("action") or ""
    vk = info.get("verify_key") or ""
    warns = info.get("warnings") or []
    if status:
        print(f"   ← {label} status={status}" + (f" verify_key={vk}" if vk else ""))
    elif vk:
        print(f"   ← {label} verify_key={vk}")
    if warns:
        print(f"⚠ 그룹웨어 warnings ({sample_no}):")
        for w in warns:
            print(f"   · {w}")
        log_message(f"groupware.warnings {sample_no}: {warns}")


def _validate_payload_for_send(payload: dict) -> list[str]:
    """전송 전 필수 필드 검사. 빈 collected_at으로 덮어쓰지 않도록."""
    errs: list[str] = []
    if not _clean_cell(payload.get("sample_no")):
        errs.append("sample_no 없음")
    if not _clean_cell(payload.get("company_name")) and not _clean_cell(payload.get("biz_no")):
        errs.append("company_name(또는 biz_no) 없음")
    if not _clean_cell(payload.get("collected_at")):
        errs.append("collected_at(채취시간) 없음 — 엑셀 채취시작/끝을 확인하세요")
    if not _clean_cell(payload.get("measure_date")):
        errs.append("measure_date(측정일) 없음")
    if not (payload.get("measurements") or []):
        errs.append("measurements 비어 있음")
    return errs


def read_biz_no(excel_path: str) -> str:
    """사업자번호 — 성적서 엑셀에는 없음. 그룹웨어 호환용 빈 문자열."""
    return ""


def _header_index_map(headers: list) -> dict[str, int]:
    m: dict[str, int] = {}
    for i, h in enumerate(headers or []):
        k = _norm_rg(h)
        if k and k not in m:
            m[k] = i
    return m


def _find_col_idx(hmap: dict[str, int], *candidates: str) -> int | None:
    for c in candidates:
        k = _norm_rg(c)
        if k in hmap:
            return hmap[k]
    for k, idx in hmap.items():
        for c in candidates:
            ck = _norm_rg(c)
            if ck and (ck in k or k in ck):
                return idx
    return None


def _cell_str(vals: list, idx: int | None) -> str:
    if idx is None or idx < 0 or idx >= len(vals):
        return ""
    v = vals[idx]
    return "" if v is None else str(v).strip()


def _conc_str(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    try:
        f = float(s.replace(",", ""))
        if f == int(f):
            return str(int(f))
        return str(f)
    except ValueError:
        return s


def _read_air_flow_rates(wb) -> tuple[str, str]:
    """
    대기측정기록부 — 배출가스유량 보정전(J14)·보정후(N14).
    excel_utils.parse_measuring_record 의 배출가스유량전/후 와 동일 셀.
    """
    ws = None
    for name in wb.sheetnames:
        if name.replace(" ", "") == "대기측정기록부":
            ws = wb[name]
            break
    if ws is None:
        return "", ""
    # J=10, N=14
    flow_pre = _conc_str(ws.cell(row=14, column=10).value)   # 보정전(미적용)
    flow_post = _conc_str(ws.cell(row=14, column=14).value)  # 보정후(적용1/적용2)
    return flow_pre, flow_post


def _norm_item_key(name: str) -> str:
    """항목명 매칭용 정규화 (공백 제거)."""
    return re.sub(r"\s+", "", str(name or "").strip())


def _read_record_item_units(wb) -> dict[str, str]:
    """
    대기측정기록부 측정결과 표 — B열 항목, F열 단위.
    일반: '측정항목' 헤더 아래 구간.
    비산먼지: 상단 요약행(예: B19/F19 = 비산먼지 / mg/S㎥)도 포함.
    반환: {정규화항목명: 단위}
    """
    ws = None
    for name in wb.sheetnames:
        if name.replace(" ", "") == "대기측정기록부":
            ws = wb[name]
            break
    if ws is None:
        return {}

    out: dict[str, str] = {}

    def _add_row(r: int) -> bool:
        raw_item = ws.cell(row=r, column=2).value  # B
        if raw_item is None or not str(raw_item).strip():
            return False
        item_s = str(raw_item).strip()
        key = _norm_item_key(item_s)
        stop_keys = ("분석기간", "종합의견", "채취일시", "방지시설", "현장기상", "측정항목")
        if any(s in key for s in stop_keys):
            return False
        unit_v = ws.cell(row=r, column=6).value  # F
        unit_s = "" if unit_v is None else str(unit_v).strip()
        if unit_s:
            out[key] = unit_s
        return True

    # 비산먼지 양식: 측정항목 헤더 위쪽(대략 15~28행) B/F
    for r in range(15, 29):
        _add_row(r)

    header_rows: list[int] = []
    for r in range(20, 60):
        b = _norm_item_key(ws.cell(row=r, column=2).value)
        if "측정항목" in b:
            header_rows.append(r)
    if header_rows:
        start = header_rows[0] + 1
        end = header_rows[1] if len(header_rows) > 1 else min(start + 30, 80)
        for r in range(start, end):
            raw_item = ws.cell(row=r, column=2).value
            if raw_item is None or not str(raw_item).strip():
                continue
            key = _norm_item_key(str(raw_item).strip())
            if any(s in key for s in ("분석기간", "종합의견", "채취일시", "방지시설", "현장기상")):
                break
            _add_row(r)

    return out


def _flow_for_air_ratio(air_ratio: str, flow_pre: str, flow_post: str) -> str:
    """공기비적용 값에 맞는 유량 선택. '적용' 포함(적용1·적용2) → 보정후, 그 외 → 보정전."""
    s = (air_ratio or "").strip()
    if "적용" in s and "미적용" not in s:
        return flow_post
    return flow_pre


def read_air_measurements(excel_path: str) -> list[dict]:
    """
    대기 — 입력(분석값)에서 항목·농도·공기비, 대기측정기록부에서 단위(B/F)·유량(J14/N14).
    농도는 '농도'/'측정농도' 열. 단위는 기록부 F열을 항목명(B열)으로 매칭.
    유량은 항목별 '공기비적용'(미적용/적용1/적용2)에 따라 J14 또는 N14.
    """
    out: list[dict] = []
    try:
        # read_only=False: 대기측정기록부 임의 셀(J14/N14) + 헤더 맵 안정 접근
        wb = load_workbook(excel_path, data_only=True, read_only=False)
    except Exception:
        return out

    try:
        flow_pre, flow_post = _read_air_flow_rates(wb)
        unit_map = _read_record_item_units(wb)

        ws = None
        for name in wb.sheetnames:
            if name.replace(" ", "") == "입력(분석값)":
                ws = wb[name]
                break
        if ws is None:
            for name in wb.sheetnames:
                if "입력" in name and "분석" in name:
                    ws = wb[name]
                    break
        if ws is None:
            return out

        header_map: dict[str, int] = {}
        for col in range(1, 40):
            v = ws.cell(row=1, column=col).value
            if v:
                header_map[str(v).strip()] = col

        def col_of(*names):
            for n in names:
                if n in header_map:
                    return header_map[n]
            return None

        c_stack = col_of("배출구", "굴뚝", "배출구번호", "시설", "시설명")
        c_unit = col_of("단위", "측정단위")
        c_conc = col_of("측정농도", "농도", "측정값", "분석결과")
        c_air = col_of("공기비적용", "공기비")
        if not c_conc:
            c_conc = 5  # E열 관례

        for r in range(2, 65):
            flag = ws.cell(row=r, column=1).value
            if flag is None or str(flag).strip() == "":
                continue
            item = ws.cell(row=r, column=2).value
            if not item:
                continue
            item_s = str(item).strip()
            if not item_s:
                continue

            stack = ""
            if c_stack:
                sv = ws.cell(row=r, column=c_stack).value
                stack = "" if sv is None else str(sv).strip()

            conc_val = ws.cell(row=r, column=c_conc).value if c_conc else None
            conc = _conc_str(conc_val)
            if not conc:
                continue

            unit = ""
            if c_unit:
                uv = ws.cell(row=r, column=c_unit).value
                unit = "" if uv is None else str(uv).strip()
            # 입력(분석값)에 단위 열이 없으면 대기측정기록부 F열 매칭
            if not unit:
                unit = unit_map.get(_norm_item_key(item_s), "")

            air_ratio = ""
            if c_air:
                av = ws.cell(row=r, column=c_air).value
                air_ratio = "" if av is None else str(av).strip()

            out.append({
                "stack": stack,
                "item_name": item_s,
                "concentration": conc,
                "unit": unit,
                "flow_rate": _flow_for_air_ratio(air_ratio, flow_pre, flow_post),
            })
    except Exception as e:
        log_error("groupware.read_air_measurements", e)
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return out


def read_water_measurements(tab4_meta: dict) -> list[dict]:
    """수질 — 탭4 매크로 headers/rows에서 항목·농도 추출."""
    headers = tab4_meta.get("headers") or []
    rows = tab4_meta.get("rows") or []
    hmap = _header_index_map(headers)

    i_item = _find_col_idx(hmap, "측정항목", "항목", "분석항목")
    i_stack = _find_col_idx(hmap, "배출구", "굴뚝", "시설", "시설명", "측정지점")
    i_conc = _find_col_idx(hmap, "측정농도", "농도", "측정값", "분석결과", "결과")
    i_unit = _find_col_idx(hmap, "단위", "측정단위")

    out: list[dict] = []
    for row_vals in rows:
        item = _cell_str(row_vals, i_item)
        if not item and row_vals:
            item = _norm_rg(row_vals[0])
            if len(row_vals) > 1 and not item:
                item = _norm_rg(row_vals[1])
        if not item:
            continue

        conc = _conc_str(_cell_str(row_vals, i_conc) if i_conc is not None else "")
        if not conc and i_conc is None and len(row_vals) > 4:
            conc = _conc_str(row_vals[4])

        if not conc:
            continue

        out.append({
            "stack": _cell_str(row_vals, i_stack),
            "item_name": item,
            "concentration": conc,
            "unit": _cell_str(row_vals, i_unit),
        })
    return out


def _first_stack(measurements: list[dict]) -> str:
    for m in measurements:
        s = (m.get("stack") or "").strip()
        if s:
            return s
    return ""


def _merge_input_keys(excel_path: str, excel_meta: dict | None) -> dict:
    """입력 시트 E2/E4/H7 우선, parse_measuring_record 업소명은 보조."""
    keys = read_input_sheet_keys(excel_path)
    if excel_meta and not keys["company_name"]:
        keys["company_name"] = _clean_cell(excel_meta.get("업소명"))
    return keys


def build_payload_air(
    sample_no: str,
    excel_path: str,
    excel_meta: dict | None = None,
    tab4_meta: dict | None = None,
) -> dict:
    """
    대기 성적서 → 그룹웨어 ReportDataIn.
    collected_at / measure_date 는 excel_meta가 비어도 엑셀에서 채움.
    facility_name 은 E4(보조키)→E2 — 굴뚝번호(stack)로 덮지 않음.
    """
    meta = _resolve_air_meta(sample_no, excel_path, excel_meta)
    input_keys = _merge_input_keys(excel_path, meta)
    fac = facility_match_keys(input_keys)

    measurements = read_air_measurements(excel_path)

    collect_start = format_time(meta.get("채취시작", "")) if meta.get("채취시작") else ""
    collect_end = format_time(meta.get("채취끝", "")) if meta.get("채취끝") else ""
    collected_at = format_collected_at(collect_start, collect_end)
    measure_date = _normalize_measure_date(meta.get("날짜"), sample_no)

    # 시설명: 그룹웨어 거래처카드와 맞출 문자열 (E4→E2). stack 코드 사용 금지.
    facility_name = fac.get("facility_name") or ""

    return {
        "media": "air",
        "sample_no": sample_no,
        "collected_at": collected_at,
        "collect_start": collect_start,
        "collect_end": collect_end,
        "measure_date": measure_date,
        "biz_no": "",
        **fac,
        "facility_name": facility_name,
        "measurements": measurements,
        "source_excel": os.path.abspath(excel_path) if excel_path else "",
    }


# ---------------------------------------------------------------------------
# 수질 — 그룹웨어 연동 미적용 (대기 E2/E4 시설매칭 규칙 해당 없음)
# 프롬프트 3 수질 경로는 eco_input._main_water 에서 호출 주석 처리됨.
# ---------------------------------------------------------------------------
def build_payload_water(
    sample_no: str,
    excel_path: str,
    tab4_meta: dict,
    company_name: str = "",
) -> dict:
    """미사용 — 수질은 그룹웨어 연동 대상 아님. 호출 시 명시적 오류."""
    raise NotImplementedError(
        "수질 성적서는 그룹웨어(프롬프트3) 연동 대상이 아닙니다. 대기만 사용하세요."
    )


# ---------------------------------------------------------------------------
# API 전송
# ---------------------------------------------------------------------------
def _request_with_retry(method: str, url: str, **kwargs) -> requests.Response | None:
    retries = max(1, int(GROUPWARE_MAX_RETRIES))
    timeout = float(kwargs.pop("timeout", GROUPWARE_TIMEOUT_SEC))
    last_err = None

    for attempt in range(1, retries + 1):
        try:
            fn = requests.post if method.upper() == "POST" else requests.get
            resp = fn(url, timeout=timeout, **kwargs)
            if resp.status_code >= 500 and attempt < retries:
                time.sleep(1.0 * attempt)
                continue
            return resp
        except Exception as e:
            last_err = e
            log_error(f"groupware.{method} {url} (attempt {attempt})", e)
            if attempt < retries:
                time.sleep(1.0 * attempt)

    if last_err:
        raise last_err
    return None


def post_report_data(payload: dict) -> tuple[bool, str, dict]:
    """POST /api/external/report_data → (ok, error, api_info)."""
    url = _url(GROUPWARE_REPORT_DATA_PATH)
    empty_info: dict = {"status": "", "warnings": [], "verify_key": ""}
    try:
        resp = _request_with_retry(
            "POST",
            url,
            headers={**_auth_headers(), "Content-Type": "application/json"},
            json=payload,
        )
        if resp is None:
            return False, "no response", empty_info
        info = _parse_api_body(resp)
        if resp.status_code >= 400:
            msg = info.get("message") or info.get("raw") or resp.text[:300]
            print(f"⚠ 그룹웨어 report_data 실패 ({resp.status_code}): {msg}")
            return False, msg, info
        print(f"✅ 그룹웨어 데이터 전송: {payload.get('sample_no')}")
        _print_api_feedback(payload.get("sample_no", ""), info, "report_data")
        return True, "", info
    except Exception as e:
        print(f"⚠ 그룹웨어 report_data 오류 ({payload.get('sample_no')}): {e}")
        log_error("groupware.post_report_data", e)
        return False, str(e), empty_info


def post_report_sync(payload: dict, pdf_path: str) -> tuple[bool, str, dict]:
    """POST /api/reports/sync — JSON + PDF 한 번에 → (ok, error, api_info)."""
    empty_info: dict = {"status": "", "warnings": [], "verify_key": ""}
    if not pdf_path or not os.path.isfile(pdf_path):
        return False, "pdf missing", empty_info

    url = _url(GROUPWARE_REPORT_SYNC_PATH)
    data_json = json.dumps(payload, ensure_ascii=False)
    try:
        with open(pdf_path, "rb") as f:
            resp = _request_with_retry(
                "POST",
                url,
                headers=_auth_headers(),
                data={"data": data_json},
                files={
                    "file": (
                        os.path.basename(pdf_path),
                        f,
                        "application/pdf",
                    )
                },
                timeout=max(float(GROUPWARE_TIMEOUT_SEC), 120.0),
            )
        if resp is None:
            return False, "no response", empty_info
        info = _parse_api_body(resp)
        if resp.status_code >= 400:
            msg = info.get("message") or info.get("raw") or resp.text[:300]
            print(f"⚠ 그룹웨어 reports/sync 실패 ({resp.status_code}): {msg}")
            return False, msg, info
        print(f"✅ 그룹웨어 sync(PDF+데이터): {payload.get('sample_no')}")
        _print_api_feedback(payload.get("sample_no", ""), info, "reports/sync")
        return True, "", info
    except Exception as e:
        print(f"⚠ 그룹웨어 reports/sync 오류 ({payload.get('sample_no')}): {e}")
        log_error("groupware.post_report_sync", e)
        return False, str(e), empty_info


def post_report_pdf(payload: dict, pdf_path: str) -> tuple[bool, str, dict]:
    """POST /api/external/report_pdf — PDF만 → (ok, error, api_info)."""
    empty_info: dict = {"status": "", "warnings": [], "verify_key": ""}
    if not pdf_path or not os.path.isfile(pdf_path):
        return False, "pdf missing", empty_info

    url = _url(GROUPWARE_REPORT_PDF_PATH)
    form = {
        "sample_no": payload.get("sample_no", ""),
        "media": payload.get("media", ""),
        "company_name": payload.get("company_name", ""),
        "sems_serial": payload.get("sems_serial", ""),
        "facility_alt": payload.get("facility_alt", ""),
        "collected_at": payload.get("collected_at", ""),
        "measure_date": payload.get("measure_date", ""),
    }
    try:
        with open(pdf_path, "rb") as f:
            resp = _request_with_retry(
                "POST",
                url,
                headers=_auth_headers(),
                data=form,
                files={
                    "file": (
                        os.path.basename(pdf_path),
                        f,
                        "application/pdf",
                    )
                },
                timeout=max(float(GROUPWARE_TIMEOUT_SEC), 120.0),
            )
        if resp is None:
            return False, "no response", empty_info
        info = _parse_api_body(resp)
        if resp.status_code >= 400:
            msg = info.get("message") or info.get("raw") or resp.text[:300]
            print(f"⚠ 그룹웨어 report_pdf 실패 ({resp.status_code}): {msg}")
            return False, msg, info
        print(f"✅ 그룹웨어 PDF 전송: {payload.get('sample_no')}")
        _print_api_feedback(payload.get("sample_no", ""), info, "report_pdf")
        return True, "", info
    except Exception as e:
        print(f"⚠ 그룹웨어 report_pdf 오류 ({payload.get('sample_no')}): {e}")
        log_error("groupware.post_report_pdf", e)
        return False, str(e), empty_info


def sync_to_groupware(payload: dict, pdf_path: str | None = None) -> dict:
    """
    그룹웨어 전송 (실패해도 예외 없음).
    반환: {data_ok, pdf_ok, pdf_attempted, error, warnings, verify_key, status}
    """
    result = {
        "data_ok": False,
        "pdf_ok": False,
        "pdf_attempted": bool(pdf_path),
        "error": "",
        "warnings": [],
        "verify_key": "",
        "status": "",
    }

    # 빈 collected_at 재전송 금지 (기존 DB 값 지우기 방지 + 최초 누락 방지)
    val_errs = _validate_payload_for_send(payload)
    if val_errs:
        msg = "; ".join(val_errs)
        print(f"❌ 그룹웨어 전송 중단 ({payload.get('sample_no')}): {msg}")
        result["error"] = msg
        result["warnings"] = val_errs
        return result

    print(
        f"▶ 그룹웨어 payload: sample_no={payload.get('sample_no')} "
        f"collected_at={payload.get('collected_at')!r} "
        f"measure_date={payload.get('measure_date')!r} "
        f"facility_name={payload.get('facility_name')!r} "
        f"company={payload.get('company_name')!r}"
    )
    log_message(
        f"groupware.send {payload.get('sample_no')} "
        f"collected_at={payload.get('collected_at')} "
        f"measure_date={payload.get('measure_date')} "
        f"facility_name={payload.get('facility_name')}"
    )

    def _merge_info(info: dict) -> None:
        result["warnings"] = list(info.get("warnings") or [])
        result["verify_key"] = info.get("verify_key") or result["verify_key"]
        result["status"] = info.get("status") or info.get("action") or result["status"]

    if pdf_path and os.path.isfile(pdf_path):
        ok, err, info = post_report_sync(payload, pdf_path)
        _merge_info(info)
        if ok:
            result["data_ok"] = True
            result["pdf_ok"] = True
            if _has_critical_field_warnings(result["warnings"]):
                result["error"] = "API warnings: " + "; ".join(result["warnings"])
            return result
        result["error"] = err

    ok, err, info = post_report_data(payload)
    _merge_info(info)
    result["data_ok"] = ok
    if not ok:
        result["error"] = err
        return result

    if _has_critical_field_warnings(result["warnings"]):
        result["error"] = "API warnings: " + "; ".join(result["warnings"])

    if pdf_path and os.path.isfile(pdf_path):
        pok, perr, pinfo = post_report_pdf(payload, pdf_path)
        # PDF 응답 warnings는 합침
        extra = list(pinfo.get("warnings") or [])
        if extra:
            result["warnings"] = list(result["warnings"]) + extra
        if pinfo.get("verify_key"):
            result["verify_key"] = pinfo["verify_key"]
        result["pdf_ok"] = pok
        if not pok:
            result["error"] = perr

    return result


def payload_to_excel_rows(payload: dict, sync_result: dict) -> list[dict]:
    """로컬 정리 엑셀용 flat 행 (항목별 1행)."""
    base = {
        "시료번호": payload.get("sample_no", ""),
        "매체": payload.get("media", ""),
        "업소명": payload.get("company_name", ""),
        "업소명_정규화": payload.get("company_name_norm", ""),
        "SEMS일련번호_E2": payload.get("sems_serial", ""),
        "시설매칭보조_E4": payload.get("facility_alt", ""),
        "채취시간": payload.get("collected_at", ""),
        "측정일": payload.get("measure_date", ""),
        "시설명": payload.get("facility_name", ""),
        "데이터전송": "OK" if sync_result.get("data_ok") else "FAIL",
        "PDF전송": (
            "OK" if sync_result.get("pdf_ok")
            else ("N/A" if not sync_result.get("pdf_attempted") else "FAIL")
        ),
        "비고": sync_result.get("error", ""),
        "원본엑셀": payload.get("source_excel", ""),
    }
    measurements = payload.get("measurements") or []
    if not measurements:
        return [dict(base, 굴뚝="", 항목="", 농도="", 단위="", 유량="")]

    rows = []
    for m in measurements:
        rows.append(dict(
            base,
            굴뚝=m.get("stack", ""),
            항목=m.get("item_name", ""),
            농도=m.get("concentration", ""),
            단위=m.get("unit", ""),
            유량=m.get("flow_rate", ""),
        ))
    return rows


def sync_tab4_complete(
    gw_log: GroupwareRunLog | None,
    *,
    media: str,
    sample_no: str,
    excel_path: str,
    tab4_meta: dict | None = None,
    excel_meta: dict | None = None,
    pdf_path: str | None = None,
    company_name: str = "",
    _log_to_gw_log: bool = True,
) -> dict:
    """
    탭4 입력완료 직후 호출. payload 조립 → API 전송 → gw_log 누적.
    **대기(air) 전용** — 수질은 연동하지 않음.
    _log_to_gw_log=False 로 호출하면 gw_log 누적을 건너뜀 (재시도 중간 호출용).
    """
    if media == "water":
        print("⚠ 그룹웨어 연동: 수질은 미지원 (스킵)")
        return {"data_ok": False, "pdf_ok": False, "pdf_attempted": False, "error": "water skipped"}

    payload = build_payload_air(
        sample_no, excel_path, excel_meta=excel_meta, tab4_meta=tab4_meta
    )

    print(
        f"▶ 그룹웨어 전송 준비: {sample_no} "
        f"(항목 {len(payload.get('measurements', []))}개, "
        f"채취={payload.get('collected_at')!r}, "
        f"측정일={payload.get('measure_date')!r})"
    )
    sync_result = sync_to_groupware(payload, pdf_path=pdf_path)

    if gw_log is not None:
        if _log_to_gw_log:
            gw_log.add_rows(payload_to_excel_rows(payload, sync_result))
        # 결과는 항상 기록 (재시도 시 마지막 결과로 덮어씀)
        gw_log.record_result(sample_no, sync_result, payload)

    # payload를 결과에 붙여 재시도 시 재사용 가능하게
    sync_result["_payload"] = payload
    return sync_result


def export_groupware_summary(
    gw_log: GroupwareRunLog | None,
    out_dir: str | None = None,
) -> str | None:
    """
    전송 기록을 로컬 엑셀로 저장 (오토핏).
    경로: {EXCEL_OUTPUT_DIR}/{YYYY}년/{M}월/그룹웨어전송_YYYYMMDD_HHMMSS.xlsx
    """
    if gw_log is None or not gw_log.rows:
        return None

    now = datetime.now()
    base_dir = out_dir or GROUPWARE_EXCEL_OUTPUT_DIR
    # 6.그룹웨어전송\2026년\7월\...
    out_dir = os.path.join(base_dir, f"{now.year}년", f"{now.month}월")
    os.makedirs(out_dir, exist_ok=True)

    stamp = now.strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(out_dir, f"그룹웨어전송_{stamp}.xlsx")

    headers = [
        "시료번호", "매체", "업소명", "업소명_정규화", "SEMS일련번호_E2", "시설매칭보조_E4",
        "채취시간", "측정일", "시설명",
        "굴뚝", "항목", "농도", "단위", "유량", "데이터전송", "PDF전송", "비고", "원본엑셀",
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "전송기록"
    ws.append(headers)
    for row in gw_log.rows:
        ws.append([row.get(h, "") for h in headers])

    last_col = chr(ord("A") + len(headers) - 1)
    autofit_columns(ws, f"A:{last_col}")
    wb.save(out_path)
    wb.close()

    print(f"✅ 그룹웨어 전송 정리 엑셀: {out_path}")
    return out_path
