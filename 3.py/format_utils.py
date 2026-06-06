# -*- coding: utf-8 -*-
"""
공통 형식 변환 함수
"""

from datetime import datetime, time, date
from typing import Union


def format_time(t: Union[str, time, datetime, None],
               include_seconds: bool = False,
               sep: str = ":") -> str:
    """시간 포맷팅"""
    if t is None or t == "":
        return ""
    
    try:
        if isinstance(t, datetime):
            t = t.time()
        
        if isinstance(t, time):
            if include_seconds:
                return f"{t.hour:02d}{sep}{t.minute:02d}{sep}{t.second:02d}"
            else:
                return f"{t.hour:02d}{sep}{t.minute:02d}"
        
        s = str(t).strip()
        parts = s.split(":")
        
        if len(parts) >= 2:
            hh = int(parts[0])
            mm = int(parts[1])
            ss = int(parts[2]) if len(parts) >= 3 else 0
            
            if include_seconds:
                return f"{hh:02d}{sep}{mm:02d}{sep}{ss:02d}"
            else:
                return f"{hh:02d}{sep}{mm:02d}"
    except:
        pass
    
    return ""


def format_float(value: Union[str, int, float, None],
                decimal_places: int = 1,
                default: str = "") -> str:
    """숫자 포맷팅"""
    if value is None or value == "":
        return default
    
    try:
        f = float(str(value).replace(",", ""))
        
        if abs(f - round(f)) < 1e-9:
            return str(int(round(f)))
        
        format_spec = f"{{:.{decimal_places}f}}"
        return format_spec.format(f)
    except:
        return default


def format_float_fixed(value, decimal_places: int = 4, default: str = "") -> str:
    """항상 고정 소수점 자릿수로 포맷 (trailing zero 유지, 예: 0.0280 → '0.0280')"""
    if value is None or value == "":
        return default
    try:
        f = float(str(value).replace(",", ""))
        return f"{f:.{decimal_places}f}"
    except:
        return default

# 호환성 래퍼
trim_hm = lambda t: format_time(t, include_seconds=False)
to_f1 = lambda v: format_float(v, decimal_places=1)
to_f2 = lambda v: format_float(v, decimal_places=2)
to_f4 = lambda v: format_float_fixed(v, decimal_places=4)
trim_time_to_hm = lambda t: format_time(t, include_seconds=False)
to_float1 = lambda v: format_float(v, decimal_places=1)
to_float2 = lambda v: format_float(v, decimal_places=2)

# ============================================================
# 날짜/숫자 파싱 유틸 (report_check.py에서 추출)
# ============================================================
import re as _re
from datetime import datetime as _datetime, timedelta as _timedelta

_num_only_re = _re.compile(r"^-?\d+(\.\d+)?$")
_extract_nums_re = _re.compile(r"-?\d+(?:\.\d+)?")
_excel_base_datetime = _datetime(1899, 12, 30)


def to_float_if_pure_number(v):
    """순수 숫자 문자열 또는 int/float이면 float 반환, 아니면 None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if not s:
        return None
    if _num_only_re.match(s):
        try:
            return float(s)
        except Exception:
            return None
    return None


def extract_numbers_from_cell(v) -> list:
    """셀값 문자열에서 숫자(float) 목록 추출."""
    if v is None:
        return []
    s = str(v).replace(",", "")
    out = []
    for x in _extract_nums_re.findall(s):
        try:
            out.append(float(x))
        except Exception:
            pass
    return out


def excel_serial_to_datetime(value):
    """엑셀 serial 숫자를 datetime 으로 변환"""
    if not isinstance(value, (int, float)):
        return None
    try:
        return _excel_base_datetime + _timedelta(days=float(value))
    except Exception:
        return None


def parse_datetime_text(value, fmts=None):
    """문자열을 지정된 datetime 포맷들로 파싱"""
    if value is None:
        return None
    if isinstance(value, _datetime):
        return value

    s = str(value).strip()
    if not s:
        return None

    formats = fmts or (
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
    )
    for fmt in formats:
        try:
            return _datetime.strptime(s, fmt)
        except Exception:
            pass
    return None


def to_datetime_if_possible(v):
    """
    다양한 형식의 값을 datetime으로 변환.
    엑셀 serial 숫자, 문자열(YYYY-MM-DD 등), datetime 객체 모두 지원.
    """
    if v is None:
        return None
    if isinstance(v, _datetime):
        return v
    if isinstance(v, (int, float)):
        return excel_serial_to_datetime(v)
    s = str(v).strip()
    if not s:
        return None
    fmts = [
        "%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d",
        "%Y-%m-%d %H:%M", "%Y.%m.%d %H:%M", "%Y/%m/%d %H:%M",
        "%Y년 %m월 %d일", "%Y년 %m월 %d일 %H:%M",
        "%H:%M",
    ]
    return parse_datetime_text(s, fmts)


def is_excel_base_date(dt_obj) -> bool:
    """엑셀 기준 날짜(1899-12-30/31) 여부 판별."""
    if dt_obj is None:
        return False
    return dt_obj.date() in (
        _excel_base_datetime.date(),
        _datetime(1899, 12, 31).date()
    )


def align_to_date(base_dt, dt_obj):
    """엑셀 기준 날짜로 파싱된 dt_obj의 날짜를 base_dt 날짜로 교정."""
    if base_dt is None or dt_obj is None:
        return dt_obj
    if is_excel_base_date(dt_obj) and not is_excel_base_date(base_dt):
        return _datetime(base_dt.year, base_dt.month, base_dt.day,
                         dt_obj.hour, dt_obj.minute, dt_obj.second)
    return dt_obj


def strip_seconds(dt_obj):
    """datetime에서 초/마이크로초 제거."""
    if dt_obj is None:
        return None
    try:
        return dt_obj.replace(second=0, microsecond=0)
    except Exception:
        return dt_obj


def fmt_hhmm(dt_obj) -> str:
    """datetime → 'HH:MM' 문자열."""
    if dt_obj is None:
        return ""
    try:
        return dt_obj.strftime("%H:%M")
    except Exception:
        return str(dt_obj)


def fmt_range(start, end) -> str:
    """두 datetime을 'HH:MM~HH:MM' 형식으로 포맷."""
    return f"{fmt_hhmm(start)}~{fmt_hhmm(end)}" if (start and end) else ""


def _strip_all_spaces(s: str) -> str:
    return "".join(str(s).split())


def staff_name_before_slash(value) -> str:
    """인력 '이름 / 팀 / 직급' → 첫 '/' 앞, 공백 제거."""
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    head = s.split("/", 1)[0].strip()
    return _strip_all_spaces(head)


def vehicle_plate_after_slash(value) -> str:
    """차량 '그랜드 스타렉스 / 81주6787' → 첫 '/' 뒤, 공백 제거. '/' 없으면 전체."""
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    tail = s.split("/", 1)[1].strip() if "/" in s else s
    return _strip_all_spaces(tail)


def equipment_name_before_slash(value) -> str:
    """장비 '대기배출가스측정기5 / Optima 7 / …' → 첫 '/' 앞, 공백 제거."""
    return staff_name_before_slash(value)


def _dedupe_tab1_tokens(normalize_fn, values) -> list:
    out, seen = [], set()
    for raw in values or []:
        tok = normalize_fn(raw)
        if tok and tok not in seen:
            seen.add(tok)
            out.append(tok)
    return out


def normalize_tab1_staff_list(values) -> list:
    return _dedupe_tab1_tokens(staff_name_before_slash, values)


def normalize_tab1_vehicle_list(values) -> list:
    return _dedupe_tab1_tokens(vehicle_plate_after_slash, values)


def normalize_tab1_equipment_list(values) -> list:
    return _dedupe_tab1_tokens(equipment_name_before_slash, values)


def normalize_tab1_select_field(field: str, values) -> list:
    """탭1 인력·차량·장비 — eco_input 등록·eco_check 비교 공통 규칙."""
    if field == "인력":
        return normalize_tab1_staff_list(values)
    if field == "차량":
        return normalize_tab1_vehicle_list(values)
    if field == "장비":
        return normalize_tab1_equipment_list(values)
    return [str(x).strip() for x in (values or []) if str(x).strip()]
