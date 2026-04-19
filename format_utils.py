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
        try:
            return _datetime(1899, 12, 30) + _timedelta(days=float(v))
        except Exception:
            return None
    s = str(v).strip()
    if not s:
        return None
    fmts = [
        "%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y%m%d",
        "%Y-%m-%d %H:%M", "%Y.%m.%d %H:%M", "%Y/%m/%d %H:%M",
        "%Y년 %m월 %d일", "%Y년 %m월 %d일 %H:%M",
        "%H:%M",
    ]
    for fmt in fmts:
        try:
            return _datetime.strptime(s, fmt)
        except Exception:
            pass
    return None


def is_excel_base_date(dt_obj) -> bool:
    """엑셀 기준 날짜(1899-12-30/31) 여부 판별."""
    if dt_obj is None:
        return False
    return dt_obj.date() in (
        _datetime(1899, 12, 30).date(),
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
