# -*- coding: utf-8 -*-
"""
공통 데이터 정규화 & 파싱 함수
- Vehicle_operation_log.py, eco_input.py, eco_check.py, receipt.py에서 추출
"""

import re
from datetime import datetime, time, date


def normalize_company(name):
    """업소명 정규화"""
    if name is None:
        return ""
    s = str(name)
    s = s.replace("(주)", "").replace("㈜", "")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^가-힣0-9A-Za-z]", "", s)
    return s


def normalize_plate(text):
    """차량번호 정규화"""
    if text is None:
        return ""
    s = str(text).strip()
    m = re.search(r"(\d{2,3})\s*([가-힣])\s*(\d{4})", s)
    if m:
        return f"{m.group(1)}{m.group(2)}{m.group(3)}"
    s2 = re.sub(r"\s+", "", s)
    m2 = re.search(r"(\d{2,3})([가-힣])(\d{4})", s2)
    if m2:
        return f"{m2.group(1)}{m2.group(2)}{m2.group(3)}"
    return ""


def normalize_name(text):
    """단일 기술인력 정규화"""
    if text is None:
        return ""
    s = str(text).strip()
    if not s:
        return ""
    first = re.split(r"\s*[/,|｜]\s*", s, maxsplit=1)[0].strip()
    name = "".join(re.findall(r"[가-힣]", first))
    return name


def normalize_names_cell(text):
    """운행일지 사용인력 정규화"""
    if text is None:
        return []
    s = str(text).strip()
    if s == "" or s.lower() == "nan":
        return []
    s = s.replace(" ", "")
    for sep in [";", "/", "／", "·", "ㆍ", "，"]:
        s = s.replace(sep, ",")
    parts = [p for p in s.split(",") if p != ""]
    names = []
    for p in parts:
        n = normalize_name(p)
        if n:
            names.append(n)
    seen = set()
    out = []
    for n in names:
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out


def parse_sn(sn: str):
    """시료번호 파싱"""
    try:
        sn = sn.strip()
        yy = int(sn[1:3])
        mm = int(sn[3:5])
        dd = int(sn[5:7])
        team = int(sn[7])
        idx = int(sn.split("-")[1])
        date_obj = date(2000 + yy, mm, dd)
        return date_obj, team, idx
    except:
        return None, None, None


def sample_to_datestr(sample_no: str) -> str:
    """시료번호에서 날짜 추출"""
    try:
        yy = int(sample_no[1:3])
        mm = int(sample_no[3:5])
        dd = int(sample_no[5:7])
        yyyy = 2000 + yy
        return f"{yyyy:04d}-{mm:02d}-{dd:02d}"
    except:
        return None


def parse_ymd_date(value):
    """YYYY-MM-DD 형식 값을 date 로 파싱"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    s = str(value).strip()
    if not s:
        return None

    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        return None


def parse_sn_date(sn: str):
    """시료번호에서 날짜만 추출"""
    m = re.match(r"^[A-Za-z](\d{2})(\d{2})(\d{2})", sn)
    if not m:
        return None
    yy, mm, dd = m.groups()
    year = 2000 + int(yy)
    month = int(mm)
    day = int(dd)
    try:
        return date(year, month, day)
    except ValueError:
        return None


def parse_sn_team(sn: str):
    """시료번호에서 팀 번호 추출"""
    if not sn:
        return None
    s = str(sn).strip()
    m = re.match(r"^[A-Za-z]\d{6}(\d)", s)
    if not m:
        return None
    try:
        t = int(m.group(1))
        return t if 1 <= t <= 9 else None
    except:
        return None


def extract_sn_text(x) -> str:
    """셀값에서 시료번호 추출"""
    if x is None:
        return ""
    s = str(x).strip()
    if not s:
        return ""
    m = re.search(r"[A-Za-z]\d{7}-\d{2,3}", s)
    return m.group(0).upper() if m else ""


def extract_sample_from_name(path_or_text: str) -> str:
    """파일명/폴더명/텍스트에서 시료번호 추출"""
    s = str(path_or_text or "").strip().upper()
    if not s:
        return ""

    s = s.replace("–", "-").replace("—", "-").replace("−", "-")

    base = s
    if "\\" in s or "/" in s:
        base = s.replace("/", "\\").rsplit("\\", 1)[-1]
    stem = re.sub(r"\.[^.]+$", "", base)
    compact = re.sub(r"\s+", "", stem)

    m = re.search(r"(?<![A-Z0-9])(A\d{7}-\d{2})(?!\d)", compact)
    if m:
        return m.group(1)

    return ""


def parse_time(s):
    """시간 파싱"""
    if s is None:
        return None
    if isinstance(s, datetime):
        return s.time()
    if isinstance(s, time):
        return s
    if isinstance(s, (int, float)) and 0 <= float(s) < 1.5:
        total_seconds = int(round(float(s) * 24 * 3600))
        h = (total_seconds // 3600) % 24
        m = (total_seconds % 3600) // 60
        return time(h, m)
    s = str(s).strip()
    if s == "" or s.lower() == "nan":
        return None
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?$", s)
    if not m:
        return None
    h = int(m.group(1))
    mi = int(m.group(2))
    if h > 23 or mi > 59:
        return None
    return time(h, mi)


def parse_time_range(text: str):
    """시간 범위 파싱"""
    if not isinstance(text, str) or "~" not in text:
        return None, None
    left, right = text.split("~")
    return parse_time(left.strip()), parse_time(right.strip())


def excel_value_to_time(value):
    """엑셀 셀값 → 시간"""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return time(hour=value.hour, minute=value.minute, second=0)
    if isinstance(value, time):
        return time(hour=value.hour, minute=value.minute, second=0)
    if isinstance(value, (int, float)):
        seconds = int(round(float(value) * 24 * 3600))
        seconds %= 24 * 3600
        h = seconds // 3600
        m = (seconds % 3600) // 60
        return time(h, m, 0)
    s = str(value).strip()
    m = re.search(r"(\d{1,2}):(\d{2})", s)
    if not m:
        return None
    h = int(m.group(1))
    mi = int(m.group(2))
    if h >= 24 or mi >= 60:
        return None
    return time(h, mi, 0)


def parse_time_range_text(s):
    """시간 범위 텍스트 파싱"""
    if not s:
        return None, None
    if "~" not in s:
        return None, None
    left, right = s.split("~")
    return excel_value_to_time(left.strip()), excel_value_to_time(right.strip())


def norm_ymd(v):
    """날짜 정규화"""
    if v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    s = s.replace(".", "-").replace("/", "-")
    if " " in s:
        s = s.split(" ")[0].strip()
    if len(s) >= 10:
        s = s[:10]
    return s


def clean_leading_mark(s):
    """선행 ×, ✕, 쉼표 제거 — eco_input, eco_check 공통"""
    if not s:
        return ""
    s = str(s).strip()
    if s.startswith(("×", "✕", ",")):
        return s[1:].strip()
    return s