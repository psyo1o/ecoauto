# -*- coding: utf-8 -*-
"""
config.ini + 기본값 로드.
- [DEFAULT]: IP, TEAM_ROOT 등 %(변수) 치환용 기본값
- [PATHS], [URLS]: 경로·URL (ini에서 수정 시 PATHS/URLS 우선)
"""
import os
import configparser

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.ini")

config = configparser.ConfigParser()

# interpolation: %(TEAM_ROOT)s 등
_DEFAULT = {
    "IP": "192.168.10.163",
    "TEAM_ROOT": r"\\%(IP)s\측정팀",
    "USER_ROOT": r"\\%(IP)s\개인업무",
    "REPORT_BASE": r"%(TEAM_ROOT)s\2.성적서",
    "REVIEW_ROOT": r"%(TEAM_ROOT)s\10.검토",
}

_PATHS = {
    "REPORT_WORKFLOW_DIRS": (
        "0 0.입력중, 0 1.완료, 0 2.검토중, 0 3.검토완료, "
        "0 4.출력완료&에코랩입력중, 0 5.최종완료"
    ),
    "REPORT_SRC": r"%(REPORT_BASE)s\0 2.검토중",
    "REPORT_DONE": r"%(REPORT_BASE)s\0 5.최종완료",
    "WATER_REPORT_BASE": r"%(REPORT_BASE)s\14.수질성적서",
    "DAEJANG_ROOT": r"%(TEAM_ROOT)s\0.시료접수발송대장",
    "RECEIPT_REVIEW": r"%(REVIEW_ROOT)s\2.발송대장 검토",
    "MEASIN_REVIEW": r"%(REVIEW_ROOT)s\3.측정인 검토",
    "MEASIN_PDF_DIR": r"%(MEASIN_REVIEW)s\PDF",
    "DRIVE_LOG_REVIEW": r"%(REVIEW_ROOT)s\4.차량운행일지 검토",
    "TOTAL_REVIEW": r"%(REVIEW_ROOT)s\5.종합 검토",
    "LOG_DIR": r"%(REVIEW_ROOT)s\_logs",
    "PDF_AIR": r"%(REPORT_BASE)s\0.PDF\2.대기pdf",
    "PDF_WATER": r"%(REPORT_BASE)s\0.PDF\1.수질pdf",
    "MOISTURE_ROOT": r"%(REPORT_BASE)s\0.수분량",
    "THC_ROOT": r"%(REPORT_BASE)s\0.THC",
    "TAB4_MACRO_FILE": r"%(REPORT_BASE)s\측정인 측정분석 입력 26.01.xlsm",
    "MOISTURE_SAMPLE": r"%(MOISTURE_ROOT)s\수분량샘플.csv",
    "THC_CSV_SAMPLE": r"%(THC_ROOT)s\FID샘플.csv",
    "THC_FID_SAMPLE": r"%(THC_ROOT)s\PF샘플.FID",
    "PDF_TMP_DIR": r"C:\measin_upload_tmp",
}

_URLS = {
    "LOGIN_URL": "https://측정인.kr/init.go",
    "FIELD_URL_AIR": "https://측정인.kr/ms/field_outair.do",
    "FIELD_URL_WATER": "https://측정인.kr/ms/field_outwater.do",
}

# DEFAULT에 경로·URL도 넣어 %(REPORT_BASE)s 등 치환이 동작하도록 함
_DEFAULT.update(_PATHS)
_DEFAULT.update(_URLS)

config.read_dict({"DEFAULT": _DEFAULT, "PATHS": _PATHS, "URLS": _URLS})

if os.path.exists(CONFIG_PATH):
    config.read(CONFIG_PATH, encoding="utf-8")
else:
    print(f"⚠ {CONFIG_PATH} 파일을 찾을 수 없습니다. 기본값을 사용합니다.")


def cfg(key: str, section: str | None = None) -> str:
    """설정값 조회. section 미지정 시 PATHS → URLS → DEFAULT 순."""
    if section:
        return config.get(section, key)
    for sec in ("PATHS", "URLS", "DEFAULT"):
        if config.has_option(sec, key):
            return config.get(sec, key)
    raise KeyError(key)


def cfg_list(key: str) -> list[str]:
    """쉼표 구분 목록 (예: REPORT_WORKFLOW_DIRS)."""
    return [part.strip() for part in cfg(key).split(",") if part.strip()]


# --- 외부 import용 상수 ---
NAS_IP = cfg("IP")
TEAM_ROOT = cfg("TEAM_ROOT")
USER_ROOT = cfg("USER_ROOT")
REVIEW_ROOT = cfg("REVIEW_ROOT")

REPORT_BASE = cfg("REPORT_BASE")
REPORT_WORKFLOW_DIRS = cfg_list("REPORT_WORKFLOW_DIRS")
REPORT_SRC = cfg("REPORT_SRC")
REPORT_DONE = cfg("REPORT_DONE")
WATER_REPORT_BASE = cfg("WATER_REPORT_BASE")

DAEJANG_ROOT = cfg("DAEJANG_ROOT")

RECEIPT_REVIEW = cfg("RECEIPT_REVIEW")
MEASIN_REVIEW = cfg("MEASIN_REVIEW")
MEASIN_PDF_DIR = cfg("MEASIN_PDF_DIR")
DRIVE_LOG_REVIEW = cfg("DRIVE_LOG_REVIEW")
TOTAL_REVIEW = cfg("TOTAL_REVIEW")
LOG_DIR = cfg("LOG_DIR")

PDF_AIR = cfg("PDF_AIR")
PDF_WATER = cfg("PDF_WATER")
MOISTURE_ROOT = cfg("MOISTURE_ROOT")
THC_ROOT = cfg("THC_ROOT")
TAB4_MACRO_FILE = cfg("TAB4_MACRO_FILE")

MOISTURE_SAMPLE = cfg("MOISTURE_SAMPLE")
THC_CSV_SAMPLE = cfg("THC_CSV_SAMPLE")
THC_FID_SAMPLE = cfg("THC_FID_SAMPLE")

PDF_TMP_DIR = cfg("PDF_TMP_DIR")

LOGIN_URL = cfg("LOGIN_URL")
FIELD_URL_AIR = cfg("FIELD_URL_AIR")
FIELD_URL_WATER = cfg("FIELD_URL_WATER")
# 하위 호환
FIELD_URL = FIELD_URL_AIR
