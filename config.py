# -*- coding: utf-8 -*-
import os
import configparser

# config.py가 위치한 폴더 기준으로 config.ini 경로 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.ini")

config = configparser.ConfigParser(interpolation=configparser.ExtendedInterpolation())

# config.ini가 없을 경우 기본값 설정을 위한 딕셔너리 (fallback)
DEFAULT_CONFIG = {
    'NAS': {
        'IP': '192.168.10.163',
        'TEAM_ROOT': r'\\192.168.10.163\측정팀',
        'USER_ROOT': r'\\192.168.10.163\개인업무'
    },
    'PATHS': {
        'REPORT_BASE': r'\\192.168.10.163\측정팀\2.성적서',
        'REPORT_SRC': r'\\192.168.10.163\측정팀\2.성적서\0 2.검토중',
        'REPORT_DONE': r'\\192.168.10.163\측정팀\2.성적서\0 5.최종완료',
        'DAEJANG_ROOT': r'\\192.168.10.163\측정팀\0.시료접수발송대장',
        'LOG_DIR': r'\\192.168.10.163\측정팀\10.검토\_logs',
        # ... 필요한 경우 추가
    }
}

if os.path.exists(CONFIG_PATH):
    config.read(CONFIG_PATH, encoding='utf-8')
else:
    print(f"⚠ {CONFIG_PATH} 파일을 찾을 수 없습니다. 기본값을 사용합니다.")
    config.read_dict(DEFAULT_CONFIG)

# 외부에서 쓰기 편하게 변수로 정의
NAS_IP = config.get("NAS", "IP")
TEAM_ROOT = config.get("NAS", "TEAM_ROOT")
USER_ROOT = config.get("NAS", "USER_ROOT")

REPORT_BASE = config.get("PATHS", "REPORT_BASE")
REPORT_SRC = config.get("PATHS", "REPORT_SRC")
REPORT_DONE = config.get("PATHS", "REPORT_DONE")
WATER_REPORT_BASE = config.get("PATHS", "WATER_REPORT_BASE")

DAEJANG_ROOT = config.get("PATHS", "DAEJANG_ROOT")

RECEIPT_REVIEW = config.get("PATHS", "RECEIPT_REVIEW")
MEASIN_REVIEW = config.get("PATHS", "MEASIN_REVIEW")
DRIVE_LOG_REVIEW = config.get("PATHS", "DRIVE_LOG_REVIEW")
TOTAL_REVIEW = config.get("PATHS", "TOTAL_REVIEW")
LOG_DIR = config.get("PATHS", "LOG_DIR")

PDF_AIR = config.get("PATHS", "PDF_AIR")
PDF_WATER = config.get("PATHS", "PDF_WATER")
MOISTURE_ROOT = config.get("PATHS", "MOISTURE_ROOT")
THC_ROOT = config.get("PATHS", "THC_ROOT")
TAB4_MACRO_FILE = config.get("PATHS", "TAB4_MACRO_FILE")
