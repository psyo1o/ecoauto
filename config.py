# -*- coding: utf-8 -*-
import os
import configparser

# config.py가 위치한 폴더 기준으로 config.ini 경로 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.ini")

# BasicInterpolation을 사용하여 %(VAR)s 형태의 치환 지원
config = configparser.ConfigParser()

# 모든 섹션의 키를 DEFAULT 섹션에 넣어 interpolation이 가능하게 함
DEFAULT_CONFIG = {
    'DEFAULT': {
        'IP': '192.168.10.163',
        'TEAM_ROOT': r'\\%(IP)s\측정팀',
        'USER_ROOT': r'\\%(IP)s\개인업무',
        'REPORT_BASE': r'%(TEAM_ROOT)s\2.성적서',
        'REVIEW_ROOT': r'%(TEAM_ROOT)s\10.검토',
        'REPORT_SRC': r'%(REPORT_BASE)s\0 2.검토중',
        'REPORT_DONE': r'%(REPORT_BASE)s\0 5.최종완료',
        'WATER_REPORT_BASE': r'%(REPORT_BASE)s\14.수질성적서',
        'DAEJANG_ROOT': r'%(TEAM_ROOT)s\0.시료접수발송대장',
        'RECEIPT_REVIEW': r'%(REVIEW_ROOT)s\2.발송대장 검토',
        'MEASIN_REVIEW': r'%(REVIEW_ROOT)s\3.측정인 검토',
        'DRIVE_LOG_REVIEW': r'%(REVIEW_ROOT)s\4.차량운행일지 검토',
        'TOTAL_REVIEW': r'%(REVIEW_ROOT)s\5.종합 검토',
        'LOG_DIR': r'%(REVIEW_ROOT)s\_logs',
        'PDF_AIR': r'%(REPORT_BASE)s\0.PDF\2.대기pdf',
        'PDF_WATER': r'%(REPORT_BASE)s\0.PDF\1.수질pdf',
        'MOISTURE_ROOT': r'%(REPORT_BASE)s\0.수분량',
        'THC_ROOT': r'%(REPORT_BASE)s\0.THC',
        'TAB4_MACRO_FILE': r'%(REPORT_BASE)s\측정인 측정분석 입력 26.01.xlsm',
    }
}

# 기본값을 먼저 로드하고, 파일이 있으면 덮어씌움
config.read_dict(DEFAULT_CONFIG)

if os.path.exists(CONFIG_PATH):
    config.read(CONFIG_PATH, encoding='utf-8')
else:
    print(f"⚠ {CONFIG_PATH} 파일을 찾을 수 없습니다. 기본값을 사용합니다.")

# 외부에서 쓰기 편하게 변수로 정의
# configparser는 DEFAULT 섹션의 값을 다른 섹션에서도 찾을 수 있게 해줌
NAS_IP = config.get("DEFAULT", "IP")
TEAM_ROOT = config.get("DEFAULT", "TEAM_ROOT")
USER_ROOT = config.get("DEFAULT", "USER_ROOT")

REPORT_BASE = config.get("DEFAULT", "REPORT_BASE")
REPORT_SRC = config.get("DEFAULT", "REPORT_SRC")
REPORT_DONE = config.get("DEFAULT", "REPORT_DONE")
WATER_REPORT_BASE = config.get("DEFAULT", "WATER_REPORT_BASE")

DAEJANG_ROOT = config.get("DEFAULT", "DAEJANG_ROOT")

RECEIPT_REVIEW = config.get("DEFAULT", "RECEIPT_REVIEW")
MEASIN_REVIEW = config.get("DEFAULT", "MEASIN_REVIEW")
DRIVE_LOG_REVIEW = config.get("DEFAULT", "DRIVE_LOG_REVIEW")
TOTAL_REVIEW = config.get("DEFAULT", "TOTAL_REVIEW")
LOG_DIR = config.get("DEFAULT", "LOG_DIR")

PDF_AIR = config.get("DEFAULT", "PDF_AIR")
PDF_WATER = config.get("DEFAULT", "PDF_WATER")
MOISTURE_ROOT = config.get("DEFAULT", "MOISTURE_ROOT")
THC_ROOT = config.get("DEFAULT", "THC_ROOT")
TAB4_MACRO_FILE = config.get("DEFAULT", "TAB4_MACRO_FILE")
