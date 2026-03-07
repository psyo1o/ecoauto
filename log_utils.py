# -*- coding: utf-8 -*-
"""
공통 로그 유틸
- 각 스크립트/GUI에서 예외 발생 시 파일로 남길 때 사용.
- 로그 저장 위치: \\192.168.10.163\측정팀\10.검토\_logs\error_log.txt
"""

import os
import traceback
from datetime import datetime


BASE_LOG_DIR = r"\\192.168.10.163\측정팀\10.검토\_logs"
ERROR_LOG_FILE = os.path.join(BASE_LOG_DIR, "error_log.txt")


def log_error(context: str, exc: BaseException) -> None:
    """
    예외 정보를 공용 에러 로그 파일에 append.
    - context: 어디서 발생했는지 간단 설명 (모듈/함수명 등)
    - exc    : 실제 Exception 객체
    """
    try:
        os.makedirs(BASE_LOG_DIR, exist_ok=True)
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        with open(ERROR_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {context}\n")
            f.write(tb)
            f.write("\n" + "=" * 80 + "\n")
    except Exception:
        # 로그 기록 중 에러가 나더라도 본 흐름을 막지 않는다.
        pass


def log_message(message: str) -> None:
    """
    단순 텍스트 메시지를 로그 파일에 남길 때 사용.
    (예: 중요 단계 시작/종료 등)
    """
    try:
        os.makedirs(BASE_LOG_DIR, exist_ok=True)
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(ERROR_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {message}\n")
    except Exception:
        pass

