# -*- coding: utf-8 -*-
"""
공통 로그 유틸 — 실행·에러 모두 프로그램루트\\4.log 에 통합.

경로: 4.log\\{YYYY}년\\{M}월\\
  - 실행 중: {YYYYMMDD_HHMMSS}_{프로그램}.log  (stdout/stderr + 예외)
  - 실행 세션 없을 때 예외만: {YYYYMMDD}_error.log
"""

from __future__ import annotations

import os
import sys
import traceback
from datetime import datetime

from config import PROJECT_ROOT

# 프로그램 루트\4.log\{연도}년\{월}월\...
APP_RUN_LOG_ROOT = os.path.join(PROJECT_ROOT, "4.log")
# 하위 호환: 예전 LOG_DIR / ERROR_LOG_FILE 참조 코드용 (실제 기록은 4.log)
BASE_LOG_DIR = APP_RUN_LOG_ROOT
ERROR_LOG_FILE = os.path.join(APP_RUN_LOG_ROOT, "error_log.txt")  # 미사용, 별도 파일 안 씀

_active_run: "RunLogSession | None" = None


def _month_folder(now: datetime | None = None) -> str:
    now = now or datetime.now()
    folder = os.path.join(APP_RUN_LOG_ROOT, f"{now.year}년", f"{now.month}월")
    os.makedirs(folder, exist_ok=True)
    return folder


def _fallback_error_path() -> str:
    """실행 세션이 없을 때 날짜별 에러 로그 (4.log 아래)."""
    now = datetime.now()
    return os.path.join(_month_folder(now), f"{now.strftime('%Y%m%d')}_error.log")


def _append_to_file(path: str, text: str) -> None:
    try:
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        with open(path, "a", encoding="utf-8") as f:
            f.write(text)
    except Exception:
        pass


def _write_log_line(text: str) -> None:
    """활성 실행 로그가 있으면 그파일, 없으면 4.log 날짜별 _error.log."""
    if _active_run is not None:
        try:
            _active_run.write(text if text.endswith("\n") else text + "\n")
            _active_run.flush()
            return
        except Exception:
            pass
    _append_to_file(_fallback_error_path(), text if text.endswith("\n") else text + "\n")


def log_error(context: str, exc: BaseException) -> None:
    """예외를 4.log 실행 로그(또는 날짜별 _error.log)에 기록."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    block = f"[{ts}] [ERROR] {context}\n{tb}\n{'=' * 80}\n"
    _write_log_line(block)


def log_message(message: str) -> None:
    """단순 메시지를 4.log 실행 로그(또는 날짜별 _error.log)에 기록."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _write_log_line(f"[{ts}] {message}\n")


class _TeeStream:
    """stdout/stderr → 콘솔 + 실행 로그 파일."""

    def __init__(self, original, session: "RunLogSession"):
        self._original = original
        self._session = session

    def write(self, text):
        try:
            self._original.write(text)
        except Exception:
            pass
        try:
            self._session.write(text)
        except Exception:
            pass
        return len(text) if text else 0

    def flush(self):
        try:
            self._original.flush()
        except Exception:
            pass
        try:
            self._session.flush()
        except Exception:
            pass

    def __getattr__(self, name):
        return getattr(self._original, name)


class RunLogSession:
    """한 번 실행분의 콘솔·에러 통합 로그."""

    def __init__(self, program: str):
        self.program = (program or "app").strip() or "app"
        self.path = ""
        self._fp = None
        self._old_out = None
        self._old_err = None
        self._started = False

    def start(self) -> "RunLogSession":
        global _active_run
        now = datetime.now()
        folder = _month_folder(now)
        safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in self.program)
        self.path = os.path.join(
            folder,
            f"{now.strftime('%Y%m%d_%H%M%S')}_{safe}.log",
        )
        self._fp = open(self.path, "a", encoding="utf-8", buffering=1)
        self._fp.write(
            f"===== {self.program} 시작 {now.strftime('%Y-%m-%d %H:%M:%S')} =====\n"
        )
        self._fp.flush()

        self._old_out = sys.stdout
        self._old_err = sys.stderr
        sys.stdout = _TeeStream(self._old_out, self)
        sys.stderr = _TeeStream(self._old_err, self)
        self._started = True
        _active_run = self
        return self

    def write(self, text: str) -> None:
        if not self._fp or not text:
            return
        try:
            self._fp.write(text if isinstance(text, str) else str(text))
        except Exception:
            pass

    def flush(self) -> None:
        if self._fp:
            try:
                self._fp.flush()
            except Exception:
                pass

    def stop(self) -> None:
        global _active_run
        if not self._started:
            return
        try:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.write(f"\n===== {self.program} 종료 {now} =====\n")
            self.flush()
        except Exception:
            pass
        if self._old_out is not None:
            sys.stdout = self._old_out
        if self._old_err is not None:
            sys.stderr = self._old_err
        try:
            if self._fp:
                self._fp.close()
        except Exception:
            pass
        self._fp = None
        self._started = False
        if _active_run is self:
            _active_run = None

    def __enter__(self) -> "RunLogSession":
        return self.start()

    def __exit__(self, exc_type, exc, tb) -> None:
        if exc is not None:
            try:
                ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.write(
                    f"[{ts}] [EXCEPTION] {exc_type.__name__ if exc_type else ''}: {exc}\n"
                )
                if tb is not None and exc is not None:
                    self.write(
                        "".join(traceback.format_exception(exc_type, exc, tb))
                    )
            except Exception:
                pass
        self.stop()


def start_run_log(program: str) -> RunLogSession:
    """실행 로그 시작 (stdout/stderr tee). stop_run_log 또는 session.stop() 호출."""
    return RunLogSession(program).start()


def stop_run_log(session: RunLogSession | None) -> None:
    if session is not None:
        session.stop()


def run_log(program: str) -> RunLogSession:
    """with run_log('eco_input'): ... 형태로 사용."""
    return RunLogSession(program)
