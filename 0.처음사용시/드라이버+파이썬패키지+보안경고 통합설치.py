# -*- coding: utf-8 -*-
"""초기 설치 통합 도구 — 개별 설치 스크립트를 순서대로 실행만 한다.

실제 설치 로직은 아래 파일에만 둔다 (여기에는 복제하지 않음).
  1) 보안경고시 실행.py
  2) 드라이버 자동설치(크롬).py
  3) 파이썬 패키지.py         (+ requirements.txt)

각 단계는 `python 스크립트.py --no-pause` 로 호출하고,
전부 끝난 뒤(또는 실패 시) 한 번만 Enter 대기한다.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent

STEPS = (
    ("1/3 보안경고 설정", "보안경고시 실행.py"),
    ("2/3 ChromeDriver", "드라이버 자동설치(크롬).py"),
    ("3/3 Python 패키지", "파이썬 패키지.py"),
)


def pause() -> None:
    input("계속하려면 Enter 키를 누르십시오...")


def run_step(label: str, script_name: str) -> int:
    script = HERE / script_name
    print()
    print("=" * 50)
    print(f"  ▶ {label}: {script_name}")
    print("=" * 50)
    print()

    if not script.is_file():
        print(f"[ERROR] 파일 없음: {script}")
        return 1

    return subprocess.call(
        [sys.executable, str(script), "--no-pause"],
        cwd=str(HERE),
    )


def main() -> int:
    print("==============================================")
    print("    보안설정 + ChromeDriver + Python 패키지")
    print("    (개별 설치 스크립트 순차 실행)")
    print("==============================================")
    print()
    print(f"[INFO] 폴더: {HERE}")
    print(f"[INFO] Python: {sys.executable}")
    print(f"[INFO] 버전: {sys.version.split()[0]}")

    failed = None
    for label, name in STEPS:
        code = run_step(label, name)
        if code != 0:
            failed = (label, name, code)
            break

    print()
    print("==============================================")
    if failed:
        label, name, code = failed
        print(f"    실패: {label} ({name})  종료코드={code}")
        print("    해당 스크립트를 단독 실행해 원인을 확인하세요.")
    else:
        print("    모든 설치가 완료되었습니다.")
        print("    ChromeDriver PATH는 재실행/재로그인 후 반영될 수 있습니다.")
    print("==============================================")
    pause()
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
