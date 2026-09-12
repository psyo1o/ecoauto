# -*- coding: utf-8 -*-
"""Python 패키지 선별 설치 도구.

Smart App Control이 .bat를 차단하는 환경을 피하기 위해
서명된 python.exe로 실행되는 .py로 제공한다. 동작은 기존 .bat와 동일.

패키지 목록: 같은 폴더의 requirements.txt (+ req_utils.py)

통합설치에서 호출 시: python 파이썬 패키지.py --no-pause
"""
from __future__ import annotations

import subprocess
import sys

from req_utils import load_pip_packages, requirements_path

USE_USER = False


def pause() -> None:
    if want_pause():
        input("계속하려면 Enter 키를 누르십시오...")


def want_pause() -> bool:
    return "--no-pause" not in sys.argv


def run_pip(args: list[str]) -> int:
    cmd = [sys.executable, "-m", "pip", *args]
    if USE_USER:
        cmd.append("--user")
    return subprocess.call(cmd)


def main() -> int:
    try:
        packages = load_pip_packages()
    except Exception as e:
        print(f"[ERROR] requirements.txt 읽기 실패: {e}")
        pause()
        return 1

    print(f"[INFO] 사용 Python: {sys.executable}")
    print(f"[INFO] 버전: {sys.version.split()[0]}")
    print(f"[INFO] 목록: {requirements_path()}")
    print()

    print("[1/3] pip 상태 확인 및 업그레이드...")
    if run_pip(["install", "--upgrade", "pip"]) != 0:
        print("[ERROR] pip 업그레이드 실패")
        pause()
        return 1
    print()

    print("[2/3] 설치가 필요한 패키지 확인 중...")
    to_install: list[str] = []
    for pkg in packages:
        show = subprocess.call(
            [sys.executable, "-m", "pip", "show", pkg],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        if show != 0:
            print(f"[MISSING] {pkg} - 설치 목록에 추가합니다.")
            to_install.append(pkg)
        else:
            print(f"[EXIST] {pkg} - 이미 설치되어 있습니다.")

    if not to_install:
        print()
        print("[INFO] 모든 패키지가 이미 설치되어 있습니다.")
    else:
        print()
        print(f"[INSTALL] 다음 패키지를 설치합니다: {' '.join(to_install)}")
        if run_pip(["install", *to_install]) != 0:
            print("[ERROR] 패키지 설치 실패")
            pause()
            return 1
    print()

    print("[3/3] 최종 설치 리스트 확인...")
    for pkg in packages:
        print(f"[PACKAGE] {pkg}")
        result = subprocess.run(
            [sys.executable, "-m", "pip", "show", pkg],
            capture_output=True,
            text=True,
        )
        for line in result.stdout.splitlines():
            if line.lower().startswith(("name:", "version:")):
                print(line)

    print()
    print("[DONE] 모든 작업이 완료되었습니다!")
    pause()
    return 0


if __name__ == "__main__":
    sys.exit(main())
