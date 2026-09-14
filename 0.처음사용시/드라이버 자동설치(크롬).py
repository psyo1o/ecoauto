# -*- coding: utf-8 -*-
"""ChromeDriver 자동 다운로드 & 설치.

Smart App Control이 .bat를 차단하는 환경을 피하기 위해
서명된 python.exe로 실행되는 .py로 제공한다. 동작은 기존 .bat와 동일.

통합설치에서 호출 시: python 드라이버 자동설치(크롬).py --no-pause
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import urllib.request
import winreg
import zipfile
from pathlib import Path

DRIVER_DIR = Path(r"C:\chromedriver")
CFT_JSON_URL = (
    "https://googlechromelabs.github.io/chrome-for-testing/"
    "last-known-good-versions-with-downloads.json"
)


def want_pause() -> bool:
    return "--no-pause" not in sys.argv


def pause() -> None:
    if want_pause():
        input("계속하려면 Enter 키를 누르십시오...")


def get_chrome_version() -> str | None:
    for chrome_path in (
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ):
        if not Path(chrome_path).is_file():
            continue
        try:
            out = subprocess.check_output(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    f"(Get-Item '{chrome_path}').VersionInfo.ProductVersion",
                ],
                text=True,
                stderr=subprocess.DEVNULL,
            ).strip()
            if out:
                return out
        except Exception:
            continue
    return None


def fetch_chromedriver_url() -> str | None:
    with urllib.request.urlopen(CFT_JSON_URL, timeout=60) as resp:
        data = json.load(resp)
    downloads = (
        data.get("channels", {})
        .get("Stable", {})
        .get("downloads", {})
        .get("chromedriver", [])
    )
    for item in downloads:
        if item.get("platform") == "win64":
            return item.get("url")
    return None


def add_to_user_path(directory: str) -> bool:
    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        "Environment",
        0,
        winreg.KEY_READ | winreg.KEY_SET_VALUE,
    ) as key:
        try:
            current, reg_type = winreg.QueryValueEx(key, "Path")
        except FileNotFoundError:
            current, reg_type = "", winreg.REG_EXPAND_SZ
        parts = [p for p in str(current).split(";") if p]
        norm = os.path.normcase(os.path.normpath(directory))
        if any(os.path.normcase(os.path.normpath(p)) == norm for p in parts):
            return False
        new_value = (
            (str(current).rstrip(";") + ";" + directory) if current else directory
        )
        winreg.SetValueEx(key, "Path", 0, reg_type, new_value)
    return True


def main() -> int:
    print("==============================================")
    print("    ChromeDriver 자동 다운로드 & 설치")
    print("==============================================")
    print()

    chrome_version = get_chrome_version()
    if not chrome_version:
        print("[ERROR] Chrome 설치 경로를 찾지 못했습니다.")
        print("Chrome 설치 후 다시 실행해주세요.")
        pause()
        return 1

    print(f"[OK] Chrome 버전: {chrome_version}")
    print(f"[OK] 메이저 버전: {chrome_version.split('.', 1)[0]}")
    print()

    print("ChromeDriver 최신 버전 정보 조회 중...")
    try:
        dl_url = fetch_chromedriver_url()
    except Exception as exc:
        print(f"[ERROR] ChromeDriver 다운로드 URL을 가져오지 못했습니다. ({exc})")
        pause()
        return 1

    if not dl_url:
        print("[ERROR] ChromeDriver 다운로드 URL을 가져오지 못했습니다.")
        pause()
        return 1

    print("[OK] ChromeDriver 다운로드 URL:")
    print(dl_url)
    print()

    DRIVER_DIR.mkdir(parents=True, exist_ok=True)
    zip_path = DRIVER_DIR / "driver.zip"

    print("----------------------------------------------")
    print("ChromeDriver 다운로드 중...")
    try:
        urllib.request.urlretrieve(dl_url, zip_path)
    except Exception as exc:
        print(f"[ERROR] 다운로드 실패! ({exc})")
        pause()
        return 1

    if not zip_path.is_file():
        print("[ERROR] 다운로드 실패!")
        pause()
        return 1

    print("[OK] 다운로드 완료")
    print("압축 해제 중...")
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(DRIVER_DIR)
    zip_path.unlink(missing_ok=True)

    chromedriver = next(DRIVER_DIR.rglob("chromedriver.exe"), None)
    if chromedriver is None or not chromedriver.is_file():
        print("[ERROR] chromedriver.exe 설치 실패")
        pause()
        return 1

    print(f"[OK] 설치 완료: {chromedriver}")
    print("PATH 등록 중...")
    if add_to_user_path(str(DRIVER_DIR)):
        print("[OK] PATH 등록 완료")
    else:
        print("[OK] 이미 PATH에 등록됨")

    print("----------------------------------------------")
    print("    ChromeDriver 자동 설치 완료!")
    print("  ▶ PC 재부팅(또는 재로그인) 후 바로 사용 가능합니다.")
    print("----------------------------------------------")
    pause()
    return 0


if __name__ == "__main__":
    sys.exit(main())
