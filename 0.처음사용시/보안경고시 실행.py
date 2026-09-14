# -*- coding: utf-8 -*-
"""Excel/NAS 보안경고 완화 (레지스트리 직접 적용).

기존 `보안경고시 실행.REG` 와 동일 키를 winreg 로 기록한다.
통합설치에서 호출 시: python 보안경고시 실행.py --no-pause
"""
from __future__ import annotations

import sys
import winreg

NAS_IP = "192.168.10.163"


def want_pause() -> bool:
    return "--no-pause" not in sys.argv


def pause() -> None:
    if want_pause():
        input("계속하려면 Enter 키를 누르십시오...")


def apply_security_settings() -> None:
    """구 REG 파일과 동일한 HKCU 항목 적용."""
    # Excel Trusted Location — NAS 루트
    nas_key = (
        r"Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\NAS"
    )
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, nas_key) as key:
        winreg.SetValueEx(key, "Path", 0, winreg.REG_SZ, f"\\\\{NAS_IP}\\")
        winreg.SetValueEx(key, "AllowSubfolders", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(
            key, "Description", 0, winreg.REG_SZ, "Trusted NAS Root Location"
        )

    # NAS → 로컬 인트라넷 영역
    zone_key = (
        rf"Software\Microsoft\Windows\CurrentVersion\Internet Settings"
        rf"\ZoneMap\Domains\{NAS_IP}"
    )
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, zone_key) as key:
        winreg.SetValueEx(key, "*", 0, winreg.REG_DWORD, 1)

    # Excel '인터넷에서 가져온 콘텐츠 차단' 해제
    excel_sec = r"Software\Microsoft\Office\16.0\Excel\Security"
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, excel_sec) as key:
        winreg.SetValueEx(
            key, "BlockContentExecutionFromInternet", 0, winreg.REG_DWORD, 0
        )


def main() -> int:
    print("==============================================")
    print("    보안경고 설정 적용")
    print("==============================================")
    print()
    print(f"[INFO] NAS IP: {NAS_IP}")
    print("  · Excel Trusted Location (NAS 루트)")
    print("  · 로컬 인트라넷 ZoneMap")
    print("  · BlockContentExecutionFromInternet = 0")
    print()

    try:
        apply_security_settings()
    except OSError as exc:
        print(f"[ERROR] 레지스트리 적용 실패: {exc}")
        pause()
        return 1

    print("[OK] 보안경고 설정 적용 완료")
    print()
    pause()
    return 0


if __name__ == "__main__":
    sys.exit(main())
