# -*- coding: utf-8 -*-
"""
측정인 로그인 정보 저장 (아이디 저장 / 비밀번호 저장)

- 프로그램은 NAS 공유 폴더에서 여러 사람이 쓰므로 NAS가 아니라
  각 PC의 Windows 사용자 폴더(%APPDATA%\\ecoauto\\measin_login.json)에 저장
- 비밀번호는 Windows DPAPI(CryptProtectData)로 암호화 → 같은 Windows 계정에서만 복호화
"""

import base64
import json
import os

_DIR = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "ecoauto")
_PATH = os.path.join(_DIR, "measin_login.json")

try:
    import win32crypt  # pywin32
except Exception:
    win32crypt = None


def can_save_password() -> bool:
    return win32crypt is not None


def _encrypt(text: str) -> str:
    blob = win32crypt.CryptProtectData(text.encode("utf-8"), "ecoauto", None, None, None, 0)
    return base64.b64encode(blob).decode("ascii")


def _decrypt(data: str) -> str:
    _, raw = win32crypt.CryptUnprotectData(base64.b64decode(data), None, None, None, 0)
    return raw.decode("utf-8")


def load() -> dict:
    """{'id': str, 'pw': str, 'save_id': bool, 'save_pw': bool}"""
    out = {"id": "", "pw": "", "save_id": False, "save_pw": False}
    try:
        with open(_PATH, encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        return out
    if d.get("id"):
        out["id"] = str(d["id"])
        out["save_id"] = True
    if d.get("pw") and win32crypt is not None:
        try:
            out["pw"] = _decrypt(d["pw"])
            out["save_pw"] = True
        except Exception:
            pass
    return out


def save(login_id: str, login_pw: str, save_id: bool, save_pw: bool) -> None:
    """체크 상태대로 저장. 둘 다 해제면 파일 삭제."""
    save_pw = save_pw and save_id and win32crypt is not None
    if not save_id or not login_id:
        clear()
        return
    d = {"id": login_id}
    if save_pw and login_pw:
        d["pw"] = _encrypt(login_pw)
    os.makedirs(_DIR, exist_ok=True)
    with open(_PATH, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False)


def clear() -> None:
    try:
        os.remove(_PATH)
    except FileNotFoundError:
        pass
    except Exception:
        pass
