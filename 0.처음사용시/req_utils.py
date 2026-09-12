# -*- coding: utf-8 -*-
"""requirements.txt 공통 로더.

pip 패키지명 → import 모듈명 별칭만 여기/텍스트로 관리한다.
"""
from __future__ import annotations

import re
from pathlib import Path

# pip 이름과 import 이름이 다른 경우만 명시 (나머지: 하이픈→언더스코어)
IMPORT_ALIASES = {
    "pywin32": "win32com",
    "pillow": "PIL",
    "webdriver-manager": "webdriver_manager",
}


def requirements_path() -> Path:
    return Path(__file__).resolve().parent / "requirements.txt"


def load_pip_packages(path: str | Path | None = None) -> list[str]:
    """requirements.txt 에서 pip 패키지명 목록을 읽는다."""
    req = Path(path) if path else requirements_path()
    if not req.is_file():
        raise FileNotFoundError(f"requirements.txt 없음: {req}")

    packages: list[str] = []
    seen: set[str] = set()
    for raw in req.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        # "pkg>=1.0 ; python_version>='3'" 등에서 이름만
        name = re.split(r"[<>=!;\[ \t]", line, maxsplit=1)[0].strip()
        if not name or name in seen:
            continue
        seen.add(name)
        packages.append(name)
    return packages


def pip_to_import(pip_name: str) -> str:
    """pip 패키지명 → importlib 검사에 쓸 모듈명."""
    key = (pip_name or "").strip()
    if key in IMPORT_ALIASES:
        return IMPORT_ALIASES[key]
    return key.replace("-", "_")


def load_required_pairs(path: str | Path | None = None) -> list[tuple[str, str]]:
    """(pip명, import명) 목록."""
    return [(p, pip_to_import(p)) for p in load_pip_packages(path)]
