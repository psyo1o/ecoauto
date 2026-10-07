# -*- coding: utf-8 -*-
"""
통합 마스터 런처 (2.검토 및 입력프로그램.pyw)
NAS UNC 경로에서 더블 클릭 실행 · 콘솔 창 없음 · 하위 GUI 무창 실행

이 파일은 무조건 Python 3.11 로 실행된다.
(.pyw → pyw.exe 기본이 3.14여도 시작 직후 3.11로 전환)

필수 패키지: `0.처음사용시/requirements.txt`
"""

# ---------------------------------------------------------------------------
# Python 3.11 고정 — 다른 import 보다 먼저 (3.14 pyw 더블클릭 대응)
# ---------------------------------------------------------------------------
import os
import sys
import subprocess

PREFERRED_PY = (3, 11)
_REEXEC_ENV = "MEASIN_LAUNCHER_PY_REEXEC"
_CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)


def _script_path_early():
    try:
        return os.path.abspath(__file__)
    except NameError:
        pass
    if sys.argv and sys.argv[0]:
        return os.path.abspath(sys.argv[0])
    return ""


def _find_py311_exes():
    """(pythonw.exe, python.exe) 후보 경로 목록."""
    roots = []
    forced = (os.environ.get("MEASIN_PYTHON") or "").strip().strip('"')
    if forced:
        roots.append(os.path.dirname(forced) if os.path.isfile(forced) else forced)

    for base in (
        os.environ.get("PROGRAMFILES", r"C:\Program Files"),
        os.environ.get("LOCALAPPDATA", ""),
        r"C:\Program Files",
        r"C:\Python311",
    ):
        if not base:
            continue
        roots.append(os.path.join(base, "Python311"))
        roots.append(os.path.join(base, "Programs", "Python", "Python311"))
    roots.append(r"C:\Program Files\Python311")

    # py -3.11 → 실제 설치 경로
    try:
        r = subprocess.run(
            ["py", "-3.11", "-c", "import sys; print(sys.executable)"],
            capture_output=True,
            text=True,
            timeout=8,
            creationflags=_CREATE_NO_WINDOW,
        )
        if r.returncode == 0:
            exe = (r.stdout or "").strip().splitlines()[-1].strip()
            if exe and os.path.isfile(exe):
                roots.insert(0, os.path.dirname(exe))
    except Exception:
        pass

    seen = set()
    out_w, out_c = [], []
    for root in roots:
        if not root:
            continue
        key = os.path.normcase(os.path.normpath(root))
        if key in seen:
            continue
        seen.add(key)
        w = os.path.join(root, "pythonw.exe")
        c = os.path.join(root, "python.exe")
        if os.path.isfile(w):
            out_w.append(w)
        if os.path.isfile(c):
            out_c.append(c)
    return out_w, out_c


def _pin_windows_py_default_311():
    """
    Windows py/pyw 런처 기본 버전을 3.11로 (사용자 LOCALAPPDATA\\py.ini).
    더블클릭 시 처음부터 3.11로 뜨게 한다.
    """
    local = os.environ.get("LOCALAPPDATA") or ""
    if not local:
        return
    ini = os.path.join(local, "py.ini")
    try:
        existing = ""
        if os.path.isfile(ini):
            with open(ini, encoding="utf-8", errors="ignore") as f:
                existing = f.read()
        # 이미 3.11 기본이면 유지
        compact = existing.lower().replace(" ", "")
        if "python=3.11" in compact:
            return
        block = "[defaults]\npython=3.11\n"
        if "[defaults]" in existing.lower():
            # defaults 섹션만 3.11로 교체
            lines = existing.splitlines()
            out = []
            in_defaults = False
            replaced = False
            for line in lines:
                s = line.strip().lower()
                if s.startswith("[") and s.endswith("]"):
                    in_defaults = s == "[defaults]"
                    out.append(line)
                    continue
                if in_defaults and s.startswith("python="):
                    out.append("python=3.11")
                    replaced = True
                    continue
                out.append(line)
            if not replaced:
                # [defaults] 끝에 추가
                rebuilt = []
                in_defaults = False
                added = False
                for line in out:
                    s = line.strip().lower()
                    if s.startswith("[") and s.endswith("]"):
                        if in_defaults and not added:
                            rebuilt.append("python=3.11")
                            added = True
                        in_defaults = s == "[defaults]"
                    rebuilt.append(line)
                if in_defaults and not added:
                    rebuilt.append("python=3.11")
                text = "\n".join(rebuilt).rstrip() + "\n"
            else:
                text = "\n".join(out).rstrip() + "\n"
        else:
            text = (existing.rstrip() + "\n\n" if existing.strip() else "") + block
        with open(ini, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
    except Exception:
        pass


def _reexec_with_py311():
    """현재가 3.11이 아니면 이 스크립트를 3.11로 다시 실행하고 종료."""
    _pin_windows_py_default_311()
    if os.environ.get(_REEXEC_ENV) == "1":
        return
    if sys.version_info[:2] == PREFERRED_PY:
        return

    script = _script_path_early()
    if not script or not os.path.isfile(script):
        return

    env = os.environ.copy()
    env[_REEXEC_ENV] = "1"
    cwd = os.path.dirname(script) or None
    extra = list(sys.argv[1:])

    pyw_list, py_list = _find_py311_exes()
    # GUI 더블클릭(pythonw)이면 pythonw 우선
    is_w = os.path.basename(sys.executable).lower() in ("pythonw.exe", "pyw.exe")
    cmds = []
    if is_w:
        for w in pyw_list:
            cmds.append([w, script, *extra])
        # Windows 런처: pyw -3.11
        cmds.append(["pyw", "-3.11", script, *extra])
        for c in py_list:
            cmds.append([c, script, *extra])
    else:
        for c in py_list:
            cmds.append([c, script, *extra])
        cmds.append(["py", "-3.11", script, *extra])
        for w in pyw_list:
            cmds.append([w, script, *extra])

    for cmd in cmds:
        try:
            subprocess.Popen(
                cmd,
                cwd=cwd,
                env=env,
                close_fds=True,
                creationflags=_CREATE_NO_WINDOW,
            )
            os._exit(0)
        except OSError:
            continue


_reexec_with_py311()

# 여기부터는 Python 3.11 에서만 실행되는 것이 정상
import importlib.util
import threading
import tkinter as tk
from tkinter import messagebox, font as tkfont, scrolledtext

# Windows: 하위 프로세스 콘솔 창 억제
CREATE_NO_WINDOW = _CREATE_NO_WINDOW

# 설치 직후 재시작 루프 방지용 (한 번만)
_DEP_RESTART_ENV = "MEASIN_LAUNCHER_AFTER_DEP_INSTALL"

# ---------------------------------------------------------------------------
# 실행 기준 디렉터리 (UNC / 더블클릭 대응)
# ---------------------------------------------------------------------------

def _resolve_base_dir():
    """런처 스크립트가 위치한 NAS 폴더 경로를 동적으로 확보한다."""
    candidates = []
    if getattr(sys, "frozen", False):
        candidates.append(os.path.dirname(os.path.abspath(sys.executable)))
    try:
        candidates.append(os.path.dirname(os.path.abspath(__file__)))
    except NameError:
        pass
    if sys.argv and sys.argv[0]:
        argv0 = sys.argv[0]
        if not os.path.isabs(argv0):
            argv0 = os.path.join(os.getcwd(), argv0)
        candidates.append(os.path.dirname(os.path.abspath(argv0)))
    candidates.append(os.getcwd())
    seen = set()
    for path in candidates:
        if not path:
            continue
        norm = os.path.normpath(path)
        if norm in seen:
            continue
        seen.add(norm)
        if os.path.isdir(norm):
            return norm
    return os.getcwd()


BASE_DIR = _resolve_base_dir()
PY_DIR = os.path.join(BASE_DIR, "3.py")
SETUP_DIR = os.path.join(BASE_DIR, "0.처음사용시")
REQUIREMENTS_TXT = os.path.join(SETUP_DIR, "requirements.txt")


def _load_req_utils():
    """0.처음사용시/req_utils.py 로드."""
    if SETUP_DIR not in sys.path:
        sys.path.insert(0, SETUP_DIR)
    import req_utils  # noqa: WPS433 — 런타임 경로 추가 후 import

    return req_utils


def _required_package_pairs():
    """(pip명, import명) — requirements.txt 기준."""
    req_utils = _load_req_utils()
    return req_utils.load_required_pairs(REQUIREMENTS_TXT)


def _launcher_script_path():
    """재시작에 쓸 이 런처 파일 경로."""
    try:
        return os.path.abspath(__file__)
    except NameError:
        if sys.argv and sys.argv[0]:
            return os.path.abspath(sys.argv[0])
    return os.path.join(BASE_DIR, "2.검토 및 입력프로그램.pyw")


def _resolve_preferred_exe(*, want_w: bool) -> str:
    """선호 3.11 pythonw/python. 없으면 빈 문자열."""
    pyw_list, py_list = _find_py311_exes()
    if want_w:
        return (pyw_list or py_list or [""])[0]
    return (py_list or pyw_list or [""])[0]


def _get_python_executable():
    """pip 설치용 — 3.11 python.exe 고정."""
    preferred = _resolve_preferred_exe(want_w=False)
    if preferred and preferred.lower().endswith("python.exe"):
        return preferred
    if preferred:
        sibling = os.path.join(os.path.dirname(preferred), "python.exe")
        if os.path.isfile(sibling):
            return sibling
    exe = sys.executable
    name = os.path.basename(exe).lower()
    if name == "pythonw.exe":
        candidate = os.path.join(os.path.dirname(exe), "python.exe")
        if os.path.isfile(candidate):
            return candidate
    return exe


def _missing_packages():
    """미설치 패키지 pip 이름 목록."""
    missing = []
    for pip_name, mod_name in _required_package_pairs():
        try:
            if importlib.util.find_spec(mod_name) is None:
                missing.append(pip_name)
        except (ImportError, ModuleNotFoundError, ValueError):
            missing.append(pip_name)
    return missing


def _center_toplevel(win, width, height):
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = max(0, (sw - width) // 2)
    y = max(0, (sh - height) // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")


def _restart_launcher():
    """설치 후 런처를 다시 띄우고 현재 프로세스는 종료."""
    script = _launcher_script_path()
    env = os.environ.copy()
    env[_DEP_RESTART_ENV] = "1"
    # GUI 재시작은 선호 3.11 pythonw 유지
    exe = _get_pythonw_executable()
    try:
        subprocess.Popen(
            [exe, script],
            cwd=BASE_DIR,
            env=env,
            creationflags=CREATE_NO_WINDOW,
        )
    except OSError:
        # fallback: python.exe
        subprocess.Popen(
            [_get_python_executable(), script],
            cwd=BASE_DIR,
            env=env,
        )
    sys.exit(0)


def _install_packages_with_ui(missing):
    """
    누락 패키지 안내 창 → pip 설치 → 성공 시 재시작.
    실패하거나 재시작 직후에도 남으면 오류 안내 후 False.
    """
    after_install = os.environ.pop(_DEP_RESTART_ENV, "") == "1"
    if after_install:
        # 방금 설치·재시작했는데도 남음 → 무한루프 방지
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "패키지 설치 실패",
            "필수 패키지 설치 후에도 일부 모듈을 불러오지 못했습니다.\n\n"
            f"미설치: {', '.join(missing)}\n\n"
            "수동 설치:\n"
            f"  {_get_python_executable()} -m pip install {' '.join(missing)}\n\n"
            "또는 `0.처음사용시\\파이썬 패키지.py` 를 실행해 주세요.",
        )
        root.destroy()
        return False

    root = tk.Tk()
    root.title("필수 패키지 설치")
    root.configure(bg="#2E3440")
    root.resizable(False, False)
    _center_toplevel(root, 520, 360)

    outer = tk.Frame(root, bg="#2E3440", padx=20, pady=16)
    outer.pack(fill=tk.BOTH, expand=True)

    tk.Label(
        outer,
        text="필수 Python 패키지가 없습니다",
        font=("맑은 고딕", 13, "bold"),
        fg="#ECEFF4",
        bg="#2E3440",
        anchor="w",
    ).pack(fill=tk.X)

    tk.Label(
        outer,
        text="아래 패키지를 자동 설치한 뒤 프로그램을 다시 시작합니다.",
        font=("맑은 고딕", 9),
        fg="#A3BE8C",
        bg="#2E3440",
        anchor="w",
    ).pack(fill=tk.X, pady=(4, 10))

    log = scrolledtext.ScrolledText(
        outer,
        height=12,
        font=("Consolas", 9),
        bg="#3B4252",
        fg="#ECEFF4",
        insertbackground="#ECEFF4",
        relief=tk.FLAT,
        wrap=tk.WORD,
    )
    log.pack(fill=tk.BOTH, expand=True)
    log.insert(tk.END, "설치 예정:\n")
    for name in missing:
        log.insert(tk.END, f"  · {name}\n")
    log.insert(tk.END, "\n")
    log.configure(state=tk.DISABLED)

    status = tk.Label(
        outer,
        text="설치를 시작합니다…",
        font=("맑은 고딕", 9),
        fg="#88C0D0",
        bg="#2E3440",
        anchor="w",
    )
    status.pack(fill=tk.X, pady=(10, 0))

    result = {"ok": False, "err": ""}

    def _append(msg):
        log.configure(state=tk.NORMAL)
        log.insert(tk.END, msg)
        log.see(tk.END)
        log.configure(state=tk.DISABLED)

    def _worker():
        py = _get_python_executable()
        cmd = [py, "-m", "pip", "install", *missing]
        try:
            root.after(0, lambda: _append(f"$ {' '.join(cmd)}\n\n"))
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=CREATE_NO_WINDOW,
            )
            assert proc.stdout is not None
            for line in proc.stdout:
                text = line
                root.after(0, lambda t=text: _append(t))
            code = proc.wait()
            if code != 0:
                result["err"] = f"pip 종료 코드 {code}"
                root.after(
                    0,
                    lambda: status.configure(
                        text="설치 실패 — 로그를 확인하세요.", fg="#BF616A"
                    ),
                )
                root.after(
                    0,
                    lambda: messagebox.showerror(
                        "설치 실패",
                        f"패키지 설치에 실패했습니다.\n\n{result['err']}\n\n"
                        f"수동: {py} -m pip install {' '.join(missing)}",
                        parent=root,
                    ),
                )
                return
            # 재확인
            still = _missing_packages()
            if still:
                result["err"] = f"설치 후에도 남음: {', '.join(still)}"
                root.after(
                    0,
                    lambda: status.configure(
                        text="설치 후에도 일부 패키지가 없습니다.", fg="#BF616A"
                    ),
                )
                root.after(
                    0,
                    lambda: messagebox.showerror(
                        "설치 확인 실패",
                        result["err"]
                        + "\n\n`0.처음사용시\\파이썬 패키지.py` 를 실행해 주세요.",
                        parent=root,
                    ),
                )
                return
            result["ok"] = True
            root.after(
                0,
                lambda: status.configure(
                    text="설치 완료 — 프로그램을 다시 시작합니다…", fg="#A3BE8C"
                ),
            )
            root.after(600, root.quit)
        except Exception as exc:
            result["err"] = str(exc)
            root.after(
                0,
                lambda: status.configure(text=f"오류: {exc}", fg="#BF616A"),
            )
            root.after(
                0,
                lambda: messagebox.showerror(
                    "설치 오류", str(exc), parent=root
                ),
            )

    threading.Thread(target=_worker, daemon=True).start()
    root.mainloop()
    try:
        root.destroy()
    except Exception:
        pass

    if result["ok"]:
        _restart_launcher()
    return False


def ensure_dependencies():
    """
    필수 패키지 검사. 누락 시 설치 UI 후 재시작(프로세스 종료).
    정상·사용자가 닫아 실패하면 False → 런처 본문 진입 여부 판단.
    """
    try:
        missing = _missing_packages()
    except Exception as exc:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "패키지 목록 오류",
            "requirements.txt / req_utils 를 읽지 못했습니다.\n\n"
            f"{exc}\n\n"
            f"경로: {REQUIREMENTS_TXT}",
        )
        root.destroy()
        return False

    if not missing:
        # 정상 기동 시 재시작 플래그 제거
        os.environ.pop(_DEP_RESTART_ENV, None)
        return True
    return _install_packages_with_ui(missing)


# ---------------------------------------------------------------------------
# 프로그램 매핑 및 시각적 그룹
# ---------------------------------------------------------------------------

PROGRAM_GROUPS = (
    {
        "title": "입력",
        "accent": "#5B8A72",
        "items": (
            ("1. 측정인 자동입력 · 검토", "eco_input_gui.py"),
            ("8. 그룹웨어 전송(재전송,직접)", "groupware_resend_gui.py"),
        ),
    },
    {
        "title": "검토",
        "accent": "#5E81AC",
        "items": (
            ("3. 발송대장 검토", "receipt.py"),
            ("5. 성적서 검토", "report_check_gui.py"),
            ("6. 차량운행일지 검토", "Vehicle_operation_log.py"),
        ),
    },
    {
        "title": "종합 · 출력",
        "accent": "#88C0D0",
        "items": (
            ("4. 종합 검토", "dash.py"),
            ("7. PDF 생성", "tab4_pdf_final_gui.py"),
        ),
    },
)

# ---------------------------------------------------------------------------
# 테마 색상 (Slate / 톤다운 Blue-Green)
# ---------------------------------------------------------------------------

COLORS = {
    "bg": "#2E3440",
    "surface": "#3B4252",
    "surface_light": "#434C5E",
    "text": "#ECEFF4",
    "text_muted": "#D8DEE9",
    "btn": "#4C566A",
    "btn_hover": "#5E81AC",
    "btn_active": "#81A1C1",
    "border": "#434C5E",
    "header": "#ECEFF4",
    "subtitle": "#A3BE8C",
    "footer": "#616E88",
}

FONT_FAMILY = "맑은 고딕"


def _get_pythonw_executable():
    """GUI 하위 프로세스는 선호 3.11 pythonw로 실행 (콘솔 없음)."""
    preferred = _resolve_preferred_exe(want_w=True)
    if preferred and preferred.lower().endswith("pythonw.exe"):
        return preferred
    if preferred:
        sibling = os.path.join(os.path.dirname(preferred), "pythonw.exe")
        if os.path.isfile(sibling):
            return sibling
    exe = sys.executable
    name = os.path.basename(exe).lower()
    if name == "pythonw.exe":
        return exe
    if name == "python.exe":
        candidate = os.path.join(os.path.dirname(exe), "pythonw.exe")
        if os.path.isfile(candidate):
            return candidate
    return exe


PYTHONW = _get_pythonw_executable()


# 스크립트 파일명 → 런처 표시명 (시작 중 안내용)
SCRIPT_LABELS = {
    "eco_input_gui.py": "1. 측정인 자동입력 · 검토",
    "groupware_resend_gui.py": "8. 그룹웨어 전송(재전송,직접)",
    "receipt.py": "3. 발송대장 검토",
    "report_check_gui.py": "5. 성적서 검토",
    "Vehicle_operation_log.py": "6. 차량운행일지 검토",
    "dash.py": "4. 종합 검토",
    "tab4_pdf_final_gui.py": "7. PDF 생성",
}


def _show_starting_toast(label: str, ms: int = 2500):
    """버튼 직후 체감용 — NAS 기동 전에 '시작 중' 안내."""
    parent = _root_ref[0]
    toast = tk.Toplevel(parent) if parent is not None else tk.Tk()
    toast.title("실행")
    toast.configure(bg="#3B4252")
    toast.resizable(False, False)
    toast.attributes("-topmost", True)
    try:
        toast.transient(parent)
    except Exception:
        pass
    frame = tk.Frame(toast, bg="#3B4252", padx=22, pady=16)
    frame.pack(fill=tk.BOTH, expand=True)
    tk.Label(
        frame,
        text="프로그램을 시작하는 중…",
        font=("맑은 고딕", 11, "bold"),
        fg="#ECEFF4",
        bg="#3B4252",
    ).pack(anchor="w")
    tk.Label(
        frame,
        text=label,
        font=("맑은 고딕", 9),
        fg="#A3BE8C",
        bg="#3B4252",
    ).pack(anchor="w", pady=(6, 0))
    tk.Label(
        frame,
        text="NAS에서 불러오는 동안 잠시 기다려 주세요.",
        font=("맑은 고딕", 8),
        fg="#88C0D0",
        bg="#3B4252",
    ).pack(anchor="w", pady=(8, 0))
    toast.update_idletasks()
    w, h = max(320, toast.winfo_reqwidth()), toast.winfo_reqheight()
    if parent is not None:
        try:
            px = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
            py = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
            toast.geometry(f"{w}x{h}+{max(0, px)}+{max(0, py)}")
        except Exception:
            pass
    toast.after(ms, toast.destroy)
    toast.update()


def _launch_script(script_name):
    """지정 스크립트를 NAS 기준 경로·무창으로 실행한다."""
    script_path = os.path.join(PY_DIR, script_name)
    if not os.path.isdir(PY_DIR):
        messagebox.showerror(
            "실행 오류",
            f"3.py 폴더를 찾을 수 없습니다.\n\n{PY_DIR}",
            parent=_root_ref[0],
        )
        return
    if not os.path.isfile(script_path):
        messagebox.showerror(
            "실행 오류",
            f"파일을 찾을 수 없습니다.\n\n{script_path}",
            parent=_root_ref[0],
        )
        return

    label = SCRIPT_LABELS.get(script_name, script_name)
    try:
        _show_starting_toast(label)
    except Exception:
        pass

    try:
        subprocess.Popen(
            [PYTHONW, script_path],
            cwd=PY_DIR,
            creationflags=CREATE_NO_WINDOW,
        )
    except OSError as exc:
        messagebox.showerror(
            "실행 오류",
            f"프로그램을 시작하지 못했습니다.\n\n{exc}",
            parent=_root_ref[0],
        )


_root_ref = [None]


class HoverButton(tk.Button):
    """마우스 오버 시 배경색이 변하는 버튼."""

    def __init__(self, master, hover_bg=None, normal_bg=None, **kwargs):
        self._normal_bg = normal_bg or COLORS["btn"]
        self._hover_bg = hover_bg or COLORS["btn_hover"]
        kwargs.setdefault("bg", self._normal_bg)
        kwargs.setdefault("activebackground", COLORS["btn_active"])
        kwargs.setdefault("relief", tk.FLAT)
        kwargs.setdefault("cursor", "hand2")
        kwargs.setdefault("bd", 0)
        kwargs.setdefault("highlightthickness", 0)
        super().__init__(master, **kwargs)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _event=None):
        self.configure(bg=self._hover_bg)

    def _on_leave(self, _event=None):
        self.configure(bg=self._normal_bg)


def _build_group(parent, group):
    """그룹 프레임(제목 + 버튼 목록)을 구성한다."""
    accent = group["accent"]
    frame = tk.Frame(
        parent,
        bg=COLORS["surface"],
        highlightbackground=accent,
        highlightthickness=2,
        padx=16,
        pady=14,
    )
    title_lbl = tk.Label(
        frame,
        text=group["title"],
        font=(FONT_FAMILY, 11, "bold"),
        fg=accent,
        bg=COLORS["surface"],
        anchor="w",
    )
    title_lbl.pack(fill=tk.X, pady=(0, 10))

    btn_frame = tk.Frame(frame, bg=COLORS["surface"])
    btn_frame.pack(fill=tk.BOTH, expand=True)

    btn_font = tkfont.Font(family=FONT_FAMILY, size=10)
    for idx, (label, script) in enumerate(group["items"]):
        btn = HoverButton(
            btn_frame,
            text=label,
            font=btn_font,
            fg=COLORS["text"],
            normal_bg=COLORS["btn"],
            hover_bg=group.get("btn_hover", COLORS["btn_hover"]),
            activeforeground=COLORS["text"],
            padx=18,
            pady=10,
            anchor="w",
            command=lambda s=script: _launch_script(s),
        )
        btn.grid(row=idx, column=0, sticky="ew", pady=4)
    btn_frame.columnconfigure(0, weight=1)
    return frame


def _fit_window(root, outer, min_w=440):
    """배치된 위젯 높이에 맞춰 창 크기를 잡고 화면 중앙에 둔다."""
    root.update_idletasks()
    win_w = max(min_w, outer.winfo_reqwidth() + 48)
    win_h = outer.winfo_reqheight() + 48
    max_h = root.winfo_screenheight() - 80
    win_h = min(win_h, max_h)
    pos_x = max(0, (root.winfo_screenwidth() - win_w) // 2)
    pos_y = max(0, (root.winfo_screenheight() - win_h) // 2)
    root.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
    root.minsize(min_w, 400)


def main():
    root = tk.Tk()
    _root_ref[0] = root
    root.title("검토 · 입력 프로그램")
    root.configure(bg=COLORS["bg"])
    root.resizable(True, True)

    outer = tk.Frame(root, bg=COLORS["bg"], padx=24, pady=20)
    outer.pack(fill=tk.BOTH, expand=True)

    header = tk.Label(
        outer,
        text="측정팀 자동화 도구",
        font=(FONT_FAMILY, 15, "bold"),
        fg=COLORS["header"],
        bg=COLORS["bg"],
    )
    header.pack(anchor="w", pady=(0, 4))

    subtitle = tk.Label(
        outer,
        text="실행할 프로그램을 선택하세요",
        font=(FONT_FAMILY, 9),
        fg=COLORS["subtitle"],
        bg=COLORS["bg"],
    )
    subtitle.pack(anchor="w", pady=(0, 16))

    groups_container = tk.Frame(outer, bg=COLORS["bg"])
    groups_container.pack(fill=tk.BOTH, expand=True)

    for g_idx, group in enumerate(PROGRAM_GROUPS):
        gf = _build_group(groups_container, group)
        gf.pack(fill=tk.X, pady=(0, 12 if g_idx < len(PROGRAM_GROUPS) - 1 else 0))

    path_hint = BASE_DIR
    if len(path_hint) > 58:
        path_hint = "…" + path_hint[-55:]
    py_ver = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    footer = tk.Label(
        outer,
        text=f"작업 경로: {path_hint}\nPython {py_ver}",
        font=(FONT_FAMILY, 8),
        fg=COLORS["footer"],
        bg=COLORS["bg"],
        anchor="w",
        justify="left",
    )
    footer.pack(fill=tk.X, pady=(14, 0))

    _fit_window(root, outer)

    root.mainloop()


if __name__ == "__main__":
    # 파일 상단 _reexec_with_py311() 에서 이미 3.11 고정
    if ensure_dependencies():
        main()
