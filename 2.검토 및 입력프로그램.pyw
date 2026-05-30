# -*- coding: utf-8 -*-
"""
통합 마스터 런처 (2.검토 및 입력프로그램.pyw)
NAS UNC 경로에서 더블 클릭 실행 · 콘솔 창 없음 · 하위 GUI 무창 실행
"""

import os
import sys
import subprocess
import tkinter as tk
from tkinter import messagebox, font as tkfont

# Windows: 하위 프로세스 콘솔 창 억제
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

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

# ---------------------------------------------------------------------------
# 프로그램 매핑 및 시각적 그룹
# ---------------------------------------------------------------------------

PROGRAM_GROUPS = (
    {
        "title": "입력",
        "accent": "#5B8A72",
        "items": (
            ("1. 측정인 자동입력", "eco_input_gui.py"),
        ),
    },
    {
        "title": "검토",
        "accent": "#5E81AC",
        "items": (
            ("2. 측정인 검토", "eco_check_gui.py"),
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
    """GUI 하위 프로세스는 pythonw로 실행하여 콘솔이 뜨지 않게 한다."""
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
    footer = tk.Label(
        outer,
        text=f"작업 경로: {path_hint}",
        font=(FONT_FAMILY, 8),
        fg=COLORS["footer"],
        bg=COLORS["bg"],
        anchor="w",
    )
    footer.pack(fill=tk.X, pady=(14, 0))

    _fit_window(root, outer)

    root.mainloop()


if __name__ == "__main__":
    main()
