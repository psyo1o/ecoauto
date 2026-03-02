# -*- coding: utf-8 -*-
"""
성적서 검토 GUI (파일 리스트 방식 + 드래그&드롭 옵션)

요구사항(사용자 피드백 반영):
- PDF 생성기처럼 "파일 리스트(Listbox)" UI로 구성
- 파일 추가 버튼 + (가능하면) 드래그&드롭
- 추가된 항목 삭제(선택 삭제/전체 비우기)
- 파일명에서 시료번호 자동 추출해서 기존 report_check.main(sample_list, user_name) 실행

주의:
- 드래그&드롭은 tkinterdnd2 설치 필요:  pip install tkinterdnd2
- Windows에서 프로그램을 '관리자 권한'으로 실행하면 탐색기에서 드롭이 차단될 수 있습니다.
  (이 경우에도 "파일 추가..." 버튼으로는 정상 동작)
"""

import os
import re
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# 실행 위치와 무관하게 같은 폴더 모듈 import 되도록 고정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

APP_VERSION = "FILELIST_v3_SAFE_DND"

# ------------------------------
# Drag & Drop (옵션)
# ------------------------------
HAS_DND = False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    HAS_DND = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    HAS_DND = False

# ------------------------------
# required modules
# ------------------------------
IMPORT_ERR = None
try:
    from gui_common import LogPanel, set_window_topmost
    import report_check
    try:
        import pythoncom  # type: ignore
    except Exception:
        pythoncom = None
except Exception as e:
    IMPORT_ERR = e
    pythoncom = None


def _parse_drop_files(data: str):
    """
    tkinterdnd2 event.data 파싱
    예) "{C:\a b\c.xlsx} {D:\d.xlsm}"
    """
    out, token, in_brace = [], "", False
    for ch in data:
        if ch == "{":
            in_brace = True
            token = ""
        elif ch == "}":
            in_brace = False
            if token:
                out.append(token)
                token = ""
        elif ch == " " and not in_brace:
            if token:
                out.append(token)
                token = ""
        else:
            token += ch
    if token:
        out.append(token)
    return [p.strip().strip('"') for p in out if p.strip()]


def extract_sample_from_name(path_or_text: str) -> str:
    """
    파일/폴더명/텍스트에서 시료번호 추출
    - 가장 흔한 패턴: A + 숫자(7~9) + - + 순번(1~2)
      예) A2512313-03, A2501011-1 등
    """
    s = os.path.basename(str(path_or_text)).strip().upper()
    s = os.path.splitext(s)[0]
    s = s.replace("–", "-").replace("—", "-").replace("−", "-")
    s = re.sub(r"\s+", "", s)

    m = re.search(r"(A\d{7,9}-\d{1,2})", s)
    if m:
        head = m.group(1)
        m2 = re.match(r"^(A\d{7,9})-(\d{1,2})$", head)
        if m2:
            return f"{m2.group(1)}-{int(m2.group(2)):02d}"
        return head

    # 하이픈이 없는 경우(마지막 2자리를 순번으로 가정)
    m = re.search(r"(A\d{7,9})(\d{2})$", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"

    # 텍스트로 들어온 시료번호 그대로도 허용 (최소 검증)
    if re.fullmatch(r"A\d{7,9}-\d{1,2}", s):
        a, b = s.split("-")
        return f"{a}-{int(b):02d}"

    return ""


class ReportCheckFileListGUI:
    def __init__(self):
        if IMPORT_ERR is not None:
            messagebox.showerror(
                "실행 오류",
                "필수 모듈 import 실패:\n"
                f"{IMPORT_ERR}\n\n"
                "report_check_gui_filelist.py, report_check.py, gui_common.py를 같은 폴더에 두고 실행해 주세요."
            )
            raise SystemExit(1)

        self.files = []      # 원본 파일 경로
        self.samples = []    # 추출된 시료번호(중복 제거)
        self._stop = False

        # tkinterdnd2(tkdnd) 로딩이 PC에 따라 실패할 수 있어 안전 fallback
        global HAS_DND
        if HAS_DND:
            try:
                self.root = TkinterDnD.Tk()
            except Exception as e:
                HAS_DND = False
                self.root = tk.Tk()
                try:
                    messagebox.showwarning(
                        '드래그&드롭 비활성화',
                        'tkinterdnd2(tkdnd) 로딩 실패로 드래그&드롭을 끕니다.\n'                        '파일 추가 버튼으로 사용하세요.\n\n'                        + ('원인: ' + str(e))
                    )
                except Exception:
                    pass
        else:
            self.root = tk.Tk()
        self.root.title("성적서 검토 - 파일 리스트 - " + APP_VERSION)
        self.root.geometry("900x650")
        self.root.minsize(860, 620)
        set_window_topmost(self.root)

        self._build_ui()
        self._setup_cleanup()

        self._log(f"[INFO] 버전: {APP_VERSION}")
        self._log(f"[INFO] DragDrop 사용: {HAS_DND}")
        self._log("[INFO] 드래그&드롭이 안 되면 '파일 추가...' 버튼으로 진행하세요.")

    # ---------------- UI ----------------
    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(4, weight=1)

        # 1) 입력
        top = ttk.LabelFrame(outer, text="입력", padding=10)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="이름(결과 저장용):").grid(row=0, column=0, sticky="w")
        self.ent_name = ttk.Entry(top)
        self.ent_name.grid(row=0, column=1, sticky="ew", padx=6)

        # 2) 파일 리스트
        mid = ttk.LabelFrame(outer, text="성적서 엑셀 파일", padding=10)
        mid.grid(row=1, column=0, sticky="nsew", pady=(10, 10))
        mid.columnconfigure(0, weight=1)
        mid.rowconfigure(1, weight=1)

        hint = "드래그&드롭 가능" if HAS_DND else "드래그&드롭 미지원(설치 필요: pip install tkinterdnd2)"
        ttk.Label(mid, text=f"※ {hint} | 허용: .xls/.xlsx/.xlsm (파일명에서 시료번호 자동 추출)", foreground="gray").grid(
            row=0, column=0, sticky="w"
        )
        if not HAS_DND:
            ttk.Button(mid, text="드래그&드롭 활성화(설치)", command=self._install_dnd).grid(row=0, column=1, sticky="e")

        self.lst = tk.Listbox(mid, height=10, selectmode="extended")
        self.lst.grid(row=1, column=0, sticky="nsew", pady=(8, 0))

        sb = ttk.Scrollbar(mid, orient="vertical", command=self.lst.yview)
        sb.grid(row=1, column=1, sticky="ns", pady=(8, 0))
        self.lst.configure(yscrollcommand=sb.set)

        # 드롭 전용 라벨(일부 환경에서 listbox/text 드롭이 씹혀도 라벨은 되는 경우가 있음)
        self.lbl_drop = ttk.Label(mid, text="여기에 엑셀 파일을 드래그&드롭", anchor="center", relief="groove")
        self.lbl_drop.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0), ipady=6)

        btns = ttk.Frame(mid)
        btns.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)
        btns.columnconfigure(2, weight=1)

        ttk.Button(btns, text="파일 추가...", command=self._add_files_dialog).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(btns, text="선택 삭제", command=self._remove_selected).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(btns, text="전체 비우기", command=self._clear_all).grid(row=0, column=2, sticky="ew", padx=4)

        # 3) 실행/진행
        runbox = ttk.Frame(outer)
        runbox.grid(row=2, column=0, sticky="ew")
        runbox.columnconfigure(0, weight=1)
        runbox.columnconfigure(1, weight=1)

        self.btn_run = ttk.Button(runbox, text="실행", command=self._on_run)
        self.btn_run.grid(row=0, column=0, sticky="ew", padx=4)

        self.btn_quit = ttk.Button(runbox, text="닫기", command=self.root.destroy)
        self.btn_quit.grid(row=0, column=1, sticky="ew", padx=4)

        self.progress = ttk.Progressbar(outer, mode="determinate")
        self.progress.grid(row=3, column=0, sticky="ew", pady=(8, 0))

        # 4) 로그
        bot = ttk.LabelFrame(outer, text="로그", padding=10)
        bot.grid(row=4, column=0, sticky="nsew", pady=(10, 0))
        bot.columnconfigure(0, weight=1)
        bot.rowconfigure(0, weight=1)

        self.log_panel = LogPanel(bot, height=22)
        self.log_panel.grid(row=0, column=0, sticky="nsew")
        self.log_panel.start_pumping()

        # DND 등록
        if HAS_DND:
            try:
                self.lst.drop_target_register(DND_FILES)
                self.lst.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass
            try:
                self.lbl_drop.drop_target_register(DND_FILES)
                self.lbl_drop.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass
            try:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

    def _setup_cleanup(self):
        def on_close():
            self._stop = True
            try:
                self.root.destroy()
            except Exception:
                pass
        self.root.protocol("WM_DELETE_WINDOW", on_close)

    def _log(self, msg: str):
        try:
            self.log_panel.write(msg + "\n")
        except Exception:
            pass

    # ---------------- file list ops ----------------
    def _refresh_listbox(self):
        self.lst.delete(0, "end")
        for p in self.files:
            sn = extract_sample_from_name(p)
            show = f"{sn}  |  {p}" if sn else p
            self.lst.insert("end", show)

    def _add_paths(self, paths):
        added = 0
        for p in paths:
            p = p.strip().strip('"')
            if not p:
                continue
            if os.path.isdir(p):
                # 폴더가 드롭되면 폴더 안의 엑셀을 1단만 스캔
                try:
                    for fn in os.listdir(p):
                        full = os.path.join(p, fn)
                        if os.path.isfile(full) and os.path.splitext(full)[1].lower() in (".xls", ".xlsx", ".xlsm"):
                            if full not in self.files:
                                self.files.append(full)
                                added += 1
                except Exception:
                    continue
            else:
                if not os.path.isfile(p):
                    continue
                if os.path.splitext(p)[1].lower() not in (".xls", ".xlsx", ".xlsm"):
                    continue
                if p not in self.files:
                    self.files.append(p)
                    added += 1

        if added:
            self._refresh_listbox()
            self._log(f"[INFO] 파일 {added}개 추가됨 (총 {len(self.files)}개)")
        else:
            self._log("[INFO] 추가된 파일이 없습니다. (중복/확장자/경로 확인)")

    def _add_files_dialog(self):
        paths = filedialog.askopenfilenames(
            title="성적서 엑셀 파일 선택",
            filetypes=[("Excel files", "*.xls;*.xlsx;*.xlsm"), ("All files", "*.*")]
        )
        if paths:
            self._add_paths(list(paths))

    def _remove_selected(self):
        idxs = list(self.lst.curselection())
        if not idxs:
            return
        for i in reversed(idxs):
            # listbox에는 "sn | path"로 표시되므로 원본 files 인덱스 그대로 사용
            try:
                del self.files[i]
            except Exception:
                pass
        self._refresh_listbox()
        self._log(f"[INFO] 선택 삭제 완료 (총 {len(self.files)}개)")

    def _clear_all(self):
        self.files = []
        self.lst.delete(0, "end")
        self._log("[INFO] 전체 비움 완료")

    # ---------------- DND ----------------
    def _on_drop(self, event):
        try:
            data = event.data
            self._log(f"[DND] drop data: {data}")
            paths = _parse_drop_files(data)
            if not paths:
                return
            self._add_paths(paths)
        except Exception as e:
            self._log(f"[DND] 오류: {e}")

    # ---------------- run ----------------
    def _on_run(self):
        user_name = self.ent_name.get().strip()
        if not user_name:
            messagebox.showerror("오류", "이름을 입력하세요.")
            return
        if not self.files:
            messagebox.showerror("오류", "엑셀 파일을 1개 이상 추가하세요.")
            return

        # 파일 목록에서 시료번호 추출
        sample_list = []
        bad = []
        for p in self.files:
            sn = extract_sample_from_name(p)
            if sn:
                if sn not in sample_list:
                    sample_list.append(sn)
            else:
                bad.append(os.path.basename(p))

        if not sample_list:
            messagebox.showerror("오류", "추출된 시료번호가 없습니다. 파일명이 시료번호 규칙인지 확인하세요.")
            return

        if bad:
            self._log("[WARN] 시료번호 추출 실패 파일(무시됨): " + ", ".join(bad))

        total = len(sample_list)
        self.btn_run.config(state="disabled", text=f"준비 중... (0/{total})")
        self.progress.config(maximum=total, value=0)

        def worker():
            try:
                if pythoncom is not None:
                    pythoncom.CoInitialize()

                # stdout 모니터로 진행률 업데이트(기존 GUI와 동일 방식)
                import sys as _sys
                original_stdout = _sys.stdout

                class ProgressMonitor:
                    def __init__(self, original, total, root, btn, prog_bar):
                        self.original = original
                        self.total = total
                        self.root = root
                        self.btn = btn
                        self.prog_bar = prog_bar
                        self.current = 0

                    def write(self, text):
                        self.original.write(text)
                        if "처리: " in text and "------" not in text:
                            self.current += 1
                            self.root.after(0, lambda c=self.current: self._update(c))

                    def _update(self, current):
                        self.btn.config(text=f"처리 중... ({current}/{self.total})")
                        self.prog_bar["value"] = current

                    def flush(self):
                        self.original.flush()

                monitor = ProgressMonitor(original_stdout, total, self.root, self.btn_run, self.progress)
                _sys.stdout = monitor

                try:
                    report_check.main(sample_list, user_name)
                finally:
                    _sys.stdout = original_stdout

                self.root.after(0, lambda: (
                    self.btn_run.config(text=f"완료! ({total}/{total})"),
                    self.progress.config(value=total)
                ))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("오류", str(e)))
                self._log("[ERROR] " + str(e))
            finally:
                try:
                    if pythoncom is not None:
                        pythoncom.CoUninitialize()
                except Exception:
                    pass
                self.root.after(0, lambda: self.btn_run.config(state="normal", text="실행"))

        threading.Thread(target=worker, daemon=True).start()

    def run(self):
        self.root.mainloop()


def main():
    ReportCheckFileListGUI().run()


if __name__ == "__main__":
    main()