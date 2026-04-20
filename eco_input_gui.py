# -*- coding: utf-8 -*-
"""
측정인 자동 입력 및 백데이터 GUI
- 매체(대기/수질) 선택에 따라 UI 동적 전환
- 대기: 기존 탭1/2/백데이터/탭4 흐름
- 수질: 시료번호 + 탭4만
"""

import datetime
import os
import re
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import builtins
import warnings

warnings.filterwarnings(
    "ignore",
    category=UserWarning,
    message="Conditional Formatting extension is not supported and will be removed"
)
warnings.filterwarnings(
    "ignore",
    category=UserWarning,
    message="Revert to STA COM threading mode"
)

# ------------------------------
# Drag & Drop (옵션) — 반드시 eco_input보다 먼저 import해야 함
# (eco_input이 pythoncom/win32com을 로드하면 tkdnd DLL과 충돌)
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
# required modules (DnD 뒤에 로드)
# eco_input은 pythoncom/win32com을 즉시 로드하므로
# 모듈 레벨에서 import하면 tkdnd DLL과 충돌함 → lazy import
# ------------------------------
from gui_common import LogPanel, set_window_topmost, set_state_recursive
from data_utils import extract_sample_from_name as common_extract_sample_from_name, parse_ymd_date


def _parse_drop_files(data: str):
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


def _extract_sample_from_name(path_or_text: str) -> str:
    return common_extract_sample_from_name(path_or_text)


class EcoInputGUI:
    """측정인 자동 입력 GUI"""

    def __init__(self):
        # tkinterdnd2(tkdnd) 로딩이 PC에 따라 실패할 수 있어 안전 fallback
        global HAS_DND
        self._dnd_fail_msg = None
        if HAS_DND:
            try:
                self.root = TkinterDnD.Tk()
            except Exception as e:
                HAS_DND = False
                self.root = tk.Tk()
                self._dnd_fail_msg = str(e)
        else:
            self.root = tk.Tk()
        self._setup_window()
        self._create_widgets()
        self._setup_dependencies()
        # DnD 실패 메시지는 UI 구성이 끝난 뒤 로그로만 표시
        if self._dnd_fail_msg:
            self.log_panel.widget.configure(state="normal")
            self.log_panel.widget.insert("end", f"[INFO] 드래그&드롭 비활성화 (원인: {self._dnd_fail_msg})\n")
            self.log_panel.widget.configure(state="disabled")

    # ──────────────────────────────────────────────
    # 윈도우 / 위젯 초기 구성
    # ──────────────────────────────────────────────
    def _setup_window(self):
        self.root.title("측정인 자동 입력 및 백데이터")
        self.root.geometry("860x1030")
        self.root.minsize(860, 930)
        self.root.resizable(True, True)
        set_window_topmost(self.root)

    def _create_widgets(self):
        outer = ttk.Frame(self.root)
        outer.pack(fill="both", expand=True)
        outer.grid_columnconfigure(0, weight=1, uniform="col")
        outer.grid_columnconfigure(1, weight=1, uniform="col")
        outer.grid_rowconfigure(0, weight=1)

        self.input_frame = ttk.Frame(outer, padding=15)
        self.input_frame.grid(row=0, column=0, sticky="nsew")
        self.input_frame.grid_columnconfigure(1, weight=1)

        self._create_input_area()

        log_pane = ttk.Frame(outer, padding=(0, 15, 15, 15))
        log_pane.grid(row=0, column=1, sticky="nsew")
        log_pane.grid_rowconfigure(0, weight=1)
        log_pane.grid_columnconfigure(0, weight=1)
        self._create_log_area(log_pane)

    def _create_input_area(self):
        self._create_media_selection()   # row 0
        self._create_login_fields()      # row 1, 2
        self._create_air_area()          # row 3~6  (대기 전용)
        self._create_water_area()        # row 7    (수질 전용)
        self._create_progress_area()     # row 8
        self._create_action_button()     # row 9

    # ── 매체 선택 ──────────────────────────────────
    def _create_media_selection(self):
        self.media_var = tk.StringVar(value="1")

        f = ttk.LabelFrame(self.input_frame, text="매체 선택", padding=10)
        f.grid(row=0, column=0, columnspan=2, sticky="we", pady=8)

        ttk.Radiobutton(f, text="대기", value="1",
                        variable=self.media_var).grid(row=0, column=0, sticky="w", padx=10)
        ttk.Radiobutton(f, text="수질 (탭4만)", value="2",
                        variable=self.media_var).grid(row=0, column=1, sticky="w", padx=10)

    # ── 로그인 (공통) ──────────────────────────────
    def _create_login_fields(self):
        ttk.Label(self.input_frame, text="측정인 아이디:").grid(
            row=1, column=0, sticky="e", pady=5)
        self.entry_id = ttk.Entry(self.input_frame, width=38)
        self.entry_id.grid(row=1, column=1, sticky="we")

        ttk.Label(self.input_frame, text="비밀번호:").grid(
            row=2, column=0, sticky="e", pady=5)
        self.entry_pw = ttk.Entry(self.input_frame, width=38, show="*")
        self.entry_pw.grid(row=2, column=1, sticky="we")

    # ── 대기 전용 영역 ─────────────────────────────
    def _create_air_area(self):
        # row 3: 작업 선택 + 모드 선택 (같은 줄)
        self.job_var = tk.StringVar(value="1")
        self.mode_var = tk.StringVar(value="2")

        job_mode_frame = ttk.Frame(self.input_frame)
        job_mode_frame.grid(row=3, column=0, columnspan=2, sticky="we", pady=4)
        job_mode_frame.grid_columnconfigure(0, weight=1)
        job_mode_frame.grid_columnconfigure(1, weight=1)

        self.air_job_frame = ttk.LabelFrame(job_mode_frame, text="작업 선택", padding=6)
        self.air_job_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        ttk.Radiobutton(self.air_job_frame,
                        text="1) 웹 자동입력",
                        value="1", variable=self.job_var).grid(row=0, column=0, sticky="w", pady=1)
        ttk.Radiobutton(self.air_job_frame,
                        text="2) 백데이터만",
                        value="2", variable=self.job_var).grid(row=1, column=0, sticky="w", pady=1)

        self.air_mode_frame = ttk.LabelFrame(job_mode_frame, text="모드 선택", padding=6)
        self.air_mode_frame.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        ttk.Radiobutton(self.air_mode_frame,
                        text="1) 시료번호 직접입력",
                        value="1", variable=self.mode_var).grid(row=0, column=0, sticky="w", pady=1)
        ttk.Radiobutton(self.air_mode_frame,
                        text="2) 팀+날짜 자동추출",
                        value="2", variable=self.mode_var).grid(row=1, column=0, sticky="w", pady=1)

        # row 4: 기능 선택
        self.tab1_var      = tk.BooleanVar(value=True)
        self.tab2_var      = tk.BooleanVar(value=True)
        self.pdf_var       = tk.BooleanVar(value=True)
        self.backdata_var  = tk.BooleanVar(value=True)
        self.tab4_var      = tk.BooleanVar(value=False)
        self.pdf_final_var = tk.BooleanVar(value=False)

        self.air_feat_frame = ttk.LabelFrame(self.input_frame, text="입력/부가기능 선택", padding=10)
        self.air_feat_frame.grid(row=4, column=0, columnspan=2, sticky="we", pady=4)

        self.chk_tab1 = ttk.Checkbutton(self.air_feat_frame,
            text="현장측정정보 (탭1)", variable=self.tab1_var)
        self.chk_tab1.grid(row=0, column=0, sticky="w", padx=5, pady=2)

        self.chk_tab2 = ttk.Checkbutton(self.air_feat_frame,
            text="시료채취/측정정보 (탭2)", variable=self.tab2_var)
        self.chk_tab2.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        self.chk_pdf = ttk.Checkbutton(self.air_feat_frame,
            text="탭2 PDF 생성/업로드", variable=self.pdf_var)
        self.chk_pdf.grid(row=1, column=0, sticky="w", padx=5, pady=2, columnspan=2)

        self.chk_backdata = ttk.Checkbutton(self.air_feat_frame,
            text="백데이터(수분/THC)", variable=self.backdata_var)
        self.chk_backdata.grid(row=2, column=0, sticky="w", padx=5, pady=2, columnspan=2)

        self.chk_tab4 = ttk.Checkbutton(self.air_feat_frame,
            text="측정분석결과 (탭4)", variable=self.tab4_var)
        self.chk_tab4.grid(row=3, column=0, sticky="w", padx=5, pady=2, columnspan=2)

        self.chk_pdf_final = ttk.Checkbutton(self.air_feat_frame,
            text="탭4 PDF 최종본 생성/업로드", variable=self.pdf_final_var)
        self.chk_pdf_final.grid(row=4, column=0, sticky="w", padx=5, pady=2, columnspan=2)

        # row 5: 모드2 입력 (팀+날짜)
        self.team_date_frame = ttk.LabelFrame(self.input_frame, text="모드2: 팀+날짜", padding=10)
        self.team_date_frame.grid(row=5, column=0, columnspan=2, sticky="we", pady=5)
        
        ttk.Label(self.team_date_frame, text="팀(미선택=전체):").grid(row=0, column=0, sticky="w", pady=3, padx=(0, 3))
        
        # 체크박스를 담을 내부 프레임
        team_chk_frame = ttk.Frame(self.team_date_frame)
        team_chk_frame.grid(row=0, column=1, sticky="w", padx=0)
        
        self.team_vars = {}
        for i in range(1, 6):
            var = tk.IntVar(value=0)
            ttk.Checkbutton(
                team_chk_frame, 
                text=f"{i}팀", 
                variable=var
            ).grid(row=0, column=i-1, padx=0)
            self.team_vars[i] = var

        ttk.Label(self.team_date_frame, text="날짜 (YYYY-MM-DD):").grid(row=1, column=0, sticky="w", pady=5, padx=(0, 3))
        self.entry_date = ttk.Entry(self.team_date_frame, width=15)
        self.entry_date.insert(0, datetime.date.today().strftime("%Y-%m-%d"))
        self.entry_date.grid(row=1, column=1, sticky="w")

        # row 6: 모드1 입력 (시료번호)
        self.air_direct_frame = ttk.LabelFrame(
            self.input_frame, text="모드1: 대기 시료번호 입력", padding=10)
        self.air_direct_frame.grid(row=6, column=0, columnspan=2, sticky="we", pady=4)
        self.air_direct_frame.grid_columnconfigure(0, weight=1)

        self.lbl_hint_air = ttk.Label(self.air_direct_frame,
                  text="시료번호 (줄바꿈 구분):  파일 드래그 또는 파일추가 버튼 사용",
                  foreground="gray")
        self.lbl_hint_air.grid(row=0, column=0, columnspan=2, sticky="w")

        self.txt_samples_air = tk.Text(self.air_direct_frame, width=40, height=5,
                                          relief="flat", highlightthickness=1,
                                          highlightbackground="black", highlightcolor="black")
        self.txt_samples_air.grid(row=1, column=0, sticky="nsew", pady=5)

        # 파일 추가/삭제/비우기 버튼 (오른쪽)
        air_btns = ttk.Frame(self.air_direct_frame)
        air_btns.grid(row=1, column=1, sticky="ns", padx=(8, 0))
        ttk.Button(air_btns, text="파일 추가...", command=self._add_files_air).pack(fill="x", pady=2)
        ttk.Button(air_btns, text="선택 삭제", command=lambda: self._remove_selected(self.txt_samples_air)).pack(fill="x", pady=2)
        ttk.Button(air_btns, text="전체 비우기", command=lambda: self._clear_all(self.txt_samples_air)).pack(fill="x", pady=2)
        if not HAS_DND:
            ttk.Button(air_btns, text="DnD 설치", command=self._install_dnd).pack(fill="x", pady=(8, 2))

        self.lbl_drop_air = ttk.Label(
            self.air_direct_frame,
            text="여기에 파일을 드래그&드롭",
            anchor="center",
            relief="groove"
        )
        self.lbl_drop_air.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 2), ipady=5)

    # ── 수질 전용 영역 ─────────────────────────────
    def _create_water_area(self):
        self.water_pdf_final_var = tk.BooleanVar(value=False)

        self.water_frame = ttk.LabelFrame(
            self.input_frame, text="수질 시료번호 입력", padding=10)
        self.water_frame.grid(row=7, column=0, columnspan=2, sticky="we", pady=4)
        self.water_frame.grid_columnconfigure(0, weight=1)

        self.lbl_hint_water = ttk.Label(self.water_frame,
                  text="시료번호 (줄바꿈 구분):  파일 드래그 또는 파일추가 버튼 사용",
                  foreground="gray")
        self.lbl_hint_water.grid(row=0, column=0, columnspan=2, sticky="w")

        self.txt_samples_water = tk.Text(self.water_frame, width=40, height=5,
                                            relief="flat", highlightthickness=1,
                                            highlightbackground="black", highlightcolor="black")
        self.txt_samples_water.grid(row=1, column=0, sticky="nsew", pady=5)

        # 파일 추가/삭제/비우기 버튼 (오른쪽)
        water_btns = ttk.Frame(self.water_frame)
        water_btns.grid(row=1, column=1, sticky="ns", padx=(8, 0))
        ttk.Button(water_btns, text="파일 추가...", command=self._add_files_water).pack(fill="x", pady=2)
        ttk.Button(water_btns, text="선택 삭제", command=lambda: self._remove_selected(self.txt_samples_water)).pack(fill="x", pady=2)
        ttk.Button(water_btns, text="전체 비우기", command=lambda: self._clear_all(self.txt_samples_water)).pack(fill="x", pady=2)
        if not HAS_DND:
            ttk.Button(water_btns, text="DnD 설치", command=self._install_dnd).pack(fill="x", pady=(8, 2))

        self.lbl_drop_water = ttk.Label(
            self.water_frame,
            text="여기에 파일을 드래그&드롭",
            anchor="center",
            relief="groove"
        )
        self.lbl_drop_water.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 5), ipady=5)

        ttk.Label(self.water_frame,
                  text="※ 수질은 탭4 PDF 생성/업로드만 수행합니다.",
                  foreground="gray").grid(row=3, column=0, sticky="w", pady=2)

        opt_frame = ttk.Frame(self.water_frame)
        opt_frame.grid(row=4, column=0, sticky="w")
        ttk.Checkbutton(opt_frame, text="PDF 생성/업로드 (탭4)",
                        variable=self.water_pdf_final_var).grid(row=0, column=0, sticky="w", padx=5)

    def _create_progress_area(self):
        self.progress_bar = ttk.Progressbar(self.input_frame, mode="determinate")
        self.progress_bar.grid(row=8, column=0, columnspan=2, sticky="ew", pady=5)

    def _create_action_button(self):
        self.start_btn = ttk.Button(
            self.input_frame, text="자동 입력 시작", command=self._on_start)
        self.start_btn.grid(row=9, column=0, columnspan=2, pady=10)

    def _create_log_area(self, parent):
        log_frame = ttk.LabelFrame(parent, text="로그", padding=10)
        log_frame.grid(row=0, column=0, sticky="nsew")
        self.log_panel = LogPanel(log_frame, height=30)
        self.log_panel.log_text.configure(width=30)
        self.log_panel.pack(fill="both", expand=True)
        self.log_panel.start_pumping()

    # ──────────────────────────────────────────────
    # 의존성 (trace)
    # ──────────────────────────────────────────────
    def _setup_dependencies(self):

        def update_media(*_):
            """매체에 따라 대기/수질 영역 전환"""
            is_water = self.media_var.get() == "2"
            air_frames = [
                self.air_job_frame, self.air_mode_frame, self.air_feat_frame,
                self.team_date_frame, self.air_direct_frame,
            ]
            if is_water:
                for f in air_frames:
                    set_state_recursive(f, "disabled")
                self.txt_samples_air.configure(
                    highlightbackground="#c0c0c0", highlightcolor="#c0c0c0",
                    bg="#f0f0f0", fg="gray")
                self.lbl_hint_air.configure(foreground="#c0c0c0")
                set_state_recursive(self.water_frame, "normal")
                self.txt_samples_water.configure(
                    highlightbackground="black", highlightcolor="black",
                    bg="white", fg="black")
                self.lbl_hint_water.configure(foreground="gray")
            else:
                set_state_recursive(self.water_frame, "disabled")
                self.txt_samples_water.configure(
                    highlightbackground="#c0c0c0", highlightcolor="#c0c0c0",
                    bg="#f0f0f0", fg="gray")
                self.lbl_hint_water.configure(foreground="#c0c0c0")
                for f in air_frames:
                    set_state_recursive(f, "normal")
                # 대기 내부 의존성 재적용
                update_pdf_state()
                update_pdf_final_state()
                update_mode_ui()
                update_job_ui()

        def update_pdf_state(*_):
            if self.tab2_var.get():
                self.chk_pdf.configure(state="normal")
            else:
                self.pdf_var.set(False)
                self.chk_pdf.configure(state="disabled")

        def update_pdf_final_state(*_):
            if self.tab4_var.get():
                self.chk_pdf_final.configure(state="normal")
            else:
                self.pdf_final_var.set(False)
                self.chk_pdf_final.configure(state="disabled")

        def update_mode_ui(*_):
            if self.mode_var.get() == "1":
                set_state_recursive(self.team_date_frame, "disabled")
                set_state_recursive(self.air_direct_frame, "normal")
                self.txt_samples_air.configure(state="normal", bg="white", fg="black",
                                               highlightbackground="black", highlightcolor="black")
                self.lbl_hint_air.configure(foreground="gray")
            else:
                set_state_recursive(self.team_date_frame, "normal")
                set_state_recursive(self.air_direct_frame, "disabled")
                self.txt_samples_air.configure(state="disabled", bg="#f0f0f0", fg="gray",
                                               highlightbackground="#c0c0c0", highlightcolor="#c0c0c0")
                self.lbl_hint_air.configure(foreground="#c0c0c0")

        def update_job_ui(*_):
            if self.job_var.get() == "2":
                self.entry_id.delete(0, "end")
                self.entry_pw.delete(0, "end")
                self.entry_id.configure(state="disabled")
                self.entry_pw.configure(state="disabled")
                self.tab1_var.set(False); self.tab2_var.set(False)
                self.pdf_var.set(False);  self.tab4_var.set(False)
                self.pdf_final_var.set(False); self.backdata_var.set(True)
                for w in [self.chk_tab1, self.chk_tab2, self.chk_pdf,
                          self.chk_tab4, self.chk_pdf_final]:
                    w.configure(state="disabled")
                self.chk_backdata.configure(state="normal")
            else:
                self.entry_id.configure(state="normal")
                self.entry_pw.configure(state="normal")
                for w in [self.chk_tab1, self.chk_tab2,
                          self.chk_backdata, self.chk_tab4]:
                    w.configure(state="normal")
                update_pdf_state()
                update_pdf_final_state()

        # 트레이스 연결
        self.media_var.trace_add("write", update_media)
        self.tab2_var.trace_add("write", update_pdf_state)
        self.tab4_var.trace_add("write", update_pdf_final_state)
        self.mode_var.trace_add("write", update_mode_ui)
        self.job_var.trace_add("write", update_job_ui)

        # 초기 상태
        update_media()
        self._setup_dnd()

    def _setup_dnd(self):
        if not HAS_DND:
            return

        for widget, handler in [
            (self.txt_samples_air, self._on_drop_air),
            (self.lbl_drop_air, self._on_drop_air),
            (self.air_direct_frame, self._on_drop_air),
            (self.txt_samples_water, self._on_drop_water),
            (self.lbl_drop_water, self._on_drop_water),
            (self.water_frame, self._on_drop_water),
        ]:
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", handler)
            except Exception:
                pass

        try:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop_by_media)
        except Exception:
            pass

    def _install_dnd(self):
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "tkinterdnd2"])
            messagebox.showinfo(
                "완료",
                "tkinterdnd2 설치가 완료되었습니다.\n프로그램을 닫고 다시 실행하면 드래그&드롭이 활성화됩니다.",
                parent=self.root
            )
        except Exception as e:
            messagebox.showerror("실패", f"설치 실패: {e}\n\n수동 설치: pip install tkinterdnd2", parent=self.root)

    def _add_files_air(self):
        self._add_files_to_text(self.txt_samples_air)

    def _add_files_water(self):
        self._add_files_to_text(self.txt_samples_water)

    def _add_files_to_text(self, text_widget):
        from tkinter import filedialog
        paths = filedialog.askopenfilenames(
            title="성적서 엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xls"), ("All files", "*.*")],
            parent=self.root
        )
        if not paths:
            return
        import eco_input as _eco
        existing = _eco.parse_sample_input(text_widget.get("1.0", "end"))
        samples = list(existing)
        added = 0
        for p in paths:
            sn = _extract_sample_from_name(p)
            if sn and sn not in samples:
                samples.append(sn)
                added += 1
        if added == 0:
            messagebox.showwarning("파일 추가", "파일명에서 시료번호를 찾지 못했습니다.", parent=self.root)
            return
        text_widget.configure(state="normal")
        text_widget.delete("1.0", "end")
        text_widget.insert("1.0", "\n".join(samples))

    def _remove_selected(self, text_widget):
        try:
            text_widget.delete("sel.first", "sel.last")
        except tk.TclError:
            messagebox.showinfo("안내", "삭제할 텍스트를 마우스로 선택하세요.", parent=self.root)

    def _clear_all(self, text_widget):
        text_widget.configure(state="normal")
        text_widget.delete("1.0", "end")

    def _append_samples_from_drop(self, text_widget, data: str):
        paths = _parse_drop_files(data)
        if not paths:
            return

        import eco_input as _eco  # lazy import
        existing = _eco.parse_sample_input(text_widget.get("1.0", "end"))
        samples = list(existing)
        added = 0

        for path in paths:
            sample_no = _extract_sample_from_name(path)
            if not sample_no or sample_no in samples:
                continue
            samples.append(sample_no)
            added += 1

        if added == 0:
            messagebox.showwarning("드래그&드롭", "드롭한 파일명에서 시료번호를 찾지 못했습니다.", parent=self.root)
            return

        text_widget.configure(state="normal")
        text_widget.delete("1.0", "end")
        text_widget.insert("1.0", "\n".join(samples))

    def _on_drop_air(self, event):
        if self.mode_var.get() != "1":
            return
        try:
            self._append_samples_from_drop(self.txt_samples_air, event.data)
        except Exception as e:
            messagebox.showerror("드래그&드롭 오류", str(e), parent=self.root)

    def _on_drop_water(self, event):
        try:
            self._append_samples_from_drop(self.txt_samples_water, event.data)
        except Exception as e:
            messagebox.showerror("드래그&드롭 오류", str(e), parent=self.root)

    def _on_drop_by_media(self, event):
        if self.media_var.get() == "2":
            return self._on_drop_water(event)
        if self.mode_var.get() == "1":
            return self._on_drop_air(event)

    # ──────────────────────────────────────────────
    # 실행 버튼
    # ──────────────────────────────────────────────
    def _on_start(self):
        media = self.media_var.get()

        if media == "2":
            answers = self._build_answers_water()
        else:
            answers = self._build_answers_air()

        if answers is None:
            return  # 검증 실패

        if not messagebox.askyesno("실행 확인", "지금 실행할까요?", parent=self.root):
            return

        self._run_automation(answers)

    def _build_answers_water(self) -> list:
        """수질 흐름 answers 구성.
        main() → "2"
        _main_water() 순서: login_id, login_pw, 시료번호, pdf(예/아니오)
        """
        login_id = self.entry_id.get().strip()
        login_pw = self.entry_pw.get().strip()
        if not login_id or not login_pw:
            messagebox.showwarning("입력 오류", "아이디/비밀번호를 입력해 주세요.", parent=self.root)
            return None

        raw = self.txt_samples_water.get("1.0", "end").strip()
        if not raw:
            messagebox.showwarning("입력 오류", "수질 시료번호를 입력해 주세요.", parent=self.root)
            return None

        return [
            "2",                                        # 매체: 수질
            login_id,
            login_pw,
            raw,                                        # 시료번호
            self._yn(self.water_pdf_final_var),         # PDF 여부
        ]

    def _build_answers_air(self) -> list:
        """대기 흐름 answers 구성.
        main() → "1"
        _main_air() 순서: job, [id, pw], mode, tab1~pdf_final, [시료번호 or 팀+날짜]
        """
        job  = self.job_var.get()
        mode = self.mode_var.get()

        do_tab1      = self.tab1_var.get()
        do_tab2      = self.tab2_var.get()
        do_pdf       = self.pdf_var.get() if do_tab2 else False
        do_backdata  = self.backdata_var.get()
        do_tab4      = self.tab4_var.get()
        do_pdf_final = self.pdf_final_var.get()

        answers = ["1", job]  # 매체=대기, 작업 선택

        if job == "1":
            login_id = self.entry_id.get().strip()
            login_pw = self.entry_pw.get().strip()
            if not login_id or not login_pw:
                messagebox.showwarning("입력 오류", "아이디/비밀번호를 입력해 주세요.", parent=self.root)
                return None
            if not any([do_tab1, do_tab2, do_backdata, do_tab4]):
                messagebox.showwarning("입력 오류", "최소 1개 이상 선택해야 합니다.", parent=self.root)
                return None
            answers += [login_id, login_pw]
        else:
            do_tab1 = do_tab2 = do_pdf = do_tab4 = do_pdf_final = False
            do_backdata = True

        answers += [
            mode,
            self._yn(do_tab1),
            self._yn(do_tab2),
            self._yn(do_pdf),
            self._yn(do_backdata),
            self._yn(do_tab4),
            self._yn(do_pdf_final),
        ]

        if mode == "1":
            raw = self.txt_samples_air.get("1.0", "end").strip()
            if not raw:
                messagebox.showwarning("입력 오류", "시료번호를 입력해 주세요.", parent=self.root)
                return None
            answers.append(raw)
        else:
            # 선택된 팀 수집하여 콤마(,)로 연결 (예: "1,3,4" 또는 미선택 시 "")
            selected_teams = [str(i) for i in range(1, 6) if self.team_vars[i].get() == 1]
            team_no = ",".join(selected_teams)
            
            date_str = self.entry_date.get().strip()
            
            if not date_str:
                messagebox.showwarning("입력 오류", "날짜를 입력해 주세요.", parent=self.root)
                return None
                
            if parse_ymd_date(date_str) is None:
                messagebox.showwarning("입력 오류", "날짜 형식: YYYY-MM-DD", parent=self.root)
                return None
                
            answers += [team_no, date_str]

        return answers

    # ──────────────────────────────────────────────
    # 백그라운드 실행
    # ──────────────────────────────────────────────
    def _run_automation(self, answers):
        self.start_btn.config(state="disabled", text="준비 중...")
        self.progress_bar.config(mode="indeterminate")
        self.progress_bar.start(10)

        def worker():
            old_input = builtins.input
            seq = iter(answers)

            def fake_input(prompt=""):
                try:
                    val = next(seq)
                    return val
                except StopIteration:
                    return ""

            builtins.input = fake_input

            try:
                import sys
                original_stdout = sys.stdout

                class ProgressMonitor:
                    def __init__(self, original, root, btn):
                        self.original = original
                        self.root     = root
                        self.btn      = btn

                    def write(self, text):
                        self.original.write(text)
                        if any(k in text for k in
                               ["입력", "생성", "업로드", "처리", "완료", "시작"]):
                            msg = text.strip()[:35]
                            if msg:
                                self.root.after(0, lambda m=msg: self.btn.config(text=m))

                    def flush(self):
                        self.original.flush()

                sys.stdout = ProgressMonitor(original_stdout, self.root, self.start_btn)
                try:
                    import eco_input as _eco  # lazy import
                    _eco.main()
                finally:
                    sys.stdout = original_stdout

            except Exception as e:
                import traceback
                traceback.print_exc()
                self.root.after(0, lambda: messagebox.showerror("오류", str(e), parent=self.root))
            finally:
                builtins.input = old_input
                self.root.after(0, lambda: (
                    self.start_btn.config(state="normal", text="자동 입력 시작"),
                    self.progress_bar.stop()
                ))

        threading.Thread(target=worker, daemon=True).start()

    @staticmethod
    def _yn(val) -> str:
        """BooleanVar 또는 bool → '예'/'아니오'"""
        v = val.get() if hasattr(val, "get") else val
        return "예" if v else "아니오"

    def run(self):
        self.root.mainloop()


def main():
    app = EcoInputGUI()
    app.run()


if __name__ == "__main__":
    main()
