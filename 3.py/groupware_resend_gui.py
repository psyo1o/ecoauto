# -*- coding: utf-8 -*-
"""
그룹웨어 전송 GUI (탭 2개)
  1) 로그 재전송 — 기간 조회 후 실패/soft/PDF 없음 건 재시도
  2) 성적서 직접 전송 — 엑셀만 골라 API 전송 (로그 조회 불필요)
"""
from __future__ import annotations

import datetime
import os
import sys
import threading
import warnings
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

warnings.filterwarnings("ignore", category=UserWarning, module=r"openpyxl")

from data_utils import extract_sample_from_name
from gui_common import LogPanel, create_scrollable_text

HAS_DND = False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    HAS_DND = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    HAS_DND = False


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


class GroupwareResendGUI:
    def __init__(self):
        global HAS_DND
        if HAS_DND:
            try:
                self.root = TkinterDnD.Tk()
            except Exception:
                HAS_DND = False
                self.root = tk.Tk()
        else:
            self.root = tk.Tk()
        self.root.title("그룹웨어 전송 (재전송 · 직접)")
        self.root.geometry("1100x740")
        self.root.minsize(960, 640)

        # 탭1: 로그 재전송
        self.pending: list[dict] = []
        self.report_paths: dict[str, str] = {}  # sample_no -> path
        # 탭2: 직접 전송
        self.direct_paths: dict[str, str] = {}
        self._busy = False

        self._build()
        self._set_default_dates()
        print("▶ 그룹웨어 전송 GUI 준비 완료")
        print("   [탭1] 기간 조회 → 성적서 추가 → 재전송")
        print("   [탭2] 성적서만 추가 → 직접 전송")

    def _set_default_dates(self):
        today = datetime.date.today()
        start = today - datetime.timedelta(days=7)
        self.entry_from.delete(0, "end")
        self.entry_from.insert(0, start.strftime("%Y-%m-%d"))
        self.entry_to.delete(0, "end")
        self.entry_to.insert(0, today.strftime("%Y-%m-%d"))

    def _build(self):
        outer = ttk.Frame(self.root)
        outer.pack(fill="both", expand=True)
        outer.grid_columnconfigure(0, weight=3)
        outer.grid_columnconfigure(1, weight=2)
        outer.grid_rowconfigure(0, weight=1)

        left = ttk.Frame(outer, padding=8)
        left.grid(row=0, column=0, sticky="nsew")
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(0, weight=1)

        right = ttk.Frame(outer, padding=(0, 8, 8, 8))
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(0, weight=1)
        right.grid_columnconfigure(0, weight=1)

        nb = ttk.Notebook(left)
        nb.grid(row=0, column=0, sticky="nsew")

        tab_resend = ttk.Frame(nb, padding=4)
        tab_direct = ttk.Frame(nb, padding=4)
        nb.add(tab_resend, text="  로그 재전송  ")
        nb.add(tab_direct, text="  성적서 직접 전송  ")

        self._build_resend_tab(tab_resend)
        self._build_direct_tab(tab_direct)
        self._build_progress(left)
        self._create_log_area(right)

    def _build_progress(self, parent):
        prog = ttk.LabelFrame(parent, text="진행 상태", padding=8)
        prog.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        prog.grid_columnconfigure(0, weight=1)

        self.progress_var = tk.StringVar(value="대기 중")
        self.lbl_progress = ttk.Label(
            prog,
            textvariable=self.progress_var,
            font=("맑은 고딕", 9, "bold"),
            foreground="darkgreen",
        )
        self.lbl_progress.grid(row=0, column=0, sticky="w")

        self.progress_bar = ttk.Progressbar(prog, mode="determinate", maximum=100)
        self.progress_bar.grid(row=1, column=0, sticky="ew", pady=(4, 0))

    def _set_progress(self, cur: int, total: int, sample_no: str = "", status: str = ""):
        total = max(int(total or 0), 0)
        cur = max(int(cur or 0), 0)
        if total <= 0:
            self.progress_bar.config(mode="indeterminate", maximum=100)
            try:
                self.progress_bar.start(12)
            except Exception:
                pass
            self.progress_var.set(status or "처리 중...")
            return

        try:
            self.progress_bar.stop()
        except Exception:
            pass
        self.progress_bar.config(mode="determinate", maximum=total, value=min(cur, total))
        sno = f" · {sample_no}" if sample_no else ""
        st = status or "처리 중"
        self.progress_var.set(f"{st}{sno}  ({cur}/{total})")

    def _reset_progress(self, msg: str = "대기 중"):
        try:
            self.progress_bar.stop()
        except Exception:
            pass
        self.progress_bar.config(mode="determinate", maximum=100, value=0)
        self.progress_var.set(msg)

    def _progress_cb(self, cur, total, sample_no, status):
        self.root.after(
            0,
            lambda c=cur, t=total, s=sample_no, st=status: self._set_progress(
                c, t, s, st
            ),
        )

    def _create_log_area(self, parent):
        log_frame = ttk.LabelFrame(parent, text="로그", padding=10)
        log_frame.grid(row=0, column=0, sticky="nsew")
        self.log_panel = LogPanel(log_frame, height=28)
        self.log_panel.log_text.configure(width=36)
        self.log_panel.pack(fill="both", expand=True)
        self.log_panel.start_pumping()

    # ═══════════════════════════════════════
    # 탭1: 로그 재전송
    # ═══════════════════════════════════════
    def _build_resend_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        hint = ttk.Label(
            parent,
            text="전송로그에서 실패·시설 soft·대기·PDF 없음 건을 모아 재시도합니다. "
            "PDF는 임시 생성 후 API로만 올리고, 0 5.최종완료·0.PDF에는 넣지 않습니다.",
            foreground="gray",
            wraplength=640,
        )
        hint.grid(row=0, column=0, sticky="w", pady=(0, 4))

        body = ttk.Frame(parent)
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(1, weight=1)

        self._build_controls(body)
        self._build_lists(body)
        self._build_actions(body)

    def _build_controls(self, parent):
        top = ttk.LabelFrame(parent, text="조회 조건", padding=8)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        ttk.Label(top, text="기간").grid(row=0, column=0, sticky="w")
        self.entry_from = ttk.Entry(top, width=12)
        self.entry_from.grid(row=0, column=1, padx=(4, 2))
        ttk.Label(top, text="~").grid(row=0, column=2)
        self.entry_to = ttk.Entry(top, width=12)
        self.entry_to.grid(row=0, column=3, padx=(2, 8))

        self.soft_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            top, text="시설 soft 포함", variable=self.soft_var
        ).grid(row=0, column=4, padx=4)

        self.pdf_missing_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            top, text="PDF 없음 포함", variable=self.pdf_missing_var
        ).grid(row=0, column=5, padx=4)

        self.btn_scan = ttk.Button(top, text="대상 조회", command=self._on_scan)
        self.btn_scan.grid(row=0, column=6, padx=4)
        self.btn_skip = ttk.Button(
            top, text="선택 스킵", command=self._on_skip_selected, state="disabled"
        )
        self.btn_skip.grid(row=0, column=7, padx=4)

    def _build_lists(self, parent):
        mid = ttk.Panedwindow(parent, orient="horizontal")
        mid.grid(row=1, column=0, sticky="nsew", pady=(0, 6))

        left = ttk.Frame(mid)
        right = ttk.Frame(mid)
        mid.add(left, weight=3)
        mid.add(right, weight=2)

        left.grid_rowconfigure(1, weight=1)
        left.grid_columnconfigure(0, weight=1)
        ttk.Label(left, text="재전송 대상 (전송로그 기간 조회)").grid(
            row=0, column=0, sticky="w"
        )
        cols = ("sample", "company", "api_company", "reason", "facility", "log", "report")
        self.tree = ttk.Treeview(
            left, columns=cols, show="headings", selectmode="extended", height=12
        )
        self.tree.heading("sample", text="시료번호")
        self.tree.heading("company", text="전송업체")
        self.tree.heading("api_company", text="API등록업체")
        self.tree.heading("reason", text="사유")
        self.tree.heading("facility", text="시설명")
        self.tree.heading("log", text="로그")
        self.tree.heading("report", text="성적서")
        self.tree.column("sample", width=108, anchor="w", stretch=False)
        self.tree.column("company", width=118, anchor="w", stretch=False)
        self.tree.column("api_company", width=138, anchor="w", stretch=False)
        self.tree.column("reason", width=72, anchor="w", stretch=False)
        self.tree.column("facility", width=96, anchor="w", stretch=False)
        self.tree.column("log", width=150, anchor="w", stretch=False)
        self.tree.column("report", width=52, anchor="center", stretch=False)
        ys = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=1, column=0, sticky="nsew")
        ys.grid(row=1, column=1, sticky="ns")
        xs.grid(row=2, column=0, sticky="ew")
        self.lbl_count = ttk.Label(left, text="대상 0건")
        self.lbl_count.grid(row=3, column=0, sticky="w", pady=(4, 0))

        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)
        ttk.Label(right, text="성적서 엑셀 (파일명에 시료번호)").grid(
            row=0, column=0, sticky="w"
        )
        txt_frame, self.txt_reports = create_scrollable_text(
            right, width=32, height=10, horizontal=True
        )
        txt_frame.grid(row=1, column=0, sticky="nsew", pady=4)

        btns = ttk.Frame(right)
        btns.grid(row=2, column=0, sticky="ew")
        ttk.Button(btns, text="파일 추가...", command=self._add_reports).pack(
            side="left", padx=(0, 4)
        )
        ttk.Button(btns, text="비우기", command=self._clear_reports).pack(side="left")

        self.lbl_drop = ttk.Label(
            right,
            text="여기에 성적서 드래그&드롭" if HAS_DND else "파일 추가로 성적서 등록",
            anchor="center",
            relief="groove",
        )
        self.lbl_drop.grid(row=3, column=0, sticky="ew", pady=(6, 0), ipady=6)
        if HAS_DND:
            try:
                self.lbl_drop.drop_target_register(DND_FILES)
                self.lbl_drop.dnd_bind("<<Drop>>", self._on_drop_reports)
                self.txt_reports.drop_target_register(DND_FILES)
                self.txt_reports.dnd_bind("<<Drop>>", self._on_drop_reports)
            except Exception:
                pass

    def _build_actions(self, parent):
        action = ttk.Frame(parent)
        action.grid(row=2, column=0, sticky="ew", pady=(0, 4))
        self.btn_resend = ttk.Button(
            action, text="선택 건 재전송", command=self._on_resend, state="disabled"
        )
        self.btn_resend.pack(side="left", padx=(0, 8))
        self.btn_resend_all = ttk.Button(
            action,
            text="성적서 있는 전부 재전송",
            command=self._on_resend_matched,
            state="disabled",
        )
        self.btn_resend_all.pack(side="left")
        ttk.Label(
            action,
            text="성공하면 로그에 '완료' 기록 → 다음 조회에서 빠짐",
            foreground="gray",
        ).pack(side="left", padx=12)

    # ═══════════════════════════════════════
    # 탭2: 성적서 직접 전송
    # ═══════════════════════════════════════
    def _build_direct_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        hint = ttk.Label(
            parent,
            text="전송로그 없이 성적서(.xlsm)만 골라 그룹웨어로 보냅니다. "
            "PDF는 임시 생성 후 API로만 올리고, 0 5.최종완료·0.PDF에는 넣지 않습니다. "
            "기록은 6.그룹웨어전송 폴더에 남습니다.",
            foreground="gray",
            wraplength=640,
        )
        hint.grid(row=0, column=0, sticky="w", pady=(0, 6))

        mid = ttk.Frame(parent)
        mid.grid(row=1, column=0, sticky="nsew")
        mid.grid_columnconfigure(0, weight=1)
        mid.grid_rowconfigure(1, weight=1)

        ttk.Label(mid, text="전송할 성적서 (파일명에 시료번호 포함)").grid(
            row=0, column=0, sticky="w"
        )
        cols = ("sample", "file", "path")
        self.tree_direct = ttk.Treeview(
            mid, columns=cols, show="headings", selectmode="extended", height=16
        )
        self.tree_direct.heading("sample", text="시료번호")
        self.tree_direct.heading("file", text="파일명")
        self.tree_direct.heading("path", text="경로")
        self.tree_direct.column("sample", width=130, anchor="w", stretch=False)
        self.tree_direct.column("file", width=280, anchor="w", stretch=False)
        self.tree_direct.column("path", width=480, anchor="w", stretch=False)
        ys = ttk.Scrollbar(mid, orient="vertical", command=self.tree_direct.yview)
        xs = ttk.Scrollbar(mid, orient="horizontal", command=self.tree_direct.xview)
        self.tree_direct.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree_direct.grid(row=1, column=0, sticky="nsew")
        ys.grid(row=1, column=1, sticky="ns")
        xs.grid(row=2, column=0, sticky="ew")

        self.lbl_direct_count = ttk.Label(mid, text="0건")
        self.lbl_direct_count.grid(row=3, column=0, sticky="w", pady=(4, 0))

        drop = ttk.Label(
            mid,
            text="여기에 성적서 드래그&드롭" if HAS_DND else "아래 버튼으로 파일 추가",
            anchor="center",
            relief="groove",
        )
        drop.grid(row=4, column=0, sticky="ew", pady=(8, 0), ipady=10)
        self.lbl_drop_direct = drop
        if HAS_DND:
            try:
                drop.drop_target_register(DND_FILES)
                drop.dnd_bind("<<Drop>>", self._on_drop_direct)
                self.tree_direct.drop_target_register(DND_FILES)
                self.tree_direct.dnd_bind("<<Drop>>", self._on_drop_direct)
            except Exception:
                pass

        btns = ttk.Frame(parent)
        btns.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(btns, text="파일 추가...", command=self._add_direct).pack(
            side="left", padx=(0, 4)
        )
        ttk.Button(btns, text="선택 제거", command=self._remove_direct_selected).pack(
            side="left", padx=(0, 4)
        )
        ttk.Button(btns, text="비우기", command=self._clear_direct).pack(side="left")
        self.btn_direct_send = ttk.Button(
            btns, text="선택 건 전송", command=self._on_direct_send_selected
        )
        self.btn_direct_send.pack(side="left", padx=(16, 4))
        self.btn_direct_send_all = ttk.Button(
            btns, text="전부 전송", command=self._on_direct_send_all
        )
        self.btn_direct_send_all.pack(side="left")

    # ── 탭1 성적서 ──
    def _add_reports(self):
        paths = filedialog.askopenfilenames(
            parent=self.root,
            title="성적서 엑셀 선택 (재전송용)",
            filetypes=[
                ("Excel", "*.xlsm *.xlsx *.xls"),
                ("All", "*.*"),
            ],
        )
        if paths:
            self._ingest_report_paths(list(paths))

    def _on_drop_reports(self, event):
        paths = _parse_drop_files(getattr(event, "data", "") or "")
        self._ingest_report_paths(paths)

    def _ingest_report_paths(self, paths: list[str]):
        added = 0
        for p in paths:
            p = os.path.normpath(p.strip().strip('"'))
            if not os.path.isfile(p):
                continue
            sno = extract_sample_from_name(p)
            if not sno:
                print(f"⚠ 시료번호 추출 실패: {os.path.basename(p)}")
                continue
            self.report_paths[sno] = p
            added += 1
        self._refresh_report_text()
        self._refresh_tree_report_col()
        print(f"✅ 성적서 등록 {added}건 (매칭 가능 {len(self.report_paths)}건)")

    def _clear_reports(self):
        self.report_paths.clear()
        self._refresh_report_text()
        self._refresh_tree_report_col()
        print("▶ 성적서 목록 비움")

    def _refresh_report_text(self):
        self.txt_reports.delete("1.0", "end")
        lines = [
            f"{sno}  ←  {p}"
            for sno, p in sorted(self.report_paths.items())
        ]
        if lines:
            self.txt_reports.insert("1.0", "\n".join(lines))

    def _refresh_tree_report_col(self):
        report_idx = 6  # sample, company, api_company, reason, facility, log, report
        for iid in self.tree.get_children():
            vals = list(self.tree.item(iid, "values"))
            sno = vals[0]
            if len(vals) <= report_idx:
                vals.extend([""] * (report_idx + 1 - len(vals)))
            vals[report_idx] = "O" if sno in self.report_paths else "X"
            self.tree.item(iid, values=vals)

    # ── 탭2 성적서 ──
    def _add_direct(self):
        paths = filedialog.askopenfilenames(
            parent=self.root,
            title="성적서 엑셀 선택 (직접 전송)",
            filetypes=[
                ("Excel", "*.xlsm *.xlsx *.xls"),
                ("All", "*.*"),
            ],
        )
        if paths:
            self._ingest_direct_paths(list(paths))

    def _on_drop_direct(self, event):
        paths = _parse_drop_files(getattr(event, "data", "") or "")
        self._ingest_direct_paths(paths)

    def _ingest_direct_paths(self, paths: list[str]):
        added = 0
        for p in paths:
            p = os.path.normpath(p.strip().strip('"'))
            if not os.path.isfile(p):
                continue
            sno = extract_sample_from_name(p)
            if not sno:
                print(f"⚠ 시료번호 추출 실패: {os.path.basename(p)}")
                continue
            self.direct_paths[sno] = p
            added += 1
        self._refresh_direct_tree()
        print(f"✅ 직접전송 목록 등록 {added}건 (총 {len(self.direct_paths)}건)")

    def _remove_direct_selected(self):
        for iid in self.tree_direct.selection():
            self.direct_paths.pop(iid, None)
        self._refresh_direct_tree()

    def _clear_direct(self):
        self.direct_paths.clear()
        self._refresh_direct_tree()
        print("▶ 직접전송 목록 비움")

    def _refresh_direct_tree(self):
        self.tree_direct.delete(*self.tree_direct.get_children())
        for sno, p in sorted(self.direct_paths.items()):
            self.tree_direct.insert(
                "",
                "end",
                iid=sno,
                values=(sno, os.path.basename(p), p),
            )
        self.lbl_direct_count.config(text=f"{len(self.direct_paths)}건")

    # ── 탭1 조회 ──
    def _on_scan(self):
        if self._busy:
            return
        d0 = self.entry_from.get().strip()
        d1 = self.entry_to.get().strip()
        soft = bool(self.soft_var.get())
        pdf_missing = bool(self.pdf_missing_var.get())

        self.btn_scan.config(state="disabled")
        self._busy = True
        self._set_progress(0, 0, "", "대상 조회 중...")
        print(
            f"▶ 대상 조회 중... {d0} ~ {d1} "
            f"(시설 soft={'포함' if soft else '제외'}, "
            f"PDF없음={'포함' if pdf_missing else '제외'})"
        )

        def worker():
            try:
                from groupware_client import collect_pending_resends
                from log_utils import start_run_log, stop_run_log

                session = None
                try:
                    session = start_run_log("gw_resend")
                except Exception:
                    pass
                pending = collect_pending_resends(
                    d0,
                    d1,
                    include_facility_soft=soft,
                    include_pdf_missing=pdf_missing,
                )
                self.pending = pending
                self.root.after(0, lambda: self._fill_tree(pending))
                try:
                    stop_run_log(session)
                except Exception:
                    pass
            except Exception as e:
                import traceback
                traceback.print_exc()
                err = str(e)
                self.root.after(
                    0,
                    lambda m=err: messagebox.showerror("조회 오류", m, parent=self.root),
                )
            finally:
                self.root.after(0, self._scan_done)

        threading.Thread(target=worker, daemon=True).start()

    def _scan_done(self):
        self._busy = False
        self.btn_scan.config(state="normal")
        has = bool(self.pending)
        st = "normal" if has else "disabled"
        self.btn_resend.config(state=st)
        self.btn_resend_all.config(state=st)
        self.btn_skip.config(state=st)
        n = len(self.pending or [])
        self._reset_progress(f"조회 완료 · 대상 {n}건" if n else "조회 완료 · 대상 없음")

    def _fill_tree(self, pending: list[dict]):
        self.tree.delete(*self.tree.get_children())
        for p in pending:
            sno = p.get("sample_no", "")
            sent_co = (p.get("company_name") or "").strip()
            api_co = (p.get("api_matched_company") or "").strip()
            if not api_co:
                api_co = "—"
            self.tree.insert(
                "",
                "end",
                iid=sno,
                values=(
                    sno,
                    sent_co,
                    api_co,
                    p.get("reason", ""),
                    p.get("facility_name", ""),
                    os.path.basename(p.get("log_path") or ""),
                    "O" if sno in self.report_paths else "X",
                ),
            )
        self.lbl_count.config(text=f"대상 {len(pending)}건")
        print(f"✅ 조회 완료: {len(pending)}건")

    def _selected_targets(self) -> list[dict]:
        ids = self.tree.selection()
        if not ids:
            return []
        by = {p["sample_no"]: p for p in self.pending}
        return [by[i] for i in ids if i in by]

    def _matched_targets(self) -> list[dict]:
        return [p for p in self.pending if p.get("sample_no") in self.report_paths]

    # ── 탭1 스킵 / 재전송 ──
    def _on_skip_selected(self):
        targets = self._selected_targets()
        if not targets:
            messagebox.showinfo("스킵", "목록에서 시료를 선택하세요.", parent=self.root)
            return
        if not messagebox.askyesno(
            "스킵",
            f"{len(targets)}건을 재전송 대상에서 제외(스킵)할까요?\n"
            "로그에 '스킵'으로 기록됩니다.",
            parent=self.root,
        ):
            return

        def worker():
            from groupware_client import mark_resend_status_in_log, RESEND_STATUS_SKIP

            n = 0
            for t in targets:
                n += mark_resend_status_in_log(
                    t.get("log_path") or "",
                    [t.get("sample_no")],
                    RESEND_STATUS_SKIP,
                    note_append="사용자 스킵",
                )
                print(f"  · 스킵: {t.get('sample_no')}")
            print(f"✅ 스킵 처리 {len(targets)}건 (로그 행 {n})")
            self.root.after(0, self._on_scan)

        threading.Thread(target=worker, daemon=True).start()

    def _on_resend(self):
        targets = self._selected_targets()
        if not targets:
            messagebox.showinfo(
                "재전송", "목록에서 재전송할 시료를 선택하세요.", parent=self.root
            )
            return
        self._run_resend(targets)

    def _on_resend_matched(self):
        targets = self._matched_targets()
        if not targets:
            messagebox.showinfo(
                "재전송",
                "성적서가 매칭된 대상이 없습니다.\n먼저 성적서 엑셀을 추가하세요.",
                parent=self.root,
            )
            return
        self._run_resend(targets)

    def _run_resend(self, targets: list[dict]):
        missing = [
            t["sample_no"]
            for t in targets
            if t.get("sample_no") not in self.report_paths
            and not (
                t.get("source_excel") and os.path.isfile(t.get("source_excel") or "")
            )
        ]
        if missing:
            if not messagebox.askyesno(
                "성적서 부족",
                f"성적서 없는 시료 {len(missing)}건:\n"
                + ", ".join(missing[:8])
                + ("…" if len(missing) > 8 else "")
                + "\n\n성적서 있는 건만 진행할까요?",
                parent=self.root,
            ):
                return
            targets = [
                t for t in targets
                if t.get("sample_no") in self.report_paths
                or (
                    t.get("source_excel")
                    and os.path.isfile(t.get("source_excel") or "")
                )
            ]
        if not targets:
            return

        if not messagebox.askyesno(
            "재전송 확인",
            f"{len(targets)}건을 그룹웨어에 재전송할까요?\n"
            "(eco_input과 동일: PDF 생성 + 데이터·PDF 전송)\n"
            "성공 건은 로그에 '완료'로 남고 다음 조회에서 빠집니다.",
            parent=self.root,
        ):
            return

        self._busy = True
        for b in (self.btn_scan, self.btn_resend, self.btn_resend_all, self.btn_skip):
            b.config(state="disabled")
        self._set_direct_busy(True)
        self._set_progress(0, len(targets), "", "재전송 준비")

        report_map = dict(self.report_paths)

        def worker():
            try:
                from groupware_client import resend_pending_batch
                from log_utils import start_run_log, stop_run_log

                session = None
                try:
                    session = start_run_log("gw_resend")
                except Exception:
                    pass
                results = resend_pending_batch(
                    targets,
                    report_excel_map=report_map,
                    mark_done=True,
                    progress_cb=self._progress_cb,
                )
                ok_n = sum(1 for r in results if r.get("ok"))
                soft_n = sum(1 for r in results if r.get("soft_only"))
                ng = [r for r in results if not r.get("ok") and not r.get("soft_only")]
                msg = (
                    f"성공 {ok_n} / 시설 soft 유지 {soft_n} / 실패 {len(ng)} "
                    f"/ 전체 {len(results)}"
                )
                print(f"✅ 재전송 결과: {msg}")
                self.root.after(
                    0,
                    lambda m=msg: self._reset_progress(f"재전송 완료 · {m}"),
                )
                title = "재전송 결과"
                if ng:
                    self.root.after(
                        0,
                        lambda: messagebox.showwarning(title, msg, parent=self.root),
                    )
                else:
                    self.root.after(
                        0,
                        lambda: messagebox.showinfo(title, msg, parent=self.root),
                    )
                try:
                    stop_run_log(session)
                except Exception:
                    pass
            except Exception as e:
                import traceback
                traceback.print_exc()
                err = str(e)
                self.root.after(0, lambda: self._reset_progress("재전송 오류"))
                self.root.after(
                    0,
                    lambda m=err: messagebox.showerror(
                        "재전송 오류", m, parent=self.root
                    ),
                )
            finally:
                self.root.after(0, self._resend_done)

        threading.Thread(target=worker, daemon=True).start()

    def _resend_done(self):
        self._busy = False
        self._set_direct_busy(False)
        self._on_scan()  # 완료 건 빠진 목록으로 다시 조회

    # ── 탭2 직접 전송 ──
    def _set_direct_busy(self, busy: bool):
        st = "disabled" if busy else "normal"
        for b in (self.btn_direct_send, self.btn_direct_send_all):
            try:
                b.config(state=st)
            except Exception:
                pass

    def _on_direct_send_selected(self):
        ids = list(self.tree_direct.selection())
        if not ids:
            messagebox.showinfo(
                "직접 전송", "목록에서 전송할 시료를 선택하세요.", parent=self.root
            )
            return
        paths = [self.direct_paths[i] for i in ids if i in self.direct_paths]
        self._run_direct_send(paths)

    def _on_direct_send_all(self):
        if not self.direct_paths:
            messagebox.showinfo(
                "직접 전송", "먼저 성적서 엑셀을 추가하세요.", parent=self.root
            )
            return
        self._run_direct_send(list(self.direct_paths.values()))

    def _run_direct_send(self, paths: list[str]):
        if not paths:
            return
        if not messagebox.askyesno(
            "직접 전송 확인",
            f"{len(paths)}건을 그룹웨어로 전송할까요?\n"
            "(eco_input과 동일: PDF 생성 + 데이터·PDF 전송)\n"
            "결과는 6.그룹웨어전송 폴더에 기록됩니다.",
            parent=self.root,
        ):
            return

        self._busy = True
        for b in (self.btn_scan, self.btn_resend, self.btn_resend_all, self.btn_skip):
            try:
                b.config(state="disabled")
            except Exception:
                pass
        self._set_direct_busy(True)
        path_list = list(paths)
        self._set_progress(0, len(path_list), "", "직접 전송 준비")

        def worker():
            try:
                from groupware_client import send_from_report_excels
                from log_utils import start_run_log, stop_run_log

                session = None
                try:
                    session = start_run_log("gw_direct")
                except Exception:
                    pass
                results = send_from_report_excels(
                    path_list,
                    write_summary=True,
                    progress_cb=self._progress_cb,
                )
                ok_n = sum(1 for r in results if r.get("ok"))
                soft_n = sum(1 for r in results if r.get("soft_only"))
                ng = [r for r in results if not r.get("ok") and not r.get("soft_only")]
                msg = (
                    f"성공 {ok_n} / 시설 soft {soft_n} / 실패 {len(ng)} "
                    f"/ 전체 {len(results)}"
                )
                print(f"✅ 직접 전송 결과: {msg}")
                self.root.after(
                    0,
                    lambda m=msg: self._reset_progress(f"직접 전송 완료 · {m}"),
                )
                title = "직접 전송 결과"
                if ng:
                    self.root.after(
                        0,
                        lambda: messagebox.showwarning(title, msg, parent=self.root),
                    )
                else:
                    self.root.after(
                        0,
                        lambda: messagebox.showinfo(title, msg, parent=self.root),
                    )
                try:
                    stop_run_log(session)
                except Exception:
                    pass
            except Exception as e:
                import traceback
                traceback.print_exc()
                err = str(e)
                self.root.after(0, lambda: self._reset_progress("직접 전송 오류"))
                self.root.after(
                    0,
                    lambda m=err: messagebox.showerror(
                        "직접 전송 오류", m, parent=self.root
                    ),
                )
            finally:
                self.root.after(0, self._direct_done)

        threading.Thread(target=worker, daemon=True).start()

    def _direct_done(self):
        self._busy = False
        self._set_direct_busy(False)
        has = bool(self.pending)
        st = "normal" if has else "disabled"
        self.btn_scan.config(state="normal")
        self.btn_resend.config(state=st)
        self.btn_resend_all.config(state=st)
        self.btn_skip.config(state=st)

    def run(self):
        self.root.mainloop()


def main():
    # 더블클릭 / 런처 실행 시 작업폴더·모듈경로 고정 (NAS 대응)
    here = os.path.dirname(os.path.abspath(__file__))
    try:
        os.chdir(here)
    except Exception:
        pass
    if here not in sys.path:
        sys.path.insert(0, here)

    try:
        GroupwareResendGUI().run()
    except Exception as e:
        import traceback
        traceback.print_exc()
        try:
            import tkinter as _tk
            from tkinter import messagebox as _mb
            r = _tk.Tk()
            r.withdraw()
            _mb.showerror("그룹웨어 전송", f"실행 오류:\n{e}")
            r.destroy()
        except Exception:
            try:
                input(f"\n실행 오류: {e}\nEnter 키를 누르면 종료...")
            except EOFError:
                pass
        raise SystemExit(1)


if __name__ == "__main__":
    main()
