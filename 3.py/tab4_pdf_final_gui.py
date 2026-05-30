# -*- coding: utf-8 -*-
"""
탭4 PDF 최종본 생성 전용 GUI
요구사항:
- 사용자가 성적서 엑셀 파일을 드래그/추가하면 리스트에 표시
- PDF 생성 위치(출력 폴더)는 사용자가 직접 지정
- 시트별 PDF 생성:
  1) "시험항목날짜서명" + "대기시료채취 및 분석일지"  -> 1개의 PDF(업로드1용)
  2) "대기측정기록부" -> 1개의 PDF(업로드2용)
  3) 최종본 PDF: 파일 이름(엑셀 파일명) 그대로, 위 3개(시험항목날짜서명/분석일지/측정기록부) 모두 병합하여 1개로 생성
- 웹 업로드는 하지 않음(생성만)

주의:
- Windows + Excel 설치 + pywin32 필요
- PDF 병합: pypdf 우선, 없으면 PyPDF2 사용
"""

import os
import sys
import time
import tempfile
import threading
import traceback
import tkinter as tk
import shutil
from tkinter import ttk, filedialog, messagebox

# ------------------------------
# Drag & Drop (옵션)
# ------------------------------
HAS_DND = False
try:
    # pip install tkinterdnd2
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    HAS_DND = True
except Exception:
    HAS_DND = False

# ------------------------------
# Excel COM
# ------------------------------
try:
    import pythoncom
    import win32com.client as win32
except Exception as e:
    pythoncom = None
    win32 = None

# ------------------------------
# PDF merger
# ------------------------------
def _get_pdf_merger():
    try:
        from pypdf import PdfMerger  # type: ignore
        return PdfMerger
    except Exception:
        try:
            from PyPDF2 import PdfMerger  # type: ignore
            return PdfMerger
        except Exception:
            return None


def merge_pdfs(pdf_list, out_pdf):
    PdfMerger = _get_pdf_merger()
    if PdfMerger is None:
        raise RuntimeError("PDF 병합 라이브러리(pypdf 또는 PyPDF2)가 필요합니다.  pip install pypdf  를 권장합니다.")

    merger = PdfMerger()
    try:
        for p in pdf_list:
            merger.append(p)
        merger.write(out_pdf)
    finally:
        try:
            merger.close()
        except Exception:
            pass


def _normalize_sheet_name(target: str) -> str:
    return (target or "").replace(" ", "").strip()


def _find_sheet(wb, candidates):
    """
    candidates: ["대기시료채취 및 분석일지", ...]
    - 정확 일치 우선
    - 공백 제거 후 일치 fallback
    """
    names = [sh.Name for sh in wb.Worksheets]
    for c in candidates:
        if c in names:
            return c

    c_norm = [_normalize_sheet_name(c) for c in candidates]
    for n in names:
        nn = _normalize_sheet_name(n)
        if nn in c_norm:
            return n
    return None


class ExcelPdfExporter:
    """Excel COM을 하나만 띄워 재사용 (속도/안정성)"""
    def __init__(self, log):
        self.log = log
        self.excel = None

    def start(self):
        if pythoncom is None or win32 is None:
            raise RuntimeError("pywin32가 필요합니다. (pythoncom/win32com import 실패)")
        pythoncom.CoInitialize()
        self.excel = win32.DispatchEx("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        try:
            self.excel.ScreenUpdating = False
        except Exception:
            pass
        try:
            self.excel.EnableEvents = False
        except Exception:
            pass

    def close(self):
        try:
            if self.excel is not None:
                try:
                    self.excel.CutCopyMode = False
                except Exception:
                    pass
                try:
                    self.excel.Quit()
                except Exception:
                    pass
        finally:
            self.excel = None
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def export_sheet_pdf(self, excel_path: str, sheet_name: str, out_pdf_path: str) -> bool:
        """단일 시트를 PDF로 내보내기"""
        os.makedirs(os.path.dirname(out_pdf_path), exist_ok=True)

        wb = None
        try:
            wb = self.excel.Workbooks.Open(excel_path, ReadOnly=True, UpdateLinks=0)
            try:
                sh = wb.Worksheets(sheet_name)
            except Exception:
                return False

            # 숨김 시트 제외
            try:
                # -1: xlSheetVisible
                if int(sh.Visible) != -1:
                    return False
            except Exception:
                pass

            sh.ExportAsFixedFormat(0, out_pdf_path)  # 0 = xlTypePDF
            return os.path.isfile(out_pdf_path)
        finally:
            try:
                if wb is not None:
                    wb.Close(False)
            except Exception:
                pass

    def build_pdfs_for_workbook(self, excel_path: str, out_dir: str, tmp_dir: str,
                                sheet_cover_cands, sheet_analy_cands, sheet_record_cands):
        base_name = os.path.splitext(os.path.basename(excel_path))[0]
        final_pdf = os.path.join(out_dir, f"{base_name}.pdf")

        wb = None
        try:
            # 1. 엑셀 파일 열기 (한 번만 열어서 모든 작업 처리)
            wb = self.excel.Workbooks.Open(excel_path, ReadOnly=True, UpdateLinks=0)

            # 2. GUI에서 입력받은 후보 이름으로 실제 시트 찾기
            sh_cover  = _find_sheet(wb, list(sheet_cover_cands))   # 선택 (없으면 패스)
            sh_analy  = _find_sheet(wb, list(sheet_analy_cands))   # 필수
            sh_record = _find_sheet(wb, list(sheet_record_cands))  # 필수

            # 필수 시트 체크
            missing = []
            if not sh_analy: missing.append("분석일지")
            if not sh_record: missing.append("측정기록부")
            if missing:
                raise RuntimeError(f"필수 시트를 찾지 못했습니다: {', '.join(missing)}")

            # 3. PDF 조각들을 담을 리스트
            pdf_parts = []
            
            # (1) 시험항목날짜서명: 일치하는 시트가 있을 때만 추출
            if sh_cover:
                p_cover = os.path.join(tmp_dir, f"{base_name}_p1.pdf")
                wb.Worksheets(sh_cover).ExportAsFixedFormat(0, p_cover)
                if os.path.isfile(p_cover):
                    pdf_parts.append(p_cover)

            # (2) 분석일지 추출
            p_analy = os.path.join(tmp_dir, f"{base_name}_p2.pdf")
            wb.Worksheets(sh_analy).ExportAsFixedFormat(0, p_analy)
            pdf_parts.append(p_analy)

            # (3) 측정기록부 추출
            p_record = os.path.join(tmp_dir, f"{base_name}_p3.pdf")
            wb.Worksheets(sh_record).ExportAsFixedFormat(0, p_record)
            pdf_parts.append(p_record)

            # 4. 추출된 조각들을 순서대로 '딱 한 번만' 병합하여 최종본 생성
            merge_pdfs(pdf_parts, final_pdf)

            # 5. 업로드용 파일 생성 (필요한 경우만 참조)
            # 업로드1: [커버(있으면) + 분석일지]
            upload1 = os.path.join(tmp_dir, f"{base_name}__업로드1(분석일지).pdf")
            merge_pdfs(pdf_parts[:-1] if sh_cover else [p_analy], upload1)
            
            # 업로드2: [측정기록부] 단독
            upload2 = os.path.join(tmp_dir, f"{base_name}__업로드2(측정기록부).pdf")
            shutil.copy2(p_record, upload2)

            return {"upload1": upload1, "upload2": upload2, "final": final_pdf}

        finally:
            # 작업이 끝나면 반드시 엑셀 닫기
            if wb is not None:
                try:
                    wb.Close(False)
                except:
                    pass



class AppBase:
    def __init__(self):
        self.files = []
        self.output_dir = ""
        self._stop = False

        # root
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        self.root.title("PDF 최종본 생성기")
        self.root.geometry("860x700")
        self.root.minsize(820, 660)

        self._build_ui()

    def _log(self, msg: str):
        ts = time.strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", line)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(3, weight=1)

        # 1) 파일 리스트 + 버튼
        top = ttk.LabelFrame(outer, text="성적서 엑셀 파일", padding=10)
        top.grid(row=0, column=0, sticky="nsew")
        top.columnconfigure(0, weight=1)
        top.columnconfigure(2, weight=0)

        # 리스트박스
        self.lst = tk.Listbox(top, height=10, selectmode="extended")
        self.lst.grid(row=0, column=0, sticky="nsew")
        top.rowconfigure(0, weight=1)

        # 스크롤
        sb = ttk.Scrollbar(top, orient="vertical", command=self.lst.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.lst.configure(yscrollcommand=sb.set)

        # Drag & Drop 안내/등록
        hint = "드래그&드롭 가능" if HAS_DND else "드래그&드롭 미지원(설치 버튼 또는 파일 추가 버튼 사용)"
        self.lbl_hint = ttk.Label(top, text=f"※ {hint}  | 허용: .xls/.xlsx/.xlsm")
        self.lbl_hint.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # 드래그&드롭 미지원이면 설치 버튼 제공(클릭 시 pip로 설치 시도)
        if not HAS_DND:
            ttk.Button(top, text="드래그&드롭 활성화(설치)", command=self._install_dnd).grid(
                row=1, column=2, sticky="e", padx=(8, 0), pady=(6, 0)
            )

        if HAS_DND:
            self.lst.drop_target_register(DND_FILES)
            self.lst.dnd_bind("<<Drop>>", self._on_drop_files)

        btns = ttk.Frame(top)
        btns.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)
        btns.columnconfigure(2, weight=1)

        ttk.Button(btns, text="파일 추가...", command=self._add_files).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(btns, text="선택 삭제", command=self._remove_selected).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(btns, text="전체 비우기", command=self._clear_all).grid(row=0, column=2, sticky="ew", padx=4)

        # 2) 출력 폴더
        mid = ttk.LabelFrame(outer, text="PDF 저장 위치", padding=10)
        mid.grid(row=1, column=0, sticky="ew", pady=(10, 10))
        mid.columnconfigure(1, weight=1)

        ttk.Label(mid, text="출력 폴더:").grid(row=0, column=0, sticky="w")
        self.ent_out = ttk.Entry(mid)
        self.ent_out.grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(mid, text="폴더 선택...", command=self._pick_output_dir).grid(row=0, column=2, sticky="e")

        
        # 2-1) 시트 이름(사용자 입력)
        mid2 = ttk.LabelFrame(outer, text="시트 이름 설정 (콤마로 여러 후보 입력 가능)", padding=10)
        mid2.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        mid2.columnconfigure(1, weight=1)

        ttk.Label(mid2, text="시험항목날짜서명(필요없으면 아무 글자):").grid(row=0, column=0, sticky="w")
        self.ent_sheet_cover = ttk.Entry(mid2)
        self.ent_sheet_cover.grid(row=0, column=1, sticky="ew", padx=6)
        self.ent_sheet_cover.insert(0, "시험항목날짜서명")

        ttk.Label(mid2, text="분석일지 또는 먼지시료채취기록지:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.ent_sheet_analy = ttk.Entry(mid2)
        self.ent_sheet_analy.grid(row=1, column=1, sticky="ew", padx=6, pady=(6, 0))
        self.ent_sheet_analy.insert(0, "대기시료채취 및 분석일지,대기시료채취및분석일지,분석일지")

        ttk.Label(mid2, text="측정기록부:").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.ent_sheet_record = ttk.Entry(mid2)
        self.ent_sheet_record.grid(row=2, column=1, sticky="ew", padx=6, pady=(6, 0))
        self.ent_sheet_record.insert(0, "대기측정기록부,측정기록부")

# 3) 진행/로그
        bot = ttk.LabelFrame(outer, text="진행 상황", padding=10)
        bot.grid(row=3, column=0, sticky="nsew")
        bot.columnconfigure(0, weight=1)
        bot.rowconfigure(1, weight=1)

        self.progress = ttk.Progressbar(bot, mode="determinate")
        self.progress.grid(row=0, column=0, sticky="ew")

        self.txt_log = tk.Text(bot, height=12, wrap="word", state="disabled")
        self.txt_log.grid(row=1, column=0, sticky="nsew", pady=(8, 0))

        # 4) 실행 버튼
        actions = ttk.Frame(outer)
        actions.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)

        self.btn_run = ttk.Button(actions, text="PDF 생성 시작", command=self._start)
        self.btn_run.grid(row=0, column=0, sticky="ew", padx=4)

        self.btn_quit = ttk.Button(actions, text="닫기", command=self.root.destroy)
        self.btn_quit.grid(row=0, column=1, sticky="ew", padx=4)

    def _parse_drop_files(self, data: str):
        # Windows DND는 중괄호로 감싸진 경로가 올 수 있음
        # 예: "{C:\a b\c.xlsx} {D:\d.xlsm}"
        out = []
        token = ""
        in_brace = False
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
        return out

    def _on_drop_files(self, event):
        paths = self._parse_drop_files(event.data)
        self._add_file_paths(paths)


    def _install_dnd(self):
        """
        드래그&드롭을 위해 tkinterdnd2 설치를 시도합니다.
        설치 성공 후에는 프로그램을 재실행해야 적용됩니다.
        """
        import subprocess
        try:
            self._log("tkinterdnd2 설치 시도: pip install tkinterdnd2")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "tkinterdnd2"])
            messagebox.showinfo("완료", "tkinterdnd2 설치가 완료되었습니다.\n프로그램을 닫고 다시 실행하면 드래그&드롭이 활성화됩니다.")
        except Exception as e:
            messagebox.showerror("실패", f"설치 실패: {e}\n\n수동 설치: pip install tkinterdnd2")
            self._log(f"❌ tkinterdnd2 설치 실패: {e}")


    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="성적서 엑셀 파일 선택",
            filetypes=[
                ("Excel files", "*.xls;*.xlsx;*.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if paths:
            self._add_file_paths(list(paths))

    def _add_file_paths(self, paths):
        added = 0
        for p in paths:
            p = p.strip().strip('"')
            if not p:
                continue
            if not os.path.isfile(p):
                continue
            ext = os.path.splitext(p)[1].lower()
            if ext not in (".xls", ".xlsx", ".xlsm"):
                continue
            if p not in self.files:
                self.files.append(p)
                self.lst.insert("end", p)
                added += 1
        if added:
            self._log(f"파일 {added}개 추가됨. (총 {len(self.files)}개)")
        else:
            self._log("추가된 파일이 없습니다. (중복/확장자/경로 확인)")

    def _remove_selected(self):
        idxs = list(self.lst.curselection())
        if not idxs:
            return
        # 뒤에서부터 지우기
        for i in reversed(idxs):
            p = self.lst.get(i)
            try:
                self.files.remove(p)
            except ValueError:
                pass
            self.lst.delete(i)
        self._log(f"선택 삭제 완료. (총 {len(self.files)}개)")

    def _clear_all(self):
        self.files = []
        self.lst.delete(0, "end")
        self._log("전체 비움 완료.")

    def _pick_output_dir(self):
        d = filedialog.askdirectory(title="PDF 저장 폴더 선택")
        if d:
            self.output_dir = d
            self.ent_out.delete(0, "end")
            self.ent_out.insert(0, d)

    def _sheet_candidates(self, s: str):
        s = (s or '').strip()
        if not s:
            return []
        parts = [p.strip() for p in s.split(',')]
        return [p for p in parts if p]

    def _validate(self):
        if not self.files:
            messagebox.showwarning("확인", "엑셀 파일을 1개 이상 추가해 주세요.")
            return False
        out = self.ent_out.get().strip()
        if not out:
            messagebox.showwarning("확인", "PDF 저장 위치(폴더)를 선택해 주세요.")
            return False
        if not os.path.isdir(out):
            try:
                os.makedirs(out, exist_ok=True)
            except Exception:
                messagebox.showwarning("확인", "PDF 저장 폴더를 만들 수 없습니다. 권한/경로를 확인해 주세요.")
                return False
        self.output_dir = out

        cover = self._sheet_candidates(self.ent_sheet_cover.get())
        analy = self._sheet_candidates(self.ent_sheet_analy.get())
        record = self._sheet_candidates(self.ent_sheet_record.get())
        if not cover or not analy or not record:
            messagebox.showwarning('확인', '시트 이름 3개(커버/분석일지/측정기록부)를 모두 입력해 주세요.')
            return False
        return True

    def _set_running(self, running: bool):
        self.btn_run.configure(state=("disabled" if running else "normal"))
        self.btn_quit.configure(state=("disabled" if running else "normal"))
        # 리스트/버튼 잠금은 간단히 생략(필요하면 추가)

    def _start(self):
        if not self._validate():
            return

        self._set_running(True)
        self._stop = False

        def worker():
            tmp_dir = tempfile.mkdtemp(prefix="tab4_pdf_")
            exporter = ExcelPdfExporter(self._log)
            ok = 0
            fail = 0

            try:
                exporter.start()
                total = len(self.files)
                self.root.after(0, lambda: self._init_progress(total))

                for i, xls in enumerate(self.files, start=1):
                    if self._stop:
                        break

                    base = os.path.basename(xls)
                    self._log(f"({i}/{total}) 처리 시작: {base}")

                    try:
                        cover_cands = self._sheet_candidates(self.ent_sheet_cover.get())
                        analy_cands = self._sheet_candidates(self.ent_sheet_analy.get())
                        record_cands = self._sheet_candidates(self.ent_sheet_record.get())
                        res = exporter.build_pdfs_for_workbook(
                            excel_path=xls,
                            out_dir=self.output_dir,
                            tmp_dir=tmp_dir,
                            sheet_cover_cands=cover_cands,
                            sheet_analy_cands=analy_cands,
                            sheet_record_cands=record_cands
                        )
                        self._log(f"  - 업로드1 생성: {os.path.basename(res['upload1'])}")
                        self._log(f"  - 업로드2 생성: {os.path.basename(res['upload2'])}")
                        self._log(f"  ✅ 최종본 저장: {res['final']}")
                        ok += 1
                    except Exception as e:
                        self._log(f"  ❌ 실패: {base} / {e}")
                        fail += 1

                    self.root.after(0, lambda v=i: self._set_progress(v))

                self._log(f"=== 완료: 성공 {ok} / 실패 {fail} ===")
                if fail:
                    self._log("실패한 파일은 시트명/숨김 여부/엑셀 보호/열린 상태(잠김) 등을 확인해 주세요.")
            except Exception as e:
                self._log("치명적 오류: " + str(e))
                self._log(traceback.format_exc())
            finally:
                try:
                    exporter.close()
                except Exception:
                    pass
                try:
                    shutil.rmtree(tmp_dir, ignore_errors=True)
                except Exception:
                    pass
                self.root.after(0, lambda: self._set_running(False))

        threading.Thread(target=worker, daemon=True).start()

    def _init_progress(self, total):
        self.progress.configure(mode="determinate", maximum=max(1, int(total)), value=0)

    def _set_progress(self, value):
        try:
            self.progress.configure(value=int(value))
        except Exception:
            pass

    def run(self):
        self.root.mainloop()


def main():
    AppBase().run()


if __name__ == "__main__":
    main()
