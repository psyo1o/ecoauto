# -*- coding: utf-8 -*-
"""
측정인 검토 GUI (진행률 표시 개선 버전)
- gui_common 모듈 활용
- 실제 진행률 표시
"""
import warnings
warnings.simplefilter("ignore")
warnings.filterwarnings("ignore")
warnings.showwarning = lambda *args, **kwargs: None
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import builtins

# 공통 모듈 import
from gui_common import LogPanel, set_window_topmost, create_labeled_entry

# 메인 로직 import
import eco_check


class EcoCheckGUI:
    """측정인 검토 GUI 클래스"""
    
    def __init__(self):
        self.root = tk.Tk()
        self._setup_window()
        self._create_widgets()

    def _setup_window(self):
        """윈도우 기본 설정"""
        self.root.title("측정인 검토")
        self.root.geometry("700x550")
        self.root.minsize(650, 500)
        set_window_topmost(self.root)

    def _create_widgets(self):
        """모든 위젯 생성"""
        frame = ttk.Frame(self.root, padding=15)
        frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self._create_login_fields(frame)
        self._create_date_field(frame)
        self._create_team_selection(frame)
        self._create_action_button(frame)      # 버튼을 여기로 이동
        self._create_progress_area(frame)
        self._create_log_area(frame)

        # 그리드 가중치 설정
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(6, weight=1)  # 로그 영역이 확장되도록

    def _create_login_fields(self, parent):
        """로그인 정보 입력 필드 생성"""
        self.entry_id = create_labeled_entry(parent, "ID", row=0, width=28)
        self.entry_pw = create_labeled_entry(parent, "PW", row=1, width=28, show="*")

    def _create_date_field(self, parent):
        """날짜 입력 필드 생성"""
        self.entry_date = create_labeled_entry(parent, "날짜", row=2, width=28)
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        self.entry_date.insert(0, today)

    def _create_team_selection(self, parent):
        """팀 선택 체크박스 생성"""
        ttk.Label(
            parent, 
            text="팀 선택(복수 가능, 미선택=전체)"
        ).grid(row=3, column=0, sticky="w", padx=5, pady=(10, 4))

        team_frame = ttk.Frame(parent)
        team_frame.grid(row=3, column=1, sticky="w", padx=5, pady=(10, 4))

        # 1~5팀 체크박스
        self.team_vars = {}
        for i in range(1, 6):
            var = tk.IntVar(value=0)
            ttk.Checkbutton(
                team_frame, 
                text=f"{i}팀", 
                variable=var
            ).grid(row=0, column=i-1, padx=4)
            self.team_vars[i] = var

    def _create_progress_area(self, parent):
        """진행률 표시 영역 생성 - 간결하게"""
        # 진행률 바만 표시 (레이블은 버튼 텍스트로 대체)
        self.progress_bar = ttk.Progressbar(parent, mode="determinate")
        self.progress_bar.grid(row=5, column=0, columnspan=2, sticky="ew", padx=5, pady=(5, 5))

    def _create_action_button(self, parent):
        """실행 버튼 생성"""
        self.start_btn = ttk.Button(
            parent,
            text="검토 시작",
            command=self._on_start
        )
        self.start_btn.grid(row=4, column=0, columnspan=2, pady=(15, 5))

    def _create_log_area(self, parent):
        """로그 영역 생성"""
        log_frame = ttk.LabelFrame(parent, text="로그", padding=10)
        log_frame.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=5, pady=(5, 5))

        self.log_panel = LogPanel(log_frame, height=14)  # 높이 증가
        self.log_panel.pack(fill="both", expand=True)
        self.log_panel.start_pumping()

    def _on_start(self):
        """검토 시작 버튼 핸들러"""
        # 입력값 수집
        login_id = self.entry_id.get().strip()
        login_pw = self.entry_pw.get().strip()
        day = self.entry_date.get().strip()

        # 선택된 팀 수집
        selected_teams = [str(i) for i in range(1, 6) if self.team_vars[i].get() == 1]
        team_input = ",".join(selected_teams)

        # 입력값 검증
        if not login_id or not login_pw or not day:
            messagebox.showwarning("입력 오류", "ID / PW / 날짜는 필수입니다.")
            return

        try:
            datetime.datetime.strptime(day, "%Y-%m-%d")
        except Exception:
            messagebox.showwarning("날짜 오류", "날짜 형식이 잘못되었습니다.\n(YYYY-MM-DD)")
            return

        # 확인 메시지
        if not messagebox.askyesno(
            "확인",
            f"날짜: {day}\n팀: {team_input if team_input else '전체'}\n\n실행할까요?"
        ):
            return

        # 버튼 상태 변경
        self.start_btn.config(state="disabled", text="실행 중...")
        self.progress_bar.config(mode="indeterminate")
        self.progress_bar.start(10)

        # input() 가로채기용 답변 준비
        answers = [login_id, login_pw, day, day, team_input]

        def worker():
            """백그라운드 작업"""
            old_input = builtins.input
            seq = iter(answers)
            
            def fake_input(prompt=""):
                try:
                    return next(seq)
                except StopIteration:
                    if "직접 입력" in prompt or "로그인 후 엔터" in prompt:
                        self.root.after(0, lambda: messagebox.showinfo(
                            "수동 확인", 
                            f"{prompt}\n\n브라우저에서 처리 후 확인을 누르세요."
                        ))
                    return ""
            
            builtins.input = fake_input
            
            try:
                # 로그 모니터링으로 버튼 텍스트 업데이트
                import sys
                original_stdout = sys.stdout
                
                class ProgressMonitor:
                    def __init__(self, original, root, btn):
                        self.original = original
                        self.root = root
                        self.btn = btn
                    
                    def write(self, text):
                        self.original.write(text)
                        
                        # 진행 상황을 버튼 텍스트로 표시
                        if "팀 데이터" in text or "처리" in text or "저장" in text:
                            msg = text.strip()[:30]  # 최대 30자
                            if msg:
                                self.root.after(0, lambda m=msg: self.btn.config(text=m))
                    
                    def flush(self):
                        self.original.flush()
                
                monitor = ProgressMonitor(original_stdout, self.root, self.start_btn)
                sys.stdout = monitor
                
                try:
                    eco_check.main()
                finally:
                    sys.stdout = original_stdout
                    
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.root.after(0, lambda: messagebox.showerror("오류", str(e)))
            finally:
                builtins.input = old_input
                self.root.after(0, lambda: (
                    self.start_btn.config(state="normal", text="검토 시작"),
                    self.progress_bar.stop()
                ))
        
        threading.Thread(target=worker, daemon=True).start()

    def run(self):
        """GUI 실행"""
        self.root.mainloop()


def main():
    """메인 함수"""
    app = EcoCheckGUI()
    app.run()


if __name__ == "__main__":
    main()
