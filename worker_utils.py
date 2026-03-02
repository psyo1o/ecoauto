# -*- coding: utf-8 -*-
"""
백그라운드 작업 실행 관련 유틸리티
- GUI에서 무거운 작업을 별도 스레드로 실행
"""

import threading
import traceback
from tkinter import messagebox

try:
    import pythoncom
except ImportError:
    pythoncom = None


class BackgroundWorker:
    """
    백그라운드에서 작업을 실행하는 워커 클래스
    
    사용 예:
        def my_task():
            # 무거운 작업
            process_data()
        
        worker = BackgroundWorker(
            root,
            task=my_task,
            on_success=lambda: print("완료!"),
            on_error=lambda e: print(f"실패: {e}")
        )
        worker.start()
    """
    
    def __init__(self, root, task, on_success=None, on_error=None, on_complete=None, use_com=False):
        """
        Args:
            root: Tkinter root 윈도우
            task: 실행할 함수 (callable)
            on_success: 성공 시 호출할 콜백
            on_error: 실패 시 호출할 콜백 (exception을 인자로 받음)
            on_complete: 성공/실패 여부와 관계없이 완료 시 호출할 콜백
            use_com: Excel COM 사용 여부 (True면 COM 초기화/해제)
        """
        self.root = root
        self.task = task
        self.on_success = on_success
        self.on_error = on_error
        self.on_complete = on_complete
        self.use_com = use_com

    def start(self):
        """백그라운드 작업 시작"""
        thread = threading.Thread(target=self._worker, daemon=True)
        thread.start()

    def _worker(self):
        """실제 작업을 수행하는 워커 함수"""
        error = None
        
        try:
            # Excel COM 사용 시 초기화
            if self.use_com and pythoncom is not None:
                pythoncom.CoInitialize()
            
            # 실제 작업 수행
            self.task()
            
        except Exception as e:
            error = e
            traceback.print_exc()
        
        finally:
            # COM 정리
            if self.use_com and pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass
            
            # UI 스레드에서 콜백 실행
            self._schedule_callbacks(error)

    def _schedule_callbacks(self, error):
        """UI 스레드에서 콜백을 실행하도록 스케줄링"""
        def callbacks():
            if error:
                if self.on_error:
                    self.on_error(error)
            else:
                if self.on_success:
                    self.on_success()
            
            if self.on_complete:
                self.on_complete()
        
        if self.root and self.root.winfo_exists():
            self.root.after(0, callbacks)


class InputInterceptor:
    """
    input() 함수 호출을 가로채서 미리 준비된 답변을 반환
    
    사용 예:
        answers = ["user123", "password", "2024-01-01"]
        
        with InputInterceptor(answers):
            # 이 블록 안에서는 input()이 answers에서 순차적으로 반환
            user = input("ID: ")  # "user123"
            pwd = input("PW: ")   # "password"
    """
    
    def __init__(self, answers, fallback_callback=None):
        """
        Args:
            answers: 미리 준비된 답변 리스트
            fallback_callback: 답변이 부족할 때 호출할 콜백 (prompt를 인자로 받음)
        """
        import builtins
        self.builtins = builtins
        self.answers = iter(answers)
        self.fallback_callback = fallback_callback
        self.original_input = None

    def __enter__(self):
        """컨텍스트 진입: input() 교체"""
        self.original_input = self.builtins.input
        self.builtins.input = self._fake_input
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """컨텍스트 종료: input() 복원"""
        self.builtins.input = self.original_input

    def _fake_input(self, prompt=""):
        """input() 대체 함수"""
        try:
            return next(self.answers)
        except StopIteration:
            # 답변이 부족한 경우
            if self.fallback_callback:
                return self.fallback_callback(prompt)
            return ""


def run_in_background(root, task, success_msg="완료", error_msg="실패", use_com=False):
    """
    간단한 백그라운드 작업 실행 헬퍼 함수
    
    Args:
        root: Tkinter root
        task: 실행할 함수
        success_msg: 성공 메시지
        error_msg: 실패 메시지 (실제 에러는 자동 추가됨)
        use_com: COM 사용 여부
    """
    def on_success():
        messagebox.showinfo("완료", success_msg)
    
    def on_error(e):
        messagebox.showerror("오류", f"{error_msg}\n\n{str(e)}")
    
    worker = BackgroundWorker(
        root,
        task=task,
        on_success=on_success,
        on_error=on_error,
        use_com=use_com
    )
    worker.start()
