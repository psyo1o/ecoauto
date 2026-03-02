# -*- coding: utf-8 -*-
"""
GUI 공통 유틸리티 모듈
- 모든 GUI에서 사용하는 공통 기능을 모음
"""

import sys
import queue
import tkinter as tk
from tkinter.scrolledtext import ScrolledText


class QueueWriter:
    """
    stdout/stderr를 queue로 모아 UI에 표시하는 Writer
    
    사용 예:
        log_q = queue.Queue()
        sys.stdout = QueueWriter(log_q)
    """
    
    def __init__(self, q: queue.Queue):
        self.q = q

    def write(self, msg: str):
        if msg:
            self.q.put(msg)

    def flush(self):
        pass


class LogPanel:
    """
    콘솔 로그를 표시하는 패널 위젯
    
    사용 예:
        log_panel = LogPanel(parent_frame)
        log_panel.pack(fill="both", expand=True)
        log_panel.start_pumping()
    """
    
    def __init__(self, parent, title="로그", height=12):
        self.parent = parent
        self.log_queue = queue.Queue()
        
        # stdout/stderr 리다이렉트
        self._orig_stdout = sys.stdout
        self._orig_stderr = sys.stderr
        sys.stdout = QueueWriter(self.log_queue)
        sys.stderr = QueueWriter(self.log_queue)
        
        # UI 구성
        self.log_text = ScrolledText(parent, height=height)
        self.log_text.pack(fill="both", expand=True)
        self.log_text.configure(state="disabled")
        
        self._pump_id = None

    def start_pumping(self):
        """로그 큐에서 메시지를 읽어 UI에 표시 (주기적 호출)"""
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", msg)
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        
        # 50ms마다 반복
        if self.parent.winfo_exists():
            self._pump_id = self.parent.after(50, self.start_pumping)

    def restore_stdio(self):
        """stdout/stderr 원복"""
        try:
            sys.stdout = self._orig_stdout
            sys.stderr = self._orig_stderr
        except:
            pass

    def pack(self, **kwargs):
        """pack 메서드 위임"""
        return self.log_text.pack(**kwargs)

    def grid(self, **kwargs):
        """grid 메서드 위임"""
        return self.log_text.grid(**kwargs)


class ProgressButton:
    """
    실행 중 상태를 표시하는 버튼
    
    사용 예:
        btn = ProgressButton(parent, text="실행", command=on_click)
        btn.pack()
        
        # 실행 시작
        btn.set_running("실행 중...")
        
        # 실행 완료
        btn.set_normal()
    """
    
    def __init__(self, parent, text="실행", command=None, **kwargs):
        import tkinter.ttk as ttk
        self.original_text = text
        self.button = ttk.Button(parent, text=text, command=command, **kwargs)

    def set_running(self, text="실행 중..."):
        """실행 중 상태로 변경"""
        self.button.config(state="disabled", text=text)

    def set_normal(self):
        """정상 상태로 복귀"""
        self.button.config(state="normal", text=self.original_text)

    def pack(self, **kwargs):
        return self.button.pack(**kwargs)

    def grid(self, **kwargs):
        return self.button.grid(**kwargs)

    def config(self, **kwargs):
        return self.button.config(**kwargs)


def set_window_topmost(window):
    """윈도우를 최상위로 설정"""
    window.attributes('-topmost', True)


def center_window(window, width=None, height=None):
    """윈도우를 화면 중앙에 배치"""
    window.update_idletasks()
    
    if width and height:
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")


def create_labeled_entry(parent, label_text, row, column=0, width=20, show=None, **kwargs):
    """
    레이블과 Entry를 함께 생성하는 헬퍼 함수
    
    Returns:
        tk.Entry: 생성된 Entry 위젯
    """
    import tkinter.ttk as ttk
    
    ttk.Label(parent, text=label_text).grid(
        row=row, column=column, sticky="w", padx=5, pady=6
    )
    
    entry = ttk.Entry(parent, width=width, show=show, **kwargs)
    entry.grid(row=row, column=column+1, sticky="ew", padx=5, pady=6)
    
    return entry


def set_state_recursive(container, state: str):
    """
    컨테이너 안의 모든 위젯의 상태를 재귀적으로 변경
    
    Args:
        container: 대상 컨테이너 위젯
        state: "normal" 또는 "disabled"
    """
    from tkinter import ttk
    
    for widget in container.winfo_children():
        if isinstance(widget, (ttk.Frame, ttk.LabelFrame, tk.Frame, tk.LabelFrame)):
            set_state_recursive(widget, state)
            continue
        
        try:
            widget.configure(state=state)
        except:
            try:
                if state == "disabled":
                    widget.state(["disabled"])
                else:
                    widget.state(["!disabled"])
            except:
                pass
