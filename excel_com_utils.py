# -*- coding: utf-8 -*-
"""
Excel COM 객체 전역 관리 매니저
(좀비 프로세스 방지용)
"""
import pythoncom
import win32com.client as win32
import atexit
import win32clipboard

_EXCEL_APP = None

def get_excel_app():
    """
    엑셀 COM 객체를 1번만 띄워 재사용.
    죽었으면 정리 후 재생성.
    """
    global _EXCEL_APP

    if _EXCEL_APP is not None:
        try:
            _ = _EXCEL_APP.Hwnd
            return _EXCEL_APP
        except:
            kill_excel_app()

    pythoncom.CoInitialize()
    app = win32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    

    _EXCEL_APP = app
    atexit.register(kill_excel_app)  # 프로그램 종료 시 안전하게 닫기
    return _EXCEL_APP

# excel_com_utils.py 의 kill_excel_app 함수 부분을 이렇게 수정하세요:

def kill_excel_app():
    """엑셀 프로세스를 확실하게 종료 (저장 묻는 팝업 100% 차단)"""
    global _EXCEL_APP
    if _EXCEL_APP is not None:
        try:
            # 1. 엑셀의 모든 경고/팝업 무시 모드 켬
            _EXCEL_APP.DisplayAlerts = False 
            _EXCEL_APP.CutCopyMode = False
            
            # 2. 열려있는 워크북이 있다면 하나씩 돌면서 '저장 안 함'으로 강제 종료
            for i in range(_EXCEL_APP.Workbooks.Count, 0, -1):
                try:
                    _EXCEL_APP.Workbooks(i).Close(SaveChanges=False)
                except:
                    pass
        except:
            pass
            
        try:
            # 3. 엑셀 앱 자체를 끄기
            _EXCEL_APP.Quit()
        except:
            pass
            
    _EXCEL_APP = None
    try:
        pythoncom.CoUninitialize()
    except:
        pass