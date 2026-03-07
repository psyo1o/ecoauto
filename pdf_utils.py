# -*- coding: utf-8 -*-
"""
PDF 생성 & 병합 통합
"""

import os
from excel_com_utils import get_excel_app

def merge_pdfs(pdf_list: list, output_pdf_path: str) -> bool:
    """여러 PDF 파일을 하나로 병합"""
    try:
        from pypdf import PdfMerger
    except:
        try:
            from PyPDF2 import PdfMerger
        except:
            print("❌ PDF 라이브러리 필요: pip install pypdf")
            return False
    
    try:
        merger = PdfMerger()
        
        for pdf_path in pdf_list:
            if os.path.isfile(pdf_path):
                merger.append(pdf_path)
        
        merger.write(output_pdf_path)
        merger.close()
        
        return True
    
    except Exception as e:
        print(f"❌ PDF 병합 실패: {e}")
        return False


class PDFExporter:
    """엑셀 → PDF 변환"""
    
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        # ✅ 전역 엑셀 앱을 가져오고 워크북을 엽니다.
        self.excel = get_excel_app()
        self.wb = self.excel.Workbooks.Open(excel_path, ReadOnly=True, UpdateLinks=0)
    
    def export_sheet(self, sheet_name: str, output_pdf_path: str) -> bool:
        """단일 시트를 PDF로 내보내기"""
        os.makedirs(os.path.dirname(output_pdf_path), exist_ok=True)
        
        try:
            # 이미 __init__에서 열어둔 wb를 그대로 사용 (불필요한 오픈/클로즈 제거로 속도 향상)
            sheet = self.wb.Worksheets(sheet_name)
        except Exception as e:
            print(f"⚠ 시트를 찾을 수 없음 ({sheet_name}): {e}")
            return False
        
        try:
            if int(sheet.Visible) != -1:
                return False
        except:
            pass
        
        try:
            # 0 = xlTypePDF
            sheet.ExportAsFixedFormat(0, output_pdf_path)
            return os.path.isfile(output_pdf_path)
        except Exception as e:
            print(f"❌ PDF 저장 실패 ({sheet_name}): {e}")
            return False
    
    def close(self):
        """워크북만 닫고, 엑셀 앱 전체 종료(Quit)는 하지 않습니다."""
        try:
            if self.wb is not None:
                self.wb.Close(SaveChanges=False)
        except:
            pass