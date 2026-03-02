# -*- coding: utf-8 -*-
"""
공통 파일 검색 함수
"""

import os
import re
from typing import List, Optional


class FileSearcher:
    """파일 검색 통합 클래스"""
    
    def __init__(self, base_dirs: List[str]):
        self.base_dirs = base_dirs
    
    def find_by_name_prefix(self, 
                           prefix: str,
                           extensions: List[str] = None,
                           sort_by: str = "mtime",
                           desc: bool = True,
                           max_results: int = None) -> List[str]:
        """파일명이 특정 문자로 시작하는 파일 찾기"""
        candidates = self._search(
            name_pattern=f"^{re.escape(prefix)}",
            extensions=extensions,
            match_mode="regex"
        )
        
        candidates = self._sort_results(candidates, sort_by, desc)
        
        if max_results:
            candidates = candidates[:max_results]
        
        return candidates
    
    def find_by_pattern(self,
                       pattern: str,
                       extensions: List[str] = None,
                       match_mode: str = "regex",
                       sort_by: str = "mtime",
                       desc: bool = True) -> List[str]:
        """정규식 패턴으로 파일 찾기"""
        candidates = self._search(
            name_pattern=pattern,
            extensions=extensions,
            match_mode=match_mode
        )
        return self._sort_results(candidates, sort_by, desc)
    
    def find_latest(self,
                   pattern: str,
                   extensions: List[str] = None,
                   match_mode: str = "regex") -> Optional[str]:
        """조건에 맞는 가장 최신 파일 1개 반환"""
        results = self.find_by_pattern(pattern, extensions, match_mode, "mtime", True)
        return results[0] if results else None
    
    def _search(self,
               name_pattern: str,
               extensions: List[str] = None,
               match_mode: str = "regex") -> List[str]:
        """내부 검색 로직"""
        candidates = []
        
        if match_mode == "regex":
            try:
                pattern = re.compile(name_pattern, re.IGNORECASE)
            except:
                return []
        
        for base_dir in self.base_dirs:
            if not os.path.isdir(base_dir):
                continue
            
            for root, dirs, files in os.walk(base_dir):
                for f in files:
                    if f.startswith("~$"):
                        continue
                    
                    if extensions:
                        if not any(f.lower().endswith(ext.lower()) for ext in extensions):
                            continue
                    
                    stem, ext = os.path.splitext(f)
                    
                    if match_mode == "regex":
                        if not pattern.match(stem):
                            continue
                    elif match_mode == "contains":
                        if name_pattern.lower() not in stem.lower():
                            continue
                    
                    full_path = os.path.join(root, f)
                    candidates.append(full_path)
        
        return candidates
    
    def _sort_results(self, candidates: List[str], sort_by: str, desc: bool) -> List[str]:
        """결과 정렬"""
        if sort_by == "mtime":
            try:
                candidates.sort(key=lambda x: os.path.getmtime(x), reverse=desc)
            except:
                pass
        elif sort_by == "name":
            candidates.sort(key=lambda x: os.path.basename(x), reverse=desc)
        
        return candidates

def find_excel_for_sample(sample_no: str,
                          nas_base: str = "",
                          nas_dirs: list = None,
                          strict: bool = True) -> str | None:
    """
    NAS_DIRS 전체에서 sample_no로 시작하는 엑셀 파일 1개 반환.
    strict=True  : sample_no 바로 뒤에 숫자가 오면 제외 (예: A-01 → A-011 제외)
    strict=False : sample_no로 시작하면 모두 허용, 우선순위 xlsm > xlsx > xls
    ~$ 임시파일 항상 제외.
    """
    import re
    exts = (".xlsm", ".xlsx", ".xls")
    pattern = re.compile(rf"^{re.escape(sample_no)}(?!\d)", re.IGNORECASE)
    candidates = []

    for d in (nas_dirs or []):
        folder = os.path.join(nas_base, d)
        if not os.path.isdir(folder):
            continue
        for root, dirs, files in os.walk(folder):
            for f in files:
                if f.startswith("~$"):
                    continue
                lf = f.lower()
                if not any(lf.endswith(e) for e in exts):
                    continue
                stem = os.path.splitext(f)[0]
                if strict:
                    if not pattern.match(stem):
                        continue
                else:
                    if not stem.lower().startswith(sample_no.lower()):
                        continue
                candidates.append(os.path.join(root, f))

    if not candidates:
        return None

    def _rank(p):
        lp = p.lower()
        if lp.endswith(".xlsm"): return 0
        if lp.endswith(".xlsx"): return 1
        return 2

    candidates.sort(key=lambda p: (_rank(p), len(p)))
    return candidates[0]
