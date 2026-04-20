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
    return find_best_matching_file(
        sample_no=sample_no,
        nas_base=nas_base,
        nas_dirs=nas_dirs,
        extensions=(".xlsm", ".xlsx", ".xls"),
        strict=strict,
    )


def find_best_matching_file(sample_no: str,
                            nas_base: str = "",
                            nas_dirs: list = None,
                            extensions: tuple | list = (),
                            strict: bool = True) -> str | None:
    """
    NAS_DIRS 전체에서 sample_no 기준으로 가장 적절한 파일 1개 반환.
    - strict=True  : sample_no 바로 뒤에 숫자가 오면 제외
    - strict=False : sample_no 로 시작하면 모두 허용
    - 확장자 우선순위와 경로 길이 우선순위는 기존 find_excel_for_sample 동작 유지
    """
    if not sample_no:
        return None

    pattern = re.compile(rf"^{re.escape(sample_no)}(?!\d)", re.IGNORECASE)
    lower_sample = sample_no.lower()
    extensions = tuple(ext.lower() for ext in (extensions or ()))
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
                if extensions and not any(lf.endswith(ext) for ext in extensions):
                    continue

                stem = os.path.splitext(f)[0]
                if strict:
                    if not pattern.match(stem):
                        continue
                else:
                    if not stem.lower().startswith(lower_sample):
                        continue

                candidates.append(os.path.join(root, f))

    if not candidates:
        return None

    def _rank(path):
        lower_path = path.lower()
        for index, ext in enumerate(extensions):
            if lower_path.endswith(ext):
                return index
        return len(extensions)

    candidates.sort(key=lambda path: (_rank(path), len(path)))
    return candidates[0]


def is_fugitive_dust_file(file_name: str, keywords: tuple | list = ("비산먼지",)) -> bool:
    """파일명에 '비산먼지'가 포함될 때만 비산먼지 성적서로 판별"""
    if not file_name:
        return False
    base_name = os.path.basename(str(file_name))
    return any(keyword in base_name for keyword in (keywords or ()))


def collect_samples_from_nas(nas_base: str,
                             nas_dirs: list = None,
                             date_str: str = "",
                             team_nos=None,
                             sample_pattern: str = r"^(A\d{6}\d-\d{2})",
                             dust_keywords: tuple | list = ("비산먼지",)) -> list:
    """
    NAS 하위 폴더를 재귀 탐색해 시료 목록을 수집.
    반환 형식: [{"sample": sample_no, "path": full_path, "dust": bool}, ...]

    기존 eco_input.extract_samples_from_nas 동작 유지:
    - 파일명 시작부분의 AYYMMDDT-XX 패턴만 사용
    - date_str 는 YYYY-MM-DD 기준으로 YYMMDD 비교
    - team_nos 가 비어있으면 전체 허용
    - dust 는 파일명에 정확히 '비산먼지' 포함 여부로 판단
    """
    sample_list = []
    pattern = re.compile(sample_pattern)
    date_key = str(date_str or "").replace("-", "")[2:]

    if isinstance(team_nos, (int, str)):
        normalized_team_nos = [str(team_nos)]
    elif team_nos:
        normalized_team_nos = [str(team_no) for team_no in team_nos]
    else:
        normalized_team_nos = []

    for nas_dir in (nas_dirs or []):
        folder = os.path.join(nas_base, nas_dir)
        if not os.path.isdir(folder):
            continue

        for root, dirs, files in os.walk(folder):
            for file_name in files:
                if file_name.startswith("~$"):
                    continue

                match = pattern.match(file_name)
                if not match:
                    continue

                sample_no = match.group(1)
                if date_key and sample_no[1:7] != date_key:
                    continue
                if normalized_team_nos and sample_no[7] not in normalized_team_nos:
                    continue

                sample_list.append({
                    "sample": sample_no,
                    "path": os.path.join(root, file_name),
                    "dust": is_fugitive_dust_file(file_name, dust_keywords),
                })

    return sample_list
