# dash_v4.py
# 발송대장(Report) + 측정인 검토(eco_check 결과, 요약 시트)들을 합쳐 대시보드 생성
# - eco_check 팀별 결과 파일(여러 개) 선택 가능
# - 장비/인력/차량 "같은 날짜에 여러 팀이 사용" 중복 검사 추가(시간 겹침은 보지 않음)
#
# 출력 시트(기존 그대로): 대시보드 / 상세 / 원인집계

import re
import os
from datetime import datetime, date, time
from typing import List, Tuple, Dict, Any, Optional, Set

import pandas as pd
import openpyxl
import threading
import queue
import tkinter as tk
from tkinter import ttk

from tkinter import Tk
from tkinter.filedialog import askopenfilename, askopenfilenames, asksaveasfilename
from tkinter.messagebox import showerror, showinfo

from data_utils import parse_sn as _parse_sn_raw, parse_time as _parse_time_raw
from log_utils import log_error


SN_RE = re.compile(r'^A\d{7}-\d{2}$')  # 예: A2512063-01  (YYMMDD + 팀1자리)


# 기본 저장 폴더(요청: 네트워크 경로로 자동 저장)
DEFAULT_OUTPUT_DIR = r"\\192.168.10.163\측정팀\10.검토\5.종합 검토"

def _extract_yymmdd_from_sample(sn: Any) -> Optional[str]:
    """시료번호/문자열에서 YYMMDD(6자리) 추출. 예: A2601053-01 -> 260105"""
    if sn is None:
        return None
    s = str(sn).strip()
    m = re.match(r"^[A-Za-z](\d{6})\d?-\d{2}$", s)
    if m:
        return m.group(1)
    m2 = re.search(r"(\d{6})", s)
    return m2.group(1) if m2 else None

def _best_date_tag(send_df: pd.DataFrame, review_paths: List[str]) -> str:
    """저장 파일명에 넣을 날짜 태그(YYMMDD). 우선순위: 발송대장 시료번호 -> 검토파일명 -> 오늘"""
    for col in ["시료번호", "시료 번호", "Sample", "샘플", "SN"]:
        if col in send_df.columns:
            vals = send_df[col].dropna().astype(str).tolist()
            for v in vals[:80]:
                tag = _extract_yymmdd_from_sample(v)
                if tag:
                    return tag
    for p in review_paths:
        tag = _extract_yymmdd_from_sample(os.path.basename(p))
        if tag:
            return tag
    return datetime.now().strftime("%y%m%d")

def _unique_path(path: str) -> str:
    """같은 이름이 있으면 _01, _02...를 붙여서 충돌 방지"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    for i in range(1, 100):
        cand = f"{base}_{i:02d}{ext}"
        if not os.path.exists(cand):
            return cand
    return f"{base}_{int(datetime.now().timestamp())}{ext}"



def _sheet_to_df(wb: openpyxl.Workbook, sheet_name: str) -> pd.DataFrame:
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"'{sheet_name}' 시트를 찾지 못했습니다.")
    ws = wb[sheet_name]
    rows = [[ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            for r in range(1, ws.max_row + 1)]
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    return df


def _parse_sn(sn: str) -> Tuple[Optional[str], Optional[str]]:
    """시료번호에서 날짜(YYYY-MM-DD), 팀 문자열을 추출. data_utils.parse_sn 래퍼."""
    d_obj, team, _ = _parse_sn_raw(str(sn).strip())
    if d_obj is None:
        return None, None
    return d_obj.isoformat(), str(team)


def _parse_time(v: Any) -> Optional[str]:
    """HH:MM 문자열로 정규화. data_utils.parse_time 래퍼."""
    t = _parse_time_raw(v)
    if t is None:
        return None
    return f"{t.hour:02d}:{t.minute:02d}"


def _parse_list(v: Any) -> List[str]:
    """'a, b, c' -> ['a','b','c']"""
    if v is None:
        return []
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return []
    # eco_check는 보통 ", "로 join함
    parts = [p.strip() for p in s.split(",")]
    return [p for p in parts if p]


def read_send_report(send_path: str) -> pd.DataFrame:
    wb = openpyxl.load_workbook(send_path, data_only=True)
    df = _sheet_to_df(wb, "Report")

    if "시료번호" not in df.columns:
        raise ValueError("발송대장 'Report' 시트에 '시료번호' 컬럼이 없습니다.")

    df["시료번호"] = df["시료번호"].astype(str)
    df = df[df["시료번호"].str.match(SN_RE)].copy()

    def send_overall(row):
        def ok(val, okset):
            if val is None:
                return False
            v = str(val).strip()
            # 값 표준화: '사용안함/미실시/해당없음' 등을 '미사용'으로 통일
            if v in {'사용안함', '미실시', '해당없음', '없음'}:
                v = '미사용'
            return v in okset

        if not ok(row.get("성적서상태"), {"OK"}):
            return "NG"
        if not ok(row.get("수분상태"), {"OK", "미사용"}):
            return "NG"
        if not ok(row.get("THC상태"), {"OK", "미사용"}):
            return "NG"
        return "OK"

    df["발송_종합"] = df.apply(send_overall, axis=1)
    return df


def read_review_summary_one(review_path: str) -> pd.DataFrame:
    wb = openpyxl.load_workbook(review_path, data_only=True)
    df = _sheet_to_df(wb, "요약")

    for col in ["시료번호", "항목", "비교"]:
        if col not in df.columns:
            raise ValueError(f"검토파일 '요약' 시트에 '{col}' 컬럼이 없습니다. ({os.path.basename(review_path)})")

    # 사이트값/엑셀값 컬럼은 있을 수도/없을 수도 있음
    if "사이트값" not in df.columns:
        df["사이트값"] = None
    if "엑셀값" not in df.columns:
        df["엑셀값"] = None

    df["비교"] = df["비교"].astype(str).str.replace("확인 불가", "확인불가")
    df["시료번호"] = df["시료번호"].astype(str)
    df["항목"] = df["항목"].astype(str)
    return df


def _agg_compare(series: pd.Series) -> str:
    s = series.astype(str).tolist()
    if any(x == "NG" for x in s):
        return "NG"
    if any(x == "확인불가" for x in s):
        return "확인불가"
    return "OK"


def read_review_summary_multi(review_paths: List[str]) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    return:
      df_norm: (시료번호, 항목, 비교, 사이트값, 엑셀값) - (시료번호,항목) 단위로 정규화/병합
      df_sum : (시료번호별 요약)
      cause_review_ng : (검토 NG 항목 집계)
    """
    dfs = []
    for p in review_paths:
        d = read_review_summary_one(p).copy()
        d["소스파일"] = os.path.basename(p)
        dfs.append(d)
    df_all = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    if df_all.empty:
        raise ValueError("검토 파일(요약)에서 읽어온 데이터가 없습니다.")

    # (시료번호, 항목) 단위로 통합(팀별 파일이 여러개여도 1줄로)
    def first_nonnull(x: pd.Series):
        for v in x.tolist():
            if v is None:
                continue
            s = str(v).strip()
            if s and s.lower() != "nan":
                return v
        return None

    df_norm = (
        df_all
        .groupby(["시료번호", "항목"], as_index=False)
        .agg({
            "비교": _agg_compare,
            "사이트값": first_nonnull,
            "엑셀값": first_nonnull,
        })
    )

    # 시료번호별 요약
    summary_rows = []
    for sn, g in df_norm.groupby("시료번호"):
        vc = g["비교"].value_counts().to_dict()
        ok_cnt = int(vc.get("OK", 0))
        ng_cnt = int(vc.get("NG", 0))
        unk_cnt = int(vc.get("확인불가", 0))
        overall = "NG" if ng_cnt > 0 else ("WARN" if unk_cnt > 0 else "OK")

        ng_items = g.loc[g["비교"] == "NG", "항목"].dropna().astype(str).tolist()
        unk_items = g.loc[g["비교"] == "확인불가", "항목"].dropna().astype(str).tolist()

        summary_rows.append({
            "시료번호": sn,
            "검토_OK": ok_cnt,
            "검토_NG": ng_cnt,
            "검토_확인불가": unk_cnt,
            "검토_종합": overall,
            "NG_항목(상위10)": ", ".join(ng_items[:10]),
            "확인불가_항목(상위10)": ", ".join(unk_items[:10]),
        })

    df_sum = pd.DataFrame(summary_rows)

    # 원인 집계(검토 NG 항목)
    cause_review_ng = df_norm[df_norm["비교"] == "NG"]["항목"].astype(str).value_counts().reset_index()
    cause_review_ng.columns = ["검토_NG항목", "건수"]

    return df_norm, df_sum, cause_review_ng


def build_dup_tables(df_norm: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame,
                                                   Set[str], Set[str], Set[str],
                                                   Set[str], Set[str], Set[str]]:
    """
    같은 날짜에 여러 팀이 같은 장비/인력/차량을 사용했는지 검사.
    - 기본은 '시간 겹침' 기준으로 NG 판정(채취시작/채취끝).
    - 시간 정보가 부족하면 '확인필요(시간미기재)'로만 표기하고 NG에는 포함하지 않음.

    return:
      eq_cases, person_cases, vehicle_cases,
      eq_dup_samples, person_dup_samples, vehicle_dup_samples,
      eq_unknown_samples, person_unknown_samples, vehicle_unknown_samples
    """
    needed = {"시료번호", "항목"}
    if not needed.issubset(df_norm.columns):
        raise ValueError("검토 요약(df_norm) 필수 컬럼이 없습니다.")

    samples = sorted(df_norm["시료번호"].dropna().astype(str).unique().tolist())

    # 시료별 자원/시간 추출
    by_sn: Dict[str, Dict[str, Any]] = {}
    for sn in samples:
        d, team = _parse_sn(sn)
        if not d or not team:
            continue

        def pick(item: str) -> Any:
            sub = df_norm[(df_norm["시료번호"] == sn) & (df_norm["항목"] == item)]
            if sub.empty:
                return None
            v = sub["엑셀값"].iloc[0] if "엑셀값" in sub.columns else None
            if v is None or str(v).strip() == "" or str(v).lower() == "nan":
                v = sub["사이트값"].iloc[0] if "사이트값" in sub.columns else None
            return v

        by_sn[sn] = {
            "날짜": d,
            "팀": team,
            "채취시작": _parse_time(pick("채취시작")),
            "채취끝": _parse_time(pick("채취끝")),
            "장비": _parse_list(pick("장비")),
            "인력": _parse_list(pick("인력")),
            "차량": _parse_list(pick("차량")),
        }

    def to_min(hhmm: Optional[str]) -> Optional[int]:
        if not hhmm:
            return None
        m = re.match(r"^(\d{2}):(\d{2})$", str(hhmm))
        if not m:
            return None
        return int(m.group(1)) * 60 + int(m.group(2))

    def norm_end(s_min: Optional[int], e_min: Optional[int]) -> Optional[int]:
        if s_min is None or e_min is None:
            return None
        # 자정 넘어가는 케이스 방어(끝<시작이면 다음날로 간주)
        if e_min < s_min:
            return e_min + 24 * 60
        return e_min

    def overlap_minutes(a_s: int, a_e: int, b_s: int, b_e: int) -> int:
        s = max(a_s, b_s)
        e = min(a_e, b_e)
        return max(0, e - s)

    def make_dups(key: str) -> Tuple[pd.DataFrame, Set[str], Set[str]]:
        """
        key: '장비'/'인력'/'차량'
        """
        # (날짜, 자원명) -> 사용 리스트
        groups: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
        for sn, info in by_sn.items():
            d = info.get("날짜")
            t = info.get("팀")
            if not d or not t:
                continue

            s = to_min(info.get("채취시작"))
            e = to_min(info.get("채취끝"))
            e = norm_end(s, e)

            for r in info.get(key, []):
                if not r:
                    continue
                k = (d, str(r))
                groups.setdefault(k, []).append({
                    "시료번호": sn,
                    "팀": t,
                    "시작": info.get("채취시작"),
                    "끝": info.get("채취끝"),
                    "s_min": s,
                    "e_min": e,
                })

        rows: List[Dict[str, Any]] = []
        dup_samples: Set[str] = set()
        unk_samples: Set[str] = set()

        for (d, r), items in groups.items():
            # 서로 다른 팀이 2개 이상일 때만 검사/표시
            teams = sorted({it["팀"] for it in items if it.get("팀")})
            if len(teams) < 2:
                continue

            # 정렬(시간 없는 건 뒤로)
            items_sorted = sorted(items, key=lambda x: (x["s_min"] is None, x["s_min"] if x["s_min"] is not None else 10**9))

            # 페어 비교(서로 다른 팀끼리만)
            n = len(items_sorted)
            for i in range(n):
                a = items_sorted[i]
                for j in range(i + 1, n):
                    b = items_sorted[j]
                    if a.get("팀") == b.get("팀"):
                        continue

                    a_s, a_e = a.get("s_min"), a.get("e_min")
                    b_s, b_e = b.get("s_min"), b.get("e_min")

                    # 시간 정보가 부족하면: 확인필요만 기록
                    if a_s is None or a_e is None or b_s is None or b_e is None:
                        rows.append({
                            "날짜": d,
                            "이름": r,
                            "시료A": a.get("시료번호"),
                            "팀A": a.get("팀"),
                            "시작A": a.get("시작"),
                            "끝A": a.get("끝"),
                            "시료B": b.get("시료번호"),
                            "팀B": b.get("팀"),
                            "시작B": b.get("시작"),
                            "끝B": b.get("끝"),
                            "겹침분": None,
                            "판정": "확인필요(시간미기재)",
                        })
                        unk_samples.add(str(a.get("시료번호")))
                        unk_samples.add(str(b.get("시료번호")))
                        continue

                    # 겹침 판정
                    if max(a_s, b_s) < min(a_e, b_e):
                        ov = overlap_minutes(a_s, a_e, b_s, b_e)
                        rows.append({
                            "날짜": d,
                            "이름": r,
                            "시료A": a.get("시료번호"),
                            "팀A": a.get("팀"),
                            "시작A": a.get("시작"),
                            "끝A": a.get("끝"),
                            "시료B": b.get("시료번호"),
                            "팀B": b.get("팀"),
                            "시작B": b.get("시작"),
                            "끝B": b.get("끝"),
                            "겹침분": ov,
                            "판정": "NG(시간겹침)",
                        })
                        dup_samples.add(str(a.get("시료번호")))
                        dup_samples.add(str(b.get("시료번호")))

        df = pd.DataFrame(rows)
        if not df.empty:
            # 보기 좋게 정렬
            df = df.sort_values(by=["날짜", "이름", "판정"], ascending=[True, True, True]).reset_index(drop=True)
        return df, dup_samples, unk_samples

    eq_cases, eq_dup, eq_unk = make_dups("장비")
    person_cases, person_dup, person_unk = make_dups("인력")
    vehicle_cases, vehicle_dup, vehicle_unk = make_dups("차량")

    # 컬럼명 구분(원인집계에서 보기 좋게)
    if not eq_cases.empty:
        eq_cases = eq_cases.rename(columns={"이름": "장비명"})
    if not person_cases.empty:
        person_cases = person_cases.rename(columns={"이름": "인력명"})
    if not vehicle_cases.empty:
        vehicle_cases = vehicle_cases.rename(columns={"이름": "차량"})

    return (eq_cases, person_cases, vehicle_cases,
            eq_dup, person_dup, vehicle_dup,
            eq_unk, person_unk, vehicle_unk)

def build_dashboard(send_df: pd.DataFrame,
                    review_sum: pd.DataFrame,
                    review_ng_causes: pd.DataFrame,
                    eq_cases: pd.DataFrame,
                    person_cases: pd.DataFrame,
                    vehicle_cases: pd.DataFrame,
                    eq_samples: Set[str],
                    person_samples: Set[str],
                    vehicle_samples: Set[str],
                    eq_unknown: Set[str],
                    person_unknown: Set[str],
                    vehicle_unknown: Set[str]) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, pd.DataFrame]]:
    samples1 = set(send_df["시료번호"].astype(str))
    samples2 = set(review_sum["시료번호"].astype(str))
    union = sorted(samples1 | samples2)

    detail = pd.DataFrame({"시료번호": union})

    keep_cols = [c for c in ["시료번호", "대장시트", "SN날짜(YYMMDD)", "대장시작", "대장끝",
                             "성적서상태", "수분상태", "THC상태", "발송_종합"] if c in send_df.columns]
    detail = detail.merge(send_df[keep_cols], on="시료번호", how="left")
    detail = detail.merge(review_sum, on="시료번호", how="left")

    # 중복 결과(검토파일이 있는 시료만 판단 가능)
    def _dup_flag(sn: str, dup_set: Set[str], unk_set: Set[str]) -> str:
        sn = str(sn)
        if sn in dup_set:
            return "NG"
        if sn in unk_set:
            return "확인필요"
        if sn in samples2:  # 검토파일에 존재하지만 중복 아님
            return "OK"
        return ""  # 검토파일이 없으면 판단 불가

    detail["장비중복"] = detail["시료번호"].apply(lambda x: _dup_flag(x, eq_samples, eq_unknown))
    detail["인력중복"] = detail["시료번호"].apply(lambda x: _dup_flag(x, person_samples, person_unknown))
    detail["차량중복"] = detail["시료번호"].apply(lambda x: _dup_flag(x, vehicle_samples, vehicle_unknown))
    detail["중복_종합"] = detail.apply(
        lambda r: "NG" if ("NG" in [r.get("장비중복"), r.get("인력중복"), r.get("차량중복")])
        else ("OK" if (r.get("장비중복") == "OK" and r.get("인력중복") == "OK" and r.get("차량중복") == "OK") else ""),
        axis=1
    )

    def final_overall(row):
        s1 = "WARN" if pd.isna(row.get("발송_종합")) else str(row.get("발송_종합"))
        s2 = "WARN" if pd.isna(row.get("검토_종합")) else str(row.get("검토_종합"))
        s3 = str(row.get("중복_종합")) if row.get("중복_종합") else "OK"  # 중복 판단 불가면 OK로 취급(최종은 따로 확인)
        if "NG" in (s1, s2, s3):
            return "NG"
        if "WARN" in (s1, s2):
            return "WARN"
        return "OK"

    detail["최종_종합"] = detail.apply(final_overall, axis=1)

    summary = pd.DataFrame([
        ["총 시료(Union)", len(union)],
        ["둘 다 존재(매칭)", len(samples1 & samples2)],
        ["발송대장만", len(samples1 - samples2)],
        ["검토요약만", len(samples2 - samples1)],
        ["최종 OK", int((detail["최종_종합"] == "OK").sum())],
        ["최종 WARN", int((detail["최종_종합"] == "WARN").sum())],
        ["최종 NG", int((detail["최종_종합"] == "NG").sum())],
        ["장비중복 NG", int((detail["장비중복"] == "NG").sum())],
        ["인력중복 NG", int((detail["인력중복"] == "NG").sum())],
        ["차량중복 NG", int((detail["차량중복"] == "NG").sum())],
    ], columns=["항목", "값"])

    
    # 원인집계 테이블들
    def _safe_series(df: pd.DataFrame, col: str) -> pd.Series:
        if col in df.columns:
            return df[col]
        return pd.Series([], dtype=str)

    def _vc_table(series: pd.Series, key_name: str) -> pd.DataFrame:
        # pandas 버전에 상관없이 [key_name, 건수] 컬럼으로 안정적으로 생성
        s = series.fillna("미기재").astype(str)
        return s.value_counts().rename_axis(key_name).reset_index(name="건수")

    causes: Dict[str, pd.DataFrame] = {
        "검토_NG항목": review_ng_causes,
        "THC상태": _vc_table(_safe_series(send_df, "THC상태"), "THC상태"),
        "수분상태": _vc_table(_safe_series(send_df, "수분상태"), "수분상태"),
        "성적서상태": _vc_table(_safe_series(send_df, "성적서상태"), "성적서상태"),
    }

    # 중복 원인(자원명별 건수 + 사례)
    if not eq_cases.empty:
        causes["장비중복(장비명별)"] = _vc_table(eq_cases["장비명"], "장비명")
        causes["장비중복(사례)"] = eq_cases
    else:
        causes["장비중복(장비명별)"] = pd.DataFrame(columns=["장비명", "건수"])
        causes["장비중복(사례)"] = pd.DataFrame(columns=["날짜", "장비명", "시료A", "팀A", "시작A", "끝A",
                                                       "시료B", "팀B", "시작B", "끝B", "겹침분", "판정"])

    if not person_cases.empty:
        causes["인력중복(인력명별)"] = _vc_table(person_cases["인력명"], "인력명")
        causes["인력중복(사례)"] = person_cases
    else:
        causes["인력중복(인력명별)"] = pd.DataFrame(columns=["인력명", "건수"])
        causes["인력중복(사례)"] = pd.DataFrame(columns=["날짜", "인력명", "시료A", "팀A", "시작A", "끝A",
                                                       "시료B", "팀B", "시작B", "끝B", "겹침분", "판정"])

    if not vehicle_cases.empty:
        causes["차량중복(차량별)"] = _vc_table(vehicle_cases["차량"], "차량")
        causes["차량중복(사례)"] = vehicle_cases
    else:
        causes["차량중복(차량별)"] = pd.DataFrame(columns=["차량", "건수"])
        causes["차량중복(사례)"] = pd.DataFrame(columns=["날짜", "차량", "시료A", "팀A", "시작A", "끝A",
                                                       "시료B", "팀B", "시작B", "끝B", "겹침분", "판정"])

    return summary, detail, causes



def write_dashboard_excel(out_path: str, summary_df: pd.DataFrame, detail_df: pd.DataFrame, causes: Dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        summary_df.to_excel(w, sheet_name="대시보드", index=False)
        detail_df.to_excel(w, sheet_name="상세", index=False)
        # -------------------------------
        # 원인집계 시트(겹치지 않게 배치)
        #   - 상단: 기존 4블록(검토/THC/수분/성적서) 가로 배치
        #   - 하단: 장비/인력/차량 중복은 세로로 섹션 분리(사례표가 가로로 길어서 겹침 방지)
        # -------------------------------
        top_blocks = [
            ("검토_NG항목", causes["검토_NG항목"], 0),
            ("THC상태", causes["THC상태"], 4),
            ("수분상태", causes["수분상태"], 7),
            ("성적서상태", causes["성적서상태"], 10),
        ]
        for _name, _df, _sc in top_blocks:
            _df.to_excel(w, sheet_name="원인집계", index=False, startcol=_sc, startrow=0)

        ws = w.sheets["원인집계"]
        top_h = max((len(_df) + 1) for _, _df, _ in top_blocks)  # header 포함
        cur_row = top_h + 2  # (0-index) 아래로 2줄 띄우기

        def _write_section(title: str, df_count: pd.DataFrame, df_cases: pd.DataFrame, startrow: int) -> int:
            # 제목(1-index 셀)
            ws.cell(row=startrow + 1, column=1).value = f"[{title}]"

            # 표는 제목 아래로
            df_count.to_excel(w, sheet_name="원인집계", index=False, startcol=0, startrow=startrow + 1)
            startcol_cases = int(df_count.shape[1]) + 2
            df_cases.to_excel(w, sheet_name="원인집계", index=False, startcol=startcol_cases, startrow=startrow + 1)

            h1 = len(df_count) + 1
            h2 = len(df_cases) + 1
            return startrow + max(h1, h2) + 3  # 다음 섹션까지 2줄 띄움

        cur_row = _write_section("장비중복", causes["장비중복(장비명별)"], causes["장비중복(사례)"], cur_row)
        cur_row = _write_section("인력중복", causes["인력중복(인력명별)"], causes["인력중복(사례)"], cur_row)
        cur_row = _write_section("차량중복", causes["차량중복(차량별)"], causes["차량중복(사례)"], cur_row)

#--------------GUI-------------
def pick_dash_inputs_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog
    from tkinter.messagebox import showerror

    result = {"send_path": "", "review_paths": []}

    root = tk.Tk()
    root.title("발송대장/측정인 종합 검토")

    var_send = tk.StringVar()
    var_review = tk.StringVar(value="0개 선택됨")

    def choose_send():
        p = filedialog.askopenfilename(
            title="발송대장 검토 파일 찾기",
            filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xls"), ("All files", "*.*")]
        )
        if p:
            var_send.set(p)

    def choose_review():
        ps = filedialog.askopenfilenames(
            title="측정인 검토 파일 선택(팀별 여러 개 선택 가능)",
            filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xls"), ("All files", "*.*")]
        )
        if ps:
            result["review_paths"] = list(ps)
            var_review.set(f"{len(result['review_paths'])}개 선택됨")

    def start():
        send_path = var_send.get().strip()
        review_paths = result["review_paths"]

        if not send_path:
            showerror("오류", "발송대장 검토 파일을 선택하세요.")
            return
        if not review_paths:
            showerror("오류", "측정인 검토 파일을 1개 이상 선택하세요.")
            return

        result["send_path"] = send_path
        root.destroy()

    frm = ttk.Frame(root, padding=15)
    frm.grid(row=0, column=0, sticky="nsew")
    frm.columnconfigure(1, weight=1)

    ttk.Label(frm, text="대장검토 파일").grid(row=0, column=0, sticky="w", padx=5, pady=6)
    ttk.Entry(frm, textvariable=var_send, width=55).grid(row=0, column=1, sticky="ew", padx=5, pady=6)
    ttk.Button(frm, text="대장파일 찾기", command=choose_send).grid(row=0, column=2, padx=5, pady=6)

    ttk.Label(frm, text="측정인 검토파일(여러개)").grid(row=1, column=0, sticky="w", padx=5, pady=6)
    ttk.Entry(frm, textvariable=var_review, width=55, state="readonly").grid(row=1, column=1, sticky="ew", padx=5, pady=6)
    ttk.Button(frm, text="검토파일 선택", command=choose_review).grid(row=1, column=2, padx=5, pady=6)

    ttk.Button(frm, text="검사 시작", width=20, command=start).grid(row=2, column=0, columnspan=3, pady=16)

    root.mainloop()

    if not result["send_path"] or not result["review_paths"]:
        return None
    return result["send_path"], result["review_paths"]
    
def run_dash_with_progress_gui(send_path: str, review_paths: List[str]):
    """
    무거운 작업을 백그라운드 스레드로 돌리고,
    진행률을 상태 라벨로 표시 (응답없음 방지).
    """
    q = queue.Queue()

    root = tk.Tk()
    root.title("종합 검토 생성 (진행률)")

    status_var = tk.StringVar(value="대기 중…")

    frm = ttk.Frame(root, padding=15)
    frm.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    ttk.Label(frm, text="상태").grid(row=0, column=0, sticky="w", padx=5, pady=6)
    lbl = ttk.Label(frm, textvariable=status_var)
    lbl.grid(row=0, column=1, sticky="ew", padx=5, pady=6)
    frm.columnconfigure(1, weight=1)

    pb = ttk.Progressbar(frm, mode="determinate", maximum=100)
    pb.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=10)

    btn_close = ttk.Button(frm, text="닫기", command=root.destroy, state="disabled")
    btn_close.grid(row=2, column=0, sticky="w", padx=5, pady=10)

    def post_prog(msg: str, pct: int):
        q.put(("PROG", msg, pct))

    def worker():
        try:
            post_prog("1/8 발송대장 읽는 중…", 10)
            send_df = read_send_report(send_path)

            post_prog("2/8 검토 요약 병합 중…", 25)
            df_norm, review_sum, review_ng_causes = read_review_summary_multi(list(review_paths))

            post_prog("3/8 중복(장비/인력/차량) 테이블 생성 중…", 45)
            eq_cases, person_cases, vehicle_cases, eq_samples, person_samples, vehicle_samples, eq_unknown, person_unknown, vehicle_unknown = build_dup_tables(df_norm)

            post_prog("4/8 대시보드 데이터 생성 중…", 65)
            summary_df, detail_df, causes = build_dashboard(
                send_df, review_sum, review_ng_causes,
                eq_cases, person_cases, vehicle_cases,
                eq_samples, person_samples, vehicle_samples,
                eq_unknown, person_unknown, vehicle_unknown
            )

            post_prog("5/8 저장 경로/파일명 결정 중…", 75)
            date_tag = _best_date_tag(send_df, list(review_paths))
            try:
                os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)
            except Exception:
                pass

            auto_path = os.path.join(DEFAULT_OUTPUT_DIR, f"{date_tag} 종합 검토.xlsx")
            out_path = _unique_path(auto_path)

            post_prog("6/8 엑셀 파일 저장 중…", 90)
            try:
                write_dashboard_excel(out_path, summary_df, detail_df, causes)
            except Exception as save_err:
                post_prog("자동 저장 실패 → 저장 위치 선택 필요", 92)
                fallback = asksaveasfilename(
                    title="종합 검토 저장 위치/파일명 선택 (자동 저장 실패)",
                    initialdir=DEFAULT_OUTPUT_DIR,
                    initialfile=os.path.basename(out_path),
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")]
                )
                if not fallback:
                    raise save_err
                out_path = fallback
                write_dashboard_excel(out_path, summary_df, detail_df, causes)

            post_prog("7/8 완료 처리 중…", 98)
            q.put(("OK", out_path))
        except Exception as e:
            q.put(("ERR", str(e)))

    def poll():
        try:
            item = q.get_nowait()
        except queue.Empty:
            root.after(150, poll)
            return

        kind = item[0]

        if kind == "PROG":
            _, msg, pct = item
            status_var.set(msg)
            try:
                pb["value"] = int(pct)
            except Exception:
                pass
            root.after(150, poll)
            return

        if kind == "OK":
            _, out_path = item
            status_var.set("완료!")
            pb["value"] = 100
            btn_close.config(state="normal")
            showinfo("완료", f"종합 검토 생성 완료!\n\n{out_path}")
            return

        if kind == "ERR":
            _, msg = item
            status_var.set("오류")
            btn_close.config(state="normal")
            showerror("오류", msg)
            return

    # 시작
    status_var.set("시작 준비 중…")
    pb["value"] = 0
    t = threading.Thread(target=worker, daemon=True)
    t.start()
    root.after(150, poll)

    root.mainloop()



def main():
    picked = pick_dash_inputs_gui()
    if not picked:
        return
    send_path, review_paths = picked

    # ✅ 무거운 작업은 진행률 GUI에서 스레드로 실행
    run_dash_with_progress_gui(send_path, list(review_paths))



if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log_error("dash.main", e)
        raise