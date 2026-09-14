# 사업장관리번호(site_no) 연동 안내

날짜: 2026-08-25 / 갱신: 2026-09-12

> **수신 매칭 정본:** `\\192.168.10.163\docker\approval_mvp\docs\ECO_INPUT_MATCHING.md`  
> (거래처 우선순위·금지사항은 위 문서. 본 파일은 H6 입력·전송만 요약)

## 사용 방법

- 성적서 「입력」시트 **H6**에 사업장관리번호(마스터 측정시설(지역)관리.xlsx **H열**)를 입력합니다.
- 주 셀은 **입력!H6** (구 K7). 라벨/구버전 폴백: J7·K7 등.
- 그룹웨어 전송 시 `site_no` 필드로 전달되며, **거래처 매칭 최우선**입니다.
- `report_data` / `reports/sync` / `report_pdf` **모두** `site_no`를 넣습니다.

## 매칭 우선순위 (수신과 동일)

1. `site_no` (입력!H6 = 마스터 H) — 있으면 **여기서 끝**
2. 없으면 `company_name` (입력!H7 = 마스터 G 전체)
3. `biz_no`는 보조만 — 복수 공장에서 `site_no` 없으면 위험

## 당분간 유지

- H7 업체명·주소 등은 **수기/기존값 유지** 가능합니다.
- 다음 단계(미구현): 마스터 Excel에서 사업장관리번호로 업체명 등 **자동조회(VBA VLOOKUP)**.
- 이번 패치에서는 입력 셀 + API 전송만 추가했습니다. VBA 자동조회는 포함하지 않습니다.

## 구현 위치

- `3.py/groupware_client.py` — `read_input_sheet_keys` / `facility_match_keys` / `post_report_pdf` Form
- 발신 보조 스펙: `그룹웨어_연동_시설매칭_스펙.md`
