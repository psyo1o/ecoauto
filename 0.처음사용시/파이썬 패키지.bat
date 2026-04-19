@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title Python 패키지 선별 설치 도구

REM =========================================================
REM 옵션 및 패키지 목록 설정
REM =========================================================
set "USE_USER=0"
set "PACKAGES=selenium webdriver-manager openpyxl pandas pywin32 requests pywinauto pyperclip PyPDF2 pypdf tkinterdnd2 pillow"

set "PIP_OPTS="
if "%USE_USER%"=="1" set "PIP_OPTS=--user"

REM =========================================================
REM Python 실행 명령 찾기
REM =========================================================
set "PY_CMD="

py -3 -c "import sys;print(sys.version)" >nul 2>&1
if !errorlevel! equ 0 (
    set "PY_CMD=py -3"
    goto :FOUND_PYTHON
)

python -c "import sys;print(sys.version)" >nul 2>&1
if !errorlevel! equ 0 (
    set "PY_CMD=python"
    goto :FOUND_PYTHON
)

:NOT_FOUND
echo [ERROR] Python을 찾지 못했습니다. 환경 변수(PATH)를 확인해주세요.
pause
exit /b 1

:FOUND_PYTHON
echo [INFO] 사용 Python: %PY_CMD%
echo.

REM =========================================================
REM 1. pip 업그레이드
REM =========================================================
echo [1/3] pip 상태 확인 및 업그레이드... [cite: 3]
%PY_CMD% -m pip install --upgrade pip %PIP_OPTS%
echo.

REM =========================================================
REM 2. 없는 패키지만 선별하여 설치
REM =========================================================
echo [2/3] 설치가 필요한 패키지 확인 중... [cite: 4]

set "TO_INSTALL="
for %%P in (%PACKAGES%) do (
    %PY_CMD% -m pip show %%P >nul 2>&1
    if !errorlevel! neq 0 (
        echo [MISSING] %%P - 설치 목록에 추가합니다.
        set "TO_INSTALL=!TO_INSTALL! %%P"
    ) else (
        echo [EXIST] %%P - 이미 설치되어 있습니다.
    )
)

if "!TO_INSTALL!"=="" (
    echo.
    echo [INFO] 모든 패키지가 이미 설치되어 있습니다. [cite: 4, 5]
) else (
    echo.
    echo [INSTALL] 다음 패키지를 설치합니다:!TO_INSTALL!
    %PY_CMD% -m pip install !TO_INSTALL! %PIP_OPTS%
    if !errorlevel! neq 0 (
        echo [ERROR] 패키지 설치 실패 [cite: 5]
        pause
        exit /b 1
    )
)
echo.

REM =========================================================
REM 3. 최종 설치 확인
REM =========================================================
echo [3/3] 최종 설치 리스트 확인...
for %%P in (%PACKAGES%) do (
    echo [PACKAGE] %%P
    %PY_CMD% -m pip show %%P | findstr /i "Name Version"
)
echo.
echo [DONE] 모든 작업이 완료되었습니다! [cite: 7]
pause
endlocal