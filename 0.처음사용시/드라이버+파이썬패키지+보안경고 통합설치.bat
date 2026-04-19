@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title 초기 설치 통합 도구

REM =========================================================
REM 옵션 및 패키지 목록 설정
REM =========================================================
set "USE_USER=0"
set "PACKAGES=selenium webdriver-manager openpyxl pandas pywin32 requests pywinauto pyperclip PyPDF2 pypdf tkinterdnd2 pillow"
set "DRIVER_DIR=C:\chromedriver"

set "PIP_OPTS="
if "%USE_USER%"=="1" set "PIP_OPTS=--user"

echo ==============================================
echo     보안설정 + ChromeDriver + Python 패키지 통합 설치
echo ==============================================
echo.

REM =========================================================
REM 1. Python 실행 명령 찾기
REM =========================================================
set "PY_CMD="

py -3 -c "import sys;print(sys.version)" >nul 2>&1
if !errorlevel! equ 0 (
    set "PY_CMD=py -3"
    goto :PYTHON_FOUND
)

python -c "import sys;print(sys.version)" >nul 2>&1
if !errorlevel! equ 0 (
    set "PY_CMD=python"
    goto :PYTHON_FOUND
)

echo [ERROR] Python을 찾지 못했습니다. 환경 변수(PATH)를 확인해주세요.
pause
exit /b 1

:PYTHON_FOUND
echo [INFO] 사용 Python: %PY_CMD%
echo.

REM =========================================================
REM 2. 보안경고 실행 REG 적용
REM =========================================================
echo [1/5] 보안경고 실행 REG 적용 중...

reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\NAS" /v "Path" /t REG_SZ /d "\\192.168.10.163\" /f >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Excel 신뢰 위치(Path) 적용 실패
    pause
    exit /b 1
)

reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\NAS" /v "AllowSubfolders" /t REG_DWORD /d 1 /f >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Excel 신뢰 위치(AllowSubfolders) 적용 실패
    pause
    exit /b 1
)

reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\NAS" /v "Description" /t REG_SZ /d "Trusted NAS Root Location" /f >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Excel 신뢰 위치(Description) 적용 실패
    pause
    exit /b 1
)

reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\192.168.10.163" /v "*" /t REG_DWORD /d 1 /f >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] 로컬 인트라넷 영역 등록 실패
    pause
    exit /b 1
)

reg add "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v "BlockContentExecutionFromInternet" /t REG_DWORD /d 0 /f >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Excel 인터넷 차단 설정 적용 실패
    pause
    exit /b 1
)

echo [INFO] REG 적용 완료
echo.

REM =========================================================
REM 3. Chrome 버전 확인
REM =========================================================
echo [2/5] ChromeDriver 설치 준비 중...

set "CHROME_VERSION="
for /f "delims=" %%a in ('powershell -command ^
    "(Get-Item 'C:\Program Files\Google\Chrome\Application\chrome.exe').VersionInfo.ProductVersion" 2^>nul') do set CHROME_VERSION=%%a

if "%CHROME_VERSION%"=="" (
    for /f "delims=" %%a in ('powershell -command ^
        "(Get-Item 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe').VersionInfo.ProductVersion" 2^>nul') do set CHROME_VERSION=%%a
)

if "%CHROME_VERSION%"=="" (
    echo [ERROR] Chrome 설치 경로를 찾지 못했습니다.
    echo Chrome 설치 후 다시 실행해주세요.
    pause
    exit /b 1
)

echo [INFO] Chrome 버전: %CHROME_VERSION%
for /f "tokens=1 delims=." %%a in ("%CHROME_VERSION%") do set MAJOR=%%a
echo [INFO] 메이저 버전: %MAJOR%
echo.

REM =========================================================
REM 4. ChromeDriver 다운로드/설치
REM =========================================================
echo [3/5] ChromeDriver 다운로드 URL 조회 중...

set "DL_URL="
for /f "delims=" %%a in ('powershell -command ^
    "(Invoke-RestMethod https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json).channels.Stable.downloads.chromedriver | Where-Object { $_.platform -eq 'win64' } | Select-Object -ExpandProperty url"') do set DL_URL=%%a

if "%DL_URL%"=="" (
    echo [ERROR] ChromeDriver 다운로드 URL을 가져오지 못했습니다.
    pause
    exit /b 1
)

echo [INFO] 다운로드 URL: %DL_URL%

if not exist "%DRIVER_DIR%" mkdir "%DRIVER_DIR%"

echo [INFO] ChromeDriver 다운로드 중...
powershell -command ^
    "Invoke-WebRequest -Uri '%DL_URL%' -OutFile '%DRIVER_DIR%\driver.zip'"

if not exist "%DRIVER_DIR%\driver.zip" (
    echo [ERROR] ChromeDriver 다운로드 실패
    pause
    exit /b 1
)

echo [INFO] 압축 해제 중...
powershell -command ^
    "Expand-Archive -Force '%DRIVER_DIR%\driver.zip' '%DRIVER_DIR%'"

del "%DRIVER_DIR%\driver.zip"

set "CHROMEDRIVER="
for /r "%DRIVER_DIR%" %%f in (chromedriver.exe) do set CHROMEDRIVER=%%f

if not exist "%CHROMEDRIVER%" (
    echo [ERROR] chromedriver.exe 설치 실패
    pause
    exit /b 1
)

echo [INFO] 설치 완료: %CHROMEDRIVER%

echo [INFO] PATH 등록 확인 중...
echo %PATH% | find /i "%DRIVER_DIR%" >nul
if %errorlevel%==0 (
    echo [INFO] 이미 PATH에 등록되어 있습니다.
) else (
    setx PATH "%PATH%;%DRIVER_DIR%" >nul
    echo [INFO] PATH 등록 완료
)
echo.

REM =========================================================
REM 5. pip 업그레이드 및 패키지 설치
REM =========================================================
echo [4/5] pip 상태 확인 및 업그레이드...
%PY_CMD% -m pip install --upgrade pip %PIP_OPTS%
if !errorlevel! neq 0 (
    echo [ERROR] pip 업그레이드 실패
    pause
    exit /b 1
)
echo.

echo [5/5] 필요한 Python 패키지 확인 및 설치...
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
    echo [INFO] 모든 Python 패키지가 이미 설치되어 있습니다.
) else (
    echo.
    echo [INSTALL] 다음 패키지를 설치합니다:!TO_INSTALL!
    %PY_CMD% -m pip install !TO_INSTALL! %PIP_OPTS%
    if !errorlevel! neq 0 (
        echo [ERROR] 패키지 설치 실패
        pause
        exit /b 1
    )
)

echo.
echo [VERIFY] 최종 설치 리스트 확인...
for %%P in (%PACKAGES%) do (
    echo [PACKAGE] %%P
    %PY_CMD% -m pip show %%P | findstr /i "Name Version"
)

echo.
echo ==============================================
echo     모든 설치가 완료되었습니다.
echo     보안경고 실행 REG도 적용되었습니다.
echo     ChromeDriver PATH 반영은 재실행/재로그인 후 적용될 수 있습니다.
echo ==============================================
pause
endlocal