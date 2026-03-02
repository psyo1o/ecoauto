@echo off
chcp 65001 >nul
title ChromeDriver 자동 설치기 (PowerShell 버전 감지)

echo ==============================================
echo     ChromeDriver 자동 다운로드 & 설치
echo ==============================================

:: ------------------------------------------------
:: PowerShell 로 Chrome 버전 가져오기
:: ------------------------------------------------
for /f "delims=" %%a in ('powershell -command ^
    "(Get-Item 'C:\Program Files\Google\Chrome\Application\chrome.exe').VersionInfo.ProductVersion" 2^>nul') do set CHROME_VERSION=%%a

if "%CHROME_VERSION%"=="" (
    for /f "delims=" %%a in ('powershell -command ^
        "(Get-Item 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe').VersionInfo.ProductVersion" 2^>nul') do set CHROME_VERSION=%%a
)

if "%CHROME_VERSION%"=="" (
    echo ❌ Chrome 설치 경로를 찾지 못했습니다.
    echo Chrome 설치 후 다시 실행해주세요.
    pause
    exit /b
)

echo ✔ Chrome 버전: %CHROME_VERSION%

:: 메이저 버전만 추출
for /f "tokens=1 delims=." %%a in ("%CHROME_VERSION%") do set MAJOR=%%a
echo ✔ 메이저 버전: %MAJOR%

:: ------------------------------------------------
:: ChromeDriver 다운로드 URL 가져오기
:: ------------------------------------------------
echo ChromeDriver 최신 버전 정보 조회 중...

for /f "delims=" %%a in ('powershell -command ^
    "(Invoke-RestMethod https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions.json).channels.stable.downloads.chromedriver | ConvertTo-Json"') do set JSON=%%a

:: URL 추출
for /f "delims=" %%a in ('powershell -command ^
    "(Invoke-RestMethod https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json).channels.Stable.downloads.chromedriver | Where-Object { $_.platform -eq 'win64' } | Select-Object -ExpandProperty url"') do set DL_URL=%%a

if "%DL_URL%"=="" (
    echo ❌ ChromeDriver 다운로드 URL을 가져오지 못했습니다.
    pause
    exit /b
)

echo ✔ ChromeDriver 다운로드 URL:
echo %DL_URL%


:: ------------------------------------------------
:: 다운로드 및 설치
:: ------------------------------------------------
set DRIVER_DIR=C:\chromedriver

if not exist "%DRIVER_DIR%" mkdir "%DRIVER_DIR%"

echo ----------------------------------------------
echo ChromeDriver 다운로드 중...

powershell -command ^
    "Invoke-WebRequest -Uri '%DL_URL%' -OutFile '%DRIVER_DIR%\driver.zip'"

if not exist "%DRIVER_DIR%\driver.zip" (
    echo ❌ 다운로드 실패!
    pause
    exit /b
)

echo ✔ 다운로드 완료
echo 압축 해제 중...

powershell -command ^
    "Expand-Archive -Force '%DRIVER_DIR%\driver.zip' '%DRIVER_DIR%'"

del "%DRIVER_DIR%\driver.zip"

for /r "%DRIVER_DIR%" %%f in (chromedriver.exe) do set CHROMEDRIVER=%%f

if not exist "%CHROMEDRIVER%" (
    echo ❌ chromedriver.exe 설치 실패
    pause
    exit /b
)

echo ✔ 설치 완료: %CHROMEDRIVER%

:: ------------------------------------------------
:: PATH 등록
:: ------------------------------------------------
echo PATH 등록 중...

echo %PATH% | find /i "%DRIVER_DIR%" >nul
if %errorlevel%==0 (
    echo ✔ 이미 PATH에 등록됨
) else (
    setx PATH "%PATH%;%DRIVER_DIR%"
    echo ✔ PATH 등록 완료
)

echo ----------------------------------------------
echo     🎉 ChromeDriver 자동 설치 완료!
echo  ▶ PC 재부팅 후 바로 사용 가능합니다.
echo ----------------------------------------------
pause
