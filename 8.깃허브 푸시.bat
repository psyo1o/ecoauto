@echo off
chcp 65001 >nul
title 깃허브 푸시 (GitHub Push)

:: UNC 네트워크 경로 문제 해결을 위해 임시 드라이브 매핑 후 이동
pushd "%~dp0"

echo =======================================
echo         GitHub 자동 푸시 프로그램
echo =======================================
echo.

:: 현재 깃 상태 보여주기
git status
echo.

:: 커밋 메시지 입력받기 (그냥 엔터치면 현재 시간으로 자동 커밋)
set /p COMMIT_MSG="커밋 메시지를 입력하세요 (엔터 누르면 자동 커밋): "
if "%COMMIT_MSG%"=="" set COMMIT_MSG=자동 커밋: %date% %time%

echo.
echo [1/3] 변경사항 추가 중... (git add .)
git add .

echo.
echo [2/3] 커밋 중... (git commit)
git commit -m "%COMMIT_MSG%"

echo.
echo [3/3] 서버로 푸시 중... (git push origin main)
git push origin main

echo.
echo =======================================
echo               푸시 완료!
echo =======================================
echo.

:: 임시 드라이브 매핑 해제
popd
pause
