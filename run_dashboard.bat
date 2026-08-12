@echo off
chcp 65001 >nul
title 이동재고 및 철거반납 통합 대시보드

where python >nul 2>nul
if errorlevel 1 (
    echo [오류] 이 컴퓨터에 파이썬이 설치되어 있지 않습니다.
    echo 먼저 파이썬^(3.x^)을 설치하고, pandas / pywin32 패키지를 설치해 주세요.
    echo   예^) pip install pandas pywin32
    pause
    exit /b 1
)

echo 이동재고 및 철거반납 통합 대시보드를 실행합니다...
echo.
python main_dashboard.py
pause
