@echo off
title 통합대시보드 EXE 빌드

set "PYCMD="
where python >nul 2>nul
if not errorlevel 1 (
    set "PYCMD=python"
    goto :python_ready
)
where py >nul 2>nul
if not errorlevel 1 (
    set "PYCMD=py"
    goto :python_ready
)

echo [오류] 이 컴퓨터에 파이썬이 설치되어 있지 않습니다.
echo        run_dashboard.bat을 먼저 한 번 실행해 파이썬을 설치한 뒤 다시 시도해 주세요.
pause
exit /b 1

:python_ready
echo 빌드에 필요한 패키지를 설치/업데이트합니다...
%PYCMD% -m pip install --quiet --upgrade pip
%PYCMD% -m pip install --quiet pandas pywin32 pyinstaller
if errorlevel 1 (
    echo [오류] 패키지 설치에 실패했습니다. 인터넷 연결을 확인해 주세요.
    pause
    exit /b 1
)

echo.
echo PyInstaller로 exe 파일을 빌드합니다. 몇 분 정도 걸릴 수 있습니다...
%PYCMD% -m PyInstaller --noconfirm main_dashboard.spec
if errorlevel 1 (
    echo [오류] 빌드에 실패했습니다. 위 오류 메시지를 확인해 주세요.
    pause
    exit /b 1
)

echo.
echo 빌드 완료! dist 폴더 안의 exe 파일을 확인하세요.
if exist dist start "" explorer dist
pause
