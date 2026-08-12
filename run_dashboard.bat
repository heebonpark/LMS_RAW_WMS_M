@echo off
title 이동재고 및 철거반납 통합 대시보드

rem ---- 파이썬(또는 py 런처)이 있는지 확인 ----
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

echo [안내] 이 컴퓨터에 파이썬이 설치되어 있지 않아 자동으로 설치를 진행합니다...
echo        ^(인터넷 연결이 필요하며, 몇 분 정도 걸릴 수 있습니다^)
echo.

where winget >nul 2>nul
if errorlevel 1 goto :manual_install

echo winget으로 파이썬을 설치합니다. 잠시만 기다려 주세요...
winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
goto :recheck

:manual_install
echo winget을 사용할 수 없어 python.org에서 설치 파일을 직접 내려받습니다...
set "PY_INSTALLER=%TEMP%\python_installer.exe"
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile '%PY_INSTALLER%' -UseBasicParsing } catch { exit 1 }"
if not exist "%PY_INSTALLER%" (
    echo [오류] 파이썬 설치 파일을 내려받지 못했습니다. 인터넷 연결을 확인하거나
    echo        https://www.python.org 에서 직접 설치한 뒤 이 파일을 다시 실행해 주세요.
    pause
    exit /b 1
)
echo 파이썬을 설치합니다. 잠시만 기다려 주세요...
echo ^(런처 등록을 위해 관리자 권한 창이 뜨면 '예'를 눌러 주세요^)
"%PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1
del "%PY_INSTALLER%" >nul 2>nul

:recheck
where py >nul 2>nul
if not errorlevel 1 (
    set "PYCMD=py"
    goto :install_done
)
where python >nul 2>nul
if not errorlevel 1 (
    set "PYCMD=python"
    goto :install_done
)

echo [오류] 파이썬 설치 후에도 인식되지 않습니다.
echo        이 창을 닫고 run_dashboard.bat을 다시 실행해 주세요.
echo        ^(그래도 안 되면 https://www.python.org 에서 수동 설치해 주세요^)
pause
exit /b 1

:install_done
echo 파이썬 설치가 완료되었습니다.
echo.

:python_ready
echo 필요한 패키지^(pandas, pywin32^) 설치 여부를 확인합니다...
%PYCMD% -c "import pandas, win32com.client" >nul 2>nul
if errorlevel 1 (
    echo 필요한 패키지를 설치합니다. 잠시만 기다려 주세요...
    %PYCMD% -m pip install --quiet --upgrade pip
    %PYCMD% -m pip install --quiet pandas pywin32
    if errorlevel 1 (
        echo [오류] 패키지 설치에 실패했습니다. 인터넷 연결을 확인해 주세요.
        pause
        exit /b 1
    )
)

echo.
echo 이동재고 및 철거반납 통합 대시보드를 실행합니다...
echo.
%PYCMD% main_dashboard.py
pause
