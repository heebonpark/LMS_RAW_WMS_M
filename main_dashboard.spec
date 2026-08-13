# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 빌드 설정 파일.
# 실행 방법: build_exe.bat (Windows에서) 또는 `pyinstaller --noconfirm main_dashboard.spec`
#
# pywin32(win32com)는 win32com.client.Dispatch()로 늦은 바인딩(late-binding)만
# 사용하므로 gen_py 캐시를 따로 번들링할 필요는 없지만, PyInstaller의 정적 분석이
# 놓치기 쉬운 pywintypes/win32timezone은 hiddenimports로 명시합니다.

hiddenimports = [
    'win32timezone',
    'win32com.client',
    'pythoncom',
    'pywintypes',
]

a = Analysis(
    ['main_dashboard.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='이동재고_철거반납_통합대시보드',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,       # GUI 앱이므로 콘솔창 없이 실행. 오류 디버깅 시 True로 바꿔 재빌드하면 콘솔에 예외가 출력됩니다.
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
