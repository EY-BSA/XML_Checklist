# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for XBRL ZIP → XLSX 변환기
빌드:  pyinstaller --clean build.spec
산출:  dist/XBRL_ZIP_to_XLSX.exe  (단일 실행 파일)
"""
import os
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

block_cipher = None

# tkinterdnd2 가 들어있는 tkdnd 네이티브 바이너리/리소스
try:
    dnd_datas = collect_data_files("tkinterdnd2")
    dnd_bins  = collect_dynamic_libs("tkinterdnd2")
except Exception:
    dnd_datas, dnd_bins = [], []


a = Analysis(
    ["main.py"],
    pathex=[os.path.abspath(".")],
    binaries=dnd_bins,
    datas=dnd_datas,
    hiddenimports=["tkinterdnd2", "openpyxl"],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="XBRL_ZIP_to_XLSX",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,         # GUI 모드 (콘솔 창 없음)
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
