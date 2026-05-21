# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec — Mac / Windows 공용
  Windows → dist/XBRL_ZIP_to_XLSX.exe
  macOS   → dist/XBRL_ZIP_to_XLSX.app
빌드: python build.py  (또는 pyinstaller --clean build.spec)
"""
import sys
import os
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

block_cipher = None

try:
    dnd_datas = collect_data_files("tkinterdnd2")
    dnd_bins  = collect_dynamic_libs("tkinterdnd2")
except Exception:
    dnd_datas, dnd_bins = [], []

a = Analysis(
    ["main.py"],
    pathex=[os.path.abspath(".")],
    binaries=dnd_bins,
    datas=dnd_datas + [("dart_taxonomy.json", ".")],
    hiddenimports=["tkinterdnd2", "openpyxl"],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
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
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# macOS 전용: .app 번들 생성
if sys.platform == "darwin":
    app = BUNDLE(
        exe,
        name="XBRL_ZIP_to_XLSX.app",
        icon=None,
        bundle_identifier="com.xbrl.zip.to.xlsx",
        info_plist={
            "CFBundleDisplayName": "XBRL ZIP to XLSX",
            "CFBundleShortVersionString": "1.0.0",
            "NSHighResolutionCapable": True,
        },
    )
