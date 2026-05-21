"""
build.py — Mac / Windows 공용 빌드 스크립트
실행: python build.py
  Windows → dist/XBRL_ZIP_to_XLSX.exe
  macOS   → dist/XBRL_ZIP_to_XLSX.app
"""
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent
VENV = ROOT / ".venv"
REQ  = ROOT / "requirements.txt"
SPEC = ROOT / "build.spec"

def run(cmd, **kw):
    print("$", " ".join(str(c) for c in cmd))
    subprocess.run(cmd, check=True, **kw)

def main():
    # ── 1. 가상환경 ──────────────────────────────────────────────
    print("\n=== [1/4] Python 확인 ===")
    run([sys.executable, "--version"])

    print("\n=== [2/4] 가상환경 생성 ===")
    if not VENV.exists():
        run([sys.executable, "-m", "venv", str(VENV)])

    py  = VENV / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")
    pip = VENV / ("Scripts/pip.exe"    if sys.platform == "win32" else "bin/pip")

    # ── 2. 의존성 ────────────────────────────────────────────────
    print("\n=== [3/4] 의존성 설치 ===")
    run([str(pip), "install", "--upgrade", "pip"])
    run([str(pip), "install", "-r", str(REQ)])

    # ── 3. PyInstaller ───────────────────────────────────────────
    print("\n=== [4/4] PyInstaller 빌드 ===")
    pyinst = VENV / ("Scripts/pyinstaller.exe" if sys.platform == "win32" else "bin/pyinstaller")
    run([str(pyinst), "--clean", "--noconfirm", str(SPEC)])

    # ── 완료 메시지 ───────────────────────────────────────────────
    print()
    if sys.platform == "win32":
        print("빌드 완료!  dist\\XBRL_ZIP_to_XLSX.exe  를 실행하세요.")
    else:
        print("빌드 완료!  dist/XBRL_ZIP_to_XLSX.app  을 실행하세요.")

if __name__ == "__main__":
    main()
