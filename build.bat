@echo off
REM ===================================================================
REM  XBRL ZIP -> XLSX 변환기  Windows 빌드 스크립트
REM  실행하면 가상환경을 만들고, 의존성을 설치한 뒤 PyInstaller 로
REM  dist\XBRL_ZIP_to_XLSX.exe 를 생성합니다.
REM ===================================================================
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
chcp 65001 >NUL
cd /d "%~dp0"

echo.
echo === [1/4] Python 확인 ===
where python >NUL 2>NUL
if errorlevel 1 (
    echo [ERROR] Python 이 PATH 에 없습니다. https://www.python.org 에서 설치 후 다시 실행하세요.
    pause
    exit /b 1
)
python --version

echo.
echo === [2/4] 가상환경 생성 / 활성화 ===
if not exist ".venv" (
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] 가상환경 생성 실패
        pause
        exit /b 1
    )
)
call ".venv\Scripts\activate.bat"

echo.
echo === [3/4] 의존성 설치 ===
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] 의존성 설치 실패
    pause
    exit /b 1
)

echo.
echo === [4/4] PyInstaller 빌드 ===
pyinstaller --clean --noconfirm build.spec
if errorlevel 1 (
    echo [ERROR] PyInstaller 빌드 실패
    pause
    exit /b 1
)

echo.
echo ====================================================================
echo  빌드 완료!  dist\XBRL_ZIP_to_XLSX.exe  를 실행하세요.
echo ====================================================================
echo.
pause
endlocal
