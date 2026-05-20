# XBRL ZIP → XLSX 변환기

DART 원문 XBRL ZIP 파일을 드래그앤드롭으로 받아서, `구조내려받기`와 동일한 다중 시트 xlsx 로 변환합니다.
모든 row 의 `DataType`, `Balance`, `Period`, `Decimal`, `Fact` 컬럼이 자동으로 채워집니다.

## 폴더 구성

```
zip_cleansing/
├── main.py            # 드래그앤드롭 GUI
├── xbrl_to_xlsx.py    # ZIP → XLSX 변환 코어 (CLI 단독 실행도 가능)
├── build.spec         # PyInstaller 사양
├── build.bat          # Windows용 빌드 스크립트 (가상환경 + 빌드 한 방)
├── requirements.txt
└── README.md
```

## 1. 개발 환경에서 바로 실행

```bash
python -m pip install -r requirements.txt
python main.py
```

창이 뜨면 `.zip` 파일을 끌어다 놓거나 `파일 선택…` 버튼을 누르세요.
변환 결과는 입력 zip 과 같은 폴더에 `<zipname>.xlsx` 로 저장됩니다.

CLI 만 쓰고 싶다면:

```bash
python xbrl_to_xlsx.py "C:\path\to\report.zip"
# 또는 출력 경로 직접 지정
python xbrl_to_xlsx.py "C:\path\to\report.zip" --out "C:\out\result.xlsx"
```

## 2. Windows .exe 빌드 (PyInstaller)

```cmd
build.bat
```

`.\dist\XBRL_ZIP_to_XLSX.exe` 가 생성됩니다. 이 exe 한 개만 가지고 다녀도 zip 을
드래그하면 같은 폴더에 xlsx 가 만들어집니다.

빌드 명령을 수동으로 돌리고 싶다면:

```cmd
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pyinstaller --clean build.spec
```

> 빌드는 반드시 **Windows** 에서 진행해야 `.exe` 가 만들어집니다.
> macOS / Linux 에서 빌드하면 해당 OS 용 실행파일이 나옵니다.

## 3. 입력 ZIP 안에 있어야 하는 파일

| 패턴 | 설명 |
| ---- | ---- |
| `*.xsd` | concept(taxonomy) 정의 — DataType, Period, Balance 정보 |
| `*_pre.xml` | presentation linkbase — 트리 구조 |
| `*_lab-ko.xml` | 한글 label |
| `*_lab-en.xml` | 영문 label |
| `*.xbrl` | 인스턴스 (선택) — Decimal, Fact 값을 채울 때 사용 |

`.xbrl` 인스턴스가 없으면 Decimal / Fact 컬럼만 비워두고 나머지는 정상 출력됩니다.

## 4. 출력 xlsx 시트 구성

- `기본정보` — 회사명, 회계기간, 작성일 등
- `재무상태표 (연결)` / `재무상태표`
- `포괄손익계산서 (연결)` / `포괄손익계산서`
- `자본변동표 (연결)` / `자본변동표`
- `현금흐름표 (연결)` / `현금흐름표`
- 각 주석 role 별 시트 (`1. 회사의 개요 (연결)`, … )

각 시트의 본문 컬럼:

```
구분 | Prefix | Name | Label(KO) | Label(EN) | Label Role | DataType | Balance | Period | Decimal | Fact
```

## 5. Decimal 채우는 규칙

한 concept 이 여러 컨텍스트(당기/전기, 연결/별도, member 별)로 출현할 수 있으므로
**당기(`contextRef` 가 `C` 로 시작) 를 우선** 사용하고, 없으면 첫 fact 의 값을 사용합니다.
값(Fact) 도 같은 기준으로 함께 표시됩니다.
