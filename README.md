# XBRL ZIP → XLSX 변환기

DART 원문 XBRL ZIP 파일을 4-시트 xlsx로 변환합니다.  
GUI(드래그앤드롭) 또는 CLI로 실행 가능하며, **Windows / macOS 모두 지원**합니다.

## 폴더 구성

```
zip_cleansing/
├── main.py               # 드래그앤드롭 GUI
├── xbrl_to_xlsx.py       # ZIP → XLSX 변환 코어 (CLI 단독 실행 가능)
├── dart_taxonomy.json    # ifrs-full / dart 표준 concept DataType·Period 조회 테이블
├── build.py              # Mac / Windows 공용 빌드 스크립트
├── build.spec            # PyInstaller 설정
├── requirements.txt
└── README.md
```

---

## 1. 개발 환경에서 바로 실행

```bash
pip install -r requirements.txt
python main.py
```

창이 뜨면 `.zip` 파일을 끌어다 놓거나 `파일 선택…` 버튼을 누르세요.  
변환 결과는 입력 zip과 같은 폴더에 `<zipname>.xlsx`로 저장됩니다.

**CLI로만 실행하고 싶다면:**

```bash
python xbrl_to_xlsx.py "path/to/report.zip"

# 출력 경로 직접 지정
python xbrl_to_xlsx.py "path/to/report.zip" --out "path/to/result.xlsx"
```

---

## 2. 실행파일 빌드 (PyInstaller)

> ⚠️ `build.bat`은 사용하지 마세요. PowerShell 환경에서 오류가 발생합니다.  
> **Windows / Mac 모두 `python build.py`를 사용하세요.**

```bash
python build.py
```

| OS | 결과물 |
|----|--------|
| Windows | `dist\XBRL_ZIP_to_XLSX.exe` |
| macOS | `dist/XBRL_ZIP_to_XLSX.app` |

빌드가 완료되면 실행파일 하나만 배포하면 됩니다. Python 설치 불필요.

---

## 3. 입력 ZIP 안에 있어야 하는 파일

| 파일 패턴 | 용도 |
|-----------|------|
| `*.xsd` | entity 확장 concept 정의 (DataType·Period·Balance) |
| `*_pre.xml` | Presentation linkbase (트리 구조) |
| `*_lab-ko.xml` | 한글 label |
| `*_lab-en.xml` | 영문 label |
| `*.xbrl` | 인스턴스 파일 (선택) — Decimal 값에 사용 |

`.xbrl`이 없으면 Decimal만 비워두고 나머지는 정상 출력됩니다.

---

## 4. 출력 xlsx 시트 구성

| 시트 | 내용 |
|------|------|
| `Concepts` | XSD element 정의 (id, name, type, balance, periodType 등) |
| `Presentation` | Presentation 트리 평탄화 (role, id, label-ko/en, depth, DataType, Period, Decimal) |
| `Label-ko` | concept별 한글 label (12개 role) |
| `Label-en` | concept별 영문 label (12개 role) |

---

## 5. DataType · Period 채우는 방식

| concept 출처 | DataType · Period 소스 |
|-------------|----------------------|
| entity 확장 (ZIP 내 XSD) | ZIP 안 `.xsd` 파일에서 직접 읽음 |
| ifrs-full / dart 표준 | `dart_taxonomy.json` 조회 (8,014개 concept) |

`dart_taxonomy.json`은 `DART_Taxonomy_Search.xlsm`의 Concepts 시트를 추출한 파일입니다.

---

## 6. Decimal 선택 규칙

같은 concept이 여러 context(당기/전기, 연결/별도, member별)로 존재할 수 있으므로  
**당기(`contextRef`가 `C`로 시작)를 우선** 사용하고, 그 중 가장 rounded된 값(최솟값)을 선택합니다.
