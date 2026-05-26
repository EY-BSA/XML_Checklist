# XBRL ZIP → XLSX 변환기

DART 원문 XBRL ZIP 파일을 Presentation 1-시트 xlsx로 변환합니다.  
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

## 1. 전체 데이터 흐름

ZIP 안의 파일들은 모두 **concept id**를 공통 키로 연결됩니다.

**① concept 정의 수집**  
`.xsd`를 파싱해 concept별 DataType·Period·Balance·substitutionGroup을 읽습니다.  
`ifrs-full` / `dart` 표준 concept은 XSD에 없으므로 `dart_taxonomy.json`으로 보완합니다.

**② label 수집**  
`*_lab-ko.xml` / `*_lab-en.xml`을 파싱해 concept별 한글·영문 label을 읽습니다.  
하나의 concept에 label role(표준·terseLabel·totalLabel 등)별로 여러 label이 존재할 수 있습니다.

**③ Presentation 트리 구성**  
`*_pre.xml`을 DFS로 순회하며 각 row를 생성합니다.  
concept id가 `Table`로 끝나는 노드를 기준으로 `table` 컬럼을 채우고,  
②의 label에서 `preferredLabel` 기준으로 표시할 label-ko / label-en을 선택합니다.

**④ Decimal 수집**  
`*.xbrl` 인스턴스에서 fact별 `decimals` 값을 읽습니다.  
당기 context(`C`로 시작) 중 가장 rounded된 값(최솟값)을 선택합니다.

**⑤ 최종 조합 → xlsx 출력**  
③의 Presentation rows에 ①의 DataType·Period와 ④의 Decimal을 합쳐 Presentation 시트를 작성합니다.  
후처리로 `order`가 없는 행과 특정 label-ko 행을 제거합니다.

| 파일 | 연결 키 | 제공하는 정보 |
|------|---------|--------------|
| `.xsd` | concept id | DataType, Period, Balance, substitutionGroup |
| `dart_taxonomy.json` | concept id | ifrs-full/dart 표준 concept의 동일 정보 |
| `*_lab-ko/en.xml` | concept id | 한글/영문 label (role별) |
| `*_pre.xml` | concept id | 트리 계층 구조, 표시 순서, preferredLabel, table 구분 |
| `*.xbrl` | concept local name | Decimal 값 |

---

## 2. 출력 xlsx 시트 구성

출력 파일은 **Presentation 시트 1장**으로 구성됩니다.

| 컬럼 | 소스 | 설명 |
|------|------|------|
| role | `*_pre.xml` | 재무제표 구분 (예: `[D210000] 재무상태표, 유동/비유동법 - 연결`) |
| table | `*_pre.xml` | concept id가 `Table`로 끝나는 노드 기준 테이블명 (role 경계에서 초기화) |
| id | `*_pre.xml` | concept id |
| label-ko | `*_lab-ko.xml` | preferredLabel 기준 한글 label |
| label-en | `*_lab-en.xml` | preferredLabel 기준 영문 label |
| depth | `*_pre.xml` | 트리 깊이 (0 = root) |
| order | `*_pre.xml` | 형제 노드 간 표시 순서 |
| pref_label | `*_pre.xml` | presentationArc의 preferredLabel role |
| DataType | `.xsd` / `dart_taxonomy.json` | monetaryItemType, stringItemType 등 |
| Period | `.xsd` / `dart_taxonomy.json` | INSTANT / DURATION |
| Decimal | `*.xbrl` | 당기 context 중 가장 rounded된 decimals 값 |

---

## 3. 개발 환경에서 바로 실행

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

## 4. 실행파일 빌드 (PyInstaller)

> ⚠️ `build.bat`은 사용하지 마세요. PowerShell 환경에서 오류가 발생합니다.  
> **Windows / Mac 모두 `python build.py`를 사용하세요.**

```bash
python build.py
```

| OS | 결과물 |
|----|--------|
| Windows | `dist\XBRL_ZIP_to_XLSX.exe` |
| macOS | `dist/XBRL_ZIP_to_XLSX.app` |

- 빌드가 완료되면 실행파일 하나만 배포하면 됩니다. Python 설치 불필요.
- 코드 수정 후에는 반드시 `python build.py`를 다시 실행해야 실행파일에 반영됩니다.
- 기존 실행파일은 새 빌드로 덮어써집니다.

---

## 5. 입력 ZIP 안에 있어야 하는 파일

| 파일 패턴 | 용도 |
|-----------|------|
| `*.xsd` | entity 확장 concept 정의 (DataType·Period·Balance) |
| `*_pre.xml` | Presentation linkbase (트리 구조) |
| `*_lab-ko.xml` | 한글 label |
| `*_lab-en.xml` | 영문 label |
| `*.xbrl` | 인스턴스 파일 (선택) — Decimal 값에 사용 |

`.xbrl`이 없으면 Decimal만 비워두고 나머지는 정상 출력됩니다.

---

## 6. DataType · Period 채우는 방식

| concept 출처 | DataType · Period 소스 |
|-------------|----------------------|
| entity 확장 (ZIP 내 XSD) | ZIP 안 `.xsd` 파일에서 직접 읽음 |
| ifrs-full / dart 표준 | `dart_taxonomy.json` 조회 (8,014개 concept) |

`dart_taxonomy.json`은 `DART_Taxonomy_Search.xlsm`의 Concepts 시트를 추출한 파일입니다.

---

## 7. Decimal 선택 규칙

같은 concept이 여러 context(당기/전기, 연결/별도, member별)로 존재할 수 있으므로  
**당기(`contextRef`가 `C`로 시작)를 우선** 사용하고, 그 중 가장 rounded된 값(최솟값)을 선택합니다.

---

## 8. 후처리 규칙

Presentation 트리 구성 후 아래 순서로 행을 정리합니다.

1. **table 컬럼 채우기** — concept id가 `Table`로 끝나는 행이 나타나면 해당 label-ko를 이후 행의 `table`로 전파합니다. role 경계에서 초기화됩니다.
2. **소급 채우기** — Table 행 바로 위에 연속된 Abstract / TextBlock 행이 있으면 동일 테이블명을 소급합니다.
3. **order 없는 행 삭제** — `order` 값이 없는 행을 제거합니다.
4. **특정 label-ko 행 삭제** — "연결 또는 별도 재무제표 [Table]" 등 불필요한 분류 행을 제거합니다.
