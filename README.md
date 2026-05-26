# XBRL ZIP 변환 · 체크리스트 도구

DART 원문 XBRL ZIP 파일을 대상으로 두 가지 기능을 제공합니다.

| 기능 | 실행 방법 | 설명 |
|------|-----------|------|
| **XLSX 변환** | `python main.py` 또는 `python xbrl_to_xlsx.py` | ZIP → Presentation 1-시트 xlsx |
| **체크리스트** | `python app.py` | 29개 검증 항목을 웹 브라우저에서 확인 |

**Windows / macOS 모두 지원합니다.**

## 폴더 구성

```
zip_cleansing/
├── main.py                # 드래그앤드롭 GUI (XLSX 변환)
├── xbrl_to_xlsx.py        # ZIP → XLSX 변환 코어 (CLI 단독 실행 가능)
├── app.py                 # Flask 웹앱 (체크리스트)
├── xbrl_zip_parser.py     # 체크리스트용 ZIP 파서
├── checklist_engine.py    # 29개 체크리스트 엔진
├── standard_taxonomy.py   # 표준 택사노미 로더
├── dart_taxonomy.json     # ifrs-full / dart 표준 concept DataType·Period 조회 테이블
├── data/
│   ├── Axis_Domain_Check.xlsx   # 축-도메인 정합성 기준 데이터
│   └── DART_Negate_Check.xlsx   # Negate 검토 기준 데이터
├── template/
│   └── XBRL_CoE_Checklist_Result.xlsx  # 체크리스트 결과 내보내기 템플릿
├── templates/
│   └── index.html         # 웹앱 프론트엔드
├── build.py               # Mac / Windows 공용 빌드 스크립트 (XLSX 변환기 전용)
├── build.spec             # PyInstaller 설정
├── requirements.txt
└── README.md
```

---

## 1. XLSX 변환기

### 1-1. 전체 데이터 흐름

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

### 1-2. 출력 xlsx 시트 구성

출력 파일은 **Presentation 시트 1장**으로 구성됩니다.

| 컬럼 | 소스 | 설명 |
|------|------|------|
| role | `*_pre.xml` | 재무제표 구분 (예: `[D210000] 재무상태표, 유동/비유동법 - 연결`) |
| table | `*_pre.xml` | concept id가 `Table`로 끝나는 노드 기준 테이블명 |
| id | `*_pre.xml` | concept id |
| label-ko | `*_lab-ko.xml` | preferredLabel 기준 한글 label |
| label-en | `*_lab-en.xml` | preferredLabel 기준 영문 label |
| depth | `*_pre.xml` | 트리 깊이 (0 = root) |
| order | `*_pre.xml` | 형제 노드 간 표시 순서 |
| pref_label | `*_pre.xml` | presentationArc의 preferredLabel role |
| DataType | `.xsd` / `dart_taxonomy.json` | monetaryItemType, stringItemType 등 |
| Period | `.xsd` / `dart_taxonomy.json` | INSTANT / DURATION |
| Decimal | `*.xbrl` | 당기 context 중 가장 rounded된 decimals 값 |

### 1-3. 실행

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

### 1-4. 실행파일 빌드 (PyInstaller)

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

## 2. 체크리스트 웹앱

### 2-1. 실행

```bash
pip install -r requirements.txt
python app.py
```

브라우저에서 `http://localhost:5000`을 열고 DART 원문 XBRL `.zip` 파일을 업로드하세요.

### 2-2. 체크리스트 항목 (29개)

| 번호 | 항목 |
|------|------|
| 1-1 | Gross 계정 사용 검토 |
| 1-2 | 초과적립액(과소적립액) 텍사노미 사용 검토 |
| 1-3 | 재고자산 세부내역 표 검토 |
| 1-4 | 유동/비유동 축 검토 |
| 2-1 | (만료) 대손충당금 멤버 사용 검토 |
| 2-2 | (만료) 금융자산 손상차손 축 사용 검토 |
| 2-3 | 대출약정 텍사노미 검토 |
| 2-4 | 미착품 텍사노미 검토 |
| 2-5 | 배당금 텍사노미 검토 |
| 2-6 | 평균유효세율 검토 (분반기) |
| 3-1 | Axis & Domain & Member 정합성 검토 |
| 3-2 | 공시금액의 사용 적정성 검토 |
| 4-1 | 현금흐름 관련 표 내에서 다른 요소 사용 |
| 4-2 | 현금흐름 관련 표의 전용요소가 다른 표에서 사용 |
| 4-3 | 판매관리비 관련 표 내에서 다른 요소 사용 |
| 4-4 | 판매비와관리비 관련 표의 전용요소가 다른 표에서 사용 |
| 4-5 | 특수관계자 관련 표 내에서 다른 요소 사용 |
| 4-6 | 특수관계자 관련 표의 전용요소가 다른 표에서 사용 |
| 5-1 | Percent 소숫점 자리수 검토 |
| 5-2 | 보유하는 주식수 속성 검토 |
| 5-3 | 이연법인세부채(자산) 텍사노미 및 부호 검토 |
| 5-4 | 기본주당이익/희석주당이익 속성 검토 |
| 5-5 | 기초/기말 영문명 검토 |
| 5-6 | 단위표시 검토 |
| 6-1 | 축 확장 검토 |
| 6-2 | 멤버 합계열 확장 검토 |
| 6-3 | Duration / Instant 속성 검토 |
| 7-1 | Client Negate 검토 |
| 7-2 | 현금흐름표 영업활동 현금흐름 검토 |

### 2-3. 결과 내보내기

체크리스트 실행 후 `결과 내보내기` 버튼을 클릭하면  
`XBRL_CoE_Checklist_Result_<회사명>.xlsx` 파일로 다운로드됩니다.

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

## 4. DataType · Period 채우는 방식

| concept 출처 | DataType · Period 소스 |
|-------------|----------------------|
| entity 확장 (ZIP 내 XSD) | ZIP 안 `.xsd` 파일에서 직접 읽음 |
| ifrs-full / dart 표준 | `dart_taxonomy.json` 조회 (8,014개 concept) |

`dart_taxonomy.json`은 `DART_Taxonomy_Search.xlsm`의 Concepts 시트를 추출한 파일입니다.

---

## 5. Decimal 선택 규칙

같은 concept이 여러 context(당기/전기, 연결/별도, member별)로 존재할 수 있으므로  
**당기(`contextRef`가 `C`로 시작)를 우선** 사용하고, 그 중 가장 rounded된 값(최솟값)을 선택합니다.

---

## 6. 후처리 규칙 (XLSX 변환기)

Presentation 트리 구성 후 아래 순서로 행을 정리합니다.

1. **table 컬럼 채우기** — concept id가 `Table`로 끝나는 행이 나타나면 해당 label-ko를 이후 행의 `table`로 전파합니다. role 경계에서 초기화됩니다.
2. **소급 채우기** — Table 행 바로 위에 연속된 Abstract / TextBlock / Explanatory 행이 있으면 동일 테이블명을 소급합니다.
3. **빈 table 채우기** — Table 개념이 없는 재무제표(재무상태표 등)는 role 이름에서 코드 부분을 제거한 값으로 채웁니다.
4. **order 없는 행 삭제** — `order` 값이 없는 행을 제거합니다.
5. **특정 label-ko 행 삭제** — "연결 또는 별도 재무제표 [Table]" 등 불필요한 분류 행을 제거합니다.
