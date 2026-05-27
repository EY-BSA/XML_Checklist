# XBRL CoE 체크리스트 도구

DART 원문 XBRL ZIP 파일을 입력받아 29개 검증 항목을 자동으로 체크하고 결과를 xlsx로 저장합니다.

**Windows / macOS 모두 지원합니다.**

---

## 폴더 구성

```
zip_cleansing/
├── main.py                # 드래그앤드롭 GUI 실행기
├── xbrl_zip_parser.py     # ZIP 파서 (XBRL 구조 분석)
├── checklist_engine.py    # 29개 체크리스트 엔진
├── checklist_export.py    # xlsx 결과 내보내기
├── standard_taxonomy.py   # 표준 택사노미 로더
├── xbrl_to_xlsx.py        # Presentation 구조 분석 도구 (CLI 단독 실행 가능)
├── dart_taxonomy.json     # ifrs-full / dart 표준 concept DataType·Period 조회 테이블
├── data/
│   ├── Axis_Domain_Check.xlsx   # 축-도메인 정합성 기준 데이터
│   └── DART_Negate_Check.xlsx   # Negate 검토 기준 데이터
├── template/
│   └── XBRL_CoE_Checklist_Result.xlsx  # 체크리스트 결과 템플릿
├── build.py               # Mac / Windows 공용 빌드 스크립트
├── build.spec             # PyInstaller 설정
├── requirements.txt
└── README.md
```

---

## 사용 방법

### 실행파일 (배포용)

빌드된 실행파일만 있으면 Python 설치 없이 바로 사용 가능합니다.

| OS | 실행파일 |
|----|---------|
| Windows | `dist\XBRL_CoE_Checklist.exe` |
| macOS | `dist/XBRL_CoE_Checklist.app` |

1. 실행파일을 실행합니다.
2. 창 안에 `.zip` 파일을 드래그하거나 `파일 선택…` 버튼을 누릅니다.
3. 완료되면 입력 zip과 같은 폴더에 `XBRL_CoE_Checklist_Result_<회사명>.xlsx`가 저장됩니다.

### 소스 코드로 실행

```bash
pip install -r requirements.txt
python main.py
```

---

## 전체 데이터 흐름

```
DART 원문 ZIP
    │
    ▼
xbrl_zip_parser.py
    ├── .xsd           → concept 정의 (DataType·Period·Balance)
    ├── dart_taxonomy.json → ifrs-full/dart 표준 concept 보완
    ├── *_lab-ko/en.xml → 한글/영문 label
    ├── *_pre.xml       → Presentation 트리 구조
    └── *.xbrl          → 인스턴스 fact (회사명·Decimal 값)
    │
    ▼
checklist_engine.py  (29개 항목 검증)
    │
    ▼
checklist_export.py
    │
    ▼
XBRL_CoE_Checklist_Result_<회사명>.xlsx
```

---

## 체크리스트 항목 (29개)

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

---

## 입력 ZIP 구성

| 파일 패턴 | 용도 |
|-----------|------|
| `*.xsd` | entity 확장 concept 정의 |
| `*_pre.xml` | Presentation linkbase (트리 구조) |
| `*_lab-ko.xml` | 한글 label |
| `*_lab-en.xml` | 영문 label |
| `*.xbrl` | 인스턴스 파일 — 회사명·Decimal 값 |

---

## 실행파일 빌드 (PyInstaller)

코드를 수정한 뒤 실행파일에 반영하려면 빌드를 다시 실행해야 합니다.

```bash
python build.py
```

- Windows / macOS 모두 동일한 명령어를 사용합니다.
- 빌드 완료 후 `dist/` 폴더에 실행파일이 생성됩니다.
- 기존 실행파일은 새 빌드로 자동 덮어써집니다.

---

## DataType · Period 조회 방식

| concept 출처 | 조회 소스 |
|-------------|---------|
| entity 확장 (ZIP 내 XSD) | ZIP 안 `.xsd` 파일에서 직접 읽음 |
| ifrs-full / dart 표준 | `dart_taxonomy.json` (8,014개 concept) |

`dart_taxonomy.json`은 `DART_Taxonomy_Search.xlsm`의 Concepts 시트를 추출한 파일입니다.
