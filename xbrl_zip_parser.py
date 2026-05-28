"""
xbrl_zip_parser.py
XBRL ZIP 파일 → checklist_engine 호환 presentation_rows 변환기
taxonomy_xlsx_parser.parse_taxonomy_xlsx() 와 동일한 출력 형태
"""
from __future__ import annotations

import io
import json
import re
import shutil
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional
from xml.etree import ElementTree as ET

# ── dart_taxonomy.json 캐시 ────────────────────────────────────────────────────
_TAXONOMY_JSON = Path(__file__).parent / "dart_taxonomy.json"
_taxonomy_cache: dict[str, dict] | None = None


def _load_dart_taxonomy() -> dict[str, dict]:
    global _taxonomy_cache
    if _taxonomy_cache is None:
        if _TAXONOMY_JSON.exists():
            _taxonomy_cache = json.loads(_TAXONOMY_JSON.read_text(encoding="utf-8"))
        else:
            _taxonomy_cache = {}
    return _taxonomy_cache


# ── 연결/별도 구분용 구조 요소 (체크리스트 불필요) ────────────────────────────────
_CONSOL_SEPARATE_NAMES: set[str] = {
    'ConsolidatedAndSeparateFinancialStatementsTable',
    'ConsolidatedAndSeparateFinancialStatementsAxis',
    'ConsolidatedAndSeparateFinancialStatementsDomain',
    'ConsolidatedMember',
    'SeparateMember',
}

# ── 네임스페이스 ──────────────────────────────────────────────────────────────
NS = {
    "xsd":   "http://www.w3.org/2001/XMLSchema",
    "xbrli": "http://www.xbrl.org/2003/instance",
    "link":  "http://www.xbrl.org/2003/linkbase",
    "xlink": "http://www.w3.org/1999/xlink",
    "xml":   "http://www.w3.org/XML/1998/namespace",
}
XLINK_HREF  = f"{{{NS['xlink']}}}href"
XLINK_LABEL = f"{{{NS['xlink']}}}label"
XLINK_FROM  = f"{{{NS['xlink']}}}from"
XLINK_TO    = f"{{{NS['xlink']}}}to"
XLINK_ROLE  = f"{{{NS['xlink']}}}role"


# ── taxonomy_xlsx_parser 와 동일한 분류 함수 ──────────────────────────────────

def _is_consol(text: str) -> Optional[bool]:
    for k in ['Consolidated', 'consolidated', '연결']:
        if k in text: return True
    for k in ['Separated', 'Separate', 'separated', '별도', 'Nonconsolidated']:
        if k in text: return False
    return None


def _extract_table_number(role_def: str) -> str:
    m = re.search(r'\[([A-Za-z]{1,3}X?\d{4,})\]', str(role_def))
    return m.group(1) if m else ''


def _extract_role_code(role_def: str, role_uri: str) -> str:
    code = _extract_table_number(role_def)
    if code: return code
    m = re.search(r'/([A-Z]{1,3}X?\d{4,})$', str(role_uri))
    return m.group(1) if m else ''


def _label_role_short(url: str) -> str:
    s = str(url).strip()
    return '' if not s else s.split('/')[-1]


def _classify_element(name: str) -> str:
    suffix = name[-4:].lower() if len(name) >= 4 else ''
    element = {
        'tory': 'Explanatory',
        'ract': 'Abstract',
        'axis': 'Axis',
        'lock': 'TextBlock',
        'able': 'Table',
        'mber': 'Member',
        'main': 'Domain',
    }.get(suffix, 'item')
    if 'lineitem' in name.lower():
        element = 'Lineitem'
    return element


def _classify_gubn(name: str) -> str:
    if name.endswith('Table'):                                   return 'TABLE'
    if name.endswith('TextBlock'):                               return 'FOOTNOTES'
    if name.endswith('Axis'):                                    return 'Axis'
    if name.endswith('Member'):                                  return 'Member'
    if name.endswith('Domain'):                                  return 'Domain'
    if name.endswith('LineItems') or name.endswith('LineItem'): return 'LINEITEM'
    return 'LINEITEM'


def _add_axis_group_fields(rows: list) -> None:
    """3-1, 3-2 체크용 축-도메인 그룹핑 필드 추가 (taxonomy_xlsx_parser 와 동일 로직)"""
    groups: dict[str, list] = defaultdict(list)
    for i, row in enumerate(rows):
        groups[row.get('role_uri', '')].append((i, row))

    for _, indexed_rows in groups.items():
        prev_element     = None
        prev_axis_domain = None
        prev_group_id    = None
        prev_axis_name   = None

        for _, (orig_idx, row) in enumerate(indexed_rows):
            element = row.get('Element', '')
            name    = row.get('Name', '')

            if prev_element == 'Axis' and element in ('Member', 'Domain'):
                axis_domain = '도메인'
            elif element == 'Axis':
                axis_domain = '축'
            elif element in ('Member', 'Domain') and prev_axis_domain in ('축', '도메인', '멤버'):
                axis_domain = '멤버'
            else:
                axis_domain = None

            axis_flag = 1 if axis_domain == '축' else 0

            if axis_domain is None:
                group_id = None
            elif axis_flag == 1:
                group_id = 1 if prev_group_id is None else prev_group_id + 1
            else:
                group_id = prev_group_id

            if axis_domain is None:
                axis_name = None
            elif prev_group_id is None or group_id != prev_group_id:
                axis_name = name if axis_domain == '축' else ''
            else:
                axis_name = name if axis_domain == '축' else prev_axis_name

            # Axis = Axis-Axis (레퍼런스 KEY 형식과 일치)
            # Domain = Axis-Domain
            # Member = Axis-Member
            if axis_domain == '축':
                key = f"{axis_name}-{axis_name}" if axis_name else None
            elif axis_domain == '도메인':
                key = f"{axis_name}-{name}" if axis_name else None
            elif axis_domain == '멤버':
                key = f"{axis_name}-{name}" if axis_name else None
            else:
                key = None

            rows[orig_idx].update({
                '축_도메인': axis_domain,
                'Axis_flag': axis_flag,
                'Axis_Name': axis_name,
                'GroupID':   group_id,
                'KEY_axis':  key,
            })

            prev_element     = element
            prev_axis_domain = axis_domain
            prev_group_id    = group_id
            prev_axis_name   = axis_name


def _remap_gubn_alteryx(rows: list) -> None:
    """Alteryx 원본 3-value 구분 체계로 재분류.
    TABLE → 'TABLE', Axis·Domain·Member → 'DOMAIN', 나머지 → 'LINEITEM'
    """
    for row in rows:
        ad = row.get('축_도메인')
        g  = row.get('구분', '')
        if g == 'TABLE':
            row['구분'] = 'TABLE'
        elif ad in ('축', '도메인', '멤버'):
            row['구분'] = 'DOMAIN'
        else:
            row['구분'] = 'LINEITEM'


# ── taxonomy_xlsx_parser.TaxonomyXlsxData 호환 클래스 ────────────────────────

class XBRLData:
    class _El:
        def __init__(self, lko: str = '', len_: str = '', lr: str = ''):
            self.label_ko   = lko
            self.label_en   = len_
            self.label_role = lr
            self.abstract   = False

    def __init__(self):
        self.company_name: str        = ''
        self.report_date:  str        = ''
        self.entity_id:    str        = ''
        self.fy:           str        = ''
        self.report_period: str       = ''
        self.presentation_rows: List[dict] = []
        self.errors:        List[str] = []
        self.axis_domain_rows: List[dict] = []
        self.elements: Dict[str, object] = {}
        self.contexts: Dict[str, object] = {}
        self.facts:    list = []
        self._fact_elements: set = set()


# ── XSD 파싱 ─────────────────────────────────────────────────────────────────

def _parse_xsd(path: str) -> tuple[dict[str, dict], dict[str, str]]:
    """
    Returns
    -------
    concept_map : {concept_id → {name, prefix, type, balance, periodType, abstract}}
    role_def    : {roleURI → definition_text}
    """
    tree = ET.parse(path)
    root = tree.getroot()

    concept_map: dict[str, dict] = {}
    for el in root.iter(f"{{{NS['xsd']}}}element"):
        cid  = el.get("id", "")
        name = el.get("name", "")
        if not cid or not name:
            continue

        # 'ifrs-full_Assets' → prefix='ifrs-full', name='Assets'
        prefix = cid.split("_", 1)[0] if "_" in cid else ""

        raw_type = el.get("type", "")
        dtype    = raw_type.split(":")[-1] if ":" in raw_type else raw_type

        concept_map[cid] = {
            "name":       name,
            "prefix":     prefix,
            "type":       dtype,
            "balance":    el.get(f"{{{NS['xbrli']}}}balance", "").lower(),
            "periodType": el.get(f"{{{NS['xbrli']}}}periodType", "").upper(),
            "abstract":   el.get("abstract", "false").lower() == "true",
        }

    role_def: dict[str, str] = {}
    for rt in root.iter(f"{{{NS['link']}}}roleType"):
        uri  = rt.get("roleURI", "")
        defn = rt.find(f"{{{NS['link']}}}definition")
        if uri and defn is not None and defn.text:
            role_def[uri] = defn.text.strip()

    # dart_taxonomy.json 으로 ifrs-full / dart 표준 concept 보완
    for cid, info in _load_dart_taxonomy().items():
        if cid not in concept_map:
            prefix = cid.split("_", 1)[0] if "_" in cid else ""
            name   = cid.split("_", 1)[1] if "_" in cid else cid
            concept_map[cid] = {
                "name":       name,
                "prefix":     prefix,
                "type":       info.get("type", ""),
                "balance":    (info.get("balance") or "").lower(),
                "periodType": (info.get("periodType") or "").upper(),
                "abstract":   False,
            }

    return concept_map, role_def


# ── Label linkbase 파싱 ───────────────────────────────────────────────────────

def _href_to_id(href: str) -> str:
    return href.split("#", 1)[1] if "#" in href else href


def _parse_labels(path: str) -> dict[str, dict[str, str]]:
    """Returns {concept_id: {role_uri: label_text}}"""
    tree = ET.parse(path)
    root = tree.getroot()

    loc_to_id:    dict[str, str]            = {}
    res_to_role:  dict[str, tuple[str, str]] = {}
    arcs:         list[tuple[str, str]]     = []

    for ll in root.iter(f"{{{NS['link']}}}labelLink"):
        for loc in ll.findall(f"{{{NS['link']}}}loc"):
            lbl = loc.get(XLINK_LABEL, "")
            cid = _href_to_id(loc.get(XLINK_HREF, ""))
            if lbl and cid:
                loc_to_id[lbl] = cid

        for lab in ll.findall(f"{{{NS['link']}}}label"):
            lbl  = lab.get(XLINK_LABEL, "")
            role = lab.get(XLINK_ROLE, "")
            text = lab.text or ""
            if lbl:
                res_to_role[lbl] = (role, text)

        for arc in ll.findall(f"{{{NS['link']}}}labelArc"):
            f = arc.get(XLINK_FROM, "")
            t = arc.get(XLINK_TO, "")
            if f and t:
                arcs.append((f, t))

    labels: dict[str, dict[str, str]] = defaultdict(dict)
    for f, t in arcs:
        cid = loc_to_id.get(f)
        rt  = res_to_role.get(t)
        if cid and rt:
            role, text = rt
            labels[cid][role] = text

    return dict(labels)


def _get_label(labels: dict[str, dict[str, str]],
               cid: str, pref_role: str | None = None) -> str:
    if not cid or cid not in labels:
        return ""
    by_role = labels[cid]
    if pref_role and pref_role in by_role:
        return by_role[pref_role]
    return by_role.get("http://www.xbrl.org/2003/role/label", "")


# ── Presentation linkbase 파싱 ────────────────────────────────────────────────

def _parse_presentation(
    path:        str,
    labels_ko:   dict[str, dict[str, str]],
    labels_en:   dict[str, dict[str, str]],
    role_def_map: dict[str, str],
    concept_map: dict[str, dict],
) -> tuple[list[dict], dict[str, object]]:
    tree = ET.parse(path)
    root = tree.getroot()

    rows:     list[dict]          = []
    elements: dict[str, object]   = {}

    for pl in root.iter(f"{{{NS['link']}}}presentationLink"):
        role_uri     = pl.get(XLINK_ROLE, "")
        role_def_str = role_def_map.get(role_uri, role_uri)

        # ── Role 메타데이터 ──
        code      = _extract_role_code(role_def_str, role_uri)
        table_num = _extract_table_number(role_def_str) or code
        parts     = role_def_str.split("|", 1)
        name_ko   = re.sub(r'^\[[^\]]+\]\s*', '', parts[0]).strip()
        name_en   = parts[1].strip() if len(parts) > 1 else ''

        is_c = _is_consol(role_def_str or role_uri)
        if is_c is None and code:
            if code[-1] == '0':   is_c = True
            elif code[-1] == '5': is_c = False

        consol_str = '-'
        if code:
            if code[-1] == '0':   consol_str = '연결'
            elif code[-1] == '5': consol_str = '별도'

        sheet_name = code or role_uri.split('/')[-1]

        # ── Locator → concept_id ──
        loc_to_id: dict[str, str] = {}
        for loc in pl.findall(f"{{{NS['link']}}}loc"):
            loc_to_id[loc.get(XLINK_LABEL, "")] = _href_to_id(loc.get(XLINK_HREF, ""))

        # ── Arc 그래프 ──
        children: dict[str, list[tuple[str, float, str]]] = defaultdict(list)
        targets:  set[str] = set()
        sources:  set[str] = set()
        for arc in pl.findall(f"{{{NS['link']}}}presentationArc"):
            f_   = arc.get(XLINK_FROM, "")
            t_   = arc.get(XLINK_TO, "")
            try:
                order = float(arc.get("order", "0"))
            except ValueError:
                order = 0.0
            pref = arc.get("preferredLabel", "") or ""
            children[f_].append((t_, order, pref))
            sources.add(f_)
            targets.add(t_)

        for f_ in children:
            children[f_].sort(key=lambda x: x[1])

        loc_order = {loc.get(XLINK_LABEL, ""): i
                     for i, loc in enumerate(pl.findall(f"{{{NS['link']}}}loc"))}
        roots = [lbl for lbl in loc_order if lbl in sources and lbl not in targets]
        roots.sort(key=lambda x: loc_order.get(x, 0))

        current_table_name_ko = ''

        def make_row(cid: str, depth: int, order_val: Any, pref_url: str,
                     par_name: str, par_lbl_ko: str, par_gubn: str) -> dict:
            nonlocal current_table_name_ko

            concept  = concept_map.get(cid, {})
            name     = concept.get("name") or (cid.split("_", 1)[-1] if "_" in cid else cid)
            prefix   = concept.get("prefix") or (cid.split("_", 1)[0] if "_" in cid else "")
            dtype    = concept.get("type", "")
            balance  = concept.get("balance", "")
            period   = concept.get("periodType", "")
            abstract = concept.get("abstract", False)

            lbl_ko   = _get_label(labels_ko, cid, pref_url or None)
            lbl_en   = _get_label(labels_en, cid, pref_url or None)
            lbl_role = _label_role_short(pref_url)

            gubn    = _classify_gubn(name)
            element = _classify_element(name)
            # Axis 바로 아래 첫 번째 자식은 Domain (이름과 무관하게 위치로 강제 분류)
            if par_gubn == 'Axis':
                gubn    = 'Domain'
                element = 'Domain'
            ext     = '확장' if prefix.startswith('entity') else '-'
            client_negate = 'negate' if 'negated' in lbl_role.lower() else '-'
            alias   = '별칭'  if 'terse'   in lbl_role.lower() else '-'

            if gubn == 'TABLE':
                current_table_name_ko = lbl_ko

            if name not in elements:
                el_obj = XBRLData._El(lbl_ko, lbl_en, lbl_role)
                el_obj.abstract = abstract
                elements[name] = el_obj

            return {
                'role_uri':        role_uri,
                'role_code':       code,
                'role_name_ko':    name_ko,
                'role_name_en':    name_en,
                'is_consolidated': is_c,
                'Role Definition': role_def_str,
                'Sheet':           sheet_name,
                '연결/별도':        consol_str,
                'Table_Number':    table_num,
                'TABLE_NUMBER':    table_num,
                'depth':           depth,
                'parent':          par_name,
                'parent_label_ko': par_lbl_ko,
                'parent_gubn':     par_gubn,
                'Prefix':          prefix,
                'Name':            name,
                'Label(KO)':       lbl_ko,
                'Label(EN)':       lbl_en,
                'Label Role':      lbl_role,
                'DataType':        dtype,
                'Balance':         balance,
                'Period':          period,
                'Decimal':         '',
                'Fact':            '',
                '구분':             gubn,
                'Element':         element,
                '확장여부':          ext,
                'Client_Negate':   client_negate,
                '별칭여부':          alias,
                'PreferredLabel':  pref_url,
                'has_fact':        False,
                'abstract':        abstract,
                'table_name_ko':   current_table_name_ko,
            }

        def dfs(loc_label: str, depth: int, order: float | None, pref_url: str,
                par_name: str, par_lbl_ko: str, par_gubn: str) -> None:
            cid = loc_to_id.get(loc_label, "")
            if order is not None:
                order_val: Any = int(order) if float(order).is_integer() else order
            else:
                order_val = ""

            row = make_row(cid, depth, order_val, pref_url, par_name, par_lbl_ko, par_gubn)

            # 연결/별도 구분용 구조 요소는 체크리스트 불필요 → 행 제외 (자식은 계속 탐색)
            if row['Name'] not in _CONSOL_SEPARATE_NAMES:
                rows.append(row)

            for to_lbl, o, p in children.get(loc_label, []):
                dfs(to_lbl, depth + 1, o, p,
                    row['Name'], row['Label(KO)'], row['구분'])

        for r in roots:
            dfs(r, 0, None, "", "", "", "")

    return rows, elements


# ── 회사명 / 기간 추출 ────────────────────────────────────────────────────────

_PERIOD_SUFFIX_MAP = {
    'FY':  '4Q',
    'FQA': '1Q', 'FQQ': '1Q',
    'HY':  '2Q',
    'TQA': '3Q', 'TQQ': '3Q',
}
_CTX_RE = re.compile(r'^CFY(\d{4})[de]([A-Z]+)')


def _extract_company_name(xbrl_path: str) -> str:
    """.xbrl 인스턴스에서 EntityRegistrantName(한글) 추출."""
    try:
        for _, el in ET.iterparse(xbrl_path, events=("end",)):
            tag = el.tag
            if tag.endswith("}EntityRegistrantName") and el.text:
                text = el.text.strip()
                if any('가' <= c <= '힣' for c in text):
                    return text
            el.clear()
    except Exception:
        pass
    return ""


def _extract_period_info(xbrl_path: str) -> tuple[str, str]:
    """
    당기(C) 컨텍스트 ID에서 회계연도·분기 추출.
    반환: (fy, period)  예) ('FY25', '1Q') / ('FY25', 'Annual')
    """
    try:
        for _, el in ET.iterparse(xbrl_path, events=("start",)):
            if el.tag.endswith('}context'):
                m = _CTX_RE.match(el.get('id', ''))
                if m:
                    year   = m.group(1)[2:]
                    suffix = m.group(2)
                    period = _PERIOD_SUFFIX_MAP.get(suffix, suffix)
                    return f'FY{year}', period
            el.clear()
    except Exception:
        pass
    return '', ''


# ── table_name_ko 후처리 ─────────────────────────────────────────────────────

def _postprocess_table_name(rows: list[dict]) -> None:
    """
    1) TABLE 행 기준으로 이후 행에 table_name_ko 전파 (role 경계에서 초기화)
    2) TABLE 행 바로 위 연속된 Abstract / TextBlock / Explanatory 행에 소급 적용
    3) 여전히 비어 있는 행은 role_name_ko (role에서 [코드] 접두어 제거한 값)으로 채움
    """
    # 1) 전방 전파 (현재 make_row 에서 이미 수행되나 role 경계 초기화 보완)
    current_role = ""
    current_table = ""
    for row in rows:
        if row.get("role_uri", "") != current_role:
            current_role = row.get("role_uri", "")
            current_table = ""
        if row.get("구분") == "TABLE":
            current_table = row.get("Label(KO)") or row.get("Name", "")
        row["table_name_ko"] = current_table

    # 2) 역소급: TABLE 행 바로 위 Abstract / TextBlock / Explanatory 연속 행
    for i, row in enumerate(rows):
        if row.get("구분") == "TABLE":
            j = i - 1
            while j >= 0 and rows[j].get("role_uri") == row.get("role_uri"):
                elem = rows[j].get("Element", "")
                if elem in ("Abstract", "TextBlock", "Explanatory"):
                    rows[j]["table_name_ko"] = row["table_name_ko"]
                    j -= 1
                else:
                    break

    # 3) 여전히 빈 행 → role_name_ko 사용
    for row in rows:
        if not row.get("table_name_ko"):
            role_def = row.get("Role Definition", "")
            name_ko = re.sub(r"^\[[^\]]+\]\s*", "", role_def.split("|", 1)[0]).strip()
            row["table_name_ko"] = name_ko


# ── 파일 자동 탐지 ────────────────────────────────────────────────────────────

def _find(directory: str, suffix: str) -> str:
    """suffix로 끝나는 파일을 directory 아래에서 재귀 탐색 (EntityTaxonomy 등 하위 폴더 지원)."""
    matches = [p for p in Path(directory).rglob(f"*{suffix}") if p.is_file()]
    if not matches:
        raise FileNotFoundError(f"'{suffix}' 로 끝나는 파일을 {directory} 에서 찾지 못했습니다.")
    # 여러 개일 경우 경로 깊이가 얕은 것(루트에 가까운 것) 우선
    matches.sort(key=lambda p: len(p.parts))
    return str(matches[0])


# ── 메인 파서 (taxonomy_xlsx_parser.parse_taxonomy_xlsx 와 동일한 인터페이스) ──

def parse_xbrl_zip(file_bytes: bytes) -> XBRLData:
    """
    ZIP 바이트 → XBRLData  (TaxonomyXlsxData 호환)

    사용 예)
        with open("entity.zip", "rb") as f:
            data = parse_xbrl_zip(f.read())
        results = run_all_checks(data)
    """
    data    = XBRLData()
    tmp_dir = tempfile.mkdtemp()

    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            zf.extractall(tmp_dir)

        xsd_path    = _find(tmp_dir, ".xsd")
        pre_path    = _find(tmp_dir, "_pre.xml")
        lab_ko_path = _find(tmp_dir, "_lab-ko.xml")
        lab_en_path = _find(tmp_dir, "_lab-en.xml")

        data.entity_id = Path(xsd_path).stem

        # 회사명 / 회계연도 / 분기: .xbrl 인스턴스에서 추출
        try:
            xbrl_path = _find(tmp_dir, ".xbrl")
            data.company_name = _extract_company_name(xbrl_path)
            data.fy, data.report_period = _extract_period_info(xbrl_path)
        except Exception:
            pass

        concept_map, role_def_map = _parse_xsd(xsd_path)
        labels_ko = _parse_labels(lab_ko_path)
        labels_en = _parse_labels(lab_en_path)

        rows, elements = _parse_presentation(
            pre_path, labels_ko, labels_en, role_def_map, concept_map
        )

        _add_axis_group_fields(rows)
        _postprocess_table_name(rows)
        _remap_gubn_alteryx(rows)

        data.presentation_rows = rows
        data.elements          = elements
        data.axis_domain_rows  = [r for r in rows if r.get('GroupID') is not None]

    except Exception as e:
        data.errors.append(str(e))
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return data
