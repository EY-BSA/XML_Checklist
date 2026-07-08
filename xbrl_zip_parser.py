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
from typing import Any, Dict, List
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
    # 4글자 대신 전체 단어로 비교 (예: '...Variable'/'...Number'/'...Contract'/
    # '...Inventory'가 'able'/'mber'/'ract'/'tory' 4글자만 보고 오매칭되는 것 방지)
    if name.endswith('Explanatory'):
        element = 'Explanatory'
    elif name.endswith('Abstract'):
        element = 'Abstract'
    elif name.endswith('Axis'):
        element = 'Axis'
    elif name.endswith('TextBlock'):
        element = 'TextBlock'
    elif name.endswith('Table'):
        element = 'Table'
    elif name.endswith('Member'):
        element = 'Member'
    elif name.endswith('Domain'):
        element = 'Domain'
    else:
        element = 'item'
    if name.lower().endswith(('lineitem', 'lineitems')):
        element = 'Lineitem'
    return element


def _classify_gubn(name: str) -> str:
    if name.endswith('Table'):                                   return 'TABLE'
    if name.endswith('TextBlock'):                               return 'FOOTNOTES'
    if name.endswith('Explanatory'):                             return 'FOOTNOTES'
    if name.endswith('Axis'):                                    return 'Axis'
    if name.endswith('Member'):                                  return 'Member'
    if name.endswith('Domain'):                                  return 'Domain'
    if name.endswith('LineItems') or name.endswith('LineItem'): return 'LINEITEM'
    return 'LINEITEM'


def _is_text_block_like_concept(name: str = '', dtype: str = '',
                                lbl_ko: str = '', lbl_en: str = '') -> bool:
    """이름이 Explanatory/TextBlock으로 끝나지 않아도
    DataType 또는 Label상 [text block]/[문장영역]이면 문장영역으로 판별.
    예: DescriptionOfManagingLiquidityRisk → Label(EN): '... [text block]'
    """
    n  = str(name  or '')
    dt = str(dtype or '').split(':')[-1]
    ko = str(lbl_ko or '').lower()
    en = str(lbl_en or '').lower()

    if n.endswith(('Explanatory', 'TextBlock')):
        return True
    if dt == 'textBlockItemType' or dt.lower().endswith('textblockitemtype'):
        return True

    label_text = f'{ko} {en}'
    text_block_markers = (
        '[text block]', 'text block',
        '[문장영역]', '문장영역',
        '[텍스트블록]', '텍스트블록',
        '[텍스트 블록]', '텍스트 블록',
    )
    return any(marker in label_text for marker in text_block_markers)


_ROLE_CODE_RE = re.compile(r'([A-Z]{1,3}X?\d{4,})[a-z]*$')


def _role_code_from_uri(role_uri: str) -> str:
    """def.xml sub-role URI에서 base role_code 추출.
    예) '.../dart_2024-06-30_role-D827585d' → 'D827585'
    """
    segment = role_uri.split('/')[-1]
    m = _ROLE_CODE_RE.search(segment)
    return m.group(1) if m else ''


def _parse_def_linkbase(path: str):
    """
    Definition linkbase(_def.xml) 파싱.

    Returns
    -------
    def_map : {role_code → {table_name → {axis_name → set[member_names]}}}
      - 표(축) 안에 어떤 멤버들이 있는지 — flat set (소속 여부만 확인 가능)
    lineitem_map : {role_code → {table_name → set[lineitem_names]}}
      - all arcrole(lineitem_abstract → table)와 domain-member 순회로 구성
      - 비어있는 경우(all arcrole 없음) → 해당 표 LINEITEM 필터링 안 함
    def_edge_map : {role_code → {table_name → {axis_name → {child_name → {유효 parent 이름들}}}}}
      - 각 멤버의 실제 부모-자식 edge(domain-member arc)를 보존
      - flat set과 달리 "이 멤버가 표에 속하는가"가 아니라
        "이 멤버가 이 부모 밑에 있는 것이 def.xml상 맞는가"를 검증할 수 있음
    """
    tree = ET.parse(path)
    root = tree.getroot()

    ARCROLE_ALL     = 'http://xbrl.org/int/dim/arcrole/all'
    ARCROLE_HC_DIM  = 'http://xbrl.org/int/dim/arcrole/hypercube-dimension'
    ARCROLE_DIM_DOM = 'http://xbrl.org/int/dim/arcrole/dimension-domain'
    ARCROLE_DOM_MEM = 'http://xbrl.org/int/dim/arcrole/domain-member'

    def_map      = defaultdict(lambda: defaultdict(lambda: defaultdict(set)))
    def_edge_map = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(set))))
    lineitem_map = defaultdict(lambda: defaultdict(set))
    arc_map      = defaultdict(dict)   # {role_code → {(from_name, to_name) → arcrole_short}}

    _ARCROLE_SHORT = {
        ARCROLE_ALL:     'all',
        ARCROLE_HC_DIM:  'hc-dim',
        ARCROLE_DIM_DOM: 'dim-dom',
        ARCROLE_DOM_MEM: 'dom-mem',
    }

    for dl in root.iter(f"{{{NS['link']}}}definitionLink"):
        role_uri  = dl.get(XLINK_ROLE, '')
        role_code = _role_code_from_uri(role_uri)

        locs = {}
        for loc in dl.findall(f"{{{NS['link']}}}loc"):
            lbl  = loc.get(XLINK_LABEL, '')
            href = loc.get(XLINK_HREF, '')
            cid  = href.split('#', 1)[1] if '#' in href else href
            name = cid.split('_', 1)[1]  if '_' in cid  else cid
            locs[lbl] = name

        hc_dim  = []
        dim_dom = []
        dom_mem = []
        all_arc = []

        for arc in dl.findall(f"{{{NS['link']}}}definitionArc"):
            arcrole = arc.attrib.get(
                '{http://www.w3.org/1999/xlink}arcrole',
                arc.attrib.get('arcrole', ''))
            frm = locs.get(arc.get(XLINK_FROM, ''), '')
            to  = locs.get(arc.get(XLINK_TO,   ''), '')
            if not frm or not to:
                continue
            if arcrole == ARCROLE_ALL:
                all_arc.append((frm, to))
            elif arcrole == ARCROLE_HC_DIM:
                hc_dim.append((frm, to))
            elif arcrole == ARCROLE_DIM_DOM:
                dim_dom.append((frm, to))
            elif arcrole == ARCROLE_DOM_MEM:
                dom_mem.append((frm, to))
            short = _ARCROLE_SHORT.get(arcrole)
            if short and frm and to:
                arc_map[role_code][(frm, to)] = short

        if not hc_dim and not all_arc:
            continue

        axis_to_dom = {ax: dom for ax, dom in dim_dom}

        children = defaultdict(list)
        for par, ch in dom_mem:
            children[par].append(ch)

        def collect(start):
            seen, stack = set(), [start]
            while stack:
                node = stack.pop()
                if node in seen:
                    continue
                seen.add(node)
                stack.extend(children.get(node, []))
            return seen

        # 축-멤버 맵 (기존, flat set) + 부모-자식 edge 맵 (신규)
        for tbl, ax in hc_dim:
            dom = axis_to_dom.get(ax, '')
            if not dom:
                continue
            members = collect(dom)
            def_map[role_code][tbl][ax].update(members)

            edge_map = def_edge_map[role_code][tbl][ax]
            edge_map[ax].add(dom)   # 축 → 도메인
            for parent, child in dom_mem:
                if parent in members and child in members:
                    edge_map[parent].add(child)

        # LINEITEM 맵 (신규): all arcrole abstract → table
        for abstract, tbl in all_arc:
            lineitem_map[role_code][tbl].update(collect(abstract))

    def_map_out = {rc: {tbl: dict(axes) for tbl, axes in tbls.items()}
                   for rc, tbls in def_map.items()}
    def_edge_map_out = {
        rc: {tbl: {ax: {par: set(chs) for par, chs in edges.items()}
                   for ax, edges in axes.items()}
             for tbl, axes in tbls.items()}
        for rc, tbls in def_edge_map.items()
    }
    lineitem_map_out = {rc: dict(tbls) for rc, tbls in lineitem_map.items()}
    arc_map_out = dict(arc_map)
    return def_map_out, lineitem_map_out, def_edge_map_out, arc_map_out


def _add_axis_group_fields(rows: list,
                           def_map: dict | None = None) -> None:
    """3-1, 3-2 체크용 축-도메인 그룹핑 필드 추가.

    def_map 있으면 표 진입 시 member→axis 역방향 인덱스를 구성하고,
    Member/Domain 행의 Axis_Name을 presentation 순서가 아닌 def.xml 기반으로
    결정한다. 여러 축이 있는 표에서 멤버가 다른 축 블록 뒤에 나타나도 정확히
    귀속된다. axis_to_group 딕셔너리로 축별 GroupID도 정확히 배정.

    def_map : {role_uri → {table_name → {axis_name → set[member_names]}}}
    def_edge_map : {role_uri → {table_name → {axis_name → {child → {유효 parent들}}}}}
    """
    groups: dict[str, list] = defaultdict(list)
    for i, row in enumerate(rows):
        groups[row.get('role_uri', '')].append((i, row))

    for _, indexed_rows in groups.items():
        if not indexed_rows:
            continue
        role_code          = indexed_rows[0][1].get('role_code', '')
        role_group_counter = 0
        current_table_name = ''
        axis_to_group: dict[str, int]        = {}  # axis_name → group_id
        mem_to_axes:   dict[str, list[str]]  = {}  # member_name → [axis_name, ...]

        prev_element     = None
        prev_axis_domain = None
        prev_group_id    = None
        prev_axis_name   = None

        for orig_idx, row in indexed_rows:
            element = row.get('Element', '')
            name    = row.get('Name', '')

            # 새 표 진입 → def_map 기반 역방향 인덱스 재구성
            if element == 'Table':
                current_table_name = name
                axis_to_group = {}
                mem_to_axes   = {}
                if def_map is not None:
                    for ax, members in def_map.get(role_code, {}).get(current_table_name, {}).items():
                        for m in members:
                            mem_to_axes.setdefault(m, []).append(ax)

            # axis_domain: presentation 순서 기반 (도메인/멤버 계층 시각 구조)
            if prev_element == 'Axis' and element in ('Member', 'Domain'):
                axis_domain = '도메인'
            elif element == 'Axis':
                axis_domain = '축'
            elif element in ('Member', 'Domain') and prev_axis_domain in ('축', '도메인', '멤버'):
                axis_domain = '멤버'
            else:
                axis_domain = None

            axis_flag = 1 if axis_domain == '축' else 0

            # axis_name: def_map 역방향 인덱스로 결정 (다중 축 정확성)
            if element == 'Axis':
                axis_name = name
            elif axis_domain in ('도메인', '멤버'):
                candidates = mem_to_axes.get(name, [])
                if len(candidates) == 1:
                    axis_name = candidates[0]
                elif prev_axis_name in candidates:
                    axis_name = prev_axis_name  # 현재 축 유지 (ambiguous)
                elif candidates:
                    axis_name = candidates[0]
                else:
                    axis_name = prev_axis_name  # def_map 미등록 → presentation fallback
            else:
                axis_name = None

            # group_id: axis_to_group 딕셔너리로 축별 배정
            if axis_domain is None:
                group_id = None
            elif element == 'Axis':
                role_group_counter += 1
                group_id = role_group_counter
                axis_to_group[name] = group_id
            else:
                group_id = axis_to_group.get(axis_name, prev_group_id) if axis_name else prev_group_id

            # ── def.xml 유효성 검증: 소속 여부 + 부모 검증 ──
            store_axis_domain = axis_domain
            store_group_id    = group_id
            if axis_domain == '멤버' and def_map is not None:
                tbl_axes = def_map.get(role_code, {}).get(current_table_name, {})
                if tbl_axes:
                    axis_members = tbl_axes.get(axis_name) if axis_name else None
                    if axis_members is not None and name not in axis_members:
                        store_axis_domain = None
                        store_group_id    = None

            # KEY 생성 ('도메인'은 Axis의 첫 번째 자식으로 참조 목록에 없으므로 제외)
            if store_axis_domain == '축':
                key = f"{axis_name}-{axis_name}" if axis_name else None
            elif store_axis_domain == '멤버':
                key = f"{axis_name}-{name}" if axis_name else None
            else:
                key = None

            rows[orig_idx].update({
                '축_도메인':        store_axis_domain,
                'Axis_flag':       axis_flag,
                'Axis_Name':       axis_name,
                'GroupID':         store_group_id,
                'KEY_axis':        key,
                'xbrl_table_name': current_table_name,
            })

            prev_element     = element
            prev_axis_domain = axis_domain
            prev_group_id    = group_id
            prev_axis_name   = axis_name


def _remap_gubn(rows: list) -> None:
    """최종 구분(gubn) 재분류.
    - DataType이 domainItemType → DOMAIN
    - Name이 'Axis'로 끝남 → DOMAIN
    - Name이 'Table'로 끝남 → TABLE
    - Name이 'TextBlock'으로 끝남 → FOOTNOTES
    - Name이 'Abstract'로 끝남 → FOOTNOTES
      단, entity 확장 항목이면서 XSD상 abstract=False인 경우는 예외로 LINEITEM 유지.
      (이름만 'Abstract'로 끝나는 실제 데이터 항목 — 구조(XSD abstract 속성) 기준 우선)
      다만 바로 다음 행이 depth+1이면서 Name이 'Table'로 끝나는 경우(= 이 항목이
      Table을 직속으로 감싸는 그룹핑 위치)라면 XSD abstract 속성이 잘못
      기재됐을 가능성이 높으므로 FOOTNOTES로 유지한다.
    - entity 확장 stringItemType이면서, 자신의 depth가 같은 role 내 직전에
      등장한 TABLE의 depth보다 얕음(= Abstract/Table 구조의 형제로 붙은
      글주석이며 실제로는 그 표 소속이 아님) → FOOTNOTES
    - 나머지 → LINEITEM
    """
    last_table_depth_by_role: dict[str, int] = {}

    for idx, row in enumerate(rows):
        name      = row.get('Name', '')
        dtype     = row.get('DataType', '')
        dtype     = dtype.split(':')[-1] if ':' in dtype else dtype
        role      = row.get('role_uri', '')
        is_entity = row.get('Prefix', '').startswith('entity')

        if dtype == 'domainItemType':
            row['구분'] = 'DOMAIN'
        elif _is_text_block_like_concept(name, dtype, row.get('Label(KO)', ''), row.get('Label(EN)', '')):
            row['구분'] = 'FOOTNOTES'
            if row.get('Element') == 'item':
                row['Element'] = 'TextBlock'
        elif name.endswith('Axis'):
            row['구분'] = 'DOMAIN'
        elif name.endswith('Table'):
            row['구분'] = 'TABLE'
            last_table_depth_by_role[role] = row.get('depth', 0)
        elif name.endswith('TextBlock'):
            row['구분'] = 'FOOTNOTES'
        elif name.endswith('Abstract'):
            next_row = rows[idx + 1] if idx + 1 < len(rows) else None
            is_table_parent = (next_row is not None
                                and next_row.get('role_uri') == role
                                and next_row.get('depth', -1) == row.get('depth', 0) + 1
                                and next_row.get('Name', '').endswith('Table'))
            if is_entity and not row.get('abstract', False) and not is_table_parent:
                row['구분'] = 'LINEITEM'
            else:
                row['구분'] = 'FOOTNOTES'
        elif (dtype == 'stringItemType'
                and is_entity
                and (role not in last_table_depth_by_role
                     or row.get('depth', 0) < last_table_depth_by_role[role])):
            row['구분'] = 'FOOTNOTES'
        else:
            row['구분'] = 'LINEITEM'


def _remove_def_invalid_rows(rows: list) -> list:
    """def.xml에 등록되지 않은 Domain/Member 행 제거.

    축_도메인=None인 Element='Domain'/'Member' 행만 제거한다.
    Lineitem은 def.xml의 domain-member 아크가 presentation보다 얕게 정의되는 경우
    (ex. UnappropriatedRetainedEarnings 하위 5개 항목)에 실제 유효 항목이
    잘못 제거되므로 필터링하지 않는다.
    """
    def _is_invalid(row):
        if row.get('축_도메인') is None and row.get('Element') in ('Domain', 'Member'):
            return True
        return False

    return [r for r in rows if not _is_invalid(r)]


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
    def_map:      dict | None = None,
    def_edge_map: dict | None = None,
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

        consol_str = '-'
        if code:
            if code[-1] == '0':   consol_str = '연결'
            elif code[-1] == '5': consol_str = '별도'

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
            # DataType/Label 기반 TextBlock 판별 (이름 suffix만으로 놓치는 경우 보완)
            if _is_text_block_like_concept(name, dtype, lbl_ko, lbl_en):
                gubn = 'FOOTNOTES'
                if element == 'item':
                    element = 'TextBlock'
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

            is_c = True if consol_str == '연결' else (False if consol_str == '별도' else None)
            return {
                'role_uri':        role_uri,
                'role_code':       code,
                'role_name_ko':    name_ko,
                'role_name_en':    name_en,
                'is_consolidated': is_c,
                'Role Definition': role_def_str,
                'Sheet':           code,
                '연결/별도':        consol_str,
                'Table_Number':    table_num,
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

        def _concept_name_from_loc(loc_label: str) -> str:
            cid = loc_to_id.get(loc_label, "")
            concept = concept_map.get(cid, {})
            return concept.get("name") or (cid.split("_", 1)[-1] if "_" in cid else cid)

        def _concept_element_from_loc(loc_label: str) -> str:
            return _classify_element(_concept_name_from_loc(loc_label))

        def _is_valid_axis_member(table_name: str, axis_name: str, member_name: str) -> bool:
            if not def_map or not code or not table_name or not axis_name or not member_name:
                return True
            tbl_axes = def_map.get(code, {}).get(table_name, {})
            if not tbl_axes:
                return True
            valid_members = tbl_axes.get(axis_name)
            if valid_members is None:
                return True
            return member_name in valid_members

        def _is_valid_axis_child(table_name: str, axis_name: str,
                                 parent_dim_name: str, child_name: str) -> bool:
            if not def_edge_map or not code:
                return True
            if not table_name or not axis_name or not parent_dim_name or not child_name:
                return True
            tbl_axes = def_edge_map.get(code, {}).get(table_name, {})
            if not tbl_axes:
                return True
            axis_edges = tbl_axes.get(axis_name)
            if not axis_edges:
                return True
            valid_children = axis_edges.get(parent_dim_name)
            if valid_children is None:
                return False
            return child_name in valid_children

        def dfs(loc_label: str, depth: int, order: float | None, pref_url: str,
                par_name: str, par_lbl_ko: str, par_gubn: str,
                current_table_name: str = "", current_axis_name: str = "",
                current_dim_parent: str = "") -> None:
            cid = loc_to_id.get(loc_label, "")
            if order is not None:
                order_val: Any = int(order) if float(order).is_integer() else order
            else:
                order_val = ""

            row = make_row(cid, depth, order_val, pref_url, par_name, par_lbl_ko, par_gubn)

            row_name    = row.get('Name', '')
            row_element = row.get('Element', '')

            next_table_name = current_table_name
            next_axis_name  = current_axis_name
            next_dim_parent = current_dim_parent

            if row_element == 'Table':
                next_table_name = row_name
                next_axis_name  = ''
                next_dim_parent = ''
            elif row_element == 'Axis':
                next_axis_name  = row_name
                next_dim_parent = row_name
            elif row_element in ('Domain', 'Member') and current_axis_name:
                next_dim_parent = row_name

            if row_name not in _CONSOL_SEPARATE_NAMES:
                rows.append(row)

            for to_lbl, o, p in children.get(loc_label, []):
                child_name    = _concept_name_from_loc(to_lbl)
                child_element = _concept_element_from_loc(to_lbl)

                if next_table_name and next_axis_name and child_element in ('Domain', 'Member'):
                    if not _is_valid_axis_member(next_table_name, next_axis_name, child_name):
                        continue
                    if next_dim_parent:
                        if not _is_valid_axis_child(next_table_name, next_axis_name,
                                                    next_dim_parent, child_name):
                            continue

                dfs(to_lbl, depth + 1, o, p,
                    row['Name'], row['Label(KO)'], row['구분'],
                    next_table_name, next_axis_name, next_dim_parent)

        for r in roots:
            dfs(r, 0, None, "", "", "", "")

    return rows, elements


# ── 회사명 / 기간 추출 ────────────────────────────────────────────────────────

_PERIOD_SUFFIX_MAP = {
    'FY':  '4Q',
    'FQA': '1Q', 'FQQ': '1Q',
    'HY':  '2Q', 'HYA': '2Q', 'HYQ': '2Q',
    'TQA': '3Q', 'TQQ': '3Q',
}
_CTX_RE = re.compile(r'^CFY(\d{4})([de])([A-Z]+)')


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
    당기(C) 컨텍스트 ID에서 회계연도·분기·결산월 추출.
    반환: (fy, period)  예) ('FY25(12)', '1Q') / ('FY25(12)', 'Annual')

    결산월은 컨텍스트의 startDate(회계연도 시작일)로 역산한다.
    DART 컨텍스트는 항상 회계연도 시작일부터의 누적기간이므로
    startDate 월의 전월이 결산월이다 (1월 시작 → 전년도 12월말 결산).
    """
    fy_str = ''
    period = ''
    try:
        for _, el in ET.iterparse(xbrl_path, events=("end",)):
            if not el.tag.endswith('}context'):
                continue
            m = _CTX_RE.match(el.get('id', ''))
            if not m:
                el.clear()
                continue

            year, de, suffix = m.group(1)[2:], m.group(2), m.group(3)
            if not fy_str:
                fy_str = f'FY{year}'
                period = _PERIOD_SUFFIX_MAP.get(suffix, suffix)

            # 'd'(duration) 컨텍스트만 startDate를 가지므로 그 경우에만 결산월 시도
            if de == 'd' and '(' not in fy_str:
                start_el = el.find(f".//{{{NS['xbrli']}}}startDate")
                if start_el is not None and start_el.text:
                    try:
                        start_month = int(start_el.text.strip()[5:7])
                        fy_end_month = 12 if start_month == 1 else start_month - 1
                        fy_str = f'FY{year}({fy_end_month:02d})'
                    except (ValueError, IndexError):
                        pass

            if '(' in fy_str:
                return fy_str, period
            el.clear()
    except Exception:
        pass
    return fy_str, period


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
        xbrl_path: str | None = None
        try:
            xbrl_path = _find(tmp_dir, ".xbrl")
            data.company_name = _extract_company_name(xbrl_path)
            data.fy, data.report_period = _extract_period_info(xbrl_path)
        except Exception:
            pass

        concept_map, role_def_map = _parse_xsd(xsd_path)
        labels_ko = _parse_labels(lab_ko_path)
        labels_en = _parse_labels(lab_en_path)

        # def.xml 기반 멤버 필터링 (presentation 파싱 전에 먼저 로드)
        def_map      = None
        def_edge_map = None
        arc_map      = {}
        try:
            def_path = _find(tmp_dir, "_def.xml")
            def_map, _, def_edge_map, arc_map = _parse_def_linkbase(def_path)
        except Exception:
            pass

        rows, elements = _parse_presentation(
            pre_path, labels_ko, labels_en, role_def_map, concept_map,
            def_map, def_edge_map
        )

        _add_axis_group_fields(rows, def_map)
        _postprocess_table_name(rows)
        _remap_gubn(rows)
        rows = _remove_def_invalid_rows(rows)

        # def.xml arcrole 부착 (parent-name 쌍으로 매칭)
        for r in rows:
            rc     = r.get('role_code', '')
            parent = r.get('parent', '')
            name   = r.get('Name', '')
            r['def_arcrole'] = arc_map.get(rc, {}).get((parent, name), '')

        data.presentation_rows = rows
        data.elements          = elements

        # GroupID가 있는 행만 추출
        # (role_uri, xbrl_table_name, Axis_Name, Name) 기준 중복 제거
        # → 같은 XBRL table이 Presentation에 두 번 나타나는 경우 방지
        seen: set[tuple] = set()
        axis_domain: list[dict] = []
        for r in rows:
            if r.get('GroupID') is None:
                continue
            dedup_key = (
                r.get('role_uri', ''),
                r.get('xbrl_table_name', ''),
                r.get('Axis_Name', ''),
                r.get('Name', ''),
            )
            if dedup_key in seen:
                continue
            seen.add(dedup_key)
            axis_domain.append(r)
        data.axis_domain_rows = axis_domain

    except Exception as e:
        data.errors.append(str(e))
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return data
