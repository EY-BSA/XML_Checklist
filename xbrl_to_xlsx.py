"""
xbrl_to_xlsx.py
────────────────────────────────────────────────────────────────────────────
XBRL 원문 ZIP (.xsd + _pre.xml + _lab-ko.xml + _lab-en.xml + .xbrl)
   ↓
'사업보고서_IFRS(원문XBRL)' 형식과 동일한 4-시트 xlsx 생성.

생성되는 시트:
  1. Concepts      — xsd 의 element 정의 1:1
                     [id, name, nillable, substitutionGroup, type, balance,
                      periodType, abstract]
  2. Presentation  — 모든 presentation row (role/depth 트리 평탄화)
                     [role, id, label-ko, label-en, depth, order, pref_label,
                      DataType, Period, Decimal]                     ← 3컬럼 추가
  3. Label-ko      — concept_id × 12 label role
                     [id, label, dart_label, terseLabel, negatedLabel,
                      verboseLabel, totalLabel, negatedTerseLabel, netLabel,
                      periodStartLabel, commentaryGuidance, periodEndLabel,
                      negatedTotalLabel]
  4. Label-en      — 위와 동일 컬럼, 영문 라벨
"""
from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# ───────────────────────────────────────────────────────────────────────────
# 네임스페이스
# ───────────────────────────────────────────────────────────────────────────
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

DEFAULT_LABEL_ROLE = "http://www.xbrl.org/2003/role/label"

# ───────────────────────────────────────────────────────────────────────────
# Label role → xlsx 컬럼 매핑 (reference xlsx 와 동일 순서)
# ───────────────────────────────────────────────────────────────────────────
LABEL_ROLE_COLUMNS: list[tuple[str, str]] = [
    # (컬럼 헤더, role URI suffix or 정확한 URI)
    ("label",              "http://www.xbrl.org/2003/role/label"),
    ("dart_label",         "/dart_label"),                                          # DART 확장 (보고서마다 prefix 다름)
    ("terseLabel",         "http://www.xbrl.org/2003/role/terseLabel"),
    ("negatedLabel",       "http://www.xbrl.org/2009/role/negatedLabel"),
    ("verboseLabel",       "http://www.xbrl.org/2003/role/verboseLabel"),
    ("totalLabel",         "http://www.xbrl.org/2003/role/totalLabel"),
    ("negatedTerseLabel",  "http://www.xbrl.org/2009/role/negatedTerseLabel"),
    ("netLabel",           "http://www.xbrl.org/2009/role/netLabel"),
    ("periodStartLabel",   "http://www.xbrl.org/2003/role/periodStartLabel"),
    ("commentaryGuidance", "http://www.xbrl.org/2003/role/commentaryGuidance"),
    ("periodEndLabel",     "http://www.xbrl.org/2003/role/periodEndLabel"),
    ("negatedTotalLabel",  "http://www.xbrl.org/2009/role/negatedTotalLabel"),
]


# ───────────────────────────────────────────────────────────────────────────
# 공통 유틸
# ───────────────────────────────────────────────────────────────────────────
def href_to_id(href: str) -> str:
    return href.split("#", 1)[1] if "#" in href else href


def _match_label_column(role_uri: str, col_uri: str) -> bool:
    """role_uri 가 col_uri 와 일치하는지 (dart_label 은 suffix 매칭)."""
    if col_uri.startswith("http"):
        return role_uri == col_uri
    # suffix 매칭 (예: '/dart_label')
    return role_uri.endswith(col_uri)


# ───────────────────────────────────────────────────────────────────────────
# 1) xsd 파싱 → Concepts 시트 데이터
# ───────────────────────────────────────────────────────────────────────────
def parse_xsd(path: str) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]], dict[str, str]]:
    """
    Returns
    -------
    concepts_rows : list of dict — Concepts 시트에 1:1
    concepts_by_id: dict[id → row] — 빠른 조회용
    role_def      : dict[roleURI → definition('[Dxxx] 한글 | English')]
    """
    tree = ET.parse(path)
    root = tree.getroot()

    rows: list[dict[str, Any]] = []
    by_id: dict[str, dict[str, Any]] = {}

    for el in root.iter(f"{{{NS['xsd']}}}element"):
        cid = el.get("id", "")
        if not cid:
            continue

        # substitutionGroup, type 는 'xbrli:xxx' / 'xbrldt:xxx' 형태 그대로 유지
        sub_group = el.get("substitutionGroup", "")
        ctype     = el.get("type", "")
        balance   = el.get(f"{{{NS['xbrli']}}}balance", "")
        period    = el.get(f"{{{NS['xbrli']}}}periodType", "")
        abstract  = el.get("abstract", "false")
        nillable  = el.get("nillable", "false")

        row = {
            "id":                cid,
            "name":              el.get("name", ""),
            "nillable":          nillable,
            "substitutionGroup": sub_group,
            "type":              ctype,
            "balance":           balance if balance else None,
            "periodType":        period if period else None,
            "abstract":          abstract,
        }
        rows.append(row)
        by_id[cid] = row

    role_def: dict[str, str] = {}
    for rt in root.iter(f"{{{NS['link']}}}roleType"):
        uri = rt.get("roleURI", "")
        defn = rt.find(f"{{{NS['link']}}}definition")
        if uri and defn is not None and defn.text:
            role_def[uri] = defn.text.strip()

    return rows, by_id, role_def


# ───────────────────────────────────────────────────────────────────────────
# 2) Label linkbase 파싱 → {concept_id: {role_uri: label_text}}
# ───────────────────────────────────────────────────────────────────────────
def parse_labels(path: str) -> dict[str, dict[str, str]]:
    tree = ET.parse(path)
    root = tree.getroot()

    loc_to_id: dict[str, str] = {}
    res_to_role_text: dict[str, tuple[str, str]] = {}
    arcs: list[tuple[str, str]] = []

    for ll in root.iter(f"{{{NS['link']}}}labelLink"):
        for loc in ll.findall(f"{{{NS['link']}}}loc"):
            href = loc.get(XLINK_HREF, "")
            lbl  = loc.get(XLINK_LABEL, "")
            cid  = href_to_id(href)
            if cid and lbl:
                loc_to_id[lbl] = cid
        for lab in ll.findall(f"{{{NS['link']}}}label"):
            lbl  = lab.get(XLINK_LABEL, "")
            role = lab.get(XLINK_ROLE, "")
            text = lab.text or ""
            if lbl:
                res_to_role_text[lbl] = (role, text)
        for arc in ll.findall(f"{{{NS['link']}}}labelArc"):
            f = arc.get(XLINK_FROM, "")
            t = arc.get(XLINK_TO, "")
            if f and t:
                arcs.append((f, t))

    labels: dict[str, dict[str, str]] = defaultdict(dict)
    for f, t in arcs:
        cid = loc_to_id.get(f)
        rt  = res_to_role_text.get(t)
        if cid and rt:
            role, text = rt
            labels[cid][role] = text
    return labels


def pick_label(labels_for_cid: dict[str, str], pref_role: str | None) -> str:
    """presentation row 의 label-ko / label-en 용 선택"""
    if not labels_for_cid:
        return ""
    if pref_role and pref_role in labels_for_cid:
        return labels_for_cid[pref_role]
    if DEFAULT_LABEL_ROLE in labels_for_cid:
        return labels_for_cid[DEFAULT_LABEL_ROLE]
    # fallback: 첫 번째 라벨
    return next(iter(labels_for_cid.values()))


# ───────────────────────────────────────────────────────────────────────────
# 3) Presentation linkbase → 평탄화된 row 리스트
# ───────────────────────────────────────────────────────────────────────────
def _role_label_only_ko(role_definition: str) -> str:
    """
    '[D210000] 재무상태표, 유동/비유동법 - 연결 | Statement ...'
    → '[D210000] 재무상태표, 유동/비유동법 - 연결'
    """
    if not role_definition:
        return ""
    return role_definition.split("|", 1)[0].strip()


def parse_presentation(
    path: str,
    labels_ko: dict[str, dict[str, str]],
    labels_en: dict[str, dict[str, str]],
    role_def_map: dict[str, str],
) -> list[dict[str, Any]]:
    tree = ET.parse(path)
    root = tree.getroot()
    rows: list[dict[str, Any]] = []

    for pl in root.iter(f"{{{NS['link']}}}presentationLink"):
        role_uri    = pl.get(XLINK_ROLE, "")
        role_def    = role_def_map.get(role_uri, role_uri)
        role_label  = _role_label_only_ko(role_def)

        loc_to_id: dict[str, str] = {}
        for loc in pl.findall(f"{{{NS['link']}}}loc"):
            loc_to_id[loc.get(XLINK_LABEL, "")] = href_to_id(loc.get(XLINK_HREF, ""))

        children: dict[str, list[tuple[str, float, str]]] = defaultdict(list)
        sources: set[str] = set()
        targets: set[str] = set()
        for arc in pl.findall(f"{{{NS['link']}}}presentationArc"):
            f = arc.get(XLINK_FROM, "")
            t = arc.get(XLINK_TO, "")
            try:
                order = float(arc.get("order", "0"))
            except ValueError:
                order = 0.0
            pref = arc.get("preferredLabel", "") or ""
            children[f].append((t, order, pref))
            sources.add(f); targets.add(t)
        for f in children:
            children[f].sort(key=lambda x: x[1])

        loc_order = {loc.get(XLINK_LABEL, ""): i
                     for i, loc in enumerate(pl.findall(f"{{{NS['link']}}}loc"))}
        roots = [lbl for lbl in loc_order
                 if lbl in sources and lbl not in targets]
        roots.sort(key=lambda x: loc_order.get(x, 0))

        def order_to_excel(o: float | None) -> Any:
            if o is None:
                return None
            if float(o).is_integer():
                return int(o)
            return o

        def emit(loc_label: str, depth: int, order_val: float | None, pref_role: str) -> None:
            cid = loc_to_id.get(loc_label, "")
            ko = pick_label(labels_ko.get(cid, {}), pref_role or None)
            en = pick_label(labels_en.get(cid, {}), pref_role or None)
            rows.append({
                "role":       role_label,
                "id":         cid,
                "label-ko":   ko,
                "label-en":   en,
                "depth":      depth,
                "order":      order_to_excel(order_val),
                "pref_label": pref_role or None,
            })

        def dfs(loc_label: str, depth: int, order_val: float | None, pref_role: str) -> None:
            emit(loc_label, depth, order_val, pref_role)
            for to_lbl, o, p in children.get(loc_label, []):
                dfs(to_lbl, depth + 1, o, p)

        for r in roots:
            dfs(r, 0, None, "")

    return rows


# ───────────────────────────────────────────────────────────────────────────
# 4) .xbrl 인스턴스 → fact 사전 (Decimal 채우기용)
# ───────────────────────────────────────────────────────────────────────────
def parse_instance_facts(xbrl_path: str) -> dict[str, list[dict[str, str]]]:
    """
    .xbrl 인스턴스에서 각 concept 의 fact 목록을 추출.

    Returns
    -------
    facts : dict[concept_name(=local name), list[{contextRef, decimals, value, unitRef}]]
    """
    facts: dict[str, list[dict[str, str]]] = defaultdict(list)
    if not xbrl_path or not Path(xbrl_path).exists():
        return facts

    NON_FACT_NS = {
        NS["xbrli"], NS["link"], NS["xlink"], NS["xsd"], NS["xml"],
        "http://xbrl.org/2006/xbrldi",
    }

    try:
        for _, el in ET.iterparse(xbrl_path, events=("end",)):
            tag = el.tag
            if not tag.startswith("{"):
                continue
            ns_uri, local = tag[1:].split("}", 1)
            if ns_uri in NON_FACT_NS:
                el.clear()
                continue
            ctx = el.get("contextRef")
            if not ctx:
                el.clear()
                continue
            facts[local].append({
                "contextRef": ctx,
                "decimals":   el.get("decimals", "") or "",
                "unitRef":    el.get("unitRef", "") or "",
                "value":      (el.text or "").strip(),
            })
            el.clear()
    except ET.ParseError:
        pass

    return facts


def pick_decimal(name: str, facts: dict[str, list[dict[str, str]]]) -> str:
    """
    당기(C로 시작) 컨텍스트 우선, 그 중 가장 작은(most rounded) decimals 반환.
    """
    fact_list = facts.get(name)
    if not fact_list:
        return ""
    candidates = [f for f in fact_list if f["contextRef"].startswith("C")] or fact_list
    numeric = []
    for f in candidates:
        try:
            numeric.append(int(f["decimals"]))
        except (ValueError, KeyError, TypeError):
            pass
    return str(min(numeric)) if numeric else ""


# ───────────────────────────────────────────────────────────────────────────
# 5) ZIP 안에서 XBRL 파일 자동 탐지
# ───────────────────────────────────────────────────────────────────────────
def _resolve_xbrl_files(directory: str) -> dict[str, str]:
    result = {"xsd": "", "pre": "", "lab_ko": "", "lab_en": "", "xbrl": ""}
    for p in Path(directory).rglob("*"):
        if not p.is_file():
            continue
        name = p.name.lower()
        if name.endswith("_pre.xml"):
            result["pre"] = str(p)
        elif name.endswith("_lab-ko.xml") or name.endswith("_lab_ko.xml"):
            result["lab_ko"] = str(p)
        elif name.endswith("_lab-en.xml") or name.endswith("_lab_en.xml"):
            result["lab_en"] = str(p)
        elif name.endswith(".xbrl"):
            result["xbrl"] = str(p)
        elif name.endswith(".xsd") and not result["xsd"]:
            result["xsd"] = str(p)

    missing = [k for k in ("xsd", "pre", "lab_ko", "lab_en") if not result[k]]
    if missing:
        raise FileNotFoundError(
            "ZIP 안에서 다음 필수 파일을 찾지 못했습니다: " + ", ".join(missing)
        )
    return result


# ───────────────────────────────────────────────────────────────────────────
# 6) xlsx 작성
# ───────────────────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="FFD9E1F2")
BOLD = Font(bold=True)


def _split_concept_id(cid: str) -> tuple[str, str]:
    if not cid or "_" not in cid:
        return "", cid or ""
    pre, rest = cid.split("_", 1)
    return pre, rest


def _write_concepts(ws, concept_rows: list[dict[str, Any]]) -> None:
    headers = ["id", "name", "nillable", "substitutionGroup",
               "type", "balance", "periodType", "abstract"]
    ws.append(headers)
    for c in ws[1]:
        c.font = BOLD
        c.fill = HEADER_FILL
    for r in concept_rows:
        ws.append([r.get(h) for h in headers])
    # 컬럼 폭
    widths = {"A": 64, "B": 50, "C": 10, "D": 22, "E": 26,
              "F": 10, "G": 12, "H": 10}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


def _write_presentation(
    ws,
    presentation_rows: list[dict[str, Any]],
    concepts_by_id: dict[str, dict[str, Any]],
    facts: dict[str, list[dict[str, str]]],
) -> None:
    headers = ["role", "id", "label-ko", "label-en", "depth", "order", "pref_label",
               "DataType", "Period", "Decimal"]   # ← 마지막 3개가 추가 컬럼
    ws.append(headers)
    for c in ws[1]:
        c.font = BOLD
        c.fill = HEADER_FILL

    for row in presentation_rows:
        cid = row["id"]
        info = concepts_by_id.get(cid, {})
        ctype = info.get("type", "") or ""
        data_type = ctype.split(":", 1)[-1] if ":" in ctype else ctype
        period = (info.get("periodType") or "").upper() or ""

        _, local_name = _split_concept_id(cid)
        decimal = pick_decimal(local_name, facts)

        ws.append([
            row["role"], cid, row["label-ko"], row["label-en"],
            row["depth"], row["order"], row["pref_label"],
            data_type, period, decimal,
        ])

    widths = {"A": 50, "B": 60, "C": 50, "D": 50, "E": 8,
              "F": 8, "G": 50, "H": 22, "I": 12, "J": 10}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


def _write_label_sheet(
    ws,
    labels: dict[str, dict[str, str]],
    ordered_ids: list[str],
) -> None:
    headers = ["id"] + [col for col, _ in LABEL_ROLE_COLUMNS]
    ws.append(headers)
    for c in ws[1]:
        c.font = BOLD
        c.fill = HEADER_FILL

    # 한 concept 의 라벨이 하나라도 있는 경우만 출력
    for cid in ordered_ids:
        by_role = labels.get(cid)
        if not by_role:
            continue
        out_row: list[Any] = [cid]
        for _, role_match in LABEL_ROLE_COLUMNS:
            # exact URI 또는 suffix 매칭
            value = None
            for ru, txt in by_role.items():
                if _match_label_column(ru, role_match):
                    value = txt
                    break
            out_row.append(value)
        # 모든 라벨이 비어있는 행은 제외
        if any(v not in (None, "") for v in out_row[1:]):
            ws.append(out_row)

    widths = {"A": 60}
    for i, _ in enumerate(LABEL_ROLE_COLUMNS, start=2):
        col_letter = chr(ord("A") + i - 1) if i <= 26 else None
        if col_letter:
            widths[col_letter] = 30
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B2"


# ───────────────────────────────────────────────────────────────────────────
# 7) 메인 엔트리
# ───────────────────────────────────────────────────────────────────────────
def convert_zip_to_xlsx(zip_path: str, out_path: str) -> str:
    """
    XBRL ZIP 파일을 4-시트 xlsx 로 변환.

    Returns: 실제로 저장된 xlsx 경로
    """
    tmp_dir = tempfile.mkdtemp(prefix="xbrl_")
    try:
        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(tmp_dir)

        files = _resolve_xbrl_files(tmp_dir)

        concept_rows, concepts_by_id, role_def = parse_xsd(files["xsd"])
        labels_ko = parse_labels(files["lab_ko"])
        labels_en = parse_labels(files["lab_en"])
        presentation_rows = parse_presentation(
            files["pre"], labels_ko, labels_en, role_def
        )
        facts = parse_instance_facts(files["xbrl"]) if files["xbrl"] else {}

        # Label 시트는 라벨에 등장하는 모든 concept_id 출력
        # 순서: presentation 등장 순서 + 그 외 concept_id 순서
        seen: set[str] = set()
        ordered_ids: list[str] = []
        for r in presentation_rows:
            cid = r["id"]
            if cid and cid not in seen:
                seen.add(cid)
                ordered_ids.append(cid)
        for cid in labels_ko:
            if cid not in seen:
                seen.add(cid)
                ordered_ids.append(cid)
        for cid in labels_en:
            if cid not in seen:
                seen.add(cid)
                ordered_ids.append(cid)

        wb = Workbook()
        wb.remove(wb.active)

        _write_concepts(wb.create_sheet("Concepts"), concept_rows)
        _write_presentation(
            wb.create_sheet("Presentation"),
            presentation_rows,
            concepts_by_id,
            facts,
        )
        _write_label_sheet(wb.create_sheet("Label-ko"), labels_ko, ordered_ids)
        _write_label_sheet(wb.create_sheet("Label-en"), labels_en, ordered_ids)

        Path(out_path).parent.mkdir(parents=True, exist_ok=True)
        wb.save(out_path)
        return out_path

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(
        description="XBRL ZIP → 4-시트 xlsx (Concepts/Presentation/Label-ko/Label-en, "
                    "Presentation 에 DataType·Period·Decimal 컬럼 추가)"
    )
    ap.add_argument("zip_path", help="입력 XBRL zip 파일")
    ap.add_argument("--out", help="출력 xlsx 경로 (기본: zip 과 동일 이름.xlsx)")
    args = ap.parse_args()

    in_path = Path(args.zip_path)
    out = args.out or str(in_path.with_suffix(".xlsx"))
    saved = convert_zip_to_xlsx(str(in_path), out)
    print(f"saved: {saved}")
