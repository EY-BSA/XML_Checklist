"""
xbrl_fact_parser_0708.py
XBRL ZIP → fact 1차원 테이블 (def.xml 유효성 검증 + 계층 분리)

def_parser와 동일하게:
  - def.xml 하이퍼큐브로 유효한 fact만 추출
  - 표/축/멤버/라인아이템 계층을 별도 컬럼으로 분리
  - 정렬: 표 등장 순서 → 멤버(presentation 순서) → 라인아이템

fact_parser 방식 유지:
  - 기간은 CY_fact / PY_fact / BY_fact 컬럼으로 나란히 배치

출력 컬럼:
  Sheet, 주석, 주석(EN), 표, 표(EN), 표ID,
  축, 축(EN), 축ID, 멤버(KO), 멤버(EN), 멤버ID,
  Label(KO), Label(EN), Name, 아이템ID, Arcrole, 연결/별도,
  contextRef, UnitRef, Decimal,
  CY_fact, PY_fact, BY_fact
"""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path
from typing import Dict, List

# def_parser에서 핵심 함수/상수 재사용
from xbrl_def_parser_0708 import (
    parse_xbrl_zip,
    _fact_candidates,
    _current_year,
    _CONSOL_AXIS,
)


# ── 유틸 ──────────────────────────────────────────────────────────────────────
def _note_no(r: dict) -> tuple:
    m = re.match(r'\s*(\d+)(?:-(\d+))?\s*\.', r.get('role_name_ko', ''))
    return (int(m.group(1)), int(m.group(2) or 0)) if m else (0, 0)


def _strip_ctx_prefix(ctx_ref: str) -> str:
    """CFY2025eFY_ 같은 기간 prefix 제거."""
    idx = ctx_ref.find('_')
    return ctx_ref[idx + 1:] if idx != -1 else ctx_ref


# ── 메인 ──────────────────────────────────────────────────────────────────────
def parse_xbrl_facts_to_rows(file_bytes: bytes) -> List[dict]:
    """
    ZIP 바이트 → fact 1차원 행 리스트

    각 행 = 특정 (lineitem, axis/member 조합)의 CY/PY/BY 값
    """
    data = parse_xbrl_zip(file_bytes)
    rows = data.presentation_rows

    max_year = _current_year(data.contexts)
    all_years = sorted({c['end'][:4] for c in data.contexts.values() if c['end']},
                       reverse=True)
    cy_year = all_years[0] if len(all_years) > 0 else ''
    py_year = all_years[1] if len(all_years) > 1 else ''
    by_year = all_years[2] if len(all_years) > 2 else ''

    # fact를 Name별로 인덱싱
    by_name: Dict[str, list] = defaultdict(list)
    for f in data.facts:
        by_name[f['Name']].append(f)

    # presentation 순서 인덱스 (정렬용)
    rows_sorted = sorted(rows, key=lambda r: (_note_no(r), r.get('Sheet', '')))

    table_order:  Dict[tuple, int] = {}
    member_order: Dict[tuple, int] = {}
    label_ko:     Dict[str, str]   = {}
    label_en:     Dict[str, str]   = {}
    name_to_id:   Dict[str, str]   = {}

    for i, r in enumerate(rows_sorted):
        label_ko.setdefault(r['Name'], r['Label(KO)'])
        label_en.setdefault(r['Name'], r['Label(EN)'])
        name_to_id.setdefault(
            r['Name'],
            f"{r['Prefix']}_{r['Name']}" if r.get('Prefix') else r['Name'])
        tkey = (r.get('Sheet', ''), r.get('xbrl_table_name', ''))
        table_order.setdefault(tkey, i)
        if r.get('Element') in ('Domain', 'Member'):
            member_order.setdefault((*tkey, r['Name']), i)

    # ── 핵심: LINEITEM별 (dims 조합) → CY/PY/BY 그룹화 ──────────────────────
    # key: (sort_key, row_dict) → 나중에 sort_key로 정렬
    keyed: list = []

    for li_idx, r in enumerate(rows_sorted):
        if r.get('구분') != 'LINEITEM':
            continue

        tkey  = (r.get('Sheet', ''), r.get('xbrl_table_name', ''))
        tname = r.get('xbrl_table_name', '')

        # def.xml 유효성 검증된 fact 후보 (전기 포함 전체)
        candidates = _fact_candidates(
            r, by_name, data.contexts, max_year,
            data.def_map, all_periods=True)

        # (dims frozenset) → {cy_year: (f,ctx), py_year: ..., by_year: ...}
        dim_groups: Dict[frozenset, Dict[str, tuple]] = defaultdict(dict)
        for f, ctx in candidates:
            dims = {k: v for k, v in ctx['dims'].items() if k != _CONSOL_AXIS}
            fs   = frozenset(dims.items())
            year = ctx['end'][:4]
            dim_groups[fs][year] = (f, ctx)

        for fs, year_map in dim_groups.items():
            dims = dict(fs)

            cy_pair = year_map.get(cy_year)
            py_pair = year_map.get(py_year)
            by_pair = year_map.get(by_year)

            if not (cy_pair or py_pair or by_pair):
                continue

            # contextRef / UnitRef / Decimal: CY 우선, 없으면 PY
            ref_pair = cy_pair or py_pair or by_pair
            ctx_ref  = _strip_ctx_prefix(ref_pair[0]['contextRef'])
            unit_ref = ref_pair[0].get('unitRef', '')
            decimal  = ref_pair[0].get('decimals', '')

            # 정렬키: 표 등장 순서 → 멤버(presentation 순서) → 라인아이템
            mem_key  = tuple(sorted(
                member_order.get((*tkey, m), 10**9) for m in dims.values()))
            sort_key = (table_order.get(tkey, 10**9), mem_key, li_idx)

            row_out = {
                'Sheet':     r.get('Sheet', ''),
                '주석':       r.get('role_name_ko', ''),
                '주석(EN)':   r.get('role_name_en', ''),
                '표':         r.get('table_name_ko', ''),
                '표(EN)':     label_en.get(tname, '') or r.get('role_name_en', ''),
                '표ID':       name_to_id.get(tname, tname),
                '축':         ' | '.join(label_ko.get(ax, ax) for ax in dims),
                '축(EN)':     ' | '.join(label_en.get(ax, ax) for ax in dims),
                '축ID':       ' | '.join(name_to_id.get(ax, ax) for ax in dims),
                '멤버(KO)':   ' | '.join(label_ko.get(m, m) for m in dims.values()),
                '멤버(EN)':   ' | '.join(label_en.get(m, m) for m in dims.values()),
                '멤버ID':     ' | '.join(name_to_id.get(m, m) for m in dims.values()),
                'Label(KO)': r.get('Label(KO)', ''),
                'Label(EN)': r.get('Label(EN)', ''),
                'Name':      r.get('Name', ''),
                '아이템ID':   name_to_id.get(r['Name'], r['Name']),
                'Arcrole':   r.get('Arcrole', ''),
                '연결/별도':   r.get('연결/별도', ''),
                'contextRef': ctx_ref,
                'UnitRef':    unit_ref,
                'Decimal':    decimal,
                'CY_fact':   cy_pair[0]['value'] if cy_pair else '',
                'PY_fact':   py_pair[0]['value'] if py_pair else '',
                'BY_fact':   by_pair[0]['value'] if by_pair else '',
            }
            keyed.append((sort_key, row_out))

    keyed.sort(key=lambda x: x[0])
    return [d for _, d in keyed]


def parse_xbrl_facts_to_df(file_bytes: bytes):
    import pandas as pd
    return pd.DataFrame(parse_xbrl_facts_to_rows(file_bytes))


if __name__ == '__main__':
    import sys
    import pandas as pd

    if len(sys.argv) < 2:
        print('Usage: python xbrl_fact_parser_0708.py <zip_path> [output.xlsx]')
        sys.exit(1)

    zip_path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else \
        str(Path(__file__).parent / f'fact_result_{Path(zip_path).stem[:20]}.xlsx')

    with open(zip_path, 'rb') as f:
        rows = parse_xbrl_facts_to_rows(f.read())

    pd.DataFrame(rows).to_excel(out_path, sheet_name='result', index=False)
    print(f'저장 완료: {out_path}  ({len(rows)}행)')
