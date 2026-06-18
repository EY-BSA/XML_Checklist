"""
web_app.py — XBRL 체크리스트 Flask 웹 애플리케이션
기존 로직(checklist_engine, checklist_export, xbrl_zip_parser, standard_taxonomy)을 그대로 사용합니다.
"""
from __future__ import annotations

import io
import os
import sys
import threading
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from xbrl_zip_parser import parse_xbrl_zip
from checklist_engine import run_all_checks, get_summary, SECTION_LABELS
from standard_taxonomy import StandardTaxonomy, enrich_axis_domain_check, enrich_dart_negate_check
from checklist_export import export_checklist

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB

# gunicorn 으로 실행될 때도 표준 택사노미를 미리 로드 (첫 요청 타임아웃 방지)
threading.Thread(target=_get_std, daemon=True).start()

_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
_TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template', 'XBRL_CoE_Checklist_Result.xlsx')

_std: StandardTaxonomy | None = None
_std_lock = threading.Lock()


def _get_std() -> StandardTaxonomy:
    global _std
    with _std_lock:
        if _std is None:
            _std = StandardTaxonomy()
            axis_path   = os.path.join(_DATA_DIR, 'Axis_Domain_Check.xlsx')
            negate_path = os.path.join(_DATA_DIR, 'DART_Negate_Check.xlsx')
            if os.path.exists(axis_path):
                enrich_axis_domain_check(_std, axis_path)
            if os.path.exists(negate_path):
                enrich_dart_negate_check(_std, negate_path)
    return _std


def _issue_to_dict(iss, chk_id: str) -> dict:
    from checklist_export import _role_def, _tbl_name
    base = {
        'role_def':    _role_def(iss),
        'table_name':  _tbl_name(iss),
        'prefix':      iss.prefix,
        'element':     iss.element_name,
        'label_ko':    iss.label_ko,
        'label_en':    iss.label_en,
        'label_role':  iss.label_role,
        'data_type':   iss.data_type,
        'period':      iss.period,
        'reason':      iss.reason,
        'is_consol':   iss.is_consolidated,
    }
    if chk_id == '7-1':
        base['dart_negate']   = iss.dart_negate
        base['client_negate'] = iss.client_negate
    return base


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/check', methods=['POST'])
def api_check():
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다.'}), 400
    f = request.files['file']
    if not f.filename or not f.filename.lower().endswith('.zip'):
        return jsonify({'error': '.zip 파일만 업로드할 수 있습니다.'}), 400

    file_bytes = f.read()
    try:
        data = parse_xbrl_zip(file_bytes)
    except Exception as e:
        return jsonify({'error': f'ZIP 파싱 오류: {e}'}), 400

    if data.errors:
        return jsonify({'error': '\n'.join(data.errors)}), 400

    std = _get_std()
    results = run_all_checks(data, std)
    summary = get_summary(results)

    company   = data.company_name or (data.entity_id or 'XBRL').split('_')[0]
    fy        = data.fy or ''
    period_   = data.report_period or ''

    checks_json = []
    for chk_id, chk in results.items():
        from checklist_export import _issue_sort_key
        sorted_issues = sorted(chk.issues, key=_issue_sort_key)
        checks_json.append({
            'id':          chk_id,
            'title':       chk.title,
            'description': chk.description,
            'category':    chk.category,
            'issue_count': chk.issue_count,
            'passed':      chk.passed,
            'consol_count': chk.consol_count,
            'sep_count':   chk.sep_count,
            'issues':      [_issue_to_dict(i, chk_id) for i in sorted_issues[:500]],
        })

    return jsonify({
        'company':  company,
        'fy':       fy,
        'period':   period_,
        'summary':  summary,
        'checks':   checks_json,
        'filename': f.filename,
    })


@app.route('/api/export', methods=['POST'])
def api_export():
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다.'}), 400
    f = request.files['file']
    if not f.filename or not f.filename.lower().endswith('.zip'):
        return jsonify({'error': '.zip 파일만 업로드할 수 있습니다.'}), 400

    file_bytes = f.read()
    try:
        data = parse_xbrl_zip(file_bytes)
    except Exception as e:
        return jsonify({'error': f'ZIP 파싱 오류: {e}'}), 400

    if data.errors:
        return jsonify({'error': '\n'.join(data.errors)}), 400

    std = _get_std()
    results = run_all_checks(data, std)

    company   = data.company_name or (data.entity_id or 'XBRL').split('_')[0]
    fy        = data.fy or ''
    period_   = data.report_period or ''
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    prefix = '_'.join(p for p in [fy, period_] if p)
    fname  = (f'{prefix}_XBRL_Checklist_Result_{company}_{timestamp}.xlsx' if prefix
              else f'XBRL_Checklist_Result_{company}_{timestamp}.xlsx')

    import tempfile
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name

    try:
        export_checklist(results, _TEMPLATE_PATH, tmp_path)
        return send_file(
            tmp_path,
            as_attachment=True,
            download_name=fname,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass


if __name__ == '__main__':
    threading.Thread(target=_get_std, daemon=True).start()
    app.run(debug=False, host='0.0.0.0', port=8080)
