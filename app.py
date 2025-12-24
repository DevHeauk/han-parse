#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
한글 파일 파싱 웹 애플리케이션
Flask 기반 웹 인터페이스
HWP (읽기 전용) 및 HWPX (편집 가능) 지원
"""

import os
import json
import tempfile
import shutil
from flask import Flask, request, jsonify, render_template, send_file
from werkzeug.utils import secure_filename
from han_parser import parse_hwp, parse_tables, save_tables_to_json
from table_reconstructor import create_hwp_from_tables_json, edit_table_data
from hwpx_parser import is_hwpx_file, parse_hwpx, save_hwpx_with_tables

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB 제한
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['TEMP_FOLDER'] = 'temp'

# 업로드 폴더 생성
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['TEMP_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'hwp', 'HWP', 'hwpx', 'HWPX'}


def allowed_file(filename):
    """허용된 파일 확장자 확인"""
    if '.' not in filename:
        return False
    ext = filename.rsplit('.', 1)[1].lower()
    return ext in ['hwp', 'hwpx']


def get_file_type(filename):
    """파일 타입 반환 (hwp 또는 hwpx)"""
    if '.' not in filename:
        return None
    return filename.rsplit('.', 1)[1].lower()


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """한글 파일 업로드 및 파싱 (HWP/HWPX 자동 감지)"""
    if 'file' not in request.files:
        return jsonify({'error': '파일이 없습니다'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '파일을 선택해주세요'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': '한글 파일(.hwp, .hwpx)만 업로드 가능합니다'}), 400
    
    try:
        # 파일 저장
        filename = secure_filename(file.filename)
        # 한글 파일명 보존
        if not filename or filename == '_':
            filename = file.filename
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        file_type = get_file_type(filename)
        
        # HWPX 파일인 경우 별도 처리
        if file_type == 'hwpx' or is_hwpx_file(filepath):
            return _process_hwpx_upload(filepath, filename)
        else:
            return _process_hwp_upload(filepath, filename)
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'파싱 중 오류 발생: {str(e)}'}), 500


def _process_hwp_upload(filepath, filename):
    """HWP 파일 업로드 처리 (읽기 전용)"""
    # 텍스트 파싱
    text = parse_hwp(filepath)
    
    # 표 파싱
    tables = parse_tables(filepath)
    
    # 표 파싱 디버깅 정보
    debug_info = {
        'table_count': len(tables),
        'parsing_method': 'pyhwp',
        'file_size': os.path.getsize(filepath) if os.path.exists(filepath) else 0
    }
    
    # 세션 데이터 저장
    session_id = os.urandom(16).hex()
    json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}.json')
    save_tables_to_json(tables, json_path)
    
    # 텍스트 데이터 저장
    text_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_text.txt')
    with open(text_path, 'w', encoding='utf-8') as f:
        f.write(text or '')
    
    # 원본 파일 경로 저장 (재구성 시 사용)
    session_data = {
        'original_file': filename,
        'original_path': filepath,
        'file_type': 'hwp',
        'text': text or '',
        'tables': tables
    }
    session_json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_session.json')
    with open(session_json_path, 'w', encoding='utf-8') as f:
        json.dump(session_data, f, ensure_ascii=False, indent=2)
    
    return jsonify({
        'success': True,
        'session_id': session_id,
        'filename': filename,
        'file_type': 'hwp',
        'editable': False,
        'text': text or '',
        'tables': tables,
        'table_count': len(tables),
        'debug_info': debug_info,
        'message': f'파싱 완료: 텍스트 {len(text) if text else 0}자, 표 {len(tables)}개 발견 (읽기 전용)' if tables else '파싱 완료: 텍스트는 추출되었지만 표를 찾을 수 없습니다. (읽기 전용)'
    })


def _process_hwpx_upload(filepath, filename):
    """HWPX 파일 업로드 처리 (편집 가능)"""
    # HWPX 파싱
    result = parse_hwpx(filepath)
    
    if not result['success']:
        return jsonify({'error': f'HWPX 파싱 오류: {result["error"]}'}), 500
    
    text = result['text']
    tables = result['tables']
    
    # 표 파싱 디버깅 정보
    debug_info = {
        'table_count': len(tables),
        'parsing_method': 'hwpx_xml',
        'file_size': os.path.getsize(filepath) if os.path.exists(filepath) else 0,
        'sections': len(result.get('sections', []))
    }
    
    # 세션 데이터 저장
    session_id = os.urandom(16).hex()
    json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(tables, f, ensure_ascii=False, indent=2)
    
    # 텍스트 데이터 저장
    text_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_text.txt')
    with open(text_path, 'w', encoding='utf-8') as f:
        f.write(text or '')
    
    # 원본 파일 경로 저장 (재구성 시 사용)
    session_data = {
        'original_file': filename,
        'original_path': filepath,
        'file_type': 'hwpx',
        'text': text or '',
        'tables': tables
    }
    session_json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_session.json')
    with open(session_json_path, 'w', encoding='utf-8') as f:
        json.dump(session_data, f, ensure_ascii=False, indent=2)
    
    return jsonify({
        'success': True,
        'session_id': session_id,
        'filename': filename,
        'file_type': 'hwpx',
        'editable': True,
        'text': text or '',
        'tables': tables,
        'table_count': len(tables),
        'debug_info': debug_info,
        'message': f'파싱 완료: 텍스트 {len(text) if text else 0}자, 표 {len(tables)}개 발견 (편집 가능!)' if tables else '파싱 완료: 텍스트는 추출되었지만 표를 찾을 수 없습니다. (편집 가능)'
    })


@app.route('/api/tables/<session_id>', methods=['GET'])
def get_tables(session_id):
    """세션의 표 데이터 가져오기"""
    json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}.json')
    
    if not os.path.exists(json_path):
        return jsonify({'error': '세션을 찾을 수 없습니다'}), 404
    
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            tables = json.load(f)
        return jsonify({'tables': tables})
    except Exception as e:
        return jsonify({'error': f'데이터 로드 중 오류: {str(e)}'}), 500


@app.route('/api/tables/<session_id>', methods=['POST'])
def update_tables(session_id):
    """표 데이터 업데이트"""
    json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}.json')
    
    if not os.path.exists(json_path):
        return jsonify({'error': '세션을 찾을 수 없습니다'}), 404
    
    try:
        data = request.get_json()
        tables = data.get('tables', [])
        
        # JSON 파일로 저장
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(tables, f, ensure_ascii=False, indent=2)
        
        return jsonify({'success': True, 'message': '표 데이터가 업데이트되었습니다'})
    
    except Exception as e:
        return jsonify({'error': f'업데이트 중 오류: {str(e)}'}), 500


@app.route('/api/tables/<session_id>/edit', methods=['POST'])
def edit_table_cell(session_id):
    """표 셀 수정"""
    json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}.json')
    
    if not os.path.exists(json_path):
        return jsonify({'error': '세션을 찾을 수 없습니다'}), 404
    
    try:
        data = request.get_json()
        table_index = data.get('table_index', 0)
        row = data.get('row', 0)
        col = data.get('col', 0)
        new_value = data.get('value', '')
        
        # 표 데이터 로드
        with open(json_path, 'r', encoding='utf-8') as f:
            tables = json.load(f)
        
        # 셀 수정
        if table_index < len(tables):
            table = tables[table_index]
            rows = table.get('rows', [])
            
            if row < len(rows):
                # 열 확장이 필요한 경우
                while len(rows[row]) <= col:
                    rows[row].append('')
                
                rows[row][col] = new_value
                
                # 저장
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(tables, f, ensure_ascii=False, indent=2)
                
                return jsonify({'success': True})
            else:
                return jsonify({'error': '행 인덱스가 범위를 벗어났습니다'}), 400
        else:
            return jsonify({'error': '표 인덱스가 범위를 벗어났습니다'}), 400
    
    except Exception as e:
        return jsonify({'error': f'수정 중 오류: {str(e)}'}), 500


@app.route('/api/download/<session_id>', methods=['POST'])
def download_hwp(session_id):
    """편집된 데이터를 ZIP 파일로 다운로드 (원본 HWP + 편집된 표/텍스트)"""
    json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}.json')
    text_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_text.txt')
    session_json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_session.json')
    
    if not os.path.exists(json_path):
        return jsonify({'error': '세션을 찾을 수 없습니다'}), 404
    
    try:
        import zipfile
        import csv
        from io import StringIO
        
        # 원본 파일 경로 가져오기
        original_path = None
        original_filename = 'original.hwp'
        if os.path.exists(session_json_path):
            with open(session_json_path, 'r', encoding='utf-8') as f:
                session_data = json.load(f)
                original_path = session_data.get('original_path')
                original_filename = session_data.get('original_file', 'original.hwp')
        
        # ZIP 파일 생성
        zip_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_package.zip')
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # 1. 원본 HWP 파일 포함
            if original_path and os.path.exists(original_path):
                zipf.write(original_path, f'원본_{original_filename}')
            
            # 2. 편집된 텍스트 포함
            if os.path.exists(text_path):
                zipf.write(text_path, '편집된_텍스트.txt')
            
            # 3. 편집된 표 데이터 (JSON)
            zipf.write(json_path, '편집된_표.json')
            
            # 4. 편집된 표 데이터 (CSV) - 각 표별로
            with open(json_path, 'r', encoding='utf-8') as f:
                tables = json.load(f)
            
            for idx, table in enumerate(tables):
                csv_content = StringIO()
                writer = csv.writer(csv_content)
                for row in table.get('rows', []):
                    writer.writerow(row)
                
                csv_bytes = csv_content.getvalue().encode('utf-8-sig')
                zipf.writestr(f'표_{idx + 1}.csv', csv_bytes)
            
            # 5. README 파일
            readme_content = f"""한글 파일 편집 결과
========================

원본 파일: {original_filename}

포함된 파일:
- 원본_{original_filename}: 원본 한글 파일
- 편집된_텍스트.txt: 편집된 텍스트 내용
- 편집된_표.json: 편집된 표 데이터 (JSON 형식)
- 표_N.csv: 각 표별 CSV 파일

사용 방법:
1. 원본 한글 파일을 한글 프로그램에서 엽니다
2. CSV 파일의 내용을 복사하여 표에 붙여넣습니다
3. 또는 텍스트 파일의 내용을 참고하여 수정합니다

참고: pyhwp 라이브러리는 읽기 전용이라 한글 파일을 직접 수정할 수 없습니다.
한글 파일을 직접 수정하려면 한글과컴퓨터 API를 사용해야 합니다.
"""
            zipf.writestr('README.txt', readme_content.encode('utf-8'))
        
        return send_file(
            zip_path,
            as_attachment=True,
            download_name='edited_package.zip',
            mimetype='application/zip'
        )
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'파일 생성 중 오류: {str(e)}'}), 500


@app.route('/api/convert-hwp-to-hwpx/<session_id>', methods=['POST'])
def convert_hwp_to_hwpx(session_id):
    """HWP 파일을 HWPX로 변환하여 다운로드"""
    session_json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_session.json')
    json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}.json')
    text_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_text.txt')
    
    if not os.path.exists(session_json_path):
        return jsonify({'error': '세션을 찾을 수 없습니다'}), 404
    
    try:
        from hwpx_parser import convert_hwp_to_hwpx
        
        # 세션 데이터 읽기
        with open(session_json_path, 'r', encoding='utf-8') as f:
            session_data = json.load(f)
        
        original_path = session_data.get('original_path')
        original_filename = session_data.get('original_file', 'original.hwp')
        
        if not original_path or not os.path.exists(original_path):
            return jsonify({'error': '원본 파일을 찾을 수 없습니다'}), 404
        
        # 파일 타입 확인
        if session_data.get('file_type') != 'hwp':
            return jsonify({'error': 'HWP 파일만 변환 가능합니다'}), 400
        
        # 템플릿 HWPX 파일 경로 (프로젝트에 있는 경우)
        template_path = None
        template_candidates = [
            os.path.join(os.path.dirname(__file__), '상담 리포트 템플릿.hwpx'),
            os.path.join(os.path.dirname(__file__), 'template.hwpx'),
        ]
        for candidate in template_candidates:
            if os.path.exists(candidate):
                template_path = candidate
                break
        
        # HWPX 파일 생성
        hwpx_filename = original_filename.replace('.hwp', '.hwpx')
        hwpx_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_converted.hwpx')
        
        # 표 데이터 로드
        tables = []
        if os.path.exists(json_path):
            with open(json_path, 'r', encoding='utf-8') as f:
                tables = json.load(f)
        
        # 텍스트 로드
        text = ''
        if os.path.exists(text_path):
            with open(text_path, 'r', encoding='utf-8') as f:
                text = f.read()
        
        # 변환 실행 (이미 파싱된 데이터 사용)
        success = convert_hwp_to_hwpx(original_path, hwpx_path, template_path, text=text, tables=tables)
        
        if not success:
            return jsonify({'error': 'HWPX 변환에 실패했습니다'}), 500
        
        # 다운로드 파일명 가져오기
        data = request.get_json() or {}
        download_name = data.get('filename', hwpx_filename)
        if not download_name.endswith('.hwpx'):
            download_name += '.hwpx'
        
        return send_file(
            hwpx_path,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.hancom.hwpml+zip'
        )
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'변환 중 오류: {str(e)}'}), 500


@app.route('/api/text/<session_id>', methods=['GET'])
def get_text(session_id):
    """세션의 텍스트 데이터 가져오기"""
    text_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_text.txt')
    
    if not os.path.exists(text_path):
        return jsonify({'error': '세션을 찾을 수 없습니다'}), 404
    
    try:
        with open(text_path, 'r', encoding='utf-8') as f:
            text = f.read()
        return jsonify({'text': text})
    except Exception as e:
        return jsonify({'error': f'텍스트 로드 중 오류: {str(e)}'}), 500


@app.route('/api/text/<session_id>', methods=['POST'])
def update_text(session_id):
    """텍스트 데이터 업데이트"""
    text_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_text.txt')
    
    if not os.path.exists(text_path):
        return jsonify({'error': '세션을 찾을 수 없습니다'}), 404
    
    try:
        data = request.get_json()
        text = data.get('text', '')
        
        # 텍스트 파일로 저장
        with open(text_path, 'w', encoding='utf-8') as f:
            f.write(text)
        
        return jsonify({'success': True, 'message': '텍스트가 업데이트되었습니다'})
    
    except Exception as e:
        return jsonify({'error': f'업데이트 중 오류: {str(e)}'}), 500


@app.route('/api/download-json/<session_id>', methods=['GET'])
def download_json(session_id):
    """표 데이터를 JSON으로 다운로드"""
    json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}.json')
    
    if not os.path.exists(json_path):
        return jsonify({'error': '세션을 찾을 수 없습니다'}), 404
    
    return send_file(
        json_path,
        as_attachment=True,
        download_name='tables.json',
        mimetype='application/json'
    )


@app.route('/api/download-hwpx/<session_id>', methods=['POST'])
def download_hwpx(session_id):
    """HWPX 파일 다운로드 (편집된 내용이 반영된 실제 hwpx 파일)"""
    json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}.json')
    session_json_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_session.json')
    
    if not os.path.exists(json_path):
        return jsonify({'error': '세션을 찾을 수 없습니다'}), 404
    
    try:
        # 세션 데이터 확인
        if not os.path.exists(session_json_path):
            return jsonify({'error': '세션 정보를 찾을 수 없습니다'}), 404
        
        with open(session_json_path, 'r', encoding='utf-8') as f:
            session_data = json.load(f)
        
        # HWPX 파일만 편집 가능
        if session_data.get('file_type') != 'hwpx':
            return jsonify({'error': 'HWPX 파일만 편집하여 저장할 수 있습니다'}), 400
        
        original_path = session_data.get('original_path')
        original_filename = session_data.get('original_file', 'output.hwpx')
        
        if not original_path or not os.path.exists(original_path):
            return jsonify({'error': '원본 파일을 찾을 수 없습니다'}), 404
        
        # 편집된 표 데이터 로드
        with open(json_path, 'r', encoding='utf-8') as f:
            tables = json.load(f)
        
        # 요청에서 파일명 가져오기
        data = request.get_json() or {}
        custom_filename = data.get('filename', '')
        if custom_filename:
            if not custom_filename.endswith('.hwpx'):
                custom_filename += '.hwpx'
            download_filename = custom_filename
        else:
            download_filename = f'edited_{original_filename}'
        
        # 출력 파일 경로
        output_path = os.path.join(app.config['TEMP_FOLDER'], f'{session_id}_edited.hwpx')
        
        # HWPX 파일 생성
        success = save_hwpx_with_tables(original_path, tables, output_path)
        
        if not success:
            return jsonify({'error': 'HWPX 파일 생성에 실패했습니다'}), 500
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=download_filename,
            mimetype='application/vnd.hancom.hwpx'
        )
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'파일 생성 중 오류: {str(e)}'}), 500


@app.route('/api/health', methods=['GET'])
def health():
    """헬스 체크"""
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)

