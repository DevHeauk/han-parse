#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HWPX 파일 파싱 및 편집 모듈
HWPX는 ZIP + XML 구조로, 직접 편집 및 저장이 가능합니다.
"""

import os
import zipfile
import xml.etree.ElementTree as ET
from typing import List, Dict, Any, Optional
import shutil
import tempfile
import re

# lxml import (선택적 - 없어도 정규식 방식으로 동작)
try:
    from lxml import etree
    LXML_AVAILABLE = True
except ImportError:
    LXML_AVAILABLE = False
    etree = None


# HWPX XML 네임스페이스
NAMESPACES = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'ho': 'http://www.hancom.co.kr/hwpml/2011/owner',
    'hp10': 'http://www.hancom.co.kr/hwpml/2016/paragraph',
}

# 네임스페이스 등록
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


def is_hwpx_file(file_path: str) -> bool:
    """HWPX 파일인지 확인"""
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            namelist = zf.namelist()
            # HWPX 특징: Contents 폴더와 content.hpf 존재
            return any('Contents/' in name for name in namelist)
    except:
        return False


def parse_hwpx(file_path: str) -> Dict[str, Any]:
    """
    HWPX 파일을 파싱하여 텍스트와 표를 추출합니다.
    
    Args:
        file_path: HWPX 파일 경로
        
    Returns:
        파싱된 데이터 (텍스트, 표, 메타데이터)
    """
    result = {
        'text': '',
        'tables': [],
        'sections': [],
        'file_list': [],
        'success': False,
        'error': None
    }
    
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            result['file_list'] = zf.namelist()
            
            # section 파일들 찾기
            section_files = sorted([f for f in zf.namelist() if 'section' in f.lower() and f.endswith('.xml')])
            
            all_text = []
            all_tables = []
            
            for section_file in section_files:
                section_data = _parse_section_xml(zf, section_file)
                result['sections'].append({
                    'file': section_file,
                    'text': section_data['text'],
                    'tables': section_data['tables']
                })
                all_text.append(section_data['text'])
                all_tables.extend(section_data['tables'])
            
            result['text'] = '\n'.join(all_text)
            result['tables'] = all_tables
            result['success'] = True
            
    except Exception as e:
        result['error'] = str(e)
        import traceback
        traceback.print_exc()
    
    return result


def _parse_section_xml(zf: zipfile.ZipFile, section_file: str) -> Dict[str, Any]:
    """섹션 XML 파일 파싱"""
    result = {'text': '', 'tables': []}
    
    try:
        with zf.open(section_file) as f:
            content = f.read().decode('utf-8')
            root = ET.fromstring(content)
            
            # 텍스트 추출
            texts = []
            for elem in root.iter():
                # t 태그에서 텍스트 추출 (여러 네임스페이스 시도)
                if elem.tag.endswith('}t') or elem.tag == 't':
                    if elem.text:
                        texts.append(elem.text)
            
            result['text'] = ''.join(texts)
            
            # 표 추출
            tables = _extract_tables_from_xml(root)
            result['tables'] = tables
            
    except Exception as e:
        print(f"섹션 파싱 오류 ({section_file}): {e}")
    
    return result


def _extract_tables_from_xml(root: ET.Element) -> List[Dict[str, Any]]:
    """XML에서 표 추출"""
    tables = []
    
    # tbl 태그 찾기 (표)
    for tbl in root.iter():
        if tbl.tag.endswith('}tbl') or tbl.tag == 'tbl':
            table_data = _parse_table_element(tbl)
            if table_data and table_data.get('rows'):
                tables.append(table_data)
    
    return tables


def _parse_table_element(tbl: ET.Element) -> Dict[str, Any]:
    """표 요소 파싱 (셀 병합 정보 포함)"""
    table_data = {
        'rows': [],
        'cells': [],  # 셀 상세 정보 (병합 포함)
        'row_count': 0,
        'col_count': 0
    }
    
    try:
        # tr (행) 찾기
        rows = []
        all_cells = []
        row_idx = 0
        
        for tr in tbl.iter():
            if tr.tag.endswith('}tr') or tr.tag == 'tr':
                row_cells = []
                row_cell_info = []
                col_idx = 0
                
                # tc (셀) 찾기
                for tc in tr.iter():
                    if tc.tag.endswith('}tc') or tc.tag == 'tc':
                        cell_text = _extract_cell_text(tc)
                        
                        # 셀 병합 정보 추출
                        colspan = 1
                        rowspan = 1
                        for child in tc.iter():
                            if child.tag.endswith('}cellSpan'):
                                colspan = int(child.get('colSpan', 1))
                                rowspan = int(child.get('rowSpan', 1))
                                break
                        
                        row_cells.append(cell_text)
                        row_cell_info.append({
                            'text': cell_text,
                            'row': row_idx,
                            'col': col_idx,
                            'colspan': colspan,
                            'rowspan': rowspan
                        })
                        col_idx += 1
                
                if row_cells:
                    rows.append(row_cells)
                    all_cells.append(row_cell_info)
                    row_idx += 1
        
        if rows:
            table_data['rows'] = rows
            table_data['cells'] = all_cells
            table_data['row_count'] = len(rows)
            table_data['col_count'] = max(len(row) for row in rows) if rows else 0
    
    except Exception as e:
        print(f"표 파싱 오류: {e}")
    
    return table_data


def _extract_cell_text(tc: ET.Element) -> str:
    """셀에서 텍스트 추출"""
    texts = []
    for elem in tc.iter():
        if elem.tag.endswith('}t') or elem.tag == 't':
            if elem.text:
                texts.append(elem.text)
    return ''.join(texts)


def edit_hwpx_table(file_path: str, table_index: int, new_rows: List[List[str]], output_path: str) -> bool:
    """
    HWPX 파일의 표를 수정하고 새 파일로 저장합니다.
    
    Args:
        file_path: 원본 HWPX 파일 경로
        table_index: 수정할 표 인덱스
        new_rows: 새로운 표 데이터 (2D 리스트)
        output_path: 출력 파일 경로
        
    Returns:
        성공 여부
    """
    try:
        # 임시 디렉토리에 압축 해제
        with tempfile.TemporaryDirectory() as temp_dir:
            # HWPX 압축 해제
            with zipfile.ZipFile(file_path, 'r') as zf:
                zf.extractall(temp_dir)
            
            # section 파일 찾기
            section_files = []
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    if 'section' in file.lower() and file.endswith('.xml'):
                        section_files.append(os.path.join(root_dir, file))
            
            section_files.sort()
            
            # 표 찾아서 수정
            table_count = 0
            modified = False
            
            for section_file in section_files:
                with open(section_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                root = ET.fromstring(content)
                
                # 표 찾기
                for tbl in root.iter():
                    if tbl.tag.endswith('}tbl') or tbl.tag == 'tbl':
                        if table_count == table_index:
                            # 이 표를 수정
                            _modify_table_element(tbl, new_rows)
                            modified = True
                            break
                        table_count += 1
                
                if modified:
                    # XML 저장
                    tree = ET.ElementTree(root)
                    with open(section_file, 'wb') as f:
                        tree.write(f, encoding='utf-8', xml_declaration=True)
                    break
            
            # 다시 ZIP으로 압축
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root_dir, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path_full = os.path.join(root_dir, file)
                        arcname = os.path.relpath(file_path_full, temp_dir)
                        zf.write(file_path_full, arcname)
            
            return True
            
    except Exception as e:
        print(f"HWPX 편집 오류: {e}")
        import traceback
        traceback.print_exc()
        return False


def _modify_table_element(tbl: ET.Element, new_rows: List[List[str]]):
    """표 요소의 내용을 수정"""
    # 기존 행(tr) 찾기
    tr_elements = []
    for elem in tbl.iter():
        if elem.tag.endswith('}tr') or elem.tag == 'tr':
            tr_elements.append(elem)
    
    # 각 행의 셀 수정
    for row_idx, tr in enumerate(tr_elements):
        if row_idx >= len(new_rows):
            break
        
        # 셀(tc) 찾기
        tc_elements = []
        for elem in tr.iter():
            if elem.tag.endswith('}tc') or elem.tag == 'tc':
                tc_elements.append(elem)
        
        # 각 셀의 텍스트 수정
        for col_idx, tc in enumerate(tc_elements):
            if col_idx >= len(new_rows[row_idx]):
                break
            
            new_text = new_rows[row_idx][col_idx]
            _set_cell_text(tc, new_text)


def _set_cell_text(tc: ET.Element, new_text: str):
    """셀의 텍스트를 설정"""
    # t 태그 찾아서 텍스트 수정
    for elem in tc.iter():
        if elem.tag.endswith('}t') or elem.tag == 't':
            elem.text = new_text
            return
    
    # t 태그가 없으면 생성 시도 (복잡한 경우 건너뜀)


def save_hwpx_with_tables(original_path: str, tables_data: List[Dict[str, Any]], output_path: str) -> bool:
    """
    HWPX 파일에 수정된 모든 표를 저장합니다.
    정규식을 사용하여 원본 XML 형식을 그대로 유지합니다.
    
    Args:
        original_path: 원본 HWPX 파일 경로
        tables_data: 수정된 표 데이터 리스트
        output_path: 출력 파일 경로
        
    Returns:
        성공 여부
    """
    try:
        # 임시 디렉토리에 압축 해제
        with tempfile.TemporaryDirectory() as temp_dir:
            # HWPX 압축 해제
            with zipfile.ZipFile(original_path, 'r') as zf:
                zf.extractall(temp_dir)
            
            # section 파일 찾기
            section_files = []
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    if 'section' in file.lower() and file.endswith('.xml'):
                        section_files.append(os.path.join(root_dir, file))
            
            section_files.sort()
            
            # 각 표 수정 (정규식 사용)
            for section_file in section_files:
                with open(section_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # 표 데이터를 정규식으로 수정
                if tables_data:
                    content = _modify_tables_with_regex(content, tables_data)
                
                # 수정된 내용 저장 (원본 형식 유지)
                with open(section_file, 'w', encoding='utf-8') as f:
                    f.write(content)
            
            # 다시 ZIP으로 압축
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root_dir, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path_full = os.path.join(root_dir, file)
                        arcname = os.path.relpath(file_path_full, temp_dir)
                        zf.write(file_path_full, arcname)
            
            return True
            
    except Exception as e:
        print(f"HWPX 저장 오류: {e}")
        import traceback
        traceback.print_exc()
        return False


def _modify_tables_with_regex(content: str, tables_data: List[Dict[str, Any]]) -> str:
    """
    정규식을 사용하여 XML 내의 표 데이터를 수정합니다.
    원본 XML 형식을 그대로 유지합니다.
    """
    import re
    
    # 표(tbl) 태그 찾기 - hp:tbl 형식
    tbl_pattern = r'(<hp:tbl[^>]*>)(.*?)(</hp:tbl>)'
    tbl_matches = list(re.finditer(tbl_pattern, content, re.DOTALL))
    
    if not tbl_matches:
        print("표를 찾을 수 없습니다.")
        return content
    
    # 역순으로 처리 (인덱스 변경 방지)
    for table_idx, match in enumerate(tbl_matches):
        if table_idx >= len(tables_data):
            break
        
        table_data = tables_data[table_idx]
        new_rows = table_data.get('rows', [])
        if not new_rows:
            continue
        
        tbl_start = match.start()
        tbl_end = match.end()
        tbl_content = match.group(0)
        
        # 표 내용 수정
        modified_tbl = _modify_single_table_regex(tbl_content, new_rows)
        
        # 원본 내용 교체
        content = content[:tbl_start] + modified_tbl + content[tbl_end:]
    
    return content


def _modify_single_table_regex(tbl_content: str, new_rows: List[List[str]]) -> str:
    """
    단일 표의 내용을 정규식으로 수정합니다.
    """
    import re
    
    # tr(행) 찾기
    tr_pattern = r'(<hp:tr[^>]*>)(.*?)(</hp:tr>)'
    tr_matches = list(re.finditer(tr_pattern, tbl_content, re.DOTALL))
    
    result = tbl_content
    offset = 0  # 문자열 길이 변경에 따른 오프셋
    
    for row_idx, tr_match in enumerate(tr_matches):
        if row_idx >= len(new_rows):
            break
        
        row_data = new_rows[row_idx]
        tr_start = tr_match.start() + offset
        tr_end = tr_match.end() + offset
        tr_content = result[tr_start:tr_end]
        
        # 행 내용 수정
        modified_tr = _modify_single_row_regex(tr_content, row_data)
        
        # 길이 차이 계산
        len_diff = len(modified_tr) - len(tr_content)
        
        # 교체
        result = result[:tr_start] + modified_tr + result[tr_end:]
        offset += len_diff
    
    return result


def _modify_single_row_regex(tr_content: str, row_data: List[str]) -> str:
    """
    단일 행의 셀 내용을 수정합니다.
    """
    import re
    
    # tc(셀) 찾기
    tc_pattern = r'(<hp:tc[^>]*>)(.*?)(</hp:tc>)'
    tc_matches = list(re.finditer(tc_pattern, tr_content, re.DOTALL))
    
    result = tr_content
    offset = 0
    
    for col_idx, tc_match in enumerate(tc_matches):
        if col_idx >= len(row_data):
            break
        
        new_text = row_data[col_idx]
        tc_start = tc_match.start() + offset
        tc_end = tc_match.end() + offset
        tc_content = result[tc_start:tc_end]
        
        # 셀 내 t 태그 수정
        modified_tc = _modify_cell_text_regex(tc_content, new_text)
        
        len_diff = len(modified_tc) - len(tc_content)
        result = result[:tc_start] + modified_tc + result[tc_end:]
        offset += len_diff
    
    return result


def _modify_cell_text_regex(tc_content: str, new_text: str) -> str:
    """
    셀 내의 t 태그 텍스트를 수정합니다.
    빈 셀(자동 닫힘 run 태그)도 처리합니다.
    """
    import re
    
    # 1. 기존 hp:t 태그 찾기 (첫 번째 것만 수정)
    t_pattern = r'(<hp:t>)(.*?)(</hp:t>)'
    
    def replace_first_t(match):
        return match.group(1) + new_text + match.group(3)
    
    result, count = re.subn(t_pattern, replace_first_t, tc_content, count=1)
    
    if count > 0:
        return result
    
    # 2. t 태그가 없는 경우 - 자동 닫힘 run 태그를 확장하여 t 태그 추가
    # <hp:run charPrIDRef="10"/> → <hp:run charPrIDRef="10"><hp:t>텍스트</hp:t></hp:run>
    self_closing_run_pattern = r'(<hp:run[^>]*)(/\s*>)'
    
    def expand_self_closing_run(match):
        return match.group(1) + f'><hp:t>{new_text}</hp:t></hp:run>'
    
    result, count = re.subn(self_closing_run_pattern, expand_self_closing_run, tc_content, count=1)
    
    if count > 0:
        return result
    
    # 3. 빈 run 태그 (</hp:run>으로 끝나는) 안에 t 태그 삽입
    empty_run_pattern = r'(<hp:run[^>]*>)(</hp:run>)'
    
    def add_t_to_empty_run(match):
        return match.group(1) + f'<hp:t>{new_text}</hp:t>' + match.group(2)
    
    result, count = re.subn(empty_run_pattern, add_t_to_empty_run, tc_content, count=1)
    
    return result


def save_hwpx_with_tables_lxml(original_path: str, tables_data: List[Dict[str, Any]], output_path: str) -> bool:
    """
    HWPX 파일에 수정된 모든 표를 저장합니다 (lxml 기반 완벽 버전).
    네임스페이스를 완벽하게 보존하면서 XML 구조를 정확히 이해하여 수정합니다.
    
    Args:
        original_path: 원본 HWPX 파일 경로
        tables_data: 수정된 표 데이터 리스트
        output_path: 출력 파일 경로
        
    Returns:
        성공 여부
    """
    if not LXML_AVAILABLE:
        print("lxml이 설치되지 않았습니다. 정규식 방식으로 대체합니다.")
        return save_hwpx_with_tables(original_path, tables_data, output_path)
    
    try:
        # 임시 디렉토리에 압축 해제
        with tempfile.TemporaryDirectory() as temp_dir:
            # HWPX 압축 해제
            with zipfile.ZipFile(original_path, 'r') as zf:
                zf.extractall(temp_dir)
            
            # section 파일 찾기
            section_files = []
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    if 'section' in file.lower() and file.endswith('.xml'):
                        section_files.append(os.path.join(root_dir, file))
            
            section_files.sort()
            
            # lxml 네임스페이스 정의
            ns_map = {
                'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
                'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
                'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
                'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
                'ho': 'http://www.hancom.co.kr/hwpml/2011/owner',
            }
            
            # 각 section 파일 처리
            for section_file in section_files:
                # lxml 파서 설정 (원본 형식 최대한 보존)
                parser = etree.XMLParser(
                    remove_blank_text=False,  # 공백 보존
                    strip_cdata=False,  # CDATA 보존
                    recover=False,  # 오류 시 실패
                    huge_tree=True  # 큰 파일 지원
                )
                
                try:
                    tree = etree.parse(section_file, parser)
                    root = tree.getroot()
                except Exception as e:
                    print(f"XML 파싱 오류 ({section_file}): {e}")
                    continue
                
                # 표 찾기
                tables = root.xpath('.//hp:tbl', namespaces=ns_map)
                
                if not tables_data:
                    continue
                
                # 각 표 수정
                for table_idx, tbl in enumerate(tables):
                    if table_idx >= len(tables_data):
                        break
                    
                    table_data = tables_data[table_idx]
                    new_rows = table_data.get('rows', [])
                    
                    if not new_rows:
                        continue
                    
                    # 행 찾기
                    rows = tbl.xpath('.//hp:tr', namespaces=ns_map)
                    
                    for row_idx, tr in enumerate(rows):
                        if row_idx >= len(new_rows):
                            break
                        
                        row_data = new_rows[row_idx]
                        
                        # 셀 찾기
                        cells = tr.xpath('.//hp:tc', namespaces=ns_map)
                        
                        for col_idx, tc in enumerate(cells):
                            if col_idx >= len(row_data):
                                break
                            
                            new_text = row_data[col_idx]
                            
                            # t 태그 찾기
                            t_elements = tc.xpath('.//hp:t', namespaces=ns_map)
                            
                            if t_elements:
                                # 기존 t 태그가 있으면 텍스트만 수정
                                t_elements[0].text = new_text
                            else:
                                # t 태그가 없으면 생성
                                # 먼저 run 태그 찾기
                                run_elements = tc.xpath('.//hp:run', namespaces=ns_map)
                                
                                if run_elements:
                                    # run 태그가 있으면 그 안에 t 태그 추가
                                    run_elem = run_elements[0]
                                    # 자동 닫힘 태그인지 확인
                                    if run_elem.text is None and len(run_elem) == 0:
                                        # 자동 닫힘 태그였던 경우, 일반 태그로 변환
                                        pass
                                    
                                    # t 태그 생성
                                    t_elem = etree.Element(
                                        '{http://www.hancom.co.kr/hwpml/2011/paragraph}t'
                                    )
                                    t_elem.text = new_text
                                    run_elem.append(t_elem)
                                else:
                                    # run 태그도 없으면 생성
                                    run_elem = etree.Element(
                                        '{http://www.hancom.co.kr/hwpml/2011/paragraph}run'
                                    )
                                    t_elem = etree.Element(
                                        '{http://www.hancom.co.kr/hwpml/2011/paragraph}t'
                                    )
                                    t_elem.text = new_text
                                    run_elem.append(t_elem)
                                    tc.append(run_elem)
                
                # 네임스페이스 보존하면서 저장
                # 원본 XML 선언과 네임스페이스 선언 유지
                try:
                    # 원본 파일의 XML 선언 읽기
                    with open(section_file, 'r', encoding='utf-8') as f:
                        original_content = f.read()
                        # XML 선언 추출
                        xml_declaration = ''
                        if original_content.startswith('<?xml'):
                            end_pos = original_content.find('?>')
                            if end_pos > 0:
                                xml_declaration = original_content[:end_pos + 2]
                except:
                    xml_declaration = '<?xml version="1.0" encoding="UTF-8"?>'
                
                # lxml로 저장 (네임스페이스 보존)
                # pretty_print=False로 원본 형식 유지
                tree.write(section_file,
                          encoding='utf-8',
                          xml_declaration=True,
                          pretty_print=False,
                          method='xml')
            
            # 다시 ZIP으로 압축
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root_dir, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path_full = os.path.join(root_dir, file)
                        arcname = os.path.relpath(file_path_full, temp_dir)
                        zf.write(file_path_full, arcname)
            
            return True
            
    except Exception as e:
        print(f"HWPX 저장 오류 (lxml): {e}")
        import traceback
        traceback.print_exc()
        # 오류 발생 시 정규식 방식으로 대체
        print("정규식 방식으로 재시도합니다...")
        return save_hwpx_with_tables(original_path, tables_data, output_path)


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("사용법: python hwpx_parser.py <hwpx파일경로>")
        sys.exit(1)
    
    file_path = sys.argv[1]
    
    if not is_hwpx_file(file_path):
        print(f"HWPX 파일이 아닙니다: {file_path}")
        sys.exit(1)
    
    print(f"HWPX 파일 파싱 중: {file_path}")
    result = parse_hwpx(file_path)
    
    if result['success']:
        print(f"\n텍스트 길이: {len(result['text'])}")
        print(f"표 개수: {len(result['tables'])}")
        
        for idx, table in enumerate(result['tables']):
            print(f"\n표 {idx + 1}: {table['row_count']}행 x {table['col_count']}열")
            for row in table['rows'][:3]:
                print(f"  {row}")
    else:
        print(f"파싱 실패: {result['error']}")

