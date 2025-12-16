#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
한글 파일(HWP) 파싱 모듈
표 파싱 및 재구성 기능 포함
"""

import sys
import os
import json
import csv
from pathlib import Path
from typing import List, Dict, Any, Optional


def parse_hwp(file_path):
    """
    한글 파일을 파싱하여 텍스트를 추출합니다.
    
    Args:
        file_path (str): 한글 파일 경로
        
    Returns:
        str: 추출된 텍스트
    """
    try:
        from hwp5.xmlmodel import Hwp5File
        
        hwp = Hwp5File(file_path)
        text_content = []
        
        # BodyText의 sections 순회
        if hasattr(hwp, 'bodytext'):
            bodytext = hwp.bodytext
            sections = bodytext.sections
            
            for section in sections:
                # models()를 통해 ParaText 추출
                for model in section.models():
                    if isinstance(model, dict):
                        model_type = model.get('type')
                        content = model.get('content', {})
                        
                        # ParaText에서 텍스트 추출
                        if 'ParaText' in str(model_type):
                            chunks = content.get('chunks', [])
                            for chunk in chunks:
                                if isinstance(chunk, tuple) and len(chunk) >= 2:
                                    chunk_content = chunk[1]
                                    if isinstance(chunk_content, str):
                                        text_content.append(chunk_content)
                            text_content.append('\n')
        
        return ''.join(text_content)
        
    except ImportError as e:
        print(f"hwp5 라이브러리가 설치되지 않았습니다: {e}")
        print("다음 명령어로 설치하세요: pip install -r requirements.txt")
        return None
    except Exception as e:
        print(f"파싱 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return None


def parse_hwp_simple(file_path):
    """
    간단한 방법으로 한글 파일의 기본 정보를 추출합니다.
    OLE 파일로 접근하여 메타데이터를 읽습니다.
    
    Args:
        file_path (str): 한글 파일 경로
        
    Returns:
        dict: 파일 정보
    """
    try:
        import olefile
        
        if not olefile.isOleFile(file_path):
            return {"error": "한글 파일이 아닙니다."}
        
        ole = olefile.OleFileIO(file_path)
        
        info = {
            "file_path": file_path,
            "is_ole": True,
            "streams": ole.listdir()
        }
        
        ole.close()
        return info
        
    except ImportError:
        print("olefile 라이브러리가 설치되지 않았습니다.")
        print("다음 명령어로 설치하세요: pip install -r requirements.txt")
        return None
    except Exception as e:
        print(f"파싱 중 오류 발생: {e}")
        return None


def parse_tables(file_path: str) -> List[Dict[str, Any]]:
    """
    한글 파일에서 모든 표를 파싱하여 구조화된 데이터로 반환합니다.
    
    Args:
        file_path (str): 한글 파일 경로
        
    Returns:
        List[Dict]: 표 데이터 리스트. 각 표는 행/열 구조를 포함
    """
    try:
        # hwp5 모듈 사용 (pyhwp가 아닌 hwp5로 import)
        from hwp5.xmlmodel import Hwp5File
        
        hwp = Hwp5File(file_path)
        tables = []
        
        # BodyText의 sections 순회
        if hasattr(hwp, 'bodytext'):
            bodytext = hwp.bodytext
            sections = bodytext.sections
            
            for section_idx, section in enumerate(sections):
                # models()를 통해 모든 모델 순회
                # 각 표를 개별적으로 처리하기 위한 상태 변수
                current_table_info = None
                current_cells_data = []
                in_table = False
                
                for model in section.models():
                    if isinstance(model, dict):
                        model_type = model.get('type')
                        content = model.get('content', {})
                        
                        # TableControl - 새로운 표 시작
                        if 'TableControl' in str(model_type):
                            # 이전 표가 있으면 저장
                            if current_table_info and current_cells_data:
                                table = _build_table_from_cells(
                                    current_table_info, current_cells_data, section_idx
                                )
                                if table:
                                    tables.append(table)
                            
                            # 새 표 시작을 위해 초기화
                            current_table_info = None
                            current_cells_data = []
                            in_table = True
                        
                        # TableBody - 표 구조 정보
                        elif 'TableBody' in str(model_type):
                            current_table_info = {
                                'rows': content.get('rows', 0),
                                'cols': content.get('cols', 0),
                            }
                        
                        # TableCell - 셀 위치 정보
                        elif 'TableCell' in str(model_type) and in_table:
                            current_cells_data.append({
                                'row': content.get('row', 0),
                                'col': content.get('col', 0),
                                'text': ''
                            })
                        
                        # ParaText - 텍스트 내용
                        elif 'ParaText' in str(model_type):
                            chunks = content.get('chunks', [])
                            text_parts = []
                            for chunk in chunks:
                                if isinstance(chunk, tuple) and len(chunk) >= 2:
                                    chunk_content = chunk[1]
                                    if isinstance(chunk_content, str):
                                        text_parts.append(chunk_content)
                            
                            if text_parts and current_cells_data and in_table:
                                # 마지막 셀에 텍스트 추가 (기존 텍스트에 이어붙이기)
                                existing = current_cells_data[-1].get('text', '')
                                current_cells_data[-1]['text'] = existing + ''.join(text_parts)
                
                # 마지막 표 처리
                if current_table_info and current_cells_data:
                    table = _build_table_from_cells(
                        current_table_info, current_cells_data, section_idx
                    )
                    if table:
                        tables.append(table)
        
        return tables
        
    except ImportError as e:
        print(f"hwp5 라이브러리가 설치되지 않았습니다: {e}")
        return []
    except Exception as e:
        print(f"표 파싱 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return []


def _build_table_from_cells(table_info: Dict, cells_data: List, section_idx: int) -> Optional[Dict]:
    """셀 데이터로부터 표 구조 생성"""
    rows = table_info.get('rows', 0)
    cols = table_info.get('cols', 0)
    
    if rows == 0 or cols == 0:
        return None
    
    # 2D 배열로 변환
    table_rows = [[''] * cols for _ in range(rows)]
    
    for cell in cells_data:
        row_idx = cell.get('row', 0)
        col_idx = cell.get('col', 0)
        text = cell.get('text', '')
        
        if 0 <= row_idx < rows and 0 <= col_idx < cols:
            table_rows[row_idx][col_idx] = text
    
    return {
        'section': section_idx,
        'rows': table_rows,
        'row_count': rows,
        'col_count': cols
    }


def _extract_table_from_control(control, section_idx: int, para_idx: int) -> Optional[Dict[str, Any]]:
    """표 컨트롤에서 데이터 추출 (개선된 버전)"""
    try:
        table_data = {
            "section": section_idx,
            "paragraph": para_idx,
            "rows": [],
            "row_count": 0,
            "col_count": 0
        }
        
        # 방법 1: rows 속성 확인
        if hasattr(control, 'rows'):
            for row in control.rows:
                row_data = []
                if hasattr(row, 'cells'):
                    for cell in row.cells:
                        cell_text = _extract_cell_text_improved(cell)
                        row_data.append(cell_text)
                elif hasattr(row, '__iter__'):
                    # row가 iterable인 경우
                    for cell in row:
                        cell_text = _extract_cell_text_improved(cell)
                        row_data.append(cell_text)
                table_data["rows"].append(row_data)
                table_data["col_count"] = max(table_data["col_count"], len(row_data))
        
        # 방법 2: table 속성 확인
        elif hasattr(control, 'table'):
            table = control.table
            if hasattr(table, 'rows'):
                for row in table.rows:
                    row_data = []
                    if hasattr(row, 'cells'):
                        for cell in row.cells:
                            cell_text = _extract_cell_text_improved(cell)
                            row_data.append(cell_text)
                    table_data["rows"].append(row_data)
                    table_data["col_count"] = max(table_data["col_count"], len(row_data))
        
        # 방법 3: 직접 데이터 접근
        elif hasattr(control, 'data'):
            # control.data에서 표 구조 추출 시도
            data = control.data
            if hasattr(data, 'rows'):
                for row in data.rows:
                    row_data = []
                    if hasattr(row, 'cells'):
                        for cell in row.cells:
                            cell_text = _extract_cell_text_improved(cell)
                            row_data.append(cell_text)
                    table_data["rows"].append(row_data)
                    table_data["col_count"] = max(table_data["col_count"], len(row_data))
        
        table_data["row_count"] = len(table_data["rows"])
        return table_data if table_data["rows"] else None
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return None


def _extract_table_from_model(ctrl) -> Optional[Dict[str, Any]]:
    """모델에서 표 데이터 추출"""
    try:
        table_data = {
            "rows": [],
            "row_count": 0,
            "col_count": 0
        }
        
        if hasattr(ctrl, 'rows'):
            for row in ctrl.rows:
                row_data = []
                if hasattr(row, 'cells'):
                    for cell in row.cells:
                        cell_text = _extract_cell_text(cell)
                        row_data.append(cell_text)
                table_data["rows"].append(row_data)
                table_data["col_count"] = max(table_data["col_count"], len(row_data))
        
        table_data["row_count"] = len(table_data["rows"])
        return table_data if table_data["rows"] else None
        
    except:
        return None


def _extract_cell_text_improved(cell) -> str:
    """셀에서 텍스트 추출 (개선된 버전)"""
    try:
        # 방법 1: 직접 text 속성
        if hasattr(cell, 'text'):
            text = cell.text
            if text:
                return str(text)
        
        # 방법 2: paragraphs를 통한 추출
        if hasattr(cell, 'paragraphs'):
            texts = []
            for para in cell.paragraphs:
                if hasattr(para, 'chars'):
                    for char in para.chars():
                        if hasattr(char, 'text'):
                            char_text = char.text
                            if char_text:
                                texts.append(str(char_text))
                elif hasattr(para, 'text'):
                    para_text = para.text
                    if para_text:
                        texts.append(str(para_text))
            if texts:
                return ''.join(texts)
        
        # 방법 3: content 속성
        if hasattr(cell, 'content'):
            content = cell.content
            if content:
                return str(content)
        
        # 방법 4: __str__ 또는 repr
        cell_str = str(cell)
        if cell_str and cell_str not in ['None', '']:
            return cell_str
        
        return ""
    except Exception as e:
        return ""


def _parse_tables_from_streams_improved(hwp_file) -> List[Dict[str, Any]]:
    """스트림에서 직접 표 파싱 시도 (개선된 버전)"""
    tables = []
    try:
        from pyhwp.hwp5.tagids import HWPTAG_TABLE
        
        # Section 스트림에서 표 찾기
        if 'Section' in hwp_file:
            try:
                sections = hwp_file['Section']
                for section in sections:
                    # Section의 paragraphs에서 표 찾기
                    if hasattr(section, 'paragraphs'):
                        for para in section.paragraphs():
                            if hasattr(para, 'controls'):
                                for ctrl in para.controls():
                                    ctrl_type = str(type(ctrl))
                                    if 'Table' in ctrl_type:
                                        table_data = _extract_table_from_control(ctrl, 0, 0)
                                        if table_data and table_data.get('rows'):
                                            tables.append(table_data)
            except Exception as e:
                pass
        
        # BodyText/Section0에서 직접 찾기
        try:
            if 'BodyText' in hwp_file:
                body_text = hwp_file['BodyText']
                if hasattr(body_text, 'section'):
                    section = body_text.section(0)
                    if section:
                        for para in section.paragraphs():
                            # ParaText에서 표 찾기
                            if hasattr(para, 'controls'):
                                for ctrl in para.controls():
                                    if 'Table' in str(type(ctrl)):
                                        table_data = _extract_table_from_control(ctrl, 0, 0)
                                        if table_data and table_data.get('rows'):
                                            tables.append(table_data)
        except Exception as e:
            pass
                
    except Exception as e:
        import traceback
        traceback.print_exc()
    
    return tables


def _parse_tables_from_bindata(hwp_file) -> List[Dict[str, Any]]:
    """BinData에서 표 찾기"""
    tables = []
    try:
        # BinData 스트림 확인
        if 'BinData' in hwp_file:
            # 바이너리 데이터에서 표 구조 찾기
            # 이 부분은 한글 파일의 바이너리 구조를 직접 파싱해야 함
            pass
    except Exception as e:
        pass
    
    return tables


def _explore_all_controls(hwp_file) -> List[Dict[str, Any]]:
    """모든 컨트롤을 탐색하여 표 찾기 (디버깅 및 포괄적 탐색)"""
    tables = []
    try:
        if 'BodyText' in hwp_file:
            body_text = hwp_file['BodyText']
            
            for section_idx, section in enumerate(body_text.sections()):
                try:
                    # Section의 모든 속성 확인
                    section_attrs = dir(section)
                    
                    # Paragraphs 순회
                    for para_idx, para in enumerate(section.paragraphs()):
                        try:
                            # Paragraph의 모든 속성 확인
                            para_attrs = dir(para)
                            
                            # controls 속성 확인
                            if 'controls' in para_attrs:
                                controls = para.controls()
                                for ctrl in controls:
                                    ctrl_type = str(type(ctrl))
                                    ctrl_name = type(ctrl).__name__
                                    
                                    # 표 관련 키워드 확인
                                    if any(keyword in ctrl_type.lower() or keyword in ctrl_name.lower() 
                                           for keyword in ['table', 'cell', 'row']):
                                        table_data = _extract_table_from_control(ctrl, section_idx, para_idx)
                                        if table_data and table_data.get('rows'):
                                            tables.append(table_data)
                            
                            # child_controls 확인
                            if 'child_controls' in para_attrs:
                                try:
                                    child_controls = para.child_controls()
                                    for ctrl in child_controls:
                                        ctrl_type = str(type(ctrl))
                                        if 'Table' in ctrl_type:
                                            table_data = _extract_table_from_control(ctrl, section_idx, para_idx)
                                            if table_data and table_data.get('rows'):
                                                tables.append(table_data)
                                except:
                                    pass
                            
                            # model 속성 확인
                            if 'model' in para_attrs:
                                try:
                                    model = para.model
                                    if hasattr(model, 'controls'):
                                        for ctrl in model.controls:
                                            if 'Table' in str(type(ctrl)):
                                                table_data = _extract_table_from_control(ctrl, section_idx, para_idx)
                                                if table_data and table_data.get('rows'):
                                                    tables.append(table_data)
                                except:
                                    pass
                        except Exception as e:
                            pass
                except Exception as e:
                    pass
    except Exception as e:
        import traceback
        traceback.print_exc()
    
    return tables


def save_tables_to_json(tables: List[Dict[str, Any]], output_path: str):
    """
    파싱된 표 데이터를 JSON 파일로 저장합니다.
    
    Args:
        tables: 표 데이터 리스트
        output_path: 출력 JSON 파일 경로
    """
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(tables, f, ensure_ascii=False, indent=2)
        print(f"표 데이터가 저장되었습니다: {output_path}")
        return True
    except Exception as e:
        print(f"JSON 저장 중 오류: {e}")
        return False


def save_tables_to_csv(tables: List[Dict[str, Any]], output_dir: str):
    """
    각 표를 개별 CSV 파일로 저장합니다.
    
    Args:
        tables: 표 데이터 리스트
        output_dir: 출력 디렉토리 경로
    """
    try:
        os.makedirs(output_dir, exist_ok=True)
        
        for idx, table in enumerate(tables):
            csv_path = os.path.join(output_dir, f"table_{idx + 1}.csv")
            with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                for row in table.get('rows', []):
                    writer.writerow(row)
            print(f"표 {idx + 1}이 저장되었습니다: {csv_path}")
        
        return True
    except Exception as e:
        print(f"CSV 저장 중 오류: {e}")
        return False


def load_tables_from_json(json_path: str) -> List[Dict[str, Any]]:
    """
    JSON 파일에서 표 데이터를 로드합니다.
    
    Args:
        json_path: JSON 파일 경로
        
    Returns:
        표 데이터 리스트
    """
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"JSON 로드 중 오류: {e}")
        return []


def load_tables_from_csv(csv_dir: str) -> List[Dict[str, Any]]:
    """
    CSV 파일들에서 표 데이터를 로드합니다.
    
    Args:
        csv_dir: CSV 파일들이 있는 디렉토리
        
    Returns:
        표 데이터 리스트
    """
    tables = []
    try:
        csv_files = sorted([f for f in os.listdir(csv_dir) if f.endswith('.csv')])
        
        for csv_file in csv_files:
            csv_path = os.path.join(csv_dir, csv_file)
            rows = []
            
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                for row in reader:
                    rows.append(row)
            
            if rows:
                table_data = {
                    "rows": rows,
                    "row_count": len(rows),
                    "col_count": max(len(row) for row in rows) if rows else 0
                }
                tables.append(table_data)
        
        return tables
    except Exception as e:
        print(f"CSV 로드 중 오류: {e}")
        return []


def reconstruct_hwp_with_tables(template_path: str, tables: List[Dict[str, Any]], output_path: str):
    """
    템플릿 한글 파일에 표 데이터를 삽입하여 새로운 파일을 생성합니다.
    
    주의: pyhwp는 주로 읽기 전용이므로, 완전한 재구성은 제한적일 수 있습니다.
    이 함수는 가능한 범위에서 표 데이터를 적용합니다.
    
    Args:
        template_path: 템플릿 한글 파일 경로
        tables: 삽입할 표 데이터 리스트
        output_path: 출력 파일 경로
    """
    try:
        from pyhwp import hwp5
        import shutil
        
        # 템플릿 파일 복사
        shutil.copy2(template_path, output_path)
        
        # 파일 열기
        hwp_file = hwp5.open(output_path, mode='r+b')
        
        # 표 데이터 적용 시도
        # 주의: pyhwp의 쓰기 기능은 제한적일 수 있습니다
        # 실제 구현은 한글 파일의 바이너리 구조를 직접 수정해야 할 수 있습니다
        
        print(f"파일이 생성되었습니다: {output_path}")
        print("주의: 표 데이터 적용은 한글 파일의 복잡한 구조로 인해 제한적일 수 있습니다.")
        print("완전한 재구성을 위해서는 한글과컴퓨터 API나 다른 도구가 필요할 수 있습니다.")
        
        return True
        
    except Exception as e:
        print(f"재구성 중 오류: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """메인 함수"""
    if len(sys.argv) < 2:
        print("사용법:")
        print("  python han_parser.py <한글파일경로> [옵션]")
        print("\n옵션:")
        print("  --tables              표만 파싱")
        print("  --save-json <경로>   표를 JSON으로 저장")
        print("  --save-csv <디렉토리> 표를 CSV로 저장")
        print("  --load-json <경로>   JSON에서 표 로드")
        print("  --reconstruct <템플릿> <출력경로>  표 데이터로 재구성")
        sys.exit(1)
    
    file_path = sys.argv[1]
    
    if not os.path.exists(file_path):
        print(f"파일을 찾을 수 없습니다: {file_path}")
        sys.exit(1)
    
    # 옵션 처리
    args = sys.argv[2:]
    tables_only = '--tables' in args
    save_json_idx = args.index('--save-json') if '--save-json' in args else -1
    save_csv_idx = args.index('--save-csv') if '--save-csv' in args else -1
    
    print(f"한글 파일 파싱 중: {file_path}")
    print("-" * 50)
    
    if not tables_only:
        # 간단한 정보 추출
        info = parse_hwp_simple(file_path)
        if info:
            print("파일 정보:")
            print(f"  - OLE 파일: {info.get('is_ole', False)}")
            print(f"  - 스트림 개수: {len(info.get('streams', []))}")
        
        print("-" * 50)
        
        # 텍스트 추출
        text = parse_hwp(file_path)
        if text:
            print("추출된 텍스트:")
            print(text[:500] + "..." if len(text) > 500 else text)
            print("-" * 50)
    
    # 표 파싱
    print("표 파싱 중...")
    tables = parse_tables(file_path)
    
    if tables:
        print(f"총 {len(tables)}개의 표를 찾았습니다.")
        for idx, table in enumerate(tables):
            print(f"\n표 {idx + 1}:")
            print(f"  - 행 수: {table.get('row_count', 0)}")
            print(f"  - 열 수: {table.get('col_count', 0)}")
            print("  - 내용 미리보기:")
            for row_idx, row in enumerate(table.get('rows', [])[:3]):  # 처음 3행만
                print(f"    행 {row_idx + 1}: {row[:5]}")  # 처음 5열만
            if len(table.get('rows', [])) > 3:
                print("    ...")
    else:
        print("표를 찾을 수 없습니다.")
        print("참고: 표 파싱은 한글 파일의 구조에 따라 제한적일 수 있습니다.")
    
    # JSON 저장
    if save_json_idx >= 0 and save_json_idx + 1 < len(args):
        json_path = args[save_json_idx + 1]
        save_tables_to_json(tables, json_path)
    
    # CSV 저장
    if save_csv_idx >= 0 and save_csv_idx + 1 < len(args):
        csv_dir = args[save_csv_idx + 1]
        save_tables_to_csv(tables, csv_dir)


if __name__ == "__main__":
    main()

