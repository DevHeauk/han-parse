#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
표 데이터 재구성 유틸리티
JSON/CSV 데이터를 읽어서 한글 파일 형식으로 재구성하는 도구
"""

import json
import csv
import os
from typing import List, Dict, Any, Optional
from han_parser import load_tables_from_json, load_tables_from_csv


def create_hwp_from_tables_json(tables_json_path: str, output_path: str, template_path: Optional[str] = None):
    """
    JSON 파일의 표 데이터로 한글 파일을 생성합니다.
    
    Args:
        tables_json_path: 표 데이터가 있는 JSON 파일 경로
        output_path: 생성할 한글 파일 경로
        template_path: (선택) 템플릿 한글 파일 경로
    """
    tables = load_tables_from_json(tables_json_path)
    
    if not tables:
        print("표 데이터를 찾을 수 없습니다.")
        return False
    
    print(f"총 {len(tables)}개의 표를 로드했습니다.")
    
    # 템플릿이 있으면 사용, 없으면 새로 생성
    if template_path and os.path.exists(template_path):
        return _reconstruct_with_template(template_path, tables, output_path)
    else:
        return _create_new_hwp(tables, output_path)


def create_hwp_from_tables_csv(csv_dir: str, output_path: str, template_path: Optional[str] = None):
    """
    CSV 파일들의 표 데이터로 한글 파일을 생성합니다.
    
    Args:
        csv_dir: CSV 파일들이 있는 디렉토리
        output_path: 생성할 한글 파일 경로
        template_path: (선택) 템플릿 한글 파일 경로
    """
    tables = load_tables_from_csv(csv_dir)
    
    if not tables:
        print("표 데이터를 찾을 수 없습니다.")
        return False
    
    print(f"총 {len(tables)}개의 표를 로드했습니다.")
    
    if template_path and os.path.exists(template_path):
        return _reconstruct_with_template(template_path, tables, output_path)
    else:
        return _create_new_hwp(tables, output_path)


def _reconstruct_with_template(template_path: str, tables: List[Dict[str, Any]], output_path: str) -> bool:
    """템플릿 파일에 표 데이터를 삽입"""
    try:
        from han_parser import reconstruct_hwp_with_tables
        return reconstruct_hwp_with_tables(template_path, tables, output_path)
    except Exception as e:
        print(f"템플릿 재구성 중 오류: {e}")
        return False


def _create_new_hwp(tables: List[Dict[str, Any]], output_path: str) -> bool:
    """새 한글 파일 생성 (제한적)"""
    print("\n주의: 완전한 한글 파일 생성은 복잡한 바이너리 구조가 필요합니다.")
    print("현재는 표 데이터를 JSON/CSV로 저장하고, 수동으로 한글에 삽입하는 것을 권장합니다.")
    print("\n표 데이터 요약:")
    for idx, table in enumerate(tables):
        print(f"표 {idx + 1}: {table.get('row_count', 0)}행 x {table.get('col_count', 0)}열")
    
    # 표 데이터를 텍스트 형식으로 저장 (참고용)
    txt_path = output_path.replace('.hwp', '_tables.txt')
    try:
        with open(txt_path, 'w', encoding='utf-8') as f:
            for idx, table in enumerate(tables):
                f.write(f"=== 표 {idx + 1} ===\n")
                f.write(f"행: {table.get('row_count', 0)}, 열: {table.get('col_count', 0)}\n\n")
                for row in table.get('rows', []):
                    f.write('\t'.join(str(cell) for cell in row) + '\n')
                f.write('\n')
        print(f"\n표 데이터가 텍스트 파일로 저장되었습니다: {txt_path}")
        return True
    except Exception as e:
        print(f"텍스트 파일 저장 중 오류: {e}")
        return False


def edit_table_data(json_path: str, table_index: int, row: int, col: int, new_value: str, output_path: Optional[str] = None):
    """
    JSON 파일의 표 데이터를 수정합니다.
    
    Args:
        json_path: JSON 파일 경로
        table_index: 수정할 표 인덱스 (0부터 시작)
        row: 행 인덱스 (0부터 시작)
        col: 열 인덱스 (0부터 시작)
        new_value: 새로운 값
        output_path: 출력 파일 경로 (없으면 원본 파일 덮어쓰기)
    """
    tables = load_tables_from_json(json_path)
    
    if table_index >= len(tables):
        print(f"표 인덱스 {table_index}가 범위를 벗어났습니다. (총 {len(tables)}개)")
        return False
    
    table = tables[table_index]
    rows = table.get('rows', [])
    
    if row >= len(rows):
        print(f"행 인덱스 {row}가 범위를 벗어났습니다. (총 {len(rows)}행)")
        return False
    
    # 열 확장이 필요한 경우
    while len(rows[row]) <= col:
        rows[row].append('')
    
    rows[row][col] = new_value
    
    # 저장
    save_path = output_path or json_path
    try:
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(tables, f, ensure_ascii=False, indent=2)
        print(f"표 데이터가 수정되었습니다: {save_path}")
        print(f"표 {table_index + 1}, 행 {row + 1}, 열 {col + 1} = '{new_value}'")
        return True
    except Exception as e:
        print(f"저장 중 오류: {e}")
        return False


def merge_tables(tables1_path: str, tables2_path: str, output_path: str):
    """
    두 개의 표 데이터 JSON 파일을 병합합니다.
    
    Args:
        tables1_path: 첫 번째 JSON 파일 경로
        tables2_path: 두 번째 JSON 파일 경로
        output_path: 병합된 데이터를 저장할 경로
    """
    tables1 = load_tables_from_json(tables1_path)
    tables2 = load_tables_from_json(tables2_path)
    
    merged = tables1 + tables2
    
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(merged, f, ensure_ascii=False, indent=2)
        print(f"표 데이터가 병합되었습니다: {output_path}")
        print(f"총 {len(merged)}개의 표 (첫 번째: {len(tables1)}개, 두 번째: {len(tables2)}개)")
        return True
    except Exception as e:
        print(f"병합 중 오류: {e}")
        return False


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("사용법:")
        print("  python table_reconstructor.py create-json <JSON경로> <출력경로> [템플릿경로]")
        print("  python table_reconstructor.py create-csv <CSV디렉토리> <출력경로> [템플릿경로]")
        print("  python table_reconstructor.py edit <JSON경로> <표인덱스> <행> <열> <새값> [출력경로]")
        print("  python table_reconstructor.py merge <JSON1> <JSON2> <출력경로>")
        sys.exit(1)
    
    command = sys.argv[1]
    
    if command == "create-json":
        if len(sys.argv) < 4:
            print("사용법: python table_reconstructor.py create-json <JSON경로> <출력경로> [템플릿경로]")
            sys.exit(1)
        template = sys.argv[4] if len(sys.argv) > 4 else None
        create_hwp_from_tables_json(sys.argv[2], sys.argv[3], template)
    
    elif command == "create-csv":
        if len(sys.argv) < 4:
            print("사용법: python table_reconstructor.py create-csv <CSV디렉토리> <출력경로> [템플릿경로]")
            sys.exit(1)
        template = sys.argv[4] if len(sys.argv) > 4 else None
        create_hwp_from_tables_csv(sys.argv[2], sys.argv[3], template)
    
    elif command == "edit":
        if len(sys.argv) < 7:
            print("사용법: python table_reconstructor.py edit <JSON경로> <표인덱스> <행> <열> <새값> [출력경로]")
            sys.exit(1)
        output = sys.argv[7] if len(sys.argv) > 7 else None
        edit_table_data(sys.argv[2], int(sys.argv[3]), int(sys.argv[4]), int(sys.argv[5]), sys.argv[6], output)
    
    elif command == "merge":
        if len(sys.argv) < 5:
            print("사용법: python table_reconstructor.py merge <JSON1> <JSON2> <출력경로>")
            sys.exit(1)
        merge_tables(sys.argv[2], sys.argv[3], sys.argv[4])
    
    else:
        print(f"알 수 없는 명령어: {command}")

